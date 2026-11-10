# @Author: YiLong Li
# @Date: 2026/11/10 12:11
# @File: merge.py
# @Software: PyCharm


"""
输入：当前脚本目录 / output / A_temp_result.xlsx、Pre_V.xlsx
输出：当前脚本目录 / output / A_result.xlsx、A_logic.xlsx、A_error.xlsx

主要功能：
    按日期匹配每日运输方量与模块数量。
    将每日方量填入模块矩阵末列，生成最终求解矩阵。

处理规则：
    仅读取“矩阵_去除全零行”工作表。
    根据日期进行一对一匹配，不自动累加重复日期方量。
    检查矩阵数据及方量数据的有效性。
    对未匹配日期、重复方量日期及非法数据进行异常记录。
    对模块数量全零、单模块数量小于10的记录进行逻辑剔除。
    根据各模块数量及方量的整体比例识别重复记录。
    比例重复记录保留首次出现的记录。

输出规则：
    A_result.xlsx：最终汇总矩阵，不含表头。
        前面各列为模块数量，最后一列为每日方量。
        保留原始模块顺序。
    A_logic.xlsx：包含逻辑删除记录及比例重复记录对照。
    A_error.xlsx：包含方量异常、日期重复、矩阵异常及日期未匹配记录。
    无逻辑删除或异常记录时，不生成对应空报告。

使用说明：
    默认读取脚本目录下的output文件夹。
    支持通过 --output-dir 自定义输入及输出目录。
    方量和模块数量必须为非负有限数值。
    不修改源文件，输出文件发布失败时执行回滚。
    Python 3.10+，依赖openpyxl。
"""


from __future__ import annotations
import argparse
import math
import re
import shutil
import sys
import tempfile
from pathlib import Path
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from functools import reduce
from collections import Counter
from collections.abc import Mapping
from contextlib import closing
from typing import Any
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
INPUT_SHEET = "矩阵_去除全零行"
SINGLE_MODULE_THRESHOLD = 10
Sheets = list


def numeric(value: Any) -> tuple[int | float, Decimal]:
    # 输入为程序已生成的纯数字矩阵。空格、文本、公式、布尔值均不默认为0。
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("矩阵单元格必须为数值，不能是空值、文本、公式或布尔值")
    try:
        decimal = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise ValueError("无法识别的数值") from error
    if not decimal.is_finite() or decimal < 0:
        raise ValueError("矩阵单元格必须为非负有限数值")
    return value, decimal


def proportion_key(values: list[Decimal]) -> tuple[int, ...]:
    """消去十进制分母，再除以整行最大公约数，得到精确比例指纹。

    不使用浮点除法，不将任一非零列省略；全零行由调用方排除。
    """
    ratios = [value.as_integer_ratio() for value in values]
    denominator = math.lcm(*(den for _, den in ratios))
    integers = [num * (denominator // den) for num, den in ratios]
    divisor = reduce(math.gcd, integers, 0)
    if divisor == 0:
        raise ValueError("全零行不应参与比例计算")
    return tuple(value // divisor for value in integers)


def parse_day(value, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.date()
        raise ValueError("日期序列值无效")
    text = str(value or "").strip()
    match = re.fullmatch(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?", text)
    if not match:
        raise ValueError(f"无法识别日期：{text!r}")
    return date(*(int(v) for v in match.groups()))


def day_text(day):
    return f"{day.year}/{day.month}/{day.day:02d}"


def volume_number(value):
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("，", "")
        try:
            number = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError("方量不是数值") from exc
        if not number.is_finite() or number < 0:
            raise ValueError("方量必须为非负有限数值")
        value = int(number) if number == number.to_integral_value() else float(number)
    return numeric(value)[0]


def write_workbook(path, sheets):
    with closing(Workbook()) as book:
        book.remove(book.active)
        for title, rows in sheets:
            ws = book.create_sheet(title)
            report = title != "汇总矩阵"
            ws.sheet_view.showGridLines = False
            ws.sheet_format.defaultColWidth = 13
            ws.freeze_panes = "F2" if report else "A1"
            for rn, row in enumerate(rows, 1):
                if rn > 1048576 or len(row) > 16384:
                    raise ValueError("输出超过Excel行列上限")
                for cn, value in enumerate(row, 1):
                    cell = ws.cell(rn, cn, value)
                    if isinstance(value, str):
                        cell.data_type = "s"
                    cell.font = Font(name="微软雅黑", size=10, bold=report and rn == 1,
                                     color="FFFFFF" if report and rn == 1 else "243746")
                    cell.alignment = Alignment(vertical="center", horizontal="left" if isinstance(value, str) else "right")
                    if report and rn == 1:
                        cell.fill = PatternFill("solid", fgColor="305496")
                    elif rn % 2 == 0:
                        cell.fill = PatternFill("solid", fgColor="F1F5FA")
                ws.row_dimensions[rn].height = 30 if report and rn == 1 else 22
            if report and ws.max_row:
                ws.auto_filter.ref = ws.dimensions
                for col, width in [(1,38),(2,24),(3,24),(4,14),(5,18)]:
                    ws.column_dimensions[get_column_letter(col)].width = width
            if not report and ws.max_column:
                ws.column_dimensions[get_column_letter(ws.max_column)].width = 19
        book.save(path)


def publish(files: Mapping[Path, Sheets | None]) -> None:
    """准备成功再替换；失败回滚。None表示删除对应旧文件。

    不保证断电或强制退出时的跨文件原子性。
    """
    stage = Path(tempfile.mkdtemp(prefix=".merge_", dir=OUTPUT_DIR))
    staged: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    installed: list[Path] = []
    preserve = False
    try:
        for index, (target, sheets) in enumerate(files.items()):
            if sheets is not None:
                temp = stage / f"new_{index}.xlsx"
                write_workbook(temp, sheets)
                staged[target] = temp
        try:
            for index, target in enumerate(files):
                if target.exists():
                    backup = stage / f"backup_{index}.xlsx"
                    target.replace(backup)
                    backups[target] = backup
            for target, temp in staged.items():
                temp.replace(target)
                installed.append(target)
        except Exception as publish_error:
            rollback_errors: list[str] = []
            for target in reversed(installed):
                try:
                    target.unlink(missing_ok=True)
                except OSError as error:
                    rollback_errors.append(f"移除{target.name}失败：{error}")
            for target, backup in backups.items():
                try:
                    backup.replace(target)
                except OSError as error:
                    rollback_errors.append(f"恢复{target.name}失败：{error}")
            if rollback_errors:
                preserve = True
                raise RuntimeError(f"发布失败：{publish_error}；{'；'.join(rollback_errors)}；恢复文件位于{stage}") from publish_error
            raise RuntimeError(f"发布失败，已恢复原文件：{publish_error}；请检查文件是否被Excel占用") from publish_error
    finally:
        if not preserve:
            try:
                shutil.rmtree(stage)
            except OSError:
                print(f"[提示] 临时目录未清理：{stage}")


def merge(output_dir=OUTPUT_DIR):
    global OUTPUT_DIR
    OUTPUT_DIR = Path(output_dir)
    matrix_path = OUTPUT_DIR / "A_temp_result.xlsx"
    volume_path = OUTPUT_DIR / "Pre_V.xlsx"
    for path in (matrix_path, volume_path):
        if not path.is_file():
            raise FileNotFoundError(f"输入文件不存在：{path}")
    errors = {}
    def report(name, header, row):
        if name not in errors:
            errors[name] = [header]
        errors[name].append(row)

    meta = ["原因", "来源文件", "来源表单", "原始行号", "日期"]
    volume_rows = []
    groups = {}
    with closing(load_workbook(volume_path, read_only=True, data_only=True)) as book:
        for ws in book:
            iterator = ws.iter_rows(values_only=True)
            header = next(iterator, ())
            names = [str(v or "").strip() for v in header]
            if "日期" not in names or "方量" not in names:
                raise ValueError(f"Pre_V.xlsx/{ws.title}缺少日期或方量表头")
            di, vi = names.index("日期"), names.index("方量")
            for line, raw in enumerate(iterator, 2):
                if all(v is None for v in raw):
                    continue
                item = dict(raw=list(raw), sheet=ws.title, line=line, day=None, volume=None, reason="")
                try:
                    item["day"] = parse_day(raw[di], book.epoch)
                    item["volume"] = volume_number(raw[vi])
                except (ValueError, TypeError, OverflowError) as exc:
                    item["reason"] = str(exc)
                volume_rows.append(item)
                if item["day"] is not None:
                    groups.setdefault(item["day"], []).append(item)
                if item["reason"]:
                    report("方量数据异常", meta + names,
                           [item["reason"], volume_path.name, ws.title, line, str(raw[di])] + list(raw))
        # 同一天多行不擅自累加或取首行，整组转人工复核。
        for day, items in groups.items():
            if len(items) > 1:
                for item in items:
                    item["reason"] = "方量日期重复，未自动合并"
                    report("方量日期重复", meta + ["原始数据"],
                           [item["reason"], volume_path.name, item["sheet"], item["line"], day_text(day), repr(item["raw"])])
    lookup = {d: items[0]["volume"] for d, items in groups.items()
              if len(items) == 1 and not items[0]["reason"]}
    matrix_dates = set()
    result, logic, duplicates = [], [], []
    seen = {}
    counts = Counter()
    with closing(load_workbook(matrix_path, read_only=True, data_only=True)) as book:
        if INPUT_SHEET not in book.sheetnames:
            raise ValueError(f"A_temp_result.xlsx缺少表单：{INPUT_SHEET}")
        ws = book[INPUT_SHEET]
        width = None
        for line, raw in enumerate(ws.iter_rows(values_only=True), 1):
            if all(v is None for v in raw):
                continue
            counts["原始记录"] += 1
            if width is None:
                width = len(raw)
                if width < 2:
                    raise ValueError("矩阵至少应有一列模块及末列日期")
            original = list(raw)
            day = None
            reason = ""
            try:
                day = parse_day(raw[-1], book.epoch)
                matrix_dates.add(day)
                if len(raw) != width:
                    raise ValueError("矩阵列数不一致")
                numbers = [numeric(v)[1] for v in raw[:-1]]
            except (ValueError, TypeError, OverflowError) as exc:
                reason = str(exc)
            headers = meta + [f"x{i}" for i in range(1, len(raw))] + ["原始日期"]
            prefix = [reason, matrix_path.name, INPUT_SHEET, line, day_text(day) if day else str(raw[-1])]
            if reason:
                report("矩阵数据异常", headers, prefix + original)
                counts["矩阵数据异常"] += 1
                continue
            if day not in lookup:
                prefix[0] = "日期未匹配方量" if day not in groups else "对应方量无效或日期重复"
                report("矩阵未匹配方量", headers, prefix + original)
                counts["矩阵未匹配方量"] += 1
                continue
            counts["匹配成功"] += 1
            values = list(raw[:-1]) + [lookup[day]]
            signature_values = numbers + [Decimal(str(lookup[day]))]
            nonzero = [v for v in numbers if v != 0]
            reason = ""
            if not nonzero:
                reason = "模块数量全部为零"
            elif len(nonzero) == 1 and 0 < nonzero[0] < SINGLE_MODULE_THRESHOLD:
                reason = f"单模块数量小于{SINGLE_MODULE_THRESHOLD}"
            if reason:
                logic.append([reason, matrix_path.name, INPUT_SHEET, line, day_text(day)] + values)
                continue
            signature = proportion_key(signature_values)
            if signature in seen:
                kept_line, kept_day, kept_values = seen[signature]
                group = len(duplicates) // 2 + 1
                duplicates.append([f"第{group}组：保留行", matrix_path.name, INPUT_SHEET, kept_line, day_text(kept_day)] + kept_values)
                duplicates.append([f"第{group}组：比例重复删除", matrix_path.name, INPUT_SHEET, line, day_text(day)] + values)
            else:
                seen[signature] = (line, day, values)
                result.append(values)
    for item in volume_rows:
        day = item["day"]
        if day is not None and day not in matrix_dates:
            report("方量未匹配矩阵", meta + ["方量", "原始数据"],
                   ["日期未匹配矩阵", volume_path.name, item["sheet"], item["line"], day_text(day), item["volume"], repr(item["raw"])])
    logic_sheets = []
    header = meta + [f"x{i}" for i in range(1, width or 1)] + ["方量"]
    if logic:
        logic_sheets.append(("逻辑删除", [header] + logic))
    if duplicates:
        logic_sheets.append(("重复删除对照", [header] + duplicates))
    publish({
        OUTPUT_DIR / "A_result.xlsx": [("汇总矩阵", result)],
        OUTPUT_DIR / "A_logic.xlsx": logic_sheets or None,
        OUTPUT_DIR / "A_error.xlsx": list(errors.items()) or None,
    })
    print("日期与方量合并 " + "=" * 48)
    for label, value in [("矩阵原始行", counts["原始记录"]), ("匹配成功行", counts["匹配成功"]),
                         ("矩阵未匹配行", counts["矩阵未匹配方量"]), ("矩阵无效行", counts["矩阵数据异常"]),
                         ("逻辑删除行", len(logic)), ("比例重复删除行", len(duplicates)//2), ("最终保留行", len(result))]:
        print(f"  {label}：{value:,}")
    print("输出文件 " + "-" * 52)
    print(f"  A_result.xlsx：{len(result):,} 行，末列为方量")
    for name, present in [("A_logic.xlsx", logic_sheets), ("A_error.xlsx", errors)]:
        print(f"  {name}：" + ("已生成" if present else "无记录，已清理同名旧文件（如有）"))
    for name, rows in errors.items():
        print(f"    {name}：{len(rows)-1:,} 条")
    return dict(counts, result=len(result), logic=len(logic), duplicates=len(duplicates)//2)


def main():
    parser = argparse.ArgumentParser(description="按日期匹配方量并生成求解矩阵")
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR, help="输入及输出文件夹，默认脚本目录/output")
    args = parser.parse_args()
    try:
        merge(args.output_dir)
        return 0
    except Exception as exc:
        print(f"处理失败：{exc}；本次未生成完整新结果，请勿继续调用旧结果。")
        return 1


if __name__ == "__main__":
    sys.exit(main())