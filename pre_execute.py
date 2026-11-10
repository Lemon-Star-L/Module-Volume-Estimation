# @Author: YiLong Li
# @Date: 2026/11/4 09:44
# @File: pre_execute.py
# @Software: PyCharm


"""
输入：当前脚本目录 / data / execute / YYYY.MM.xlsx
映射：当前脚本目录 / data / Resource.xlsx
输出：当前脚本目录 / output / A_temp_result.xlsx、A_temp_erro.xlsx

主要功能：
    按年月读取业务数据，提取汽运收货记录。
    根据货品编码匹配SKU及备件模块，汇总各模块每日数量。
    生成模块数量矩阵及数据异常报告。

处理规则：
    按Resource中的备件信息及模块信息完成SKU和模块映射。
    仅统计汽运记录，按箱号日期进行每日汇总。
    对必填字段缺失、非汽运、SKU未匹配、模块异常及重复记录进行识别。
    跨文件、跨工作表进行重复记录检查。
    非法日期、非法数量、月份不符及表头异常列为严重错误。

输出规则：
    A_temp_result.xlsx：包含两个工作表。
        矩阵_去除全零行：删除模块数量全部为零的日期。
        矩阵_保留全零行：保留全部有效日期。
    矩阵不含表头，每行代表一个日期。
    前面各列为模块数量，最后一列为日期。
    A_temp_erro.xlsx：记录异常明细、处理统计及数据核对结果。

使用说明：
    月度文件命名格式为 YYYY.MM.xlsx。
    Resource.xlsx包含“备件信息”和“模块信息”两个工作表。
    模块顺序按照Resource中“模块信息”的原始顺序排列。
    支持普通及Strict格式XLSX，数量支持文本数字及中英文逗号。
    公式按Excel缓存值读取，运行前需计算并保存。
    不修改输入文件。
"""


from pathlib import Path
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, localcontext
from collections import defaultdict, Counter
from contextlib import closing
import json
import math
import re
import sqlite3
import tempfile
import sys
import warnings
import unicodedata
from time import perf_counter
from zipfile import ZipFile, ZIP_DEFLATED

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.datetime import from_excel
except ImportError:
    raise SystemExit("缺少依赖，请运行：pip install openpyxl（或 uv add openpyxl）")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data" / "execute"
RESOURCE_FILE = BASE_DIR / "data" / "Resource.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
REQUIRED = ("箱号日期", "发运方式", "货品编码", "数量")
PARTS_SHEET = "备件信息"
MODULES_SHEET = "模块信息"
MAX_DATA_ROWS = 1_048_575
PROGRESS_SECONDS = 2.0  # 按时间刷新实际读取条数，不按固定条数触发

class ManagedWorkbook:
    """关闭读取器时一并清理格式兼容副本。"""
    def __init__(self, book, temporary=None):
        self.book = book
        self.temporary = temporary

    def __getattr__(self, name):
        return getattr(self.book, name)

    def __getitem__(self, name):
        return self.book[name]

    def close(self):
        try:
            self.book.close()
        finally:
            if self.temporary is not None:
                self.temporary.cleanup()
                self.temporary = None


def open_checked(path):
    """支持普通/Strict XLSX；仅在临时副本中转换命名空间，不修改源文件。"""
    temporary = None
    actual_path = path
    try:
        with ZipFile(path) as source:
            workbook_xml = source.read("xl/workbook.xml")
            if b"http://purl.oclc.org/ooxml/spreadsheetml/main" in workbook_xml:
                temporary = tempfile.TemporaryDirectory(prefix="pre_execute_strict_")
                actual_path = Path(temporary.name) / "compatible.xlsx"
                namespaces = {
                    b"http://purl.oclc.org/ooxml/spreadsheetml/main":
                        b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    b"http://purl.oclc.org/ooxml/officeDocument/relationships":
                        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    b"http://purl.oclc.org/ooxml/drawingml/main":
                        b"http://schemas.openxmlformats.org/drawingml/2006/main",
                }
                # 只改命名空间声明和关系Type属性，绝不替换业务单元格文本。
                patterns = [
                    (re.compile(rb'((?:xmlns(?::[\w.-]+)?|Type)\s*=\s*["\x27])'
                                + re.escape(old) + rb'(?=[/"\x27])'), new)
                    for old, new in namespaces.items()
                ]
                with ZipFile(actual_path, "w", ZIP_DEFLATED) as target:
                    for entry in source.infolist():
                        content = source.read(entry.filename)
                        if entry.filename.endswith((".xml", ".rels")):
                            for pattern, replacement in patterns:
                                content = pattern.sub(
                                    lambda match, value=replacement: match[1] + value, content)
                        target.writestr(entry, content)
                print(f"  已启用Strict格式兼容读取：{Path(path).name}（源文件不变）", flush=True)
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            book = load_workbook(actual_path, read_only=True, data_only=True)
        messages = [str(item.message) for item in caught]
        invalid = [message for message in messages if "invalid specification" in message.lower()]
        if invalid or not book.worksheets:
            book.close()
            detail = "；".join(invalid) or "没有可读取的工作表"
            raise ValueError(
                f"工作簿结构异常，可能导致表单被跳过：{detail}。"
                "请用Excel打开该文件并另存为新的.xlsx后重试；不要直接改扩展名。"
            )
        for message in messages:
            print(f"  文件提示：{message}", flush=True)
        return ManagedWorkbook(book, temporary)
    except Exception:
        if temporary is not None:
            temporary.cleanup()
        raise


EXCLUSION_CATEGORIES = (
    "非汽运", "发运方式为空", "其他必填字段为空", "日期或数量非法",
    "SKU未匹配", "模块映射异常", "重复记录", "其他异常",
)


def exclusion_category(reasons):
    """按优先级唯一归类；完整原因仍保留在异常明细中。"""
    if any(reason.startswith("非汽运：") for reason in reasons):
        return "非汽运"
    if "发运方式为空" in reasons:
        return "发运方式为空"
    if any(f"{field}为空" in reasons for field in REQUIRED):
        return "其他必填字段为空"
    if any(reason.startswith(("数量", "无法识别箱号日期", "无效Excel日期序号", "箱号日期"))
           for reason in reasons):
        return "日期或数量非法"
    if "货品编码无法匹配SKU" in reasons:
        return "SKU未匹配"
    if any(reason.startswith(("SKU对应", "备件模块")) for reason in reasons):
        return "模块映射异常"
    if "除箱号日期外所有字段重复" in reasons:
        return "重复记录"
    return "其他异常"


def new_sheet_stats(file, sheet):
    return {"文件": file, "表单": sheet, "状态": "读取中",
            "原始记录数": 0, "参与汇总记录数": 0, "异常记录数": 0,
            "重复记录数": 0, "文本数量转换数": 0, "原始数量": Decimal(0), "汽运原始数量": Decimal(0),
            "纳入矩阵数量": Decimal(0), "剔除数量": Decimal(0), "无效数量记录数": 0,
            **{f"剔除_{name}": 0 for name in EXCLUSION_CATEGORIES}}


def summarize(items):
    total = new_sheet_stats("全部", "合计")
    for item in items:
        for field in list(total)[3:]:
            total[field] += item[field]
    return total


def quantity_text(value):
    text = format(value, ",f")
    return text.rstrip("0").rstrip(".") if "." in text else text


CONSOLE_WIDTH = 76


def display_width(text):
    return sum(2 if unicodedata.east_asian_width(char) in ("W", "F") else 1
               for char in str(text))


def pad_label(text, width=18):
    return str(text) + " " * max(0, width - display_width(text))


def separator(title="", major=False):
    char = "=" if major else "-"
    lead = f" {title} " if title else ""
    print(lead + char * max(3, CONSOLE_WIDTH - display_width(lead)), flush=True)


def metric(label, value, unit="", indent=4):
    print(" " * indent + pad_label(label) + str(value).rjust(14)
          + (f" {unit}" if unit else ""), flush=True)


def print_summary(label, item):
    separator(label, major="文件合计" in label or "总计" in label)
    metric("原始记录", f"{item['原始记录数']:,}", "条")
    metric("纳入汇总", f"{item['参与汇总记录数']:,}", "条")
    metric("剔除合计", f"{item['异常记录数']:,}", "条")
    separator("剔除明细")
    found = False
    for name in EXCLUSION_CATEGORIES:
        count = item[f"剔除_{name}"]
        if count:
            metric(name, f"{count:,}", "条")
            found = True
    if not found:
        print("    无剔除记录", flush=True)
    separator("数量核对")
    for label, field in [("原始数量", "原始数量"), ("汽运原始数量", "汽运原始数量"),
                         ("纳入数量", "纳入矩阵数量"), ("剔除数量", "剔除数量")]:
        metric(label, quantity_text(item[field]))
    if item['文本数量转换数']:
        metric("文本数量转换", f"{item['文本数量转换数']:,}", "条")
    if item['无效数量记录数']:
        metric("数量不可核对", f"{item['无效数量记录数']:,}", "条")


def key(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def blank(value):
    return value is None or (isinstance(value, str) and not value.strip())

def quantity(value):
    if blank(value) or isinstance(value, bool):
        raise ValueError("数量为空或不是有效数值")
    text = str(value).strip().replace("，", ",")
    if "," in text:
        text = re.sub(r"\s*,\s*", ",", text)
        if not re.fullmatch(r"[+]?[0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]+)?", text):
            raise ValueError("数量千位分隔格式错误")
        text = text.replace(",", "")
    try:
        result = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("数量不是有效数字") from exc
    if not result.is_finite() or result < 0:
        raise ValueError("数量必须为非负有限数字")
    if not math.isfinite(float(result)):
        raise ValueError("数量超出Excel支持范围")
    return result

def parse_day(value, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        parsed = from_excel(value, epoch)
        if isinstance(parsed, datetime):
            return parsed.date()
        raise ValueError("无效Excel日期序号")
    text = key(value).replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date()
    except ValueError as exc:
        raise ValueError("无法识别箱号日期，需要包含年份") from exc

def day_text(day):
    return f"{day.year:04d}/{day.month}/{day.day:02d}"

def token(value):
    # 用完整带类型值做唯一键，不使用有碰撞风险的短哈希。
    if value is None:
        return ["empty", ""]
    if isinstance(value, bool):
        return ["bool", value]
    if isinstance(value, (int, float, Decimal)):
        return ["number", str(Decimal(str(value)).normalize())]
    if isinstance(value, (date, datetime)):
        return ["date", value.isoformat()]
    return ["text", str(value)]

def plain(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    return value

def read_header(sheet, required):
    rows = iter(sheet.iter_rows(values_only=True))
    first = next(rows, ())
    if not any(not blank(v) for v in first):
        if any(any(not blank(v) for v in row) for row in rows):
            raise ValueError("第一行无表头但后续有数据")
        return None
    names, positions, counts = [], {}, Counter()
    for i, value in enumerate(first):
        name = key(value)
        if name:
            counts[name] += 1
        names.append(name or f"未命名列{i + 1}")
        positions[name] = i
    duplicates = [name for name, count in counts.items() if count > 1]
    if duplicates:
        raise ValueError("重复字段名：" + "、".join(duplicates))
    if len(set(names)) != len(names):
        raise ValueError("实际字段名与自动补充的未命名列冲突")
    missing = [name for name in required if name not in positions]
    if missing:
        raise ValueError("缺少字段：" + "、".join(missing))
    return names, positions, rows

class Audit:
    def __init__(self, path):
        self.path = path
        self.stream = path.open("w", encoding="utf-8")
        self.fields = {}
        self.count = 0
        self.fatal = 0
        self.reasons = Counter()
        self.sheet_summaries = []

    def add(self, level, file, sheet="", row="", reasons=(), first="", fields=None):
        self.count += 1
        self.fatal += level == "错误"
        self.reasons.update(reasons)
        fields = fields or {}
        for name in fields:
            self.fields.setdefault(name, None)
        item = [level, "；".join(reasons), file, sheet, row, first,
                {name: plain(value) for name, value in fields.items()}]
        self.stream.write(json.dumps(item, ensure_ascii=False) + "\n")

    def finish(self):
        self.stream.close()

def load_mapping(audit):
    parts = defaultdict(set)
    modules = defaultdict(list)
    module_names = []
    try:
        with closing(open_checked(RESOURCE_FILE)) as book:
            for title, required in [(PARTS_SHEET, ("SKU", "备件模块")),
                                    (MODULES_SHEET, ("备件模块", "自变量"))]:
                if title not in book.sheetnames:
                    raise ValueError(f"缺少表单：{title}")
                parsed = read_header(book[title], required)
                if parsed is None:
                    raise ValueError(f"{title}没有数据")
                names, columns, rows = parsed
                for line, row in enumerate(rows, 2):
                    if all(blank(v) for v in row):
                        continue
                    module = key(row[columns["备件模块"]])
                    if title == PARTS_SHEET:
                        sku = key(row[columns["SKU"]])
                        if sku:
                            parts[sku].add(module)
                    else:
                        index = len(module_names)
                        var = key(row[columns["自变量"]])
                        if not module or var != f"x{index + 1}":
                            audit.add("错误", RESOURCE_FILE.name, title, line,
                                      [f"模块为空或自变量顺序错误，应为x{index + 1}"],
                                      fields=dict(zip(names, row)))
                        module_names.append(module)
                        modules[module].append(index)
        if not module_names:
            raise ValueError("模块信息没有有效模块行")
        if len(module_names) + 1 > 16384:
            raise ValueError("模块数加日期列超过Excel列数上限")
    except Exception as exc:
        audit.add("错误", RESOURCE_FILE.name, reasons=[str(exc)])
    return parts, modules, module_names

def resolve(sku, parts, modules):
    choices = parts.get(sku, set())
    if not choices:
        return None, "货品编码无法匹配SKU"
    if len(choices) > 1:
        return None, "SKU对应多个不同模块"
    module = next(iter(choices))
    if not module:
        return None, "SKU对应的备件模块为空"
    indexes = modules.get(module, [])
    if not indexes:
        return None, f"备件模块无法匹配自变量：{module}"
    if len(indexes) > 1:
        return None, f"备件模块对应多个自变量：{module}"
    return indexes[0], ""

def discover():
    if not DATA_DIR.is_dir():
        raise FileNotFoundError(f"找不到业务目录：{DATA_DIR}")
    files = []
    for path in DATA_DIR.iterdir():
        match = re.fullmatch(r"(\d{4})\.(\d{2})\.xlsx", path.name, re.IGNORECASE)
        if path.is_file() and match:
            year, month = map(int, match.groups())
            date(year, month, 1)
            files.append((year, month, path))
    if not files:
        raise ValueError("未找到YYYY.MM.xlsx文件，例如2026.07.xlsx")
    periods = Counter((y, m) for y, m, _ in files)
    if any(count > 1 for count in periods.values()):
        raise ValueError("同一个月份存在多个文件，请先核实")
    return sorted(files, key=lambda item: (item[0], item[1], item[2].name))

def collect(files, parts, modules, audit, stage):
    stats = Counter()
    daily = defaultdict(lambda: [Decimal(0), Decimal(0), Decimal(0)])
    matrix = defaultdict(lambda: defaultdict(Decimal))
    dates = set()
    db = sqlite3.connect(stage / "duplicates.sqlite")
    # 去重存磁盘，避免数十万条完整记录长期占据Python内存。
    db.execute("CREATE TABLE seen (signature TEXT PRIMARY KEY, origin TEXT NOT NULL)")
    try:
        for year, month, path in files:
            print(flush=True)
            separator(f"读取文件  {path.name}", major=True)
            stats["文件数"] += 1
            file_sheets = []
            current = None
            try:
                with closing(open_checked(path)) as book:
                    for sheet_index, sheet in enumerate(book.worksheets):
                        current = new_sheet_stats(path.name, sheet.title)
                        file_sheets.append(current)
                        audit.sheet_summaries.append(current)
                        if sheet_index > 0:
                            print(flush=True)
                        separator(f"读取表单  {sheet.title}" + ("（隐藏表）" if sheet.sheet_state != "visible" else ""))
                        try:
                            parsed = read_header(sheet, REQUIRED)
                            if parsed is None:
                                stats["空表跳过数"] += 1
                                current["状态"] = "空表跳过"
                                print_summary("表单小计（空表）", current)
                                continue
                            names, columns, rows = parsed
                        except ValueError as exc:
                            audit.add("错误", path.name, sheet.title, 1, [str(exc)])
                            current["状态"] = "表头错误，未读取"
                            print(f"  表单读取失败：{exc}", flush=True)
                            continue
                        stats["工作表数"] += 1
                        progress_at = perf_counter()
                        for line, row in enumerate(rows, 2):
                            if all(blank(v) for v in row):
                                continue
                            stats["读取记录数"] += 1
                            current["原始记录数"] += 1
                            now = perf_counter()
                            if now - progress_at >= PROGRESS_SECONDS:
                                print(f"    读取中  本表 {current['原始记录数']:,} 条  |  全部累计 {stats['读取记录数']:,} 条", flush=True)
                                progress_at = now
                            raw = dict(zip(names, row))
                            # 字段之外有数据时仍保留，且参与去重。
                            for i in range(len(names), len(row)):
                                if row[i] is not None:
                                    raw[f"未命名列{i + 1}"] = row[i]
                            reasons = [f"{name}为空" for name in REQUIRED if blank(raw[name])]
                            day, qty, fatal = None, None, False
                            if not blank(raw["箱号日期"]):
                                try:
                                    day = parse_day(raw["箱号日期"], book.epoch)
                                    if (day.year, day.month) != (year, month):
                                        raise ValueError(f"箱号日期{day}不属于文件月份{year}.{month:02d}")
                                    dates.add(day)
                                except (ValueError, TypeError, OverflowError) as exc:
                                    reasons.append(str(exc))
                                    day, fatal = None, True
                            if not blank(raw["数量"]):
                                try:
                                    qty = quantity(raw["数量"])
                                except (ValueError, TypeError, OverflowError) as exc:
                                    reasons.append(str(exc))
                                    fatal = True
                            if not blank(raw["发运方式"]) and key(raw["发运方式"]) != "汽运":
                                reasons.append(f"非汽运：{key(raw['发运方式'])}")
                            index = None
                            if not blank(raw["货品编码"]):
                                index, error = resolve(key(raw["货品编码"]), parts, modules)
                                if error:
                                    reasons.append(error)
                            if qty is not None:
                                if isinstance(raw["数量"], str):
                                    current["文本数量转换数"] += 1
                                current["原始数量"] += qty
                                if key(raw["发运方式"]) == "汽运":
                                    current["汽运原始数量"] += qty
                            else:
                                current["无效数量记录数"] += 1
                            if fatal:
                                current["状态"] = "含严重数据错误"
                            # 不同表列序不影响判断；格式产生的未命名全空尾列不参与。
                            comparable = [(name, token(value)) for name, value in raw.items()
                                          if name != "箱号日期" and not (name.startswith("未命名列") and value is None)]
                            signature = json.dumps(sorted(comparable), ensure_ascii=False, separators=(",", ":"))
                            found = db.execute("SELECT origin FROM seen WHERE signature=?", (signature,)).fetchone()
                            first = found[0] if found else ""
                            if found:
                                reasons.append("除箱号日期外所有字段重复")
                            if day is not None and qty is not None:
                                daily[day][0] += qty
                            if reasons:
                                audit.add("错误" if fatal else "剔除", path.name, sheet.title,
                                          line, reasons, first, raw)
                                stats["异常记录数"] += 1
                                current["异常记录数"] += 1
                                current[f"剔除_{exclusion_category(reasons)}"] += 1
                                if qty is not None:
                                    current["剔除数量"] += qty
                                if day is not None and qty is not None:
                                    daily[day][2] += qty
                                if qty is None:
                                    stats["数量不可核对记录数"] += 1
                                if found:
                                    stats["重复剔除数"] += 1
                                    current["重复记录数"] += 1
                            else:
                                origin = f"{path.name} / {sheet.title} / 第{line}行 / {day_text(day)}"
                                db.execute("INSERT INTO seen VALUES (?,?)", (signature, origin))
                                matrix[day][index] += qty
                                daily[day][1] += qty
                                stats["参与汇总数"] += 1
                                current["参与汇总记录数"] += 1
                                current["纳入矩阵数量"] += qty
                        print(f"    读取完成  本表 {current['原始记录数']:,} 条  |  全部累计 {stats['读取记录数']:,} 条", flush=True)
                        if current["状态"] == "读取中":
                            current["状态"] = "完成"
                        if current["原始数量"] != current["纳入矩阵数量"] + current["剔除数量"]:
                            audit.add("错误", path.name, sheet.title, reasons=["表单数量核对失败"])
                            current["状态"] = "数量核对失败"
                        print_summary("表单小计", current)
                        db.commit()
            except Exception as exc:
                audit.add("错误", path.name, reasons=[f"文件读取或处理失败：{exc}"])
                if current is None:
                    current = new_sheet_stats(path.name, "无法读取")
                    file_sheets.append(current)
                    audit.sheet_summaries.append(current)
                current["状态"] = "文件读取失败（可能不完整）"
                print(f"  文件读取失败：{exc}", flush=True)
            print_summary(f"{path.name} 文件合计", summarize(file_sheets))
        for day, (source, included, excluded) in daily.items():
            if source != included + excluded or sum(matrix[day].values(), Decimal(0)) != included:
                audit.add("错误", "", reasons=[f"{day}数量核对失败"])
        if stats["读取记录数"] != stats["参与汇总数"] + stats["异常记录数"]:
            audit.add("错误", "", reasons=["记录条数核对失败"])
    finally:
        db.close()
    stats["日期数"] = len(dates)
    stats["全零日期数"] = sum(not any(matrix[d].values()) for d in dates)
    return stats, matrix, sorted(dates), daily

HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10, color="243746")
BLUE = PatternFill("solid", fgColor="305496")
STRIPE = PatternFill("solid", fgColor="F1F5FA")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def append_row(sheet, values, header=False, stripe=False):
    cells = []
    for value in values:
        if isinstance(value, Decimal):
            if not math.isfinite(float(value)):
                raise ValueError("汇总数值超出Excel支持范围")
            value = int(value) if value == value.to_integral_value() else float(value)
        cell = WriteOnlyCell(sheet, value=value)
        if isinstance(value, str):
            cell.data_type = "s"
        cell.font = HEADER_FONT if header else BODY_FONT
        cell.alignment = CENTER
        if header:
            cell.fill = BLUE
        elif stripe:
            cell.fill = STRIPE
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            cell.number_format = "0.############" if value % 1 else "0"
        cells.append(cell)
    sheet.append(cells)

def report_sheet(book, name, headers, widths=None):
    sheet = book.create_sheet(name)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "C2"
    sheet.sheet_format.defaultRowHeight = 24
    sheet.row_dimensions[1].height = 34
    for i, header in enumerate(headers, 1):
        sheet.column_dimensions[get_column_letter(i)].width = min(32, max(16, len(header) * 2 + 6))
    for col, width in (widths or {}).items():
        sheet.column_dimensions[col].width = width
    append_row(sheet, headers, header=True)
    return sheet

def write_report(target, audit, stats, daily, module_names):
    book = Workbook(write_only=True)
    fields = list(audit.fields)
    headers = ["级别", "异常原因", "来源文件", "来源表单", "原始行号", "首次保留位置"] + fields
    if len(headers) > 16384:
        raise ValueError("错误报告字段超过Excel列数上限")
    categories = report_sheet(book, "剔除分类", ["来源文件", "来源表单", "剔除类别", "记录数"])
    category_counts = Counter()
    for item in audit.sheet_summaries:
        for category in EXCLUSION_CATEGORIES:
            count = item[f"剔除_{category}"]
            if count:
                append_row(categories, [item["文件"], item["表单"], category, count])
                category_counts[category] += count
    for category, count in category_counts.items():
        append_row(categories, ["全部文件", "合计", category, count], stripe=True)
    append_row(categories, ["说明", "各类别互斥", "按优先级归类，完整原因见异常明细", None])
    review_headers = fields + ["来源文件", "来源表单", "原始行号", "完整异常原因"]
    review = report_sheet(book, "SKU未匹配复核", review_headers,
                          {get_column_letter(len(review_headers)): 50})
    review_count, review_part = 0, 1
    sheet = report_sheet(book, "异常明细", headers, {"B": 46, "F": 52})
    count, part = 0, 1
    with audit.path.open(encoding="utf-8") as stream:
        for line in stream:
            if count == MAX_DATA_ROWS:
                sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{count + 1}"
                part += 1
                sheet = report_sheet(book, f"异常明细_{part}", headers, {"B": 46, "F": 52})
                count = 0
            item = json.loads(line)
            # 同时含其他原因的记录也收录，每条源记录只复制一次，不按SKU去重。
            if "货品编码无法匹配SKU" in item[1].split("；"):
                if review_count == MAX_DATA_ROWS:
                    review.auto_filter.ref = f"A1:{get_column_letter(len(review_headers))}{review_count + 1}"
                    review_part += 1
                    review = report_sheet(book, f"SKU未匹配复核_{review_part}", review_headers,
                                          {get_column_letter(len(review_headers)): 50})
                    review_count = 0
                append_row(review, [item[6].get(name) for name in fields]
                           + [item[2], item[3], item[4], item[1]], stripe=review_count % 2 == 0)
                review_count += 1
            append_row(sheet, item[:6] + [item[6].get(name) for name in fields], stripe=count % 2 == 0)
            count += 1
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{count + 1}"
    review.auto_filter.ref = f"A1:{get_column_letter(len(review_headers))}{review_count + 1}"
    summary = report_sheet(book, "处理统计", ["项目", "结果"], {"A": 32, "B": 80})
    summary_rows = list(stats.items()) + [
        ("严重错误数", audit.fatal),
        ("结果状态", "失败：未发布新矩阵，旧矩阵已移除" if audit.fatal else "成功"),
        ("去重范围", "所有文件所有表单；只排除箱号日期，保留首条可汇总记录"),
        ("日期范围", "源数据中可解析且属于文件月份的箱号日期，不补自然日"),
        ("数量核对范围", "日期与数量有效的记录；无效数量不以0代替"),
        ("异常计数说明", "一条记录可有多个异常原因，原因计数不可直接相加")]
    for item in summary_rows:
        append_row(summary, item)
    for reason, count in audit.reasons.items():
        append_row(summary, [reason, count])
    sheet_headers = list(new_sheet_stats("", ""))
    overview = report_sheet(book, "表单汇总", sheet_headers, {"A": 22, "B": 22, "C": 34})
    for item in audit.sheet_summaries:
        append_row(overview, [item[name] for name in sheet_headers])
    overall = summarize(audit.sheet_summaries)
    overall["状态"] = "未发布矩阵" if audit.fatal else "完成"
    append_row(overview, [overall[name] for name in sheet_headers], stripe=True)
    file_view = report_sheet(book, "文件汇总", sheet_headers, {"A": 22, "B": 22, "C": 34})
    grouped = defaultdict(list)
    for item in audit.sheet_summaries:
        grouped[item["文件"]].append(item)
    for name, items in grouped.items():
        subtotal = summarize(items)
        subtotal.update({"文件": name, "表单": "文件合计",
                         "状态": "完成" if all(i["状态"] in ("完成", "空表跳过") for i in items) else "存在错误"})
        append_row(file_view, [subtotal[field] for field in sheet_headers])
    append_row(file_view, [overall[name] for name in sheet_headers], stripe=True)
    check = report_sheet(book, "日期核对", ["日期", "可核对原始数量", "纳入矩阵数量", "异常剔除数量", "差额"])
    for day, values in sorted(daily.items()):
        append_row(check, [day_text(day), *values, values[0] - values[1] - values[2]])
    mapping = report_sheet(book, "模块列对应", ["矩阵列序号", "自变量", "备件模块"])
    for i, module in enumerate(module_names, 1):
        append_row(mapping, [i, f"x{i}", module])
    if module_names:
        append_row(mapping, [len(module_names) + 1, "日期", "箱号日期（YYYY/M/DD）"])
    book.save(target)
    book.close()

def write_matrix(target, matrix, dates, module_count):
    if len(dates) > 1_048_576:
        raise ValueError("日期数超过Excel行数上限")
    book = Workbook(write_only=True)
    for title, keep_zero in [("矩阵_去除全零行", False), ("矩阵_保留全零行", True)]:
        sheet = book.create_sheet(title)
        sheet.sheet_view.showGridLines = False
        sheet.sheet_format.defaultColWidth = 12
        sheet.sheet_format.defaultRowHeight = 22
        sheet.column_dimensions[get_column_letter(module_count + 1)].width = 20
        for day in dates:
            values = matrix[day]
            if not keep_zero and not any(values.values()):
                continue
            append_row(sheet, [values.get(i, 0) for i in range(module_count)] + [day_text(day)])
    book.save(target)
    book.close()

def publish(prepared, stage):
    """先备份目标，再替换；可捕获的失败回滚，恢复失败保留备份目录。"""
    backups, installed = {}, []
    try:
        for i, target in enumerate(prepared):
            if target.exists():
                backup = stage / f"backup_{i}.xlsx"
                target.replace(backup)
                backups[target] = backup
        for target, source in prepared.items():
            if source is not None:
                source.replace(target)
                installed.append(target)
    except Exception as exc:
        errors = []
        for target in reversed(installed):
            try:
                target.unlink(missing_ok=True)
            except OSError as err:
                errors.append(str(err))
        for target, backup in backups.items():
            try:
                backup.replace(target)
            except OSError as err:
                errors.append(str(err))
        if errors:
            raise RuntimeError(f"发布失败且回滚不完整，恢复文件在{stage}：{errors}") from exc
        raise RuntimeError("输出发布失败，已恢复旧文件。请关闭Excel后重试；旧文件不是本次结果。") from exc

def run():
    started_at = perf_counter()
    separator("月度模块数量预处理", major=True)
    print(f"输入：{DATA_DIR}")
    print(f"输出：{OUTPUT_DIR}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    # 临时目录仅在成功/完整回滚后清理，遇到无法恢复的情况留存备份。
    import shutil
    stage = Path(tempfile.mkdtemp(prefix=".pre_execute_", dir=OUTPUT_DIR))
    audit = Audit(stage / "errors.jsonl")
    stats, matrix, dates, daily, module_names = Counter(), {}, [], {}, []
    try:
        with localcontext() as ctx:
            ctx.prec = 50
            try:
                parts, modules, module_names = load_mapping(audit)
                if not audit.fatal:
                    stats, matrix, dates, daily = collect(discover(), parts, modules, audit, stage)
            except Exception as exc:
                audit.add("错误", "", reasons=[str(exc)])
            finally:
                audit.finish()
            report = stage / "report.xlsx"
            if audit.count:
                write_report(report, audit, stats, daily, module_names)
            result = stage / "result.xlsx"
            if not audit.fatal:
                write_matrix(result, matrix, dates, len(module_names))
            publish({
                OUTPUT_DIR / "A_temp_erro.xlsx": report if audit.count else None,
                OUTPUT_DIR / "A_temp_result.xlsx": None if audit.fatal else result,
            }, stage)
    except Exception:
        audit.stream.close()
        print(f"处理未完成；请勿使用旧输出。恢复用临时目录：{stage}")
        raise
    else:
        shutil.rmtree(stage, ignore_errors=True)
    separator("运行结束  |  " + ("失败" if audit.fatal else "完成"), major=True)
    print(f"    文件 {stats['文件数']} 个  |  表单 {stats['工作表数']} 个  |  模块 {len(module_names):,} 个")
    print(f"    耗时 {perf_counter() - started_at:.1f} 秒  |  严重错误 {audit.fatal} 项")
    print_summary("全体文件总计", summarize(audit.sheet_summaries))
    if not audit.fatal:
        metric("源数据日期", len(dates), "天")
        metric("去除全零行后", len(dates) - stats['全零日期数'], "天")
    separator("输出文件")
    if not audit.fatal:
        print(f"    矩阵：{OUTPUT_DIR / 'A_temp_result.xlsx'}")
    if audit.count:
        print(f"    报告：{OUTPUT_DIR / 'A_temp_erro.xlsx'}")
        print("    人工复核：报告中的“SKU未匹配复核”表。")
    else:
        print("    无异常，不生成错误报告；如有旧报告，已删除。")
    separator("统计说明")
    print("    每条剔除记录只归入一个类别；全部原因保留在报告中。")
    print("    数量缺失或非法的记录不参与数量核对。")
    separator(major=True)
    return 1 if audit.fatal else 0

if __name__ == "__main__":
    try:
        sys.exit(run())
    except Exception as exc:
        print(f"[运行失败] {exc}")
        sys.exit(1)