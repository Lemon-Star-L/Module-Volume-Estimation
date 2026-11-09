# @Author: YiLong Li
# @Date: 2026/11/9 13:52
# @File: pre_execute_V.py
# @Software: PyCharm


"""
输入：当前脚本目录 / data / Record.xlsx
输出：当前脚本目录 / output / Pre_V.xlsx、Pre_V_report.xlsx

主要功能：
    读取可见工作表，跳过隐藏表，清理重复及缺失记录，按日汇总方量。

处理规则：
    除序号外，整行所有字段完全一致时仅保留第一条。
    字段按原列位置和读取值比较，不去除文本空格或转换日期格式。
    日期或方量任一为空时不纳入统计，并记录原因。
    零方量不纳入统计；非空但无效的日期或方量仍报错停止。

输出规则：
    Pre_V.xlsx：日期、方量，日期升序，不补无业务日期。
    Pre_V_report.xlsx：异常明细、处理统计，每次成功运行刷新。
    明细保留来源、行号、原因、首次位置及原始字段。

使用说明：
    依赖 openpyxl，运行：python pre_execute_V.py。
    支持 --input、--output、--report 自定义路径。
    公式按缓存值读取，需先在Excel中计算并保存。
    不修改源文件，失败时请勿使用旧结果。
"""


import argparse
import logging
import os
import re
import tempfile
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, List, Tuple
from contextlib import closing

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

LOGGER = logging.getLogger(__name__)
BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR / 'data' / 'Record.xlsx'
OUTPUT_PATH = BASE_DIR / 'output' / 'Pre_V.xlsx'
REQUIRED = ('发运日期', '装载号', '方量')


def label(value):
    return re.sub(r'\s+', '', str(value or ''))


def normalize(value):
    if value is None:
        return ''
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def parse_date(value, sheet_name, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime):
            return converted.date()
        raise ValueError('无效Excel日期序号')
    s = str(value or '').strip()
    match = re.fullmatch(r'(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[ T].*)?', s)
    if match:
        return date(*map(int, match.groups()))
    match = re.fullmatch(r'(\d{1,2})[月/.-](\d{1,2})日?', s)
    if match:
        year = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', sheet_name)
        if not year:
            raise ValueError('日期缺少年份，工作表名也没有年份')
        return date(int(year[1]), int(match[1]), int(match[2]))
    raise ValueError(f'无法解析发运日期：{value!r}')


def parse_volume(value):
    if value is None or isinstance(value, bool):
        raise ValueError('方量为空或不是数值')
    try:
        result = Decimal(str(value).strip().replace(',', '').replace('，', ''))
    except InvalidOperation as exc:
        raise ValueError(f'无效方量：{value!r}') from exc
    if not result.is_finite() or result < 0:
        raise ValueError(f'方量必须为非负有限数值：{value!r}')
    return result


def collect_records(book: Workbook) -> Tuple[Dict[date, Decimal], List[Dict[str, Any]], Dict[str, int]]:
    """读取清洗数据；正常完成时返回汇总、异常记录及统计。"""
    reports: List[Dict[str, Any]] = []
    totals: Dict[date, Decimal] = defaultdict(Decimal)
    seen = {}
    stats: Dict[str, int] = {
        '工作表数': 0, '合计行数': 0, '读取记录数': 0,
        '重复删除数': 0, '日期或方量为空数': 0, '剔除记录数': 0,
        '零方量删除数': 0, '参与汇总数': 0,
    }
    for sheet in book.worksheets:
        if sheet.sheet_state != 'visible':
            stats['跳过隐藏表数'] = stats.get('跳过隐藏表数', 0) + 1
            print(f'跳过隐藏工作表：{sheet.title}')
            continue
        rows = iter(sheet.iter_rows(values_only=True))
        leading = []
        columns = None
        header_row = None
        for row_no, row in enumerate(rows, 1):
            leading.append(row)
            names = [label(v) for v in row]
            if all(name in names for name in REQUIRED):
                if any(names.count(name) != 1 for name in REQUIRED):
                    raise ValueError(f'{sheet.title}: 必要字段存在重名列')
                columns = {name: names.index(name) for name in REQUIRED}
                ignored = {i for i, name in enumerate(names) if name == '序号'}
                header_row = row_no
                source_headers = [str(v).strip() if v is not None else '' for v in row]
                break
            if row_no >= 20:
                break
        if columns is None:
            if any(any(v is not None for v in r) for r in leading) or any(
                any(v is not None for v in r) for r in rows
            ):
                raise ValueError(f'{sheet.title}: 前20行未找到必要表头 {REQUIRED}')
            continue
        stats['工作表数'] += 1
        started = False
        for row_no, row in enumerate(rows, header_row + 1):
            location = f'{sheet.title}!第{row_no}行'
            if not any(v is not None and str(v).strip() for v in row):
                continue
            raw_date, raw_load, raw_volume = [row[columns[n]] for n in REQUIRED]
            # 截图的第2行属于子表头，三个必要字段对应格全部为空。
            if not started and all(v is None for v in (raw_date, raw_load, raw_volume)):
                parent = ''
                for i, sub in enumerate(row):
                    if i < len(source_headers) and source_headers[i]:
                        parent = source_headers[i]
                    if sub is not None and str(sub).strip():
                        source_headers[i] = f'{parent}-{str(sub).strip()}' if parent else str(sub).strip()
                continue
            started = True
            if all(label(row[columns[n]]) == n for n in REQUIRED):
                continue
            # 明确的合计行不当作发运记录。
            if any(label(v) in ('合计', '总计', '小计') for v in row[:columns['发运日期'] + 1]) and raw_load is None:
                stats['合计行数'] += 1
                continue
            stats['读取记录数'] += 1
            # 排除序号，其余字段按读取值及原始列位置精确比较。
            # 忽略工作表格式导致的末尾全空列，不改变有效字段内容。
            values = [v for i, v in enumerate(row) if i not in ignored]
            while values and values[-1] is None:
                values.pop()
            signature = tuple((type(v).__name__, v) for v in values)
            first = seen.get(signature)
            missing = []
            if raw_date is None or (isinstance(raw_date, str) and not raw_date.strip()):
                missing.append('发运日期为空')
            if raw_volume is None or (isinstance(raw_volume, str) and not raw_volume.strip()):
                missing.append('方量为空')
            reasons = []
            if first is not None:
                reasons.append('整行重复（不含序号）')
                stats['重复删除数'] += 1
            else:
                seen[signature] = location
            if missing:
                reasons.extend(missing)
                stats['日期或方量为空数'] += 1
            if reasons:
                fields = {}
                used = {}
                for i, value in enumerate(row):
                    name = source_headers[i] if i < len(source_headers) else ''
                    if not name and value is None:
                        continue
                    name = name or f'未命名列{i + 1}'
                    used[name] = used.get(name, 0) + 1
                    if used[name] > 1:
                        name = f'{name}（{used[name]}）'
                    if (i == columns['发运日期'] or any(key in name for key in
                            ('应该到达时间', '实际到达时间', '应该卸货时间', '实际卸货时间'))) and value is not None:
                        try:
                            value = parse_date(value, sheet.title, book.epoch)
                        except (ValueError, OverflowError, TypeError):
                            pass
                    fields[name] = value
                reports.append({'fields': fields, 'sheet': sheet.title,
                                'row': row_no, 'reason': '；'.join(reasons),
                                'first': first or ''})
                stats['剔除记录数'] += 1
                continue
            try:
                volume = parse_volume(raw_volume)
                if volume == 0:
                    stats['零方量删除数'] += 1
                    continue
                day = parse_date(raw_date, sheet.title, book.epoch)
            except (ValueError, OverflowError) as exc:
                raise ValueError(f'{location}: {exc}') from exc
            totals[day] += volume
            stats['参与汇总数'] += 1
    return dict(totals), reports, stats


def style_table(sheet, widths, freeze='A2', filters=True):
    """统一蓝白表格样式，保留首行字段与原有数据结构。"""
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 100
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.sheet_properties.outlinePr.summaryRight = False
    sheet.freeze_panes = freeze
    sheet.sheet_properties.tabColor = '305496'
    sheet.row_dimensions[1].height = 34
    if filters:
        sheet.auto_filter.ref = sheet.dimensions
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for cell in sheet[1]:
        cell.font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='305496')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(right=Side(style='thin', color='627FA9'))
    for row in sheet.iter_rows(min_row=2):
        sheet.row_dimensions[row[0].row].height = 26
        for cell in row:
            cell.font = Font(name='微软雅黑', size=10, color='243746')
            cell.fill = PatternFill('solid', fgColor='F1F5FA' if cell.row % 2 == 0 else 'FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(bottom=Side(style='hair', color='DCE5EF'))
    sheet.print_title_rows = '1:1'
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.orientation = 'landscape' if sheet.max_column > 5 else 'portrait'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0


def write_results(totals: Dict[date, Decimal], reports: List[Dict[str, Any]],
                  stats: Dict[str, int], destination: Path,
                  report_destination: Path) -> None:
    """独立生成每日汇总及处理报告。"""
    result = Workbook()
    ws = result.active
    ws.title = 'Sheet1'
    ws.append(['日期', '方量'])
    for day, volume in sorted(totals.items()):
        ws.append([f'{day.year:04d}/{day.month}/{day.day:02d}', float(volume)])
        ws.cell(ws.max_row, 1).number_format = '@'
        ws.cell(ws.max_row, 2).number_format = '0.############'
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 18
    ws.freeze_panes = 'A2'
    stats['输出天数'] = len(totals)
    stats['报告记录数'] = len(reports)
    report = Workbook()
    detail = report.active
    detail.title = '异常明细'
    field_names = list(dict.fromkeys(name for record in reports for name in record['fields']))
    has_duplicates = any(record['first'] for record in reports)
    metadata = ['剔除原因', '来源工作表', '原始行号']
    if has_duplicates:
        metadata.append('重复记录首次位置')
    short_names = {
        '司机信息-姓名': '司机姓名', '司机信息-联系方式': '联系方式',
        '车辆发运信息-车牌号': '车牌号', '车辆发运信息-挂牌号': '挂牌号',
        '车辆发运信息-车辆规格': '车辆规格', '车辆发运信息-是否整车': '是否整车',
        '车辆到达信息-合同时间（天）': '合同时间（天）',
        '车辆到达信息-应该到达时间': '应该到达时间',
        '车辆到达信息-实际到达时间': '实际到达时间',
        '车辆到达信息-延迟天数': '到达延迟天数',
        '车辆卸货信息-应该卸货时间': '应该卸货时间',
        '车辆卸货信息-实际卸货时间': '实际卸货时间',
        '车辆卸货信息-延迟天数': '卸货延迟天数',
    }
    headers = metadata + [short_names.get(name, name) for name in field_names]
    detail.append(headers)
    for record in reports:
        values = [record['reason'], record['sheet'], record['row']]
        if has_duplicates:
            values.append(record['first'])
        values += [record['fields'].get(name) for name in field_names]
        detail.append(values)
        for cell in detail[detail.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'
            elif isinstance(cell.value, (date, datetime)):
                cell.number_format = 'yyyy/m/dd'
            elif isinstance(cell.value, (float, int)):
                cell.number_format = '0.############' if cell.value % 1 else '0'
    detail.freeze_panes = f'{get_column_letter(len(metadata) + 1)}2'
    detail.auto_filter.ref = detail.dimensions
    widths = {}
    for i, name in enumerate(headers, 1):
        # 预留筛选按钮宽度，表头不与按钮挤在一起。
        width = max(14, min(24, len(name) * 2 + 6))
        if name == '剔除原因':
            width = 26
        elif name == '重复记录首次位置':
            width = 30
        elif name in ('联系方式', '装载号', '车牌号', '挂牌号'):
            width = 19
        elif name in ('发运日期', '应该到达时间', '实际到达时间', '应该卸货时间', '实际卸货时间'):
            width = 20
        widths[get_column_letter(i)] = width
    style_table(detail, widths, f'{get_column_letter(len(metadata) + 1)}2')
    for row in detail.iter_rows(min_row=2):
        row[0].fill = PatternFill('solid', fgColor='FFF4D6')
        row[0].alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
        if row[0].value and len(str(row[0].value)) > 14:
            detail.row_dimensions[row[0].row].height = 40
    summary = report.create_sheet('处理统计')
    summary.append(['项目', '数量'])
    for key, value in stats.items():
        summary.append([key, value])
    summary.append(['说明', '重复与缺失可能同时命中，分类计数不可直接相加'])
    summary.column_dimensions['A'].width = 26
    summary.column_dimensions['B'].width = 65
    style_table(ws, {'A': 22, 'B': 26})
    for row in ws.iter_rows(min_row=2):
        row[1].alignment = Alignment(horizontal='right', vertical='center', indent=1)
        # 至少显示3位小数，其余有效小数保留，不改变底层数值。
        row[1].number_format = '#,##0.000#########'
    style_table(summary, {'A': 28, 'B': 60}, filters=False)
    summary.row_dimensions[summary.max_row].height = 40
    summary.cell(summary.max_row, 2).alignment = Alignment(
        horizontal='left', vertical='center', wrap_text=True, indent=1)
    # 两个文件先全部写入临时文件，再替换正式文件。
    pending = []
    try:
        for workbook, target in [(report, report_destination), (result, destination)]:
            target.parent.mkdir(parents=True, exist_ok=True)
            fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=target.parent)
            os.close(fd)
            pending.append((temp, target))
            workbook.save(temp)
        for temp, target in pending:
            os.replace(temp, target)
    finally:
        result.close()
        report.close()
        for temp, _ in pending:
            if os.path.exists(temp):
                os.unlink(temp)


def preprocess(input_path=INPUT_PATH, output_path=OUTPUT_PATH, report_path=None) -> Dict[str, int]:
    """依次执行读取、校验、输出；错误继续向调用方抛出。"""
    source, destination = Path(input_path), Path(output_path)
    if source.resolve() == destination.resolve():
        raise ValueError('输入和输出不能是同一个文件')
    report_destination = Path(report_path) if report_path else destination.with_name('Pre_V_report.xlsx')
    if len({source.resolve(), destination.resolve(), report_destination.resolve()}) != 3:
        raise ValueError('源文件、汇总文件、报告文件不能使用相同路径')
    with closing(load_workbook(source, data_only=True, read_only=True)) as book:
        totals, reports, stats = collect_records(book)
    if stats.get('工作表数', 0) == 0:
        raise ValueError('未找到可处理的工作表')
    write_results(totals, reports, stats, destination, report_destination)
    return stats


def main():
    parser = argparse.ArgumentParser(description='每日方量汇总预处理')
    parser.add_argument('--input', default=str(INPUT_PATH), help='源Excel路径')
    parser.add_argument('--output', default=str(OUTPUT_PATH), help='结果Excel路径')
    parser.add_argument('--report', default=None, help='报告路径，默认与汇总文件同目录')
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    try:
        stats = preprocess(args.input, args.output, args.report)
    except Exception:
        LOGGER.exception('处理失败，请勿使用本次输出或旧结果；请修正后重新运行')
        return 1
    destination = Path(args.output).resolve()
    report_destination = (Path(args.report).resolve() if args.report
                          else destination.with_name('Pre_V_report.xlsx'))
    print('\n每日方量预处理完成')
    print(f"  工作表：{stats.get('工作表数', 0)} 个")
    print(f"  跳过隐藏表：{stats.get('跳过隐藏表数', 0)} 个")
    print(f"  读取记录：{stats.get('读取记录数', 0):,} 条")
    print(f"  参与汇总：{stats.get('参与汇总数', 0):,} 条")
    print(f"  输出日期：{stats.get('输出天数', 0):,} 天")
    print('\n剔除情况')
    print(f"  整行重复（不含序号）：{stats.get('重复删除数', 0):,} 条")
    print(f"  日期或方量为空：{stats.get('日期或方量为空数', 0):,} 条")
    print(f"  零方量：{stats.get('零方量删除数', 0):,} 条")
    print(f"  跳过合计行：{stats.get('合计行数', 0):,} 条")
    print(f"  报告记录：{stats.get('报告记录数', 0):,} 条")
    print('  注：重复与缺失可能同时命中，不重复计入报告行数。')
    print('\n输出文件')
    print(f'  汇总：{destination}')
    print(f'  报告：{report_destination}')
    if stats.get('报告记录数', 0):
        print('\n提示：存在剔除记录，请查看报告中的“异常明细”。')
    print()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())