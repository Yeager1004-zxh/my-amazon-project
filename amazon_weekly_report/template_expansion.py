from __future__ import annotations

import re
from copy import copy
from datetime import date

from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string

from .template_inspection import DAY_PATTERN, TemplateInspection, TemplateInspectionError
from .timeline import LaunchDay


class TemplateExpansionError(ValueError):
    """The template cannot be extended without risking its layout."""


def _day_rows(worksheet, summary: TemplateInspection) -> list[tuple[int, int]]:
    day_col = column_index_from_string(summary.day_column)
    rows: list[tuple[int, int]] = []
    for row in range(summary.header_row + 1, worksheet.max_row + 1):
        value = worksheet.cell(row, day_col).value
        match = DAY_PATTERN.fullmatch(str(value or "").strip())
        if match:
            rows.append((row, int(match.group(1))))
    expected = list(range(1, len(rows) + 1))
    if [day for _, day in rows] != expected:
        raise TemplateExpansionError("模板 D 编号不连续或顺序异常，停止扩行。")
    if any(rows[i + 1][0] != rows[i][0] + 1 for i in range(len(rows) - 1)):
        raise TemplateExpansionError("模板日期行不连续，停止扩行以免改坏其他内容。")
    return rows


def _style_source_row(worksheet, summary: TemplateInspection, day_number: int, current_count: int) -> int:
    # The observed template alternates weekly shading. Reuse the matching day
    # from week 1 or 2; when only one week exists, use week 1 throughout.
    week = (day_number - 1) // 7 + 1
    day_in_week = (day_number - 1) % 7
    style_week = 2 if week % 2 == 0 and current_count >= 14 else 1
    return summary.first_day_row + (style_week - 1) * 7 + day_in_week


def _copy_row_style(worksheet, source_row: int, target_row: int) -> None:
    source_dimension = worksheet.row_dimensions[source_row]
    target_dimension = worksheet.row_dimensions[target_row]
    target_dimension.height = source_dimension.height
    target_dimension.hidden = source_dimension.hidden
    target_dimension.outlineLevel = source_dimension.outlineLevel
    target_dimension.collapsed = source_dimension.collapsed
    for column in range(1, worksheet.max_column + 1):
        source = worksheet.cell(source_row, column)
        target = worksheet.cell(target_row, column)
        if isinstance(target, MergedCell):
            continue
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.data_type == "f" and isinstance(source.value, str):
            try:
                target.value = Translator(source.value, origin=source.coordinate).translate_formula(target.coordinate)
            except Exception as exc:
                raise TemplateExpansionError(
                    f"公式 {source.coordinate} 无法安全扩展到 {target.coordinate}。"
                ) from exc
        else:
            target.value = None
        target.comment = copy(source.comment) if source.comment else None


def _extend_conditional_formatting(worksheet, first_day_row: int, old_last_row: int, new_last_row: int) -> None:
    formatting = worksheet.conditional_formatting
    additions: list[tuple[str, object]] = []
    for conditional in list(formatting):
        coverage: dict[tuple[int, int], int] = {}
        for cell_range in conditional.sqref.ranges:
            if cell_range.min_row >= first_day_row and cell_range.max_row >= old_last_row:
                key = (cell_range.min_col, cell_range.max_col)
                coverage[key] = max(coverage.get(key, 0), cell_range.max_row)
        for (min_col, max_col), covered_to in coverage.items():
            if covered_to >= new_last_row:
                continue
            start = f"{get_column_letter(min_col)}{covered_to + 1}"
            end = f"{get_column_letter(max_col)}{new_last_row}"
            for rule in formatting[conditional]:
                additions.append((f"{start}:{end}", copy(rule)))
    for sqref, rule in additions:
        formatting.add(sqref, rule)


def _extend_data_validation(worksheet, first_day_row: int, old_last_row: int, new_last_row: int) -> None:
    for validation in worksheet.data_validations.dataValidation:
        coverage: dict[tuple[int, int], int] = {}
        for cell_range in validation.sqref.ranges:
            if cell_range.min_row >= first_day_row and cell_range.max_row >= old_last_row:
                key = (cell_range.min_col, cell_range.max_col)
                coverage[key] = max(coverage.get(key, 0), cell_range.max_row)
        additions: list[str] = []
        for (min_col, max_col), covered_to in coverage.items():
            if covered_to < new_last_row:
                start = f"{get_column_letter(min_col)}{covered_to + 1}"
                end = f"{get_column_letter(max_col)}{new_last_row}"
                additions.append(f"{start}:{end}")
        for target_range in additions:
            validation.add(target_range)


def ensure_day_rows(worksheet, summary: TemplateInspection, required_days: int) -> list[int]:
    """Extend a weekly template with recurring styles and formula translation.

    New rows follow the template's seven-day week layout. The former final row
    loses its table-ending border; that border is moved to the new final row.
    """
    if required_days < 1:
        raise TemplateExpansionError("周报至少需要 D01。")
    rows = _day_rows(worksheet, summary)
    if not rows:
        raise TemplateExpansionError("模板中没有可扩展的 D 日期行。")
    current_count = len(rows)
    if required_days <= current_count:
        return [row for row, _ in rows]

    old_last_row = rows[-1][0]
    first_day_row = rows[0][0]
    if any(row != first_day_row + index for index, (row, _day) in enumerate(rows)):
        raise TemplateExpansionError("日期行与模板连续行位置不一致，停止扩行。")
    new_last_row = first_day_row + required_days - 1
    if current_count >= 14:
        style_rows = (first_day_row, first_day_row + 13)
        if style_rows[1] > worksheet.max_row:
            raise TemplateExpansionError("模板缺少第二周格式，无法安全延续交替周次样式。")
    for merged in worksheet.merged_cells.ranges:
        cell_range = CellRange(str(merged))
        if cell_range.max_row >= first_day_row and cell_range.min_row <= new_last_row:
            raise TemplateExpansionError(
                f"日期区包含合并单元格 {merged}，当前不能安全扩展。"
            )

    terminal_day = max(current_count, summary.day_count)
    terminal_row = first_day_row + terminal_day - 1
    if terminal_row > worksheet.max_row:
        terminal_row = old_last_row
    final_borders = [copy(worksheet.cell(terminal_row, column).border.bottom) for column in range(1, worksheet.max_column + 1)]
    # Normalize the old endpoint to its regular weekly style before moving its
    # bottom border to the new endpoint.
    manual_column = column_index_from_string(summary.columns["日目标销量"]) if "日目标销量" in summary.columns else None

    def normalize_row_style(row_number: int, day_number: int) -> None:
        source_row = _style_source_row(worksheet, summary, day_number, current_count)
        for column in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row_number, column)
            source = worksheet.cell(source_row, column)
            if not isinstance(cell, MergedCell) and source.has_style:
                cell._style = copy(source._style)

    normalize_row_style(old_last_row, current_count)
    if terminal_row > old_last_row and terminal_row != new_last_row:
        normalize_row_style(terminal_row, terminal_day)

    for day_number in range(current_count + 1, required_days + 1):
        target_row = first_day_row + day_number - 1
        if target_row <= worksheet.max_row:
            day_cell = worksheet.cell(target_row, column_index_from_string(summary.day_column))
            if day_cell.value not in (None, ""):
                raise TemplateExpansionError(f"第 {target_row} 行已有非日期内容，不能安全写入 D{day_number:02d}。")
            occupied = []
            for column in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(target_row, column)
                if column == manual_column or cell.data_type == "f":
                    continue
                if cell.value not in (None, ""):
                    occupied.append(cell.coordinate)
            if occupied:
                raise TemplateExpansionError(
                    f"D{day_number:02d} 的预留行含有其他内容（{', '.join(occupied[:5])}），停止以免覆盖。"
                )
            if any(worksheet.cell(target_row, column).has_style for column in range(1, worksheet.max_column + 1)):
                continue  # Reuse the blank template row and preserve manual target values.
        source_row = _style_source_row(worksheet, summary, day_number, current_count)
        _copy_row_style(worksheet, source_row, target_row)

    for column, bottom in enumerate(final_borders, start=1):
        cell = worksheet.cell(new_last_row, column)
        if isinstance(cell, MergedCell):
            continue
        border = copy(cell.border)
        border.bottom = copy(bottom)
        cell.border = border

    _extend_conditional_formatting(worksheet, first_day_row, old_last_row, new_last_row)
    _extend_data_validation(worksheet, first_day_row, old_last_row, new_last_row)
    return [first_day_row + index for index in range(required_days)]


def write_timeline_rows(
    worksheet,
    summary: TemplateInspection,
    timeline: list[LaunchDay],
    row_numbers: list[int],
    *,
    clear_after: bool,
) -> None:
    from openpyxl.utils.cell import column_index_from_string

    week_column = column_index_from_string(summary.columns["周次"])
    day_column = column_index_from_string(summary.columns["天数"])
    week_start_column = column_index_from_string(summary.columns["周开始日期"])
    for index, row in enumerate(row_numbers):
        if index >= len(timeline):
            if clear_after:
                for column in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row, column)
                    if cell.data_type != "f":
                        cell.value = None
            continue
        item = timeline[index]
        week_cell = worksheet.cell(row, week_column)
        day_cell = worksheet.cell(row, day_column)
        if week_cell.data_type != "f":
            week_cell.value = item.week_label if item.is_week_start else None
        if day_cell.data_type != "f":
            day_cell.value = item.day_label
        week_date = item.date if item.is_week_start else None
        date_cell = worksheet.cell(row, week_start_column)
        if date_cell.data_type != "f":
            date_cell.value = week_date
        if week_date and date_cell.data_type != "f":
            date_cell.number_format = "yyyymmdd"


def parse_week_start_date(value: object) -> date | None:
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip())
        except ValueError:
            match = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", value.strip())
            if match:
                try:
                    return date(int(match[1]), int(match[2]), int(match[3]))
                except ValueError:
                    return None
    return None
