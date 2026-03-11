import pytest
from parser.xlsx_parser import XLSXParser, WorkbookData


def test_parse_returns_workbook_data(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    assert isinstance(result, WorkbookData)


def test_parse_detects_formula(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    formula_cells = [c for c in result.cells.values() if c.formula is not None]
    assert len(formula_cells) > 0


def test_parse_is_hardcoded_correct(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    hardcoded = [c for c in result.cells.values() if c.is_hardcoded]
    assert len(hardcoded) > 0
    for cell in hardcoded:
        assert cell.formula is None
        assert isinstance(cell.value, (int, float))


def test_parse_named_range_found(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    assert len(result.named_ranges) > 0
    assert "discount_rate" in result.named_ranges


def test_parse_empty_file_no_crash(empty_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(empty_xlsx_path)
    assert isinstance(result, WorkbookData)
    assert len(result.cells) == 0


def test_parse_sheets_populated(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    assert "Assumptions" in result.sheets
    assert "DCF" in result.sheets


def test_parse_cell_address_format(sample_xlsx_path):
    parser = XLSXParser()
    result = parser.parse(sample_xlsx_path)
    for addr in result.cells.keys():
        assert "!" in addr, f"Cell address {addr} missing sheet prefix"
