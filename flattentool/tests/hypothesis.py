from __future__ import unicode_literals
import pytest
import os
from flattentool import output, schema
from flattentool.sheet import Sheet
import openpyxl
from hypothesis import given, assume, strategies
import csv


class MockParser(object):
    def __init__(self, main_sheet, sub_sheets):
        self.main_sheet = Sheet(main_sheet)
        self.sub_sheets = {k:Sheet(v) for k,v in sub_sheets.items()}


@given(column_name=strategies.text(), cell_value=strategies.text())
def test_text_xlsx(tmpdir, column_name, cell_value):
    " Test the basic one column case with lots of interesting text (xlsx) "

    for s in [column_name, cell_value]:
        assume(s != '')
        assume('\r' not in s)

    parser = MockParser([column_name], {})
    parser.main_sheet.lines = [{column_name: cell_value}]
    try:
        spreadsheet_output = output.FORMATS['xlsx'](
            parser=parser,
            main_sheet_name='release',
            output_name=os.path.join(tmpdir.strpath, 'release'+output.FORMATS_SUFFIX['xlsx']))
        spreadsheet_output.write_sheets()
    except openpyxl.exceptions.IllegalCharacterError:
        return

    wb = openpyxl.load_workbook(tmpdir.join('release.xlsx').strpath)
    assert wb.get_sheet_names() == ['release']
    assert len(wb['release'].rows) == 2
    assert [ x.value for x in wb['release'].rows[0] ] == [ column_name ]
    assert [ x.value for x in wb['release'].rows[1] ] == [ cell_value ]


@given(column_name=strategies.text(), cells=strategies.lists(strategies.text()))
def test_text_csv(tmpdir, column_name, cells):
    " Test the basic one column case with lots of interesting text (csv) "

    for s in [column_name] + cells:
        assume('\r' not in s)
        assume('\0' not in s)

    parser = MockParser([column_name], {})
    parser.main_sheet.lines = [{column_name: x} for x in cells]
    spreadsheet_output = output.FORMATS['csv'](
        parser=parser,
        main_sheet_name='release',
        output_name=os.path.join(tmpdir.strpath, 'release'+output.FORMATS_SUFFIX['csv']))
    spreadsheet_output.write_sheets()

    # Check CSV
    assert set(tmpdir.join('release').listdir()) == set([
        tmpdir.join('release').join('release.csv'),
    ])
    release_csv_list = list(csv.reader(tmpdir.join('release', 'release.csv').open()))
    assert [x[0] for x in release_csv_list] == [column_name] + cells
