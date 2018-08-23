from zipfile import ZipFile
import re


def get_dimensions(xlsx_file, sheet_name):
    sheet_id = get_sheet_names(xlsx_file).get(sheet_name)
    full_sheet_name = f'xl/worksheets/sheet{sheet_id}.xml'
    dimension_tag = re.compile(r'<dimension\sref="(.*?)".*?\/>')
    with ZipFile(xlsx_file) as zipfile:
        with zipfile.open(full_sheet_name) as sheet:
            dimensions = _get_dim_recursive_(sheet, dimension_tag)
    return _parse_dimensions_(dimensions)


def _get_dim_recursive_(sheet, dimension_tag):
    chunk = sheet.read(1000)
    try:
        return dimension_tag.search(str(chunk)).group(1)
    except AttributeError:
        sheet.seek(-500, 1)
        return _get_dim_recursive_(sheet, dimension_tag)


def _parse_dimensions_(dim):
    dim = dim.split(':')
    dim_regex = re.compile(r'(?P<col>[A-Z]+)(?P<row>\d+)')
    if len(dim) == 1:
        match = dim_regex.match(dim[0])
        dimensions = dict(
            start_column=letters_to_number(match.group('col')),
            start_row=int(match.group('row')),
            end_column=letters_to_number(match.group('col')),
            end_row=int(match.group('row'))
        )
    elif len(dim) == 2:
        match_start = dim_regex.match(dim[0])
        match_end = dim_regex.match(dim[1])
        dimensions = dict(
            start_column=letters_to_number(match_start.group('col')),
            start_row=int(match_start.group('row')),
            end_column=letters_to_number(match_end.group('col')),
            end_row=int(match_end.group('row'))
        )
    else:
        raise AttributeError('The dimensions tag is invalid')
    return dimensions


def letters_to_number(letters):
    value_table = {a: n+1 for n, a in enumerate('ABCDEFGHIJKLMNOPQRSTUVWXYZ')}
    value_list = reversed([value_table[letter] for letter in letters])
    letter_value = sum([i * 26 ** p for p, i in enumerate(value_list)])
    return letter_value


def get_sheet_names(xlsx_file):
    sheet_name_regex = re.compile(
        r'<sheet\sname="(?P<name>.+?)"\ssheetId="(?P<id>\d+)"'
    )
    with ZipFile(xlsx_file) as zipfile:
        with zipfile.open('xl/workbook.xml') as wb:
            book_data = wb.read()
    matches = sheet_name_regex.finditer(str(book_data))
    return {match.groupdict()['name']: i+1 for i, match in enumerate(matches)}
