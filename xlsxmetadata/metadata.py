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
    dim_regex = re.compile(
        r'(?P<startcol>[A-Z]+)(?P<startrow>\d+):(?P<endcol>[A-Z]+)(?P<endrow>\d+)'
    )
    match = dim_regex.match(dim)
    return dict(
        start_column=letters_to_number(match.group('startcol')),
        start_row=int(match.group('startrow')),
        end_column=letters_to_number(match.group('endcol')),
        end_row=int(match.group('endrow'))
    )


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
    all_sheets_with_ids = {}
    for match in matches:
        dct = match.groupdict()
        all_sheets_with_ids.update({dct['name']: int(dct['id'])})
    return all_sheets_with_ids
