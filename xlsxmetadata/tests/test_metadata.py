import pytest
import metadata as m


def test_get_dimensions():
    assert m.get_dimensions('./test_file.xlsx', 'test_sheet') == dict(
        start_column=1,
        start_row=1,
        end_column=16384,
        end_row=1048576
    )


def test__parse_dimensions_():
    test_cases = {
        'A1:B2': dict(
            start_column=1,
            start_row=1,
            end_column=2,
            end_row=2
        ),
        'A1:XFD22000': dict(
            start_column=1,
            start_row=1,
            end_column=16384,
            end_row=22000
        ),
        'AAA22:XFD22000': dict(
            start_column=703,
            start_row=22,
            end_column=16384,
            end_row=22000
        )
    }

    for k, v in test_cases.items():
        assert m._parse_dimensions_(k) == v


def test_letters_to_number():
    test_cases = dict(
        XFD=16384,
        A=1,
        AA=27,
        ZA=677,
        AAB=704
    )
    for k, v in test_cases.items():
        assert m.letters_to_number(k) == v


def test_get_sheet_names():
    assert m.get_sheet_names('./test_file.xlsx') == {'test_sheet': 1}
