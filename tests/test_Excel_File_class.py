import pytest
from excel_lib import ExcelFile
import io
import pandas as pd

# Path to test files directory
TEST_FILES_DIR = "tests/files/"

# Configuration for section names
SECTIONS_CONFIG = {
    "SECTION_STATION_TAKEOVER_DIVIDER": ["STACJA ŁADOWANIA – DANE"],
    "SECTION_CONTACT_PERSON": ["OSOBA KONTAKTOWA - EKSPOLATACJA STACJI"],
    "SECTION_RESPONSIBLE_PERSON": ["OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA"]
}

@pytest.fixture
def load_excel_file(request):
    """Loads a real Excel file from the test directory."""
    file_name = request.param
    file_path = f"{TEST_FILES_DIR}{file_name}"
    
    with open(file_path, "rb") as f:
        return io.BytesIO(f.read())

@pytest.mark.parametrize("load_excel_file", ["valid.xlsx"], indirect=True)
def test_valid_excel(load_excel_file):
    """Tests a valid Excel file."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert excel.worksheet_count == 1
    assert not excel.non_cell_objects  # No non-cell objects should be detected

@pytest.mark.parametrize("load_excel_file", ["image_in_cell.xlsx"], indirect=True)
def test_image_in_cell(load_excel_file):
    """Tests an Excel file with an image anchored in a cell."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) > 0

@pytest.mark.parametrize("load_excel_file", ["image_outside_cell.xlsx"], indirect=True)
def test_image_outside_cell(load_excel_file):
    """Tests an Excel file with an image not anchored in any cell."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) > 0

@pytest.mark.parametrize("load_excel_file", ["image_in_and_outside_cell.xlsx"], indirect=True)
def test_image_in_and_outside_cell(load_excel_file):
    """Tests an Excel file with both anchored and unanchored images."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) > 1

@pytest.mark.parametrize("load_excel_file", ["multiple_sheets.xlsx"], indirect=True)
def test_multiple_sheets(load_excel_file):
    """Tests an Excel file containing multiple sheets."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert excel.worksheet_count > 1

@pytest.mark.parametrize("load_excel_file", ["all_in_one.xlsx"], indirect=True)
def test_all_in_one(load_excel_file):
    """Tests an Excel file containing multiple sheets and images."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) > 1
    assert excel.worksheet_count > 1

@pytest.mark.parametrize("load_excel_file", ["2_image_in_cells.xlsx"], indirect=True)
def test_two_images_in_cells(load_excel_file):
    """Tests an Excel file with two images anchored in different cells."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) == 2

@pytest.mark.parametrize("load_excel_file", ["3_image_in_cells_one_far.xlsx"], indirect=True)
def test_three_images_one_far(load_excel_file):
    """Tests an Excel file with three images, one far from data range."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert len(excel.non_cell_objects) == 3

@pytest.fixture
def sample_excel_with_sections():
    """Creates an in-memory Excel file with multiple sections for testing _identify_sections."""
    output = io.BytesIO()
    df = pd.DataFrame({
        "A": [
            "liczba",
            1,  # Initial numbering for Lp
            2,
            "STACJA ŁADOWANIA – DANE",  # Section header
            1,  # Restart numbering within section
            2,
            3,
            "Adres",
            4,
            5,
            "TERMINAL",  # Section header
            None,  # Empty section between headers
            1,  # Restart numbering within section
            2,
            3,
            4,
            "Adres",  # Repeated key in a different section
            5,
            "OSOBA KONTAKTOWA - EKSPLOATACJA STACJI",  # Section header
            1,  # Restart numbering within section
            2,
            3,
            4,
            5,
            "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA",  # Mandatory section
            1,
            2,
            3,
            4,
            5
        ],
        "B": [
            "Lp",
            "Osoba odpowiedzialna",
            "Numer jobu",
            None,  # Section header
            "Deadline",
            "Pełna Nazwa Klienta",
            "Model",
            "Adres",
            "Rodzaj stacji (DC / AC)",
            "Adres",
            None,  # Section header
            None,  # Empty section simulated
            "Czy stacja będzie z terminalem?",
            "Kto dostarcza terminal?",
            "Model terminala",
            "Numer seryjny terminala",
            None,
            None,
            None,  # Section header
            "Imię i nazwisko",
            "Numer telefonu",
            "Adres e-mail",
            None,
            None,
            None,  # Mandatory section header
            "Imię i nazwisko",
            "Numer telefonu",
            "Adres e-mail",
            "Test Key 4",
            "Test Key 5"
        ],
        "C": [
            "1",
            "Adam Nijaki",
            None,
            None,  # Merged cells for section header
            "11.11.2024",
            "ELOCITY",
            "LS4",
            "Testowa Ulica 1",  # First occurrence of "Adres"
            "AC",
            "Testowa Ulica 2",  # Repeated key "Adres" in a different section
            None,
            None, 
            "Tak",
            "elocity",
            "PAX IM 30",
            "123456789",
            None,
            None, 
            None,
            "Jan Kowalski",
            "987654321",
            "jan.kowalski@mail.com",
            None,
            None,
            None,  # Mandatory section header
            "Test Value 1",
            "Test Value 2",
            "Test Value 3",
            "Test Value 4",
            "Test Value 5"
        ]
    })
    df.to_excel(output, index=False, header=False)
    output.seek(0)
    return output

@pytest.fixture
def excel_instance_with_sections(sample_excel_with_sections):
    """Creates an instance of ExcelFile for testing create_template_structure."""
    sections_config = SECTIONS_CONFIG
    return ExcelFile(sample_excel_with_sections, sections_config)

# Tests for find_row_for_key

def test_find_existing_keys_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key finds the correct indices for multiple existing keys."""
    test_cases = {
        "Model": 5,
        "Osoba odpowiedzialna": 0,
        "Numer seryjny terminala": 14
    }
    
    for key, expected_index in test_cases.items():
        row_index = excel_instance_with_sections.find_row_for_key(key)
        assert row_index == expected_index, f"Expected index {expected_index} for key '{key}', but got {row_index}"

def test_find_non_existing_key_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key returns -1 for a non-existing key."""
    row_index = excel_instance_with_sections.find_row_for_key("NonExistingKey")
    assert row_index == -1, f"Expected -1, but got {row_index}"

def test_find_key_with_whitespace_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key handles keys with additional spaces correctly."""
    row_index = excel_instance_with_sections.find_row_for_key(" Model ")  # Extra spaces
    assert row_index == 5, f"Expected index 5, but got {row_index}"


def test_find_key_outside_specified_section(excel_instance_with_sections):
    """Tests if find_row_for_key returns -1 for a key that exists but is outside the specified section."""
    row_index = excel_instance_with_sections.find_row_for_key("Adres", "OSOBA KONTAKTOWA - EKSPLOATACJA STACJI")
    assert row_index == -1, f"Expected -1 for 'Adres' in section 'OSOBA KONTAKTOWA - EKSPLOATACJA STACJI', but got {row_index}"


def test_find_key_in_global_section(excel_instance_with_sections):
    """Tests if find_row_for_key correctly handles keys in global_data."""
    row_index = excel_instance_with_sections.find_row_for_key("Numer jobu", "global_data")
    assert row_index == 1, f"Expected index 1 for 'Numer jobu' in global_data, but got {row_index}"


def test_find_key_in_section_not_identified(excel_instance_with_sections):
    """Tests if find_row_for_key raises a ValueError when the section is not found."""
    row_index = excel_instance_with_sections.find_row_for_key("Model", "NonExistingSection")
    assert row_index == 5, f"Expected index 5, but got {row_index}"

def test_ignore_section_headers(excel_instance_with_sections):
    """Tests if find_row_for_key ignores section headers when searching for keys."""
    row_index = excel_instance_with_sections.find_row_for_key("STACJA ŁADOWANIA – DANE")
    assert row_index == -1, f"Expected -1 for section header, but got {row_index}"

def test_identify_sections(excel_instance_with_sections):
    """Tests if _identify_sections correctly identifies sections in the Excel file."""
    sections = excel_instance_with_sections._identify_sections()
    expected_sections = {
        "STACJA ŁADOWANIA – DANE": [3, 8],
        "TERMINAL": [10, 16],
        "OSOBA KONTAKTOWA - EKSPLOATACJA STACJI": [18, 22],
        "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA": [24, 28]
    }
    assert sections == expected_sections, f"Expected {expected_sections}, but got {sections}"

def test_empty_section_between_headers(excel_instance_with_sections):
    """Tests behavior when a section header is followed by no data."""
    sections = excel_instance_with_sections._identify_sections()
    assert "OSOBA KONTAKTOWA - EKSPLOATACJA STACJI" in sections, "Section header identified incorrectly."

def test_repeated_keys_sections(excel_instance_with_sections):
    """Tests behavior when the same key appears in different sections."""
    row_index = excel_instance_with_sections.find_row_for_key("Adres")
    assert row_index == [6, 8], f"Expected index 9 for first occurrence of 'Adres', but got {row_index}"

def test_section_with_lowercase_header(excel_instance_with_sections):
    """Tests if _identify_sections correctly ignores lowercase section headers."""
    sections = excel_instance_with_sections._identify_sections()
    assert all(header.isupper() for header in sections.keys()), "Lowercase section headers should not be identified."

def test_find_key_with_empty_value(excel_instance_with_sections):
    """Tests if find_row_for_key correctly handles keys where the corresponding value is empty."""
    row_index = excel_instance_with_sections.find_row_for_key("Model terminala")
    assert row_index == 13, f"Expected index 13 for 'Model terminala', but got {row_index}"

# Tests for create_template_structure

def test_create_template_structure(excel_instance_with_sections):
    """Tests if create_template_structure correctly extracts the expected structure."""
    template_structure = excel_instance_with_sections.create_template_structure()
    
    assert "takeover" in template_structure, "Template structure should contain 'takeover'."
    assert "stations" in template_structure, "Template structure should contain 'stations'."
    assert isinstance(template_structure["takeover"].get("global_data"), dict), "Global data should be a dictionary."
    assert isinstance(template_structure["takeover"].get("contact_person"), (dict, type(None))), "Contact person should be a dictionary or None."
    assert isinstance(template_structure["takeover"].get("responsible_person"), (dict, type(None))), "Responsible person should be a dictionary or None."
    assert isinstance(template_structure["stations"], dict), "Stations should be a dictionary."

def test_template_structure_integrity(excel_instance_with_sections):
    """Tests if create_template_structure extracts the exact content from the sections."""
    template_structure = excel_instance_with_sections.create_template_structure()
    
    for section, fields in template_structure["stations"].items():
        for key, row_index in fields.items():
            assert excel_instance_with_sections.worksheet.iat[row_index, 1] == key, \
                f"Expected key '{key}' in section '{section}' at row {row_index}, but found different value."
    
    for key, values in template_structure["takeover"].items():
        if isinstance(values, dict):
            for field_key, row_index in values.items():
                assert excel_instance_with_sections.worksheet.iat[row_index, 1] == field_key, \
                    f"Expected key '{field_key}' in takeover '{key}' at row {row_index}, but found different value."


# Tests for compare_structure_with_file

@pytest.fixture
def standard_template():
    """Standard template with full structure based on dataset."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 0, "Numer jobu": 1},
            "contact_person": {"Imię i nazwisko": 18, "Numer telefonu": 19},
            "responsible_person": {"Imię i nazwisko": 24, "Numer telefonu": 25}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 5, "Adres": 8},
            "TERMINAL": {"Czy stacja będzie z terminalem?": 11}
        }
    }

@pytest.fixture
def missing_values_template():
    """Template with missing sections and values."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 1},
            "contact_person": None,  # Missing contact_person
            "responsible_person": {"Imię i nazwisko": 25}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 6},  # Missing "Adres"
            "TERMINAL": {}  # Empty section
        }
    }

@pytest.fixture
def wrong_value_in_mandatory_section_template():
    """Standard template with full structure based on dataset."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 1, "Numer jobu": 2},
            "contact_person": 1,
            "responsible_person": {"Imię i nazwisko": 25, "Numer telefonu": 26}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 6, "Adres": 8},
            "TERMINAL": {"Czy stacja będzie z terminalem?": 13}
        }
    }

@pytest.fixture
def swapped_values_template():
    """Template with incorrect values in some fields."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 3},  # Wrong row number
            "contact_person": {"Imię i nazwisko": 18},  # Incorrect index
            "responsible_person": {"Imię i nazwisko": 24}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 10},  # Incorrect row index
            "TERMINAL": {"Czy stacja będzie z terminalem?": 12}  # Wrong row
        }
    }

@pytest.fixture
def duplicate_value_template():
    """Template with duplicate values that should trigger an error."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 1, "Numer jobu": 2},
            "contact_person": {"Imię i nazwisko": 19}, 
            "responsible_person": {"Imię i nazwisko": 19}  # Duplicate row index
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Adres": 9}  # Wrong value for a duplicate
        }
    }

@pytest.fixture
def empty_template():
    """Empty template structure."""
    return {
        "takeover": {
            "global_data": {},
            "contact_person": {}, 
            "responsible_person": {}
        },
        "stations": {}
    }

# 1. Test ensuring that the function correctly updates the row indices.
def test_compare_structure_with_file(excel_instance_with_sections, standard_template):
    """Tests if compare_structure_with_file correctly updates the template structure."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(standard_template)
    
    for section, fields in updated_structure["stations"].items():
        for key, row_index in fields.items():
            assert excel_instance_with_sections.worksheet.iat[row_index, 1] == key, \
                f"Expected key '{key}' in section '{section}' at row {row_index}, but found different value."
    
    for key, values in updated_structure["takeover"].items():
        if isinstance(values, dict):
            for field_key, row_index in values.items():
                assert excel_instance_with_sections.worksheet.iat[row_index, 1] == field_key, \
                    f"Expected key '{field_key}' in takeover '{key}' at row {row_index}, but found different value."

# 2. Test case where multiple matches exist for a key.
def test_compare_structure_with_multiple_matches(excel_instance_with_sections, duplicate_value_template):
    """Tests if compare_structure_with_file raises an error when multiple matches are found for a key."""
    # Checks, if error is raised
    with pytest.raises(ValueError, match="Multiple matches found for key 'Adres'"):
        excel_instance_with_sections.compare_structure_with_file(duplicate_value_template)

# 3. Test handling of missing keys in the template.
def test_compare_structure_with_missing_keys(excel_instance_with_sections, missing_values_template):
    """Tests if compare_structure_with_file correctly handles missing keys in the template."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(missing_values_template)
    assert updated_structure["takeover"]["contact_person"] is None
    assert "Adres" not in updated_structure["stations"]["STACJA ŁADOWANIA – DANE"], "Expected 'Model' to be removed from structure."

def test_compare_structure_with_wrong_value_in_mandatory_sections(excel_instance_with_sections, wrong_value_in_mandatory_section_template):
    """Tests if compare_structure_with_file raises an error when multiple matches are found for a key."""
    # Checks, if error is raised
    with pytest.raises(ValueError, match="Expected a dictionary or None, but got int"):
        excel_instance_with_sections.compare_structure_with_file(wrong_value_in_mandatory_section_template)

# 4. Test case where the function is run with an empty template.
def test_compare_structure_with_empty_template(excel_instance_with_sections, empty_template):
    """Tests if compare_structure_with_file correctly handles an empty template."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(empty_template)
    assert updated_structure == empty_template, "Expected empty template structure to remain unchanged."

# 5. Test verifying that no changes occur when the template matches the file exactly.
def test_compare_structure_no_changes(excel_instance_with_sections, standard_template):
    """Tests if compare_structure_with_file does not modify a template that already matches the file."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(standard_template)
    assert updated_structure == standard_template, "Expected no changes in structure when template already matches the file."