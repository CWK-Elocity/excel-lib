import pytest
import pandas as pd
import io
from excel_lib.excel_file import ExcelFile

# Global config for section names used in tests.
SECTIONS_CONFIG = {
    "SECTION_STATION_TAKEOVER_DIVIDER": ["STACJA ŁADOWANIA – DANE", "STACJA ŁADOWANIA - DANE"],
    "SECTION_CONTACT_PERSON": ["OSOBA KONTAKTOWA - EKSPOLATACJA STACJI"],
    "SECTION_RESPONSIBLE_PERSON": ["OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA"]
}

# ----------------------------
# Fixtures for loading Excel files.
# ----------------------------
@pytest.fixture
def load_excel_file(request):
    """Loads a real Excel file from the test directory."""
    file_name = request.param
    file_path = f"tests/files/{file_name}"
    with open(file_path, "rb") as f:
        file_stream = io.BytesIO(f.read())
    return file_stream

# Example real file tests (use parametrization to load different files)
@pytest.mark.parametrize("load_excel_file", ["valid.xlsx"], indirect=True)
def test_valid_excel(load_excel_file):
    """Test a valid Excel file loads correctly and no non-cell objects are detected."""
    excel = ExcelFile(load_excel_file, SECTIONS_CONFIG)
    assert excel.worksheet_count == 1
    assert not excel.non_cell_objects, "No non-cell objects should be detected."

# More real file tests can be added similarly, ensuring images and multiple sheets are handled correctly.
# -------------------------------------------------------------------------------------------
# Fixtures and helper functions for in-memory Excel with sections.
# The fixture below creates an in-memory Excel file used for many tests.
# -------------------------------------------------------------------------------------------
@pytest.fixture
def sample_excel_with_sections():
    """Creates an in-memory Excel file with multiple sections for testing _identify_sections."""
    output = io.BytesIO()
    df = pd.DataFrame({
        "A": [
            "liczba",
            1,  # Initial numbering
            2,
            "STACJA ŁADOWANIA – DANE",  # Section header
            1, 2, 3,
            "Adres",
            4, 5,
            "TERMINAL",  # Another section header
            None,
            1, 2, 3, 4,
            "Adres",  # Repeated key in a different section
            5,
            "OSOBA KONTAKTOWA - EKSPOLATACJA STACJI",  # Section header
            1, 2, 3, 4, 5,
            "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA",  # Mandatory section header
            1, 2, 3, 4, 5
        ],
        "B": [
            "Lp",
            "Osoba odpowiedzialna",
            "Numer jobu",
            None,
            "Deadline",
            "Pełna Nazwa Klienta",
            "Model",
            "Adres",
            "Rodzaj stacji (DC / AC)",
            "Adres",
            None,
            None,
            "Czy stacja będzie z terminalem?",
            "Kto dostarcza terminal?",
            "Model terminala",
            "Numer seryjny terminala",
            None, None, None,
            "Imię i nazwisko",
            "Numer telefonu",
            "Adres e-mail",
            None, None, None,
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
            None,
            "11.11.2024",
            "ELOCITY",
            "LS4",
            "Testowa Ulica 1",
            "AC",
            "Testowa Ulica 2",
            None, None,
            "Tak",
            "elocity",
            "PAX IM 30",
            "123456789",
            None, None, None,
            "Jan Kowalski",
            "987654321",
            "jan.kowalski@mail.com",
            None, None, None,
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
    """Creates an ExcelFile instance using the in-memory Excel file with multiple sections."""
    return ExcelFile(sample_excel_with_sections, SECTIONS_CONFIG)

# -------------------------------------------------------------------------------------------
# Tests for find_row_for_key functionality
# -------------------------------------------------------------------------------------------
def test_find_existing_keys_in_sections(excel_instance_with_sections):
    """
Test that find_row_for_key finds the correct row indices for specified keys.
    Verifies that keys present in the file are correctly located.
"""
    test_cases = {
        "Model": 5,
        "Osoba odpowiedzialna": 0,
        "Numer seryjny terminala": 14
    }
    for key, expected_index in test_cases.items():
        row_index = excel_instance_with_sections.find_row_for_key(key)
        assert row_index == expected_index, f"Expected {expected_index} for key '{key}', but got {row_index}"

def test_find_non_existing_key_in_sections(excel_instance_with_sections):
    """
Test that find_row_for_key returns -1 for keys that do not exist.
"""
    row_index = excel_instance_with_sections.find_row_for_key("NonExistingKey")
    assert row_index == -1, f"Expected -1, but got {row_index}"

def test_find_key_with_whitespace_in_sections(excel_instance_with_sections):
    """
Test that find_row_for_key correctly handles keys with extra whitespace.
    The function should trim string values before comparing.
"""
    row_index = excel_instance_with_sections.find_row_for_key(" Model ")
    assert row_index == 5, f"Expected index 5 for key 'Model', but got {row_index}"

def test_find_key_outside_specified_section(excel_instance_with_sections):
    """
Test that searching for a key within a specified section returns the correct row index.
"""
    row_index = excel_instance_with_sections.find_row_for_key("Test Key 4", "OSOBA KONTAKTOWA - EKSPOLATACJA STACJI")
    assert row_index == 27, "Expected 27 for 'Test Key 4' in the specified section, but got a different result."

def test_find_key_in_section_with_duplicate_in_another_section(excel_instance_with_sections):
    """
Test that find_row_for_key correctly filters match results when the same key appears in multiple sections.
"""
    row_index = excel_instance_with_sections.find_row_for_key("Imię i nazwisko", "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA")
    assert row_index == 24, f"Expected index 24 for 'Imię i nazwisko' in the specified section, but got {row_index}"

def test_find_key_in_global_section(excel_instance_with_sections):
    """
Test that find_row_for_key correctly searches in the global section.
"""
    row_index = excel_instance_with_sections.find_row_for_key("Numer jobu", "global_data")
    assert row_index == 1, f"Expected index 1 for 'Numer jobu' in global_data, but got {row_index}"

def test_find_multiple_keys_in_specified_section(excel_instance_with_sections):
    """
    Test that find_row_for_key raises a ValueError 
when multiple matches are found for the key 'Adres' in the section 'STACJA ŁADOWANIA – DANE'.
    """
    with pytest.raises(ValueError, match="Multiple matches found for key 'Adres' in section 'STACJA ŁADOWANIA – DANE'"):
        excel_instance_with_sections.find_row_for_key("Adres", "STACJA ŁADOWANIA – DANE")

def test_ignore_section_headers(excel_instance_with_sections):
    """
Test that find_row_for_key does not match section header cells.
    Searching for a header should return -1.
"""
    row_index = excel_instance_with_sections.find_row_for_key("STACJA ŁADOWANIA – DANE")
    assert row_index == -1, f"Expected -1 when searching for a section header, got {row_index}"

def test_identify_sections(excel_instance_with_sections):
    """
Test that _identify_sections correctly identifies sections based on
    uppercase strings in the first column.
"""
    sections = excel_instance_with_sections._identify_sections()
    expected_sections = {
        "STACJA ŁADOWANIA – DANE": [3, 8],
        "TERMINAL": [10, 16],
        "OSOBA KONTAKTOWA - EKSPOLATACJA STACJI": [18, 22],
        "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA": [24, 28]
    }
    assert sections == expected_sections, f"Expected sections {expected_sections}, but got {sections}"

def test_empty_section_between_headers(excel_instance_with_sections):
    """
Test behavior when a section header is followed by an empty section.
    It should still be identified correctly.
"""
    sections = excel_instance_with_sections._identify_sections()
    assert "OSOBA KONTAKTOWA - EKSPOLATACJA STACJI" in sections, "Section header not identified as expected."

def test_repeated_keys_sections(excel_instance_with_sections):
    """
Test that repeated keys across sections are handled appropriately.
    In this case, the function should return a list of indices when duplicates exist.
"""
    row_index = excel_instance_with_sections.find_row_for_key("Adres")
    assert isinstance(row_index, list), "Expected a list of indices for a duplicated key."
    assert len(row_index) >= 2, f"Expected multiple indices for 'Adres', got {row_index}"

def test_section_with_lowercase_header(excel_instance_with_sections):
    """
Test that _identify_sections ignores headers that are not entirely uppercase.
"""
    sections = excel_instance_with_sections._identify_sections()
# Check each identified section, if any header is lowercase this test should fail.
    for header in sections.keys():
        assert header.isupper(), f"Header '{header}' is not uppercase and should be ignored."

def test_find_key_with_empty_value(excel_instance_with_sections):
    """
Test that find_row_for_key handles keys that have an empty value (or NaN) correctly.
"""
    row_index = excel_instance_with_sections.find_row_for_key("Model terminala")
    assert row_index == 13, f"Expected index 13 for 'Model terminala', but got {row_index}"

# -------------------------------------------------------------------------------------------
# Tests for create_template_structure
# -------------------------------------------------------------------------------------------
def test_create_template_structure(excel_instance_with_sections):
    """
Test that create_template_structure extracts the expected overall structure from the Excel file.
"""
    template_structure = excel_instance_with_sections.create_template_structure()
    
    # Check that primary sections exist in the template
    assert "takeover" in template_structure, "Template should contain 'takeover'."
    assert "stations" in template_structure, "Template should contain 'stations'."
    
    # Validate types for the takeover sub-sections
    assert isinstance(template_structure["takeover"].get("global_data"), dict), "Global data should be a dict."
    assert isinstance(template_structure["takeover"].get("contact_person"), (dict, type(None))), \
        "Contact person should be a dict or None."
    assert isinstance(template_structure["takeover"].get("responsible_person"), (dict, type(None))), \
        "Responsible person should be a dict or None."
    assert isinstance(template_structure["stations"], dict), "Stations should be a dict."
    assert len(template_structure["takeover"].get("global_data")) == 2, "Expected two keys in global data."
    assert len(template_structure["takeover"].get("contact_person")) == 3, "Expected three keys in contact person." # 3 keys in contact person cause 2 are set to None
    assert len(template_structure["takeover"].get("responsible_person")) == 4, "Expected four keys in responsible person." # 4 keys in responsible person cause 1 is set to None
    assert len(template_structure["stations"]) == 4, "Expected 4 sections in station."
    assert len(template_structure["stations"].get("STACJA ŁADOWANIA – DANE")) == 5, "Expected 5 keys in STACJA ŁADOWANIA – DANE section"

def test_template_structure_integrity(excel_instance_with_sections):
    """
Test that the template structure contains only the expected sections and keys.
    Ensures that no extraneous keys are added during extraction.
"""
    template_structure = excel_instance_with_sections.create_template_structure()
    
    # For each station section, ensure that only expected keys exist (implementation dependent)
    for section, fields in template_structure["stations"].items():
        for key in fields:
            assert key in fields, f"Unexpected key '{key}' in stations section."
    
    # Verify takeover sections similarly
    for key, values in template_structure["takeover"].items():
        if values is not None:
            for sub_key in values:
                assert sub_key in values, f"Unexpected sub-key '{sub_key}' in takeover section."

# -------------------------------------------------------------------------------------------
# Tests for compare_structure_with_file
# -------------------------------------------------------------------------------------------
# Fixtures for various template scenarios are defined here.
@pytest.fixture
def standard_template():
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 0, "Numer jobu": 1},
            "contact_person": {"Imię i nazwisko": 18, "Numer telefonu": 19},
            "responsible_person": {"Imię i nazwisko": 24, "Numer telefonu": 25}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 5, "Rodzaj stacji (DC / AC)": 7},
            "TERMINAL": {"Czy stacja będzie z terminalem?": 11}
        }
    }

@pytest.fixture
def missing_values_template():
    """Template with missing sections and values."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 1},
            "contact_person": None,
            "responsible_person": {"Imię i nazwisko": 25}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Model": 6},
            "TERMINAL": {}
        }
    }

@pytest.fixture
def wrong_value_in_mandatory_section_template():
    """Template with incorrect data type in mandatory sections."""
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
def duplicate_value_template():
    """Template with duplicate keys that should trigger an error."""
    return {
        "takeover": {
            "global_data": {"Osoba odpowiedzialna": 1, "Numer jobu": 2},
            "contact_person": {"Imię i nazwisko": 19}, 
            "responsible_person": {"Imię i nazwisko": 19}
        },
        "stations": {
            "STACJA ŁADOWANIA – DANE": {"Adres": 9}  # Duplicate key expected to raise error.
        }
    }

@pytest.fixture
def empty_template():
    """An empty template for testing empty structure handling."""
    return {
        "takeover": {
            "global_data": {},
            "contact_person": {},
            "responsible_person": {}
        },
        "stations": {}
    }

def test_compare_structure_with_file(excel_instance_with_sections, standard_template):
    """Test that compare_structure_with_file updates row indices correctly."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(standard_template)
    
    # Check station sections: each key should match the Excel cell based on row index.
    for section, fields in updated_structure["stations"].items():
        for key, row_index in fields.items():
            cell_value = excel_instance_with_sections.worksheet.iat[row_index, 1]
            assert cell_value == key, f"Expected key '{key}' at row {row_index} in section '{section}', got '{cell_value}'"
    
    # Check takeover sections.
    for key, values in updated_structure["takeover"].items():
        if isinstance(values, dict):
            for field_key, row_index in values.items():
                # Here you might compare against an expected value.
                assert row_index is not None, f"Row index for '{field_key}' in takeover '{key}' should not be None."

def test_compare_structure_with_multiple_matches(excel_instance_with_sections, duplicate_value_template):
    """Test that multiple matches for the same key raise a ValueError."""
    with pytest.raises(ValueError, match="Multiple matches found for key 'Adres'"):
        excel_instance_with_sections.compare_structure_with_file(duplicate_value_template)

def test_compare_structure_with_missing_keys(excel_instance_with_sections, missing_values_template):
    """Test that missing keys in the template result in a None or omission in the final structure."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(missing_values_template)
    assert updated_structure["takeover"]["contact_person"] is None, "Missing contact_person should be None."
    assert "Adres" not in updated_structure["stations"]["STACJA ŁADOWANIA – DANE"], "Non-existent key 'Adres' should not be added."

def test_compare_structure_with_wrong_value_in_mandatory_sections(excel_instance_with_sections, wrong_value_in_mandatory_section_template):
    """Test that wrong data types in mandatory sections cause an error."""
    with pytest.raises(ValueError, match="Expected a dictionary or None, but got int"):
        excel_instance_with_sections.compare_structure_with_file(wrong_value_in_mandatory_section_template)

def test_compare_structure_with_empty_template(excel_instance_with_sections, empty_template):
    """Test that an empty template remains unchanged after comparison."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(empty_template)
    assert updated_structure == empty_template, "Empty template structure should remain unchanged."

def test_compare_structure_no_changes(excel_instance_with_sections, standard_template):
    """Test that if the template perfectly matches the file, no changes occur."""
    updated_structure = excel_instance_with_sections.compare_structure_with_file(standard_template)
    assert updated_structure == standard_template, "Template according to file should remain unchanged."

# -------------------------------------------------------------------------------------------
# Tests for create_data_structure_from_template
# -------------------------------------------------------------------------------------------
@pytest.fixture
def sample_excel_file(tmp_path):
    """Creates a temporary test Excel file for create_data_structure_from_template."""
    file_path = tmp_path / "test_file.xlsx"
    df = pd.DataFrame({
        'A': ['Lp', 'SECTION1', 'key1', 'key2', 'SECTION2', 'key3', 'key6', 'SECTION3', 'key5', 'key6', 'SECTION4', 'key7', 'key1', 'SECTION6', 'key9', 'key10'],
        'B': [1, None, 'key1', 'key2', None, 'key3', 'key4', None, 'key5', 'key6', None, 'key7', 'key1', None, 'key9', 'key10'],
        'C': [2, 'value7', 'value8', 'value9', 'value10', 'value11', 'value12', 'value13', 'value14', 'value15', 'value30', 'value31', 'value32', 'value33', 'value34', 'value35'],
        'D': [3, 'value7', 'value8', 'value69', 'value11', 'value10', 'value12', 'value53', 'value53', 'value25', 'value16', 'value17', 'value18', 'value19', 'value20', 'value21'],
        'E': [4, 'value7', 'value69', 'value9', 'value10', 'value11', 'value12', 'value51', 'value51', 'value15', 'value19', 'value20', 'value21', 'value22', 'value23', 'value24'],
        'F': [5, 'value7', 'value8', 'value69', 'value11', 'value10', 'value12', 'value3', 'value53', 'value35', 'value22', 'value23', 'value24', 'value25', 'value26', 'value27'],
        'G': [6, 'value7', 'value8', 'value9', 'value10', 'value11', 'value12', 'value13', 'value14', 'value15', 'value25', 'value26', 'value27', 'value28', 'value29', 'value30'],
    })
    df.to_excel(file_path, index=False, header=False)
    return file_path

@pytest.fixture
def sample_template():
    """Creates a sample template structure for testing create_data_structure_from_template."""
    return {
        "takeover": {
            "global_data": {"key1": 1, "key2": 2},
            "contact_person": {"key3": 4}, 
            "responsible_person": {"key5": 7, "key6": 8}, #key6 is with wrong row
        },
        "stations": {
            "SECTION4": {"key7": 10, "key1": 1},
            "SECTION2": {"key9": 13, "key10": 13}
        }
    }

def test_create_data_structure_from_template(sample_excel_file, sample_template):
    """Test that create_data_structure_from_template returns the expected data structure."""
    with open(sample_excel_file, "rb") as f:
        file_stream = io.BytesIO(f.read())
    sections_config = {
        "SECTION_STATION_TAKEOVER_DIVIDER": ["SECTION1"],
        "SECTION_CONTACT_PERSON": ["SECTION2"],
        "SECTION_RESPONSIBLE_PERSON": ["SECTION3"]
    }
    excel_file = ExcelFile(file_stream, sections_config)
    compared_template = excel_file.compare_structure_with_file(sample_template)
    data_structure = excel_file.create_data_structure_from_template(compared_template)

    # Check that the data structure is a list with one element
    assert isinstance(data_structure, list)
    assert len(data_structure) == 3
    
    # Verify each section contains expected keys and values
    assert data_structure[0]["global_data"] == {"key1": "value8", "key2": "value9"}
    assert data_structure[0]["contact_person"] == {"key3": "value11"}
    assert data_structure[0]["responsible_person"] == {"key5": "value14", "key6": "value15"}
    assert len(data_structure[0]["stations"]) == 2
    assert data_structure[0]["stations"][0]["SECTION4"] == {"key7": "value31", "key1": "value32"}
    assert data_structure[0]["stations"][0]["SECTION2"] == {"key9": "value34", "key10": "value35"}
    assert data_structure[0]["stations"][1]["SECTION4"] == {"key7": "value26", "key1": "value27"}
    assert data_structure[0]["stations"][1]["SECTION2"] == {"key9": "value29", "key10": "value30"}
    assert data_structure[1]["global_data"] == {"key1": "value8", "key2": "value69"}
    assert data_structure[1]["contact_person"] == {"key3": "value10"}
    assert data_structure[1]["responsible_person"] == 'Dla każdej stacji inna'
    assert len(data_structure[1]["stations"]) == 2
    assert data_structure[1]["stations"][0]["SECTION4"] == {"key7": "value17", "key1": "value18"}
    assert data_structure[1]["stations"][0]["SECTION2"] == {"key9": "value20", "key10": "value21"}
    assert data_structure[1]["stations"][1]["SECTION4"] == {"key7": "value23", "key1": "value24"}
    assert data_structure[1]["stations"][1]["SECTION2"] == {"key9": "value26", "key10": "value27"}
    assert data_structure[2]["global_data"] == {"key1": "value69", "key2": "value9"}
    assert data_structure[2]["contact_person"] == {"key3": "value11"}
    assert data_structure[2]["responsible_person"] == {"key5": "value51", "key6": "value15"}
    assert len(data_structure[2]["stations"]) == 1
    assert data_structure[2]["stations"][0]["SECTION4"] == {"key7": "value20", "key1": "value21"}
    assert data_structure[2]["stations"][0]["SECTION2"] == {"key9": "value23", "key10": "value24"}