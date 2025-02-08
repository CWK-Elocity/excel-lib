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

# Tests for find_row_for_key

@pytest.fixture
def sample_excel_with_sections():
    """Creates an in-memory Excel file with sections for testing find_row_for_key."""
    output = io.BytesIO()
    df = pd.DataFrame({
        "A": [1, 2, 3, 4, 5, 6, 7, 8],  # Lp (auxiliary numbering)
        "B": [
            "STACJA ŁADOWANIA – DANE",  # Section header
            "Osoba odpowiedzialna",
            "Numer jobu",
            "Deadline",
            "Pełna Nazwa Klienta",
            "Model",
            "Numer seryjny",
            "Rodzaj stacji (DC / AC)"
        ],
        "C": [
            None,  # Merged cells for section header
            "Adam Nijaki",
            "TEST",
            "11.11.2024",
            "ELOCITY",
            "LS4",
            "M4756023-3",
            "AC"
        ]
    })
    df.to_excel(output, index=False, header=False)
    output.seek(0)
    return output

@pytest.fixture
def excel_instance_with_sections(sample_excel_with_sections):
    """Creates an instance of ExcelFile for testing data with sections."""
    sections_config = {}
    return ExcelFile(sample_excel_with_sections, sections_config)

def test_find_existing_key_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key finds the correct index for an existing key."""
    row_index = excel_instance_with_sections.find_row_for_key("Model")
    assert row_index == 4, f"Expected index 4 (Excel row 6), but got {row_index}"

def test_find_non_existing_key_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key returns -1 for a non-existing key."""
    row_index = excel_instance_with_sections.find_row_for_key("NonExistingKey")
    assert row_index == -1, f"Expected -1, but got {row_index}"

def test_find_key_with_whitespace_in_sections(excel_instance_with_sections):
    """Tests if find_row_for_key handles keys with additional spaces correctly."""
    row_index = excel_instance_with_sections.find_row_for_key(" Model ")  # Extra spaces
    assert row_index == 4, f"Expected index 4 (Excel row 6), but got {row_index}"

def test_ignore_section_headers(excel_instance_with_sections):
    """Tests if find_row_for_key ignores section headers when searching for keys."""
    row_index = excel_instance_with_sections.find_row_for_key("STACJA ŁADOWANIA – DANE")
    assert row_index == -1, f"Expected -1 for section header, but got {row_index}"

