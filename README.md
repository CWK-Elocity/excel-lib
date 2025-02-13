# Excel Processing Library

This library provides tools to process Excel files in memory, validate their structure, and extract specific information such as images, charts, and sectioned data.

## Features

- **Validation**: Ensures the file is a valid Excel `.xlsx` file using `zipfile`.
- **Image and Chart Detection**: Extracts references to images and charts, including their placement information.
- **Sectioned Data Extraction**: Handles Excel files with section headers, merged cells, and auxiliary numbering.
- **Single Worksheet Processing**: Reads only the first worksheet and warns if multiple sheets are present.
- **Error Handling for Duplicate Keys**: If multiple matches for the same key are found during structure comparison, a `ValueError` is raised, ensuring data consistency and validity.

## Installation

To install the library locally, use:

```bash
pip install .
```

If you experience issues where the library is not recognized during testing, ensure it is reinstalled after every update.

## Tests

The library includes comprehensive tests for its functionality. Key test cases include:

1. **Validation Tests**: Confirm the library properly validates `.xlsx` files.
2. **Image Detection Tests**: Verify that images and charts are correctly detected in various scenarios:
   - Single image in a cell
   - Image outside a cell
   - Multiple images in different cells and locations
   - Updated image detection using `zipfile` to parse `xl/media/` and `xl/drawings/` files for better accuracy.
3. **Section Handling Tests**: Ensure section headers and merged cells are processed correctly.
4. **Row Index Detection**:
   - Tests for finding specific keys in the second column of the Excel file.
   - Handles extra whitespace around keys.
   - Ignores section headers during key search.
5. **New Tests for Class `ExcelFile`**:
   - Tests for `find_row_for_key` to ensure proper handling of auxiliary numbering (`Lp`) and section headers.
   - Updated to reflect zero-based indexing in pandas.
   - Extended handling of duplicate keys by returning all occurrences.
   - Raises `ValueError` when multiple matches are found for a key during structure comparison.
6. **Dynamic Section Configuration**:
   - The function `create_template_structure` now dynamically retrieves section keys from `section_config` instead of using hardcoded values.
   - This makes the library more flexible and adaptable to different Excel formats.

To run tests, execute the following:

```bash
pytest tests/
```

### Key Updates to Testing

- **Zero-Based Indexing**: Test cases now account for `pandas` indexing starting at `0` while Excel rows start at `1`. For example, the key "Model" located in Excel row `6` will correspond to index `4` in `pandas`.
- **Fixtures for Realistic Data**: Tests simulate real-world Excel structures with auxiliary numbering (`Lp`), section headers, and merged cells.
- **Improved Image Detection**: Image detection no longer relies on `openpyxl` and instead uses `zipfile` to directly parse Excel archives.
- **Dynamic Section Configuration**: `section_config` must contain keys for defining section mappings:
  ```python
  sections_config = {
      "SECTION_STATION_TAKEOVER_DIVIDER": ["STACJA ŁADOWANIA – DANE"],
      "SECTION_CONTACT_PERSON": ["OSOBA KONTAKTOWA - EKSPLOATACJA STACJI"],
      "SECTION_RESPONSIBLE_PERSON": ["OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA"]
  }
  ```

## Usage

To properly process an Excel file using this library, follow these steps:

1. Open your Excel file as a binary stream.
2. Provide the desired section configuration via a dictionary. For example:

   ```python
   sections_config = {
       "SECTION_STATION_TAKEOVER_DIVIDER": ["STACJA ŁADOWANIA – DANE"],
       "SECTION_CONTACT_PERSON": ["OSOBA KONTAKTOWA - EKSPLOATACJA STACJI"],
       "SECTION_RESPONSIBLE_PERSON": ["OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA"]
  }
   ```

3. Create an instance of `ExcelFile` using the file stream and configuration.
4. Generate a template based on the file's identified sections with `create_template_structure()`.
5. Update the template structure to map actual row indices via `compare_structure_with_file()`.
6. Finally, extract the data in a structured format using `create_data_structure_from_template()`.

Below is an example:

```python
from excel_lib.excel_file import ExcelFile

sections_config = {
    "SECTION_STATION_TAKEOVER_DIVIDER": ["STACJA ŁADOWANIA – DANE"],
    "SECTION_CONTACT_PERSON": ["OSOBA KONTAKTOWA - EKSPLOATACJA STACJI"],
    "SECTION_RESPONSIBLE_PERSON": ["OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA"]
}

# Open the Excel file as a binary stream
with open("example.xlsx", "rb") as file_stream:
    # Initialize ExcelFile with the file stream and section configuration
    excel = ExcelFile(file_stream, sections_config)

# Generate the initial template structure based on the file's sections
template = excel.create_template_structure()
print("Template Structure:", template)

# Update the template structure by comparing it with the actual file content
updated_template = excel.compare_structure_with_file(template)
print("Updated Template Structure:", updated_template)

# Extract data using the updated template structure
data_structure = excel.create_data_structure_from_template(updated_template)
print("Extracted Data Structure:", data_structure)
```

