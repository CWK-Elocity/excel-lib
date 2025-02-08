# Excel Processing Library

This library provides tools to process Excel files in memory, validate their structure, and extract specific information such as images, charts, and sectioned data.

## Features

- **Validation**: Ensures the file is a valid Excel `.xlsx` file using `zipfile`.
- **Image and Chart Detection**: Extracts references to images and charts, including their placement information.
- **Sectioned Data Extraction**: Handles Excel files with section headers, merged cells, and auxiliary numbering.
- **Single Worksheet Processing**: Reads only the first worksheet and warns if multiple sheets are present.

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

To run tests, execute the following:

```bash
pytest tests/
```

### Key Updates to Testing

- **Zero-Based Indexing**: Test cases now account for `pandas` indexing starting at `0` while Excel rows start at `1`. For example, the key "Model" located in Excel row `6` will correspond to index `4` in `pandas`.
- **Fixtures for Realistic Data**: Tests simulate real-world Excel structures with auxiliary numbering (`Lp`), section headers, and merged cells.
- **Improved Image Detection**: Image detection no longer relies on `openpyxl` and instead uses `zipfile` to directly parse Excel archives.

## Usage


