import zipfile

def inspect_excel_structure(file_path):
    """Inspect the internal structure of an Excel (.xlsx) file."""
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        file_list = zip_ref.namelist()
        print("\n".join(file_list))  # Display all internal files

# Podmień 'your_file.xlsx' na ścieżkę do testowego pliku
inspect_excel_structure("tests/files/image_outside_cell.xlsx")