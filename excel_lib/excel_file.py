import pandas as pd
import io
import zipfile

def file_to_io_stream(path):
    with open(path, "rb") as file:
        file_stream = io.BytesIO(file.read())
    return file_stream

def is_match(value1, value2):
    """Compares two values, stripping strings if applicable."""
    if isinstance(value1, str) and isinstance(value2, str):
        return value1.strip() == value2.strip()
    return value1 == value2

class ExcelFile:
    def __init__(self, file_stream, sections_config):
        """Initializes the ExcelFile class, validates the file, and checks for non-cell objects."""
        self.sections_config = sections_config
        self.worksheet_count = 0
        self.worksheet = None
        self.identified_sections = None
        self._validate_excel_file(file_stream)
        self.non_cell_objects = self._check_for_non_cell_objects(file_stream)
        self._load_first_worksheet(file_stream)  # Load only the first worksheet

    def _validate_excel_file(self, file_stream):
        """Validates if the file is a correct Excel (.xlsx) file."""
        file_stream.seek(0)  # Ensure stream starts at the beginning
        try:
            with zipfile.ZipFile(file_stream, 'r') as zip_ref:
                if "xl/workbook.xml" not in zip_ref.namelist():
                    raise ValueError("Invalid Excel file: missing xl/workbook.xml")
        except zipfile.BadZipFile:
            raise ValueError("Invalid Excel file: unable to open as ZIP archive")

    def _check_for_non_cell_objects(self, file_stream):
        """Extracts images and chart references from an Excel file."""
        non_cell_objects = []
        file_stream.seek(0)  # Ensure the stream is at the beginning
        with zipfile.ZipFile(file_stream, 'r') as zip_ref:
            # Check for media files (images)
            media_files = [f for f in zip_ref.namelist() if f.startswith("xl/media/")]
            for media_file in media_files:
                non_cell_objects.append(f"Image found: {media_file}")
            # Check for drawings
            drawing_files = [f for f in zip_ref.namelist() if f.startswith("xl/drawings/drawing")]
            for drawing_file in drawing_files:
                with zip_ref.open(drawing_file) as f:
                    content = f.read().decode("utf-8")
                    if "<xdr:twoCellAnchor>" in content:
                        non_cell_objects.append(f"Image anchored in {drawing_file}")
                    elif "<xdr:absoluteAnchor>" in content:
                        non_cell_objects.append(f"Image not anchored in {drawing_file}")
        return non_cell_objects

    def _load_first_worksheet(self, file_stream):
        """Loads only the first worksheet and warns if there are multiple sheets."""
        file_stream.seek(0)  # Reset stream position
        excel_file = pd.ExcelFile(file_stream)
        self.worksheet_count = len(excel_file.sheet_names)  # Get sheet count
        if self.worksheet_count > 1:
            print(f"Warning: The Excel file contains {self.worksheet_count} sheets. Only the first sheet will be used.")
        self.worksheet = pd.read_excel(file_stream, sheet_name=excel_file.sheet_names[0])

    def find_row_for_key(self, key, section_name=None):
        """
        Finds all row indices for a given key in the worksheet.
        Optionally filters matches by a specific section.
        Section filter only works when there is more tan one match.

        Args:
            key (str): The key to search for.
            section_name (str, optional): The name of the section to restrict the search.

        Returns:
            int or list: Single index if one match is found, list of indices if multiple matches are found,
                        or -1 if no match is found.

        Raises:
            ValueError: If multiple matches are found within the specified section.
        """
        # Find all matching indices
        matching_indices = [
            row_index for row_index, value in self.worksheet.iloc[:, 1].items()
            if pd.notna(value) and is_match(value, key)
        ]

        if not matching_indices:
            return -1  # No matches found

        if len(matching_indices) == 1:
            return matching_indices[0]  # Single match in the entire worksheet

        section_keys = {
            "global_data": self.sections_config.get("SECTION_STATION_TAKEOVER_DIVIDER", []),
            "contact_person": self.sections_config.get("SECTION_CONTACT_PERSON", []),
            "responsible_person": self.sections_config.get("SECTION_RESPONSIBLE_PERSON", [])
        }

        if section_name in section_keys:
            section_name = section_keys[section_name][0]

        if section_name:
            # Get section ranges
            if self.identified_sections is None:
                self._identify_sections()
            if self.identified_sections is None:
                raise ValueError(f"No sections have been identified.")
            if section_name not in self.identified_sections:
                return -1

            section_start, section_end = self.identified_sections[section_name]
            # Filter matches by section range
            section_matches = [index for index in matching_indices if section_start <= index <= section_end]

            if len(section_matches) == 1:
                return section_matches[0]  # Single match in the section
            elif len(section_matches) > 1:
                raise ValueError(f"Multiple matches found for key '{key}' in section '{section_name}': {section_matches}")
            return -1  # No matches in the specified section

        return matching_indices  # Return all matches if no section specified


    def _identify_sections(self):
        """Identify sections based on first column in workbook

        Returns:
            list: list of section names
        """
        sections = {}
        current_section = None
        for row_index, value in enumerate(self.worksheet.iloc[:, 0]):
            if isinstance(value, str) and value.isupper():
                if current_section:
                    sections[current_section][1] = row_index -1
                if isinstance(value, str) and isinstance(value, str):
                    current_section = value.strip()
                else:
                    current_section = value
                sections[current_section] = [row_index + 1, None]

        if current_section:
            sections[current_section][1] = self.worksheet.iloc[:, 0].last_valid_index()
        self.identified_sections = sections
        return sections
    
    def create_template_structure(self):
        """Creates a template structure based on the Excel file."""
        template_structure = {
            "takeover": {
                "global_data": None,
                "contact_person": None,
                "responsible_person": None
            },
            "stations": {}
        }

        # Identify sections
        sections = self._identify_sections()

        # Dynamically retrieve section keys from self.sections_config
        section_keys = {
            "contact_person": self.sections_config.get("SECTION_CONTACT_PERSON", []),
            "responsible_person": self.sections_config.get("SECTION_RESPONSIBLE_PERSON", [])
        }

        takeover_divider_key = self.sections_config.get("SECTION_STATION_TAKEOVER_DIVIDER", [])
        if takeover_divider_key:
            divider = next((name.strip() for name in takeover_divider_key if name.strip() in sections), None)
            if divider:
                global_data = {}
                for row_index, row in self.worksheet.iloc[:sections[divider][0]-1, :2].iterrows():
                    value, key = row
                    if pd.notna(key) and pd.notna(value):
                        global_data[key] = row_index
                template_structure["takeover"]["global_data"] = global_data

        # Populate takeover sections
        for key, section_names in section_keys.items():
            section_match = next((name for name in section_names if name in sections), None)
            if section_match:
                section_data = {}
                for row_index, row in self.worksheet.iloc[sections[section_match][0]:sections[section_match][1]+2, :2].iterrows():
                    value, key_name = row
                    if pd.notna(key_name) and pd.notna(value):
                        section_data[key_name] = row_index
                template_structure["takeover"][key] = section_data

        # Populate stations
        for section, section_range in sections.items():
            station_data = {}
            for row_index in range(section_range[0], section_range[1]+1):
                key = self.worksheet.iat[row_index, 1]
                if pd.notna(key):
                    if isinstance(key, str):
                        key = key.strip()
                    station_data[key] = row_index
            template_structure["stations"][section] = station_data

        return template_structure

    def compare_structure_with_file(self, template):
        """Compares actual working file with given template to obtain rows that will be used to determine value

        Args:
            template (dictionary): a dict containing template which user wants to use

        Returns:
            dict: similar to template but updated for that file
        """
        updated_structure = {
            "takeover": {
                "global_data": {},
                "contact_person": None,
                "responsible_person": None
            },
            "stations": {}
        }

        def _update_rows_in_structure(self, data_section, name_of_the_section):
            """Checks if the value is in the same row in file and in template.
            Otherwise looks for that specific value in all rows, and if found then updates row number.

            Args:
                data_section (dict or None): Section of whole data.

            Returns:
                dict or None: Updated section or None if input is None.

            Raises:
                InvalidDataSectionError: If the data_section is not a dict or None.
            """
            if data_section is None:
                return None
            if not isinstance(data_section, dict):
                raise ValueError(
                    f"Expected a dictionary or None, but got {type(data_section).__name__}."
                )

            updated_section = {}
            for key, expected_row in data_section.items():
                actual_label = (
                    self.worksheet.iloc[expected_row, 1] 
                    if expected_row < len(self.worksheet) else None
                )
                row_matches = self.find_row_for_key(key, name_of_the_section)

                if pd.notna(actual_label) and is_match(actual_label, key) and isinstance(row_matches, int):
                    updated_section[key] = row_matches
                else:
                    if isinstance(row_matches, list) and len(row_matches) > 1:
                        raise ValueError(f"Multiple matches found for key '{key}': {row_matches}")
                    updated_section[key] = row_matches if isinstance(row_matches, int) else -1

            return updated_section

        for section_name, takeover_section in template["takeover"].items():
            updated_structure["takeover"][section_name] = _update_rows_in_structure(self, takeover_section, section_name)

        for section_name, station_section in template["stations"].items():
            updated_structure["stations"][section_name] = _update_rows_in_structure(self, station_section, section_name)

        return updated_structure

    def create_data_structure_from_template(self, template):
        """Gather data from file based on template and structurize them in one dict object

        Args:
            template (dictionary): a dict containg template wich user wants to use

        Returns:
            dict: containg ale data categorised into sections
        """
        data_structure = self.compare_structure_with_file(template)

        collected_takeover_structures = []

        for column in range(2, self.worksheet.shape[1]):
            # Take global data from column
            global_data_section = data_structure["takeover"]["global_data"] or {}
            current_global_data = {
                key: self.worksheet.iloc[row, column] if row >= 0 and row < len(self.worksheet) else None 
                for key, row in global_data_section.items()
            }

            # Skip columns where all values are None
            if all(pd.isna(value) for value in current_global_data.values()):
                continue

            # Check whether there is a struct with this global data
            matching_group = None
            for group in collected_takeover_structures:
                if group["global_data"] == current_global_data:
                    matching_group = group
                    break

            # If there is none, create new one
            if not matching_group:
                matching_group = {
                    "global_data": current_global_data,
                    "contact_person": None,
                    "responsible_person": None,
                    "stations": []
                }
                collected_takeover_structures.append(matching_group)

            # Compare contact person - add fallback for None
            contact_person_section = data_structure["takeover"]["contact_person"] or {}
            current_contact_person = {
                key: self.worksheet.iloc[row, column] if row >= 0 and row < len(self.worksheet) else None 
                for key, row in contact_person_section.items()
            }
            if matching_group["contact_person"] is None:
                matching_group["contact_person"] = current_contact_person
            elif matching_group["contact_person"] != current_contact_person:
                matching_group["contact_person"] = "Dla każdej stacji inna"

            # Compare responsible person - add fallback for None
            responsible_person_section = data_structure["takeover"]["responsible_person"] or {}
            current_responsible_person = {
                key: self.worksheet.iloc[row, column] if row >= 0 and row < len(self.worksheet) else None 
                for key, row in responsible_person_section.items()
            }
            if matching_group["responsible_person"] is None:
                matching_group["responsible_person"] = current_responsible_person
            elif matching_group["responsible_person"] != current_responsible_person:
                matching_group["responsible_person"] = "Dla każdej stacji inna"

            # Add station data - ensure each section and fields are properly handled
            station_data = {}
            for section, fields in data_structure["stations"].items():
                if fields is not None:  # Skip sections with None fields
                    section_data = {
                        field: self.worksheet.iloc[row, column] if row >= 0 and row < len(self.worksheet) else None
                        for field, row in fields.items()
                    }
                    station_data[section] = section_data
            
            matching_group["stations"].append(station_data)

        return collected_takeover_structures
