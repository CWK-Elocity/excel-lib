# Documentation for Handling Excel File to Create Data Structure

---

## Requirements for the Excel File:

1. **Sections Written in Uppercase in the First Column:**
   - Each section in the file must be clearly marked and written in uppercase in the first column.
   - Default sections (their names can be modified in the `env.py` file):
     - **`SECTION_STATION_TAKEOVER_DIVIDER`**:
       - This is a list of possible section names separating takeover data from individual charging station data.
       - Default values:
         ```python
         SECTION_STATION_TAKEOVER_DIVIDER = [
             "STACJA ŁADOWANIA - DANE",
             "STACJA ŁADOWANIA - PODSTAWOWE DANE"
         ]
         ```
     - **`SECTION_CONTACT_PERSON`**:
       - This is a list of possible section names containing contact details for station operation.
       - Default values:
         ```python
         SECTION_CONTACT_PERSON = [
             "OSOBA KONTAKTOWA - EKSPLOATACJA STACJI",
             "KONTAKT TECHNICZNY"
         ]
         ```
     - **`SECTION_RESPONSIBLE_PERSON`**:
       - This is a list of possible section names defining the person responsible for taking over the station on the client's side.
       - Default values:
         ```python
         SECTION_RESPONSIBLE_PERSON = [
             "OSOBA ODPOWIEDZIALNA ZA PRZEJĘCIE STACJI PO STRONIE KLIENTA",
             "MANAGER STACJI"
         ]
         ```

2. **Excel File Structure:**
   - **First Column:** Contains section headers or numbering.
   - **Second Column:** Data keys (e.g., "Name and Surname", "Phone Number").
   - **Third and Subsequent Columns:** Data related to individual stations.

---

## Assumptions and Constraints:

- **Sections in Uppercase:**
  - Values in the first column written in uppercase (`isupper()`) indicate the start of new sections.
- **Flexibility in Section Names:**
  - The names of sections are now stored as lists in the `env.py` file, allowing for multiple possible names for each section.
- **Data Consistency:**
  - This function **does not assume data consistency** between columns (e.g., for contacts or responsibilities).
  - If data is inconsistent, each column (station) can have its own data.

---
