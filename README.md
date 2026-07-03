# SmartDocCreator-Edu

# School Notification Automation System

This project is a Python-based automation tool designed to streamline the administrative process of creating and organizing official school assignment notifications (e.g., teaching hours, placements, and reassignments) for teachers in a school district.

## Features

- **Document Generation**: Automatically generates personalized `.docx` documents based on Excel templates for various categories of teachers (e.g., newly appointed, substitute, or teachers with excess hours).
- **PDF Conversion**: Automatically converts generated documents to PDF format.
- **Placement Summary**: Creates a consolidated `Τοποθέτηση.docx` document that lists all assignments and placements in a table format, grouped by status (Entry, Modification, etc.).
- **Smart Organization**: Automatically sorts generated files into specific folders named after the relevant school, reading metadata (keywords) directly from the generated Word documents.
- **Batch Processing**: Handles Excel files with multiple sheets to process thousands of entries efficiently.

## Project Structure

- `main.py`: The entry point. Uses a GUI file picker to select the input Excel file and runs the full pipeline.
- `notification3.py`: Handles the generation of individual teacher notifications (`.docx` + `.pdf`).
- `placement.py`: Generates the summary document showing all placements in tables.
- `sort.py`: Scans the `notifications/` directory and organizes the files into subfolders based on school names extracted from document properties.
- `shared.py`: Contains common configurations, constants, and helper functions.

## Prerequisites

- Python 3.7+
- Libraries: `pandas`, `python-docx`, `openpyxl` (for Excel processing), and a PDF conversion engine (as configured in your `shared.py`).

## Installation

1. Clone this repository or copy the files to your local machine.
2. Install the required dependencies:
   ```bash
   pip install pandas python-docx openpyxl
   ```
3. Ensure your directory contains:
   - `TEMPLATE_ΚΟΙΝΟΠΟΙΗΤΗΡΙΑ.docx` (Template for individual notifications)
   - `TEMPLATE_ΤΟΠΟΘΕΤΗΣΕΙΣ.docx` (Template for the summary document)
   - A properly formatted Excel file with columns matching your `NOTIFICATION_COLUMNS` and `PLACEMENT_COLUMNS` definitions.

## Usage

1. Run the main script:
   ```bash
   python main.py
   ```
2. A file dialog will appear; select your data Excel file.
3. The system will:
   - Generate individual documents in the `notifications/` folder.
   - Create a summary `Τοποθέτηση.docx` file.
   - Organize all files into the `my_school_pdfs/` directory, structured by school name.

## Configuration

- **Excel Structure**: Ensure your Excel sheets have headers corresponding to the logic defined in `notification3.py` and `placement.py`.
- **Customization**: Update `shared.py` if column names, paths, or document templates change.
