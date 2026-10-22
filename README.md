# SBS Evaluation Automation

A Streamlit-based web application for Strathmore Business School to automate the generation of faculty evaluation reports, programme reports, and departmental reports.

## Features

- **Data Ingestion**: Upload Excel/CSV files or connect to public Google Sheets
- **Single Programme Mode**: Generate individual faculty reports and overall programme reports
- **Master Departmental Report Mode**: Aggregate all programmes into a comprehensive periodic departmental review
- **Automated Report Generation**: Creates Word documents with tables, charts, and formatted reports
- **ZIP Export**: Individual faculty reports are bundled into a downloadable ZIP file

## Requirements

- Python 3.8 or higher
- pip package manager

## Installation

1. Navigate to the project directory:
   ```bash
   cd C:\Users\pakaki\CascadeProjects\sbs_evaluation_automation
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit application:
   ```bash
   streamlit run app.py
   ```

2. The application will open in your default web browser (typically at http://localhost:8501)

3. **Data Source Selection**:
   - Choose "Upload Excel / CSV File" to upload your evaluation dataset
   - Or choose "Google Sheets Public Link" to connect to a public Google Sheet

4. **Report Generation**:
   - **Single Programme Mode**: Select a specific programme, enter participant numbers, and generate individual faculty reports and/or overall programme reports
   - **Master Departmental Report Mode**: Generate a comprehensive report aggregating all programmes in the dataset

## Data Format

The Excel/CSV file should contain a sheet named "Comments & Scores" with the following columns:
- Programme Name
- Session
- Faculty Name
- Category
- Punctuality
- Clarity of presentation of concepts, course material, content and use of practical examples.
- Degree of preparation for the lessons
- Course design and organisation
- Content and methodology
- Quality of discussions
- The information presented was relevant to my work
- Comments

## Output

Generated reports are saved in the `Generated_Reports_GUI` directory:
- Individual faculty reports (ZIP file)
- Overall programme reports (Word document)
- Master departmental reports (Word document)

## Optional Logo

Add your institution logo by placing one of the following files in the project directory:
- `logo.png`
- `input_file_0.png`
- `SBS_Logo.png`

## Troubleshooting

- If the application fails to load, ensure all dependencies are installed correctly
- For Google Sheets integration, ensure the sheet is set to "Anyone with the link can view"
- Make sure the Excel file has the correct sheet name ("Comments & Scores") and column headers
