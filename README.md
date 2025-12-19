# Layer Audit Automation System

A Flask-based web application for automating layer audits, image capture, and report generation using Python, SQL Server, and Microsoft Excel automation.

## Features
- **Audit Data Entry**: Web interface for inputting audit details (Area, Line No, Die No, etc.).
- **Image Capture**: Integrated camera support to capture audit evidence directly.
- **Data Persistence**: Stores audit records in a SQL Server database.
- **Automated Reporting**: Generates Excel reports with embedded images and exports them as PDFs using `win32com`.

## Prerequisites

This application relies on Windows-specific technologies and local software installations:
- **Operating System**: Windows 10/11 (Required for `pywin32` COM automation).
- **Python**: 3.10 or higher.
- **Microsoft Excel**: Must be installed and activated (used for report generation).
- **SQL Server**: A running instance (e.g., SQL Server Express) with the schema configured.
- **ODBC Driver**: "ODBC Driver 17 for SQL Server" must be installed.

## Installation

1.  **Clone the repository**:
    ```bash
    git clone <repository-url>
    cd Audit-System
    ```

2.  **Create a Virtual Environment**:
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: If `requirements.txt` is missing, install: `flask pyodbc pywin32 pillow`)*

4.  **Configure the Application**:
    - Open `example.ini`.
    - Update the `[database_details]` section with your SQL Server credentials.
    - Update the `[path]` section with valid absolute paths on your machine.
    
    ```ini
    [path]
    uploadPath = C:\Path\To\Your\Uploads
    excelSave = C:\Path\To\Your\SaveDir
    excelPath = C:\Path\To\Source\ExcelTemplate
    ```

## How to Run

1.  **Start the Server**:
    ```bash
    python main.py
    ```

2.  **Access the App**:
    - Open your browser and navigate to `http://127.0.0.1:5080`.

## Project Structure

- `main.py`: Core Flask application logic and routing.
- `static/`: CSS styles, images, and generated assets.
- `templates/`: HTML templates for the UI.
- `example.ini`: Configuration file for DB and file paths.

## Troubleshooting

- **Database Error**: Ensure the Connection String in `main.py` matches your ODBC driver version.
- **Excel Error**: Ensure Excel is not open in a blocking mode or showing a popup dialog when the script runs.
