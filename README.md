# Anexo II Report Generator

A simple desktop application to fill in data and automatically generate an "Anexo II" Word document from a template.

## Features

-   Easy-to-use graphical user interface for data input.
-   Dynamic forms that adjust the number of inputs for equipment and probes.
-   Auto-filling for some fields based on other inputs.
-   Ability to insert images from files or via a built-in screenshot tool.
-   Input validation for required fields.
-   Generates a ready-to-use `.docx` document.

## Prerequisites

Before running the application, ensure you have Python 3.x installed on your system.

## Installation

1.  **Clone this repository (or download as a ZIP):**
    ```bash
    git clone [YOUR_REPOSITORY_URL]
    cd [PROJECT_DIRECTORY_NAME]
    ```

2.  **Create and activate a virtual environment (recommended):**
    ```bash
    # For Windows
    python -m venv venv
    .\venv\Scripts\activate

    # For macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: You need to create the `requirements.txt` file first. See below)*

## How to Create `requirements.txt`

If you don't have one, you can create it with the following command after installing all necessary libraries (`PyQt5`, `python-docx`, `Pillow`):

```bash
pip freeze > requirements.txt
```

Your `requirements.txt` file will look something like this:

```
Pillow==10.1.0
PyQt5==5.15.10
python-docx==1.1.0
# add other library versions you use
```

## Usage

1.  **Prepare the Template:**
    Ensure the template file `New_Template2.docx` is located inside the `templates` folder. If the folder doesn't exist, the application will create it on the first run.

2.  **Run the Application:**
    ```bash
    python main_app.py
    ```

## Distribution (Creating a Standalone .exe)

You can package the application into a single executable file using PyInstaller. This allows users to run the application without installing Python or any dependencies.

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Build the Executable:**
    Run the following command from the project's root directory:
    ```bash
    pyinstaller --name "AnexoIIGenerator" --onefile --windowed --add-data "templates;templates" main_app.py
    ```

3.  **Find the Executable:**
    The generated `AnexoIIGenerator.exe` file will be located in the `dist` folder. You can distribute this file to your users.

### Running the .exe
Simply double-click the `AnexoIIGenerator.exe` file to run the application. Make sure the user has permissions to run executables in their system.

3.  **Fill in the Data:**
    Fill out all the fields in the application's interface.

4.  **Generate the Document:**
    Click the "GENERAR DOCUMENTO DE WORD (.docx)" button and choose a location to save the generated file.

## Project Structure

```
.
├── main_app.py             # Main script to run the application
├── ui_builder.py           # Module for building the user interface (GUI)
├── document_processor.py   # Module for processing and generating the Word document
├── fields.py               # Definitions for all input fields and placeholders
├── config.py               # Configuration for paths and template filename
├── utils.py                # Utility functions (e.g., validation)
├── screenshot.py           # Dialog for the screenshot feature
├── templates/
│   └── New_Template2.docx  # The Word template file
└── README.md               # This file
```
