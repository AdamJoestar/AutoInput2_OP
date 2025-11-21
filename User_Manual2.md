# User Manual - Anexo II Generator

## Chapter 1: Introduction

### What is this application?
The Anexo II Generator application is a PyQt5-based tool designed to help users create Anexo II documents automatically. This application allows you to fill out forms with the required data, then generates a Word document (.docx) based on the provided template.

### Who is this application for?
This application is intended for users who need to create reports or Anexo II documents, such as in the context of product testing or technical documentation. It is suitable for new users who are not familiar with manual document creation processes.

## Chapter 2: Installation

This application is provided as a single executable file (.exe) and does not require a formal installation process.

1. **Download the file:** Obtain the `AnexoIIGenerator.exe` file.
2. **Run the application:** Simply double-click the `AnexoIIGenerator.exe` file to start the program. No further installation is needed.

**System Requirements:**
- Windows 10 or Windows 11.
- No additional software required.

## Chapter 3: Getting Started

After launching the application, you will see the main application window. Below is an explanation of the main parts:

- **Logo:** Company logo image at the top.
- **Title:** "Ingresar Datos Para el Anexo" indicating the application's purpose.
- **Spin Boxes for Row Selection:**
  - "EQUIPOS Y MÉTODOS UTILIZADOS (max 12)": Select the number of equipment (1-12).
  - "SONDA TOTAL (max 10)": Select the number of probes (1-10).
  - "Número de fotos adicionales (max 10)": Select the number of additional photos (0-10).
- **Date and Time for Excel Data Filtering:** Select start and end date range to load data from Excel files.
- **Scrollable Form Area:** Where you fill in all the required data.
- **"CARGAR DATOS DE EXCEL" Button:** To load data from Excel based on the selected date range.
- **"GENERAR DOCUMENTO DE WORD (.docx)" Button:** To generate the final document.
- **Template Information:** Displays the template name and folder location.

[Screenshot: Main application window]

## Chapter 4: Filling the Form

The form is divided into several sections according to the Anexo II document structure. Each section has input fields that need to be filled. Below is an explanation of how to fill each type of field:

### Form Sections:
1. **Encabezado - Información del documento:** Basic document information such as test number, revision, and emission date.
2. **0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO:** Information about the test requester.
3. **1. INFORMACIÓN GENERAL DEL PRODUCTO:** General product information.
4. **1.1. CONDICIONES DEL ENSAYO:** Test conditions.
5. **2. EQUIPOS Y MÉTODOS UTILIZADOS:** List of equipment and methods used (number according to spin box).
6. **3. TEMPERATURAS REGISTRADAS:** Recorded temperatures (number according to probe spin box).
7. **3.1. GRÁFICA GENERADA:** Generated graphs.
8. **4. ESTABILIZACIÓN TÉRMICA:** Thermal stabilization.
9. **5. RESULTADOS:** Final results.
10. **6. FOTOGRAFIAS:** Documentation photos.
11. **7. FOTO MONTAJE FINAL:** Final montage photos (including additional photos).

### How to Fill Text Fields:
- **Simple Text Fields (QLineEdit):** Click on the field and type text directly. For example, for names or numbers.
- **Long Text Fields (QTextEdit):** Click and type longer text, such as descriptions or notes.

### How to Select a Date:
- Click on the date field (QDateEdit) to open the calendar popup.
- Select the date by clicking on the desired date in the calendar.

### How to Select Options from Dropdown:
- Click on the dropdown field (QComboBox) to see the list of options.
- Scroll and select the appropriate option, such as equipment type or temperature unit.

### How to Insert Images:
- **"Browse" Button:** Click to open the file selection dialog. Select an image file (formats: PNG, JPG, JPEG, etc.) from your computer.
- **"Screenshot" Button:** Click to take a screenshot. Drag the mouse to select the screen area you want to capture, then click OK.

**Note:** Some fields will be filled automatically based on your selections, such as autofill for certain equipment or deviation calculations.

## Chapter 5: Generating the Document

1. Make sure all required fields have been filled.
2. Click the "GENERAR DOCUMENTO DE WORD (.docx)" button.
3. The "Save As" window will appear. Select the save location and file name for the Word document.
4. Click "Save".
5. A confirmation message will appear after the document is successfully saved.

[Screenshot: Save process]

## Chapter 6: Frequently Asked Questions (FAQ) & Troubleshooting

**Q: The application won't open.**  
A: Make sure you have followed the installation steps correctly. Try running as administrator if necessary.

**Q: I see a "Plantilla no encontrada" (Template not found) message.**  
A: Make sure the `New_Template2.docx` file is inside the `templates` folder, which should be in the same location as the application.

**Q: I get an error when saving the document.**  
A: Make sure you have permission to write files in the selected folder. Try selecting another folder like Desktop.

**Q: How do I load data from Excel?**  
A: Select the start and end date range, then click the "CARGAR DATOS DE EXCEL" button. Select the appropriate Excel file.

**Q: Certain fields don't appear or can't be filled.**  
A: Make sure the number of rows selected in the spin boxes matches your needs. If still problematic, restart the application.

**Q: The application is slow or crashes.**  
A: Make sure the system meets the minimum requirements. Close other unnecessary applications.

If problems persist, contact technical support with the error details that appear.
