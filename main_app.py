import sys
import os
import json
import shutil
import tempfile
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QMessageBox, QMenuBar, QAction, QFileDialog, QMainWindow, QLineEdit, QTextEdit, QDateEdit, QComboBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from config import TEMPLATES_DIR, TEMPLATE_FILENAME, TEMPLATE_PATH
from fields import FIELD_DEFINITIONS
from ui_builder import UIBuilder
from document_processor import DocumentProcessor
import os


class DocumentGeneratorApp(QMainWindow):
    """Application to input data and generate Word documents from a template."""
    def __init__(self):
        """
        Initializes the main application.

        Sets window properties, stylesheet, and initializes main components
        like UIBuilder and DocumentProcessor. Also loads default text templates
        from external files.
        """
        super().__init__()
        self.templates_dir = TEMPLATES_DIR
        self.template_filename = TEMPLATE_FILENAME
        self.template_path = TEMPLATE_PATH
        self.setWindowTitle("Generador de Anexo al Informe")
        self.setWindowIcon(QIcon("logo vibia.png"))
        self.setStyleSheet("""
            QWidget {
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f5f5f5;
                color: #333;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #ddd;
                border-radius: 8px;
                margin-top: 10px;
                background-color: #ffffff;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
                font-size: 16px;
            }
            QLabel {
                color: #555;
            }
            QLineEdit, QTextEdit, QDateEdit, QComboBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
                background-color: #fff;
            }
            QLineEdit:focus, QTextEdit:focus, QDateEdit:focus, QComboBox:focus {
                border-color: #3498db;
            }
            QPushButton {
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QMessageBox QPushButton {
                background-color: #808080;
                color: #FFFFFF;
            }
            QMessageBox QPushButton:hover {
                background-color: #909090;
            }
        """)
        self.ui_builder = UIBuilder(self)
        self.document_processor = DocumentProcessor(self)
        self.load_stabilization_template()
        self.init_menu()
        self.init_ui()

    def closeEvent(self, event):
        """
        Handles the application window's close event.

        Displays a confirmation dialog to the user before exiting.

        Args:
            event (QCloseEvent): The event received when the window is about to close.
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('Confirmar salida')
        msg_box.setText('¿Estás seguro de que quieres salir?')
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.No)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #808080;
                color: #FFFFFF;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #909090;
            }
        """)
        reply = msg_box.exec_()

        if reply == QMessageBox.Yes:
            # Membersihkan file-file sementara sebelum keluar
            if hasattr(self.ui_builder, 'temp_files'):
                for temp_path in self.ui_builder.temp_files:
                    try:
                        os.remove(temp_path)
                    except OSError:
                        # Abaikan error jika file tidak ada atau tidak bisa dihapus
                        pass
            event.accept()
        else:
            event.ignore()

    def init_menu(self):
        """Initializes the menu bar with File menu for save/load project."""
        self.menu_bar = self.menuBar()  # This creates the menu bar
        file_menu = self.menu_bar.addMenu('Archivo')

        # Save Project action
        save_action = QAction('Guardar Proyecto', self)
        save_action.setShortcut('Ctrl+S')
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)

        # Load Project action
        load_action = QAction('Cargar Proyecto', self)
        load_action.setShortcut('Ctrl+O')
        load_action.triggered.connect(self.load_project)
        file_menu.addAction(load_action)

        file_menu.addSeparator()

        # Exit action
        exit_action = QAction('Salir', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def init_ui(self):
        """Initializes and builds the main user interface."""
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        self.main_layout = QVBoxLayout(central_widget)
        central_widget.setLayout(self.main_layout)
        self.ui_builder.init_ui()

    def load_method_template(self):
        """Load the method template and set it as default for TEXT12."""
        template_path = os.path.join(os.getcwd(), 'method_template.txt')
        if os.path.exists(template_path):
            with open(template_path, 'r', encoding='utf-8') as f:
                self.method_template = f.read()
        else:
            self.method_template = "Método de ensayo no disponible. Por favor, verifique el archivo method_template.txt."

    def load_stabilization_template(self):
        """Load the stabilization templat2e and set it as default for TEXT_EST."""
        template_path = os.path.join(os.getcwd(), 'stabilization_template.txt')
        if os.path.exists(template_path):
            with open(template_path, 'r', encoding='utf-8') as f:
                self.stabilization_template = f.read()
        else:
            self.stabilization_template = "Descripción de estabilización térmica no disponible."

    def load_description_template(self):
        """Load the description template and set it as default for TEXT14."""
        template_path = os.path.join(os.getcwd(), 'description_template.txt')
        if os.path.exists(template_path):
            with open(template_path, 'r', encoding='utf-8') as f:
                self.description_template = f.read()
        else:
            self.description_template = "Descripción no disponible. Por favor, verifique el archivo description_template.txt."

    def save_project(self):
        """Saves the current project data to a JSON file."""
        # Open save dialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Proyecto", "", "Archivos de Proyecto (*.json);;Todos los archivos (*)"
        )

        if file_path:
            if not file_path.lower().endswith('.json'):
                file_path += '.json'
            
            # Create a directory for the project next to the .json file
            project_base_dir = os.path.dirname(file_path)
            project_name = os.path.splitext(os.path.basename(file_path))[0]
            project_files_dir = os.path.join(project_base_dir, f"{project_name}_files")
            os.makedirs(project_files_dir, exist_ok=True)

            # Collect all input data
            project_data = {
                'spin_equipment': self.ui_builder.spin_equipment.value(),
                'spin_sonda': self.ui_builder.spin_sonda.value(),
                'input_data': {},
                'saved_files': {} # Stores relative paths
            }

            # Collect values from all input widgets and handle file paths
            for key, widget in self.ui_builder.input_widgets.items():
                definition = FIELD_DEFINITIONS.get(key, {})
                if definition.get('type') == 'file':
                    value = widget.text()
                    if value and os.path.exists(value):
                        # Copy any file (temp or browsed) to the project directory
                        filename = os.path.basename(value)
                        saved_path = os.path.join(project_files_dir, filename)
                        shutil.copy2(value, saved_path)
                        
                        # Store the relative path for portability
                        relative_path = os.path.join(f"{project_name}_files", filename)
                        project_data['input_data'][key] = relative_path
                        project_data['saved_files'][key] = relative_path
                    else:
                        project_data['input_data'][key] = "" # Store empty if path is invalid
                elif isinstance(widget, QLineEdit):
                    project_data['input_data'][key] = widget.text()
                elif isinstance(widget, QTextEdit):
                    project_data['input_data'][key] = widget.toPlainText()
                elif isinstance(widget, QDateEdit):
                    project_data['input_data'][key] = widget.date().toString("dd/MM/yyyy")
                elif isinstance(widget, QComboBox):
                    project_data['input_data'][key] = widget.currentText()

            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(project_data, f, ensure_ascii=False, indent=2)
                QMessageBox.information(self, "Éxito", f"Proyecto guardado exitosamente en:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al guardar el proyecto: {e}")

    def load_project(self):
        """Loads project data from a JSON file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Proyecto", "", "Archivos de Proyecto (*.json);;Todos los archivos (*)"
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    project_data = json.load(f)

                # Set spin box values
                if 'spin_equipment' in project_data:
                    self.ui_builder.spin_equipment.setValue(project_data['spin_equipment'])
                if 'spin_sonda' in project_data:
                    self.ui_builder.spin_sonda.setValue(project_data['spin_sonda'])

                # Rebuild form with loaded spin values
                self.ui_builder.rebuild_form()

                # Load input data
                if 'input_data' in project_data:
                    for key, value in project_data['input_data'].items():
                        definition = FIELD_DEFINITIONS.get(key, {})
                        if key in self.ui_builder.input_widgets:
                            widget = self.ui_builder.input_widgets[key]
                            if definition.get('type') == 'file' and value:
                                # Reconstruct the absolute path from the relative path
                                project_base_dir = os.path.dirname(file_path)
                                full_path = os.path.join(project_base_dir, value)
                                
                                if os.path.exists(full_path):
                                    widget.setText(full_path)
                                    # We don't need to add this to temp_files for cleanup
                                    # as it's now a permanent project file.
                                else:
                                    # Path is invalid, inform the user
                                    widget.setText(f"File not found: {value}")
                            elif isinstance(widget, QLineEdit):
                                widget.setText(value)
                            elif isinstance(widget, QTextEdit):
                                widget.setPlainText(value)
                            elif hasattr(widget, 'setDate') and value:
                                from PyQt5.QtCore import QDate
                                date = QDate.fromString(value, "dd/MM/yyyy")
                                if date.isValid():
                                    widget.setDate(date)
                            elif isinstance(widget, QComboBox):
                                index = widget.findText(value)
                                if index >= 0:
                                    widget.setCurrentIndex(index)

                QMessageBox.information(self, "Éxito", f"Proyecto cargado exitosamente desde:\n{file_path}")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar el proyecto: {e}")

    def load_excel_data(self):
        """Loads temperature data from an Excel file and populates the form fields, filtering by selected time range."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo Excel con datos de temperatura", "", "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )

        if not file_path:
            return

        try:
            # Read Excel file
            df = pd.read_excel(file_path)

            # Check if 'Waktu' column exists, if not, rename first column
            if 'Waktu' not in df.columns:
                df = df.rename(columns={df.columns[0]: 'Waktu'})

            # Convert 'Waktu' to datetime
            df['Waktu'] = pd.to_datetime(df['Waktu'])

            # Set 'Waktu' as index
            df = df.set_index('Waktu')

            # Ensure index is timezone-naive for comparison
            if df.index.tz is not None:
                df.index = df.index.tz_localize(None)

            # Get selected time range from UI
            start_datetime = self.ui_builder.start_datetime.dateTime().toPyDateTime()
            end_datetime = self.ui_builder.end_datetime.dateTime().toPyDateTime()

            # Ensure datetime objects are timezone-naive for comparison
            start_datetime = start_datetime.replace(tzinfo=None)
            end_datetime = end_datetime.replace(tzinfo=None)

            # Use exact selected datetime for precise filtering

            # Debug: Print time range and sample data
            print(f"Start datetime: {start_datetime}")
            print(f"End datetime: {end_datetime}")
            print(f"Data index type: {type(df.index[0])}")
            print(f"Sample index values: {df.index[:5].tolist()}")
            print(f"Total records before filter: {len(df)}")

            # Filter data by time range (inclusive)
            df_filtered = df.loc[(df.index >= start_datetime) & (df.index <= end_datetime)]

            print(f"Total records after filter: {len(df_filtered)}")
            if not df_filtered.empty:
                print(f"Filtered index range: {df_filtered.index.min()} to {df_filtered.index.max()}")

            if df_filtered.empty:
                QMessageBox.warning(self, "Advertencia", f"No se encontraron datos en el rango de tiempo seleccionado:\nDe {start_datetime} a {end_datetime}\n\nTotal registros en archivo: {len(df)}")
                return

            # Get the last row (final temperature) from filtered data
            final_temps = df_filtered.iloc[-1]

            # Populate TEMPERATURAS REGISTRADAS fields
            num_sensors = min(len(final_temps), 10)  # Max 10 sensors
            self.ui_builder.spin_sonda.setValue(num_sensors)
            self.ui_builder.rebuild_form()

            # Auto-fill Punto de Medición based on column names
            column_names = df_filtered.columns.tolist()
            for i in range(1, num_sensors + 1):
                punto_key = f"PUNTO{i}"
                temp_key = f"TEMP{i}"
                title_key = f"TITLE{i+2}"  # TITLE3 for sensor 1, TITLE4 for sensor 2, etc.
                if punto_key in self.ui_builder.input_widgets and i-1 < len(column_names):
                    column_name = column_names[i-1]
                    mapped_punto = self.map_column_to_punto(column_name)
                    if mapped_punto:
                        self.ui_builder.input_widgets[punto_key].setCurrentText(mapped_punto)
                        # Auto-fill corresponding TITLE field in Fotografías section
                        if title_key in self.ui_builder.input_widgets:
                            self.ui_builder.input_widgets[title_key].setText(mapped_punto)

                if temp_key in self.ui_builder.input_widgets:
                    temp_value = str(final_temps.iloc[i-1])
                    self.ui_builder.input_widgets[temp_key].setText(temp_value)

            # Calculate min, max, and deviation for ESTABILIZACIÓN TÉRMICA from filtered data
            if len(df_filtered) > 1:
                for i in range(1, num_sensors + 1):
                    sensor_data = df_filtered.iloc[:, i-1]
                    valmin_key = f"VALMIN{i}"
                    valmax_key = f"VALMAX{i}"
                    desvi_key = f"DESVI{i}"

                    if valmin_key in self.ui_builder.input_widgets:
                        min_val = sensor_data.min()
                        self.ui_builder.input_widgets[valmin_key].setText(f"{min_val:.2f}")

                    if valmax_key in self.ui_builder.input_widgets:
                        max_val = sensor_data.max()
                        self.ui_builder.input_widgets[valmax_key].setText(f"{max_val:.2f}")

                    if desvi_key in self.ui_builder.input_widgets:
                        desvi_val = max_val - min_val
                        self.ui_builder.input_widgets[desvi_key].setText(f"{desvi_val:.2f}")

            # Populate RESULTADOS with final temperatures
            for i in range(1, num_sensors + 1):
                tempe_key = f"TEMPE{i}"
                if tempe_key in self.ui_builder.input_widgets:
                    temp_value = str(final_temps.iloc[i-1])
                    self.ui_builder.input_widgets[tempe_key].setText(temp_value)

            QMessageBox.information(self, "Éxito", f"Datos cargados exitosamente desde:\n{file_path}\n\nRango de tiempo: {start_datetime} - {end_datetime}\nSe encontraron {num_sensors} sensores de temperatura.\nTotal registros filtrados: {len(df_filtered)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar el archivo Excel: {e}")

    def map_column_to_punto(self, column_name):
        """Maps Excel column names to Punto de Medición options using fuzzy matching."""
        import re
        from difflib import get_close_matches

        # Normalize column name: remove numbers, dots, and extra spaces
        normalized = re.sub(r'[0-9\.\s]+', '', column_name.lower())

        # Define mapping patterns
        mappings = {
            'tcled': 'Tc LED',
            'tc led': 'Tc LED',
            'tcledcalido': 'Tc LED Calido',
            'tc led calido': 'Tc LED Calido',
            'tcledfrio': 'Tc LED Frio',
            'tc led frio': 'Tc LED Frio',
            'carcasaint': 'Carcasa int.',
            'carcasa int': 'Carcasa int.',
            'carcasaext': 'Carcasa ext.',
            'carcasa ext': 'Carcasa ext.',
            'disipador': 'Disipador',
            'difusor': 'Difusor',
            'pcb': 'PCB',
            'driver': 'Driver',
            'tambiente': 'T. Ambiente',
            't ambiente': 'T. Ambiente',
            'tamb': 'T. Ambiente',
            'reflector': 'Reflector',
            'lente': 'Lente'
        }

        # Try exact match first
        if normalized in mappings:
            return mappings[normalized]

        # Try fuzzy matching
        close_matches = get_close_matches(normalized, mappings.keys(), n=1, cutoff=0.6)
        if close_matches:
            return mappings[close_matches[0]]

        # If no match found, return None
        return None

    def generate_document(self):
        """Delegates the document generation task to the document processor."""
        self.document_processor.generate_document()


if __name__ == '__main__':
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
        print(f"La carpeta 'templates' acaba de ser creada. Por favor, coloque el archivo '{TEMPLATE_FILENAME}' Dentro de él, luego vuelve a ejecutar la aplicación.")
        sys.exit()

    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    window.resize(600, 700)
    sys.exit(app.exec_())
