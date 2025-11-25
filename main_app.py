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
from styles import LIGHT_THEME # Impor stylesheet
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
        self.setStyleSheet(LIGHT_THEME) # Terapkan stylesheet dari file terpisah
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
                'num_sonda': self.ui_builder.num_sonda,
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
                if 'num_sonda' in project_data:
                    self.ui_builder.num_sonda = project_data['num_sonda']

                # Rebuild form with loaded values
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
        """
        Loads temperature data from an Excel file, automatically finds the most stable 1-hour window,
        and populates the form fields with the results.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo Excel con datos de temperatura", "", "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if not file_path:
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)

        try:
            # 1. Baca dan Bersihkan Data
            df = pd.read_excel(file_path, header=0)
            if 'Waktu' not in df.columns:
                df = df.rename(columns={df.columns[0]: 'Waktu'})

            QApplication.processEvents() # Allow GUI to update

            df['Waktu'] = pd.to_datetime(df['Waktu'], errors='coerce')
            df.dropna(subset=['Waktu'], inplace=True)
            df = df.set_index('Waktu')

            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].str.replace(',', '.', regex=False).astype(float)

            QApplication.processEvents() # Allow GUI to update

            # 2. Cari Jendela Waktu 1 Jam
            one_hour = pd.Timedelta(hours=1)
            tolerance = pd.Timedelta(seconds=30)
            candidates = []
            for i in range(len(df)):
                for j in range(i + 1, len(df)):
                    duration = df.index[j] - df.index[i]
                    if one_hour - tolerance <= duration <= one_hour + tolerance:
                        candidates.append((df.index[i], df.index[j]))

            if not candidates:
                QMessageBox.warning(self, "Tidak Ditemukan", "Tidak ada periode data berdurasi 1 jam yang ditemukan di file Excel.")
                return

            QApplication.processEvents() # Allow GUI to update

            # 3. Cari Jendela Paling Stabil
            best_window = None
            min_total_deviation = float('inf')

            for start, end in candidates:
                window_df = df.loc[start:end]
                deviations = window_df.max() - window_df.min()
                total_deviation = deviations.sum()

                if total_deviation < min_total_deviation:
                    min_total_deviation = total_deviation
                    best_window = window_df

            if best_window is None:
                QMessageBox.warning(self, "Error", "Gagal menemukan jendela data terbaik.")
                return

            # 4. Ekstrak Data dari Jendela Terbaik
            final_temps = best_window.iloc[-1]
            min_vals = best_window.min()
            max_vals = best_window.max()
            deviations = max_vals - min_vals

            num_sensors = min(len(df.columns), 10)
            self.ui_builder.num_sonda = num_sensors  # Update num_sonda to match detected sensors
            self.ui_builder.rebuild_form()

            # 5. Isi Form UI
            column_names = df.columns.tolist()
            highlight_style = "background-color: #fff9c4;"  # light yellow

            for i in range(1, num_sensors + 1):
                col_name = column_names[i-1]

                # --- Mengisi Bagian 3: TEMPERATURAS REGISTRADAS ---
                punto_key = f"PUNTO{i}"
                temp_key = f"TEMP{i}"
                title_key = f"TITLE{i+2}"  # TITLE3 for sensor 1, TITLE4 for sensor 2, etc.

                # Auto-fill Punto de Medición dari nama kolom
                if punto_key in self.ui_builder.input_widgets:
                    mapped_punto = self.map_column_to_punto(col_name)
                    if mapped_punto:
                        widget = self.ui_builder.input_widgets[punto_key]
                        widget.setCurrentText(mapped_punto)
                        widget.setStyleSheet(highlight_style)
                        # Auto-fill juga field TITLE di bagian Fotografi
                        if title_key in self.ui_builder.input_widgets:
                            title_widget = self.ui_builder.input_widgets[title_key]
                            title_widget.setText(mapped_punto)
                            title_widget.setStyleSheet(highlight_style)

                # Isi Temperatur Medida (nilai akhir)
                if temp_key in self.ui_builder.input_widgets:
                    temp_value = str(round(final_temps[col_name], 2))
                    widget = self.ui_builder.input_widgets[temp_key]
                    widget.setText(temp_value)
                    widget.setStyleSheet(highlight_style)

                # --- Mengisi Bagian 4: ESTABILIZACIÓN TÉRMICA ---
                valmin_key = f"VALMIN{i}"
                valmax_key = f"VALMAX{i}"
                desvi_key = f"DESVI{i}"

                if valmin_key in self.ui_builder.input_widgets:
                    widget = self.ui_builder.input_widgets[valmin_key]
                    widget.setText(f"{min_vals[col_name]:.2f}")
                    widget.setStyleSheet(highlight_style)

                if valmax_key in self.ui_builder.input_widgets:
                    widget = self.ui_builder.input_widgets[valmax_key]
                    widget.setText(f"{max_vals[col_name]:.2f}")
                    widget.setStyleSheet(highlight_style)

                if desvi_key in self.ui_builder.input_widgets:
                    widget = self.ui_builder.input_widgets[desvi_key]
                    widget.setText(f"{deviations[col_name]:.2f}")
                    widget.setStyleSheet(highlight_style)

                # --- Mengisi Bagian 5: RESULTADOS ---
                tempe_key = f"TEMPE{i}"
                if tempe_key in self.ui_builder.input_widgets:
                    temp_value = str(round(final_temps[col_name], 2))
                    widget = self.ui_builder.input_widgets[tempe_key]
                    widget.setText(temp_value)
                    widget.setStyleSheet(highlight_style)

            QMessageBox.information(
                self, "Éxito",
                f"Datos cargados y analizados con éxito dari:\n{file_path}\n\n"
                f"Se encontró un período estable desde {best_window.index.min().strftime('%Y-%m-%d %H:%M:%S')} "
                f"hasta {best_window.index.max().strftime('%Y-%m-%d %H:%M:%S')}.\n"
                f"Encontrado {num_sensors} sensor de temperatura."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar el archivo Excel: {e}")
        finally:
            QApplication.restoreOverrideCursor()

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
