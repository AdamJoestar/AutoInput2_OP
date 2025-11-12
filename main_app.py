import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
from config import TEMPLATES_DIR, TEMPLATE_FILENAME, TEMPLATE_PATH
from fields import FIELD_DEFINITIONS
from ui_builder import UIBuilder
from document_processor import DocumentProcessor
import os


class DocumentGeneratorApp(QWidget):
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
        self.setWindowTitle("Generador de Anexo II al Informe")
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
        """)
        self.ui_builder = UIBuilder(self)
        self.document_processor = DocumentProcessor(self)
        self.load_stabilization_template()
        self.init_ui()

    def closeEvent(self, event):
        """
        Handles the application window's close event.

        Displays a confirmation dialog to the user before exiting.

        Args:
            event (QCloseEvent): The event received when the window is about to close.
        """
        reply = QMessageBox.question(self, 'Confirmar salida', '¿Estás seguro de que quieres salir?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def init_ui(self):
        """Initializes and builds the main user interface."""
        main_layout = QVBoxLayout(self)
        self.setLayout(main_layout)
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
        """Load the stabilization template and set it as default for TEXT_EST."""
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
