import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
from config import TEMPLATES_DIR, TEMPLATE_FILENAME, TEMPLATE_PATH
from fields import FIELD_DEFINITIONS
from ui_builder import UIBuilder
from document_processor import DocumentProcessor


class DocumentGeneratorApp(QWidget):
    """Aplikasi untuk menginput data dan menghasilkan dokumen Word dari template."""
    def __init__(self):
        super().__init__()
        self.templates_dir = TEMPLATES_DIR
        self.template_filename = TEMPLATE_FILENAME
        self.template_path = TEMPLATE_PATH
        self.setWindowTitle("Generador de Anexo II al Informe")
        self.setStyleSheet("font-size: 14px; font-family: Arial;")
        self.ui_builder = UIBuilder(self)
        self.document_processor = DocumentProcessor(self)
        self.init_ui()

    def closeEvent(self, event):
        """Menampilkan popup konfirmasi sebelum menutup aplikasi."""
        reply = QMessageBox.question(self, 'Confirmar salida', '¿Estás seguro de que quieres salir?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def init_ui(self):
        """Membangun antarmuka pengguna."""
        main_layout = QVBoxLayout(self)
        self.setLayout(main_layout)
        self.ui_builder.init_ui()

    def generate_document(self):
        """Delegate to document processor."""
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
