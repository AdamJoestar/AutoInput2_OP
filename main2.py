import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QMessageBox, 
    QScrollArea, QGridLayout, QTextEdit, QGroupBox, QFileDialog
)
from PyQt5.QtCore import Qt
from docx import Document
from datetime import date
import os
import re 

# --- Konfigurasi File ---
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
TEMPLATE_FILENAME = "New_Template2.docx" 
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, TEMPLATE_FILENAME)

# --- Definisi Placeholders & Input Fields ---
FIELD_DEFINITIONS = {
    # 0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO
    "TEXT1": {"placeholder": "[TEXT1]", "label": "Solicitante", "type": "text"},
    "TEXT4": {"placeholder": "[TEXT4]", "label": "Operario del ensayo", "type": "text"},
    "TEXT2": {"placeholder": "[TEXT2]", "label": "Departamento", "type": "text"},
    "TEXT5": {"placeholder": "[TEXT5]", "label": "Responsable del ensayo", "type": "text"},
    "TEXT3": {"placeholder": "[TEXT3]", "label": "Fecha de solicitud (DD/MM/YYYY)", "type": "text"},
    
    # 1. INFORMACIÓN GENERAL DEL PRODUCTO
    "TEXT6": {"placeholder": "[TEXT6]", "label": "Referencia del modelo ensayado", "type": "text"},
    "TEXT7": {"placeholder": "[TEXT7]", "label": "Aplicación", "type": "text"},
    "TEXT8": {"placeholder": "[TEXT8]", "label": "Fuente de luz", "type": "text"},
    
    # 1.1. CONDICIONES DEL ENSAYO
    "TEXT9": {"placeholder": "[TEXT9]", "label": "Ensayo térmico realizado en", "type": "text"},
    "TEXT10": {"placeholder": "[TEXT10]", "label": "Temperatura de color ensayada (CCT)", "type": "text"},
    "TEXT11": {"placeholder": "[TEXT11]", "label": "Luminaria alimentada a", "type": "text"},
    
    # 2. EQUIPOS Y MÉTODOS UTILIZADOS
    "EQUIPO1": {"placeholder": "[EQUIPO1]", "label": "Equipo 1", "type": "text"},
    "MARCA1": {"placeholder": "[MARCA1]", "label": "Marca/Modelo 1", "type": "text"},
    "TIPO1": {"placeholder": "[TIPO1]", "label": "Tipo/Aplicación 1", "type": "text"},
    "FECHA1": {"placeholder": "[FECHA1]", "label": "Fecha de calibración 1", "type": "text"},
    "OBSER1": {"placeholder": "[OBSER1]", "label": "Observaciones 1", "type": "text"},
    "EQUIPO2": {"placeholder": "[EQUIPO2]", "label": "Equipo 2", "type": "text"},
    "MARCA2": {"placeholder": "[MARCA2]", "label": "Marca/Modelo 2", "type": "text"},
    "TIPO2": {"placeholder": "[TIPO2]", "label": "Tipo/Aplicación 2", "type": "text"},
    "FECHA2": {"placeholder": "[FECHA2]", "label": "Fecha de calibración 2", "type": "text"},
    "OBSER2": {"placeholder": "[OBSER2]", "label": "Observaciones 2", "type": "text"},
    "EQUIPO3": {"placeholder": "[EQUIPO3]", "label": "Equipo 3", "type": "text"},
    "MARCA3": {"placeholder": "[MARCA3]", "label": "Marca/Modelo 3", "type": "text"},
    "TIPO3": {"placeholder": "[TIPO3]", "label": "Tipo/Aplicación 3", "type": "text"},
    "FECHA3": {"placeholder": "[FECHA3]", "label": "Fecha de calibración 3", "type": "text"},
    "OBSER3": {"placeholder": "[OBSER3]", "label": "Observaciones 3", "type": "text"},
    "EQUIPO4": {"placeholder": "[EQUIPO4]", "label": "Equipo 4", "type": "text"},
    "MARCA4": {"placeholder": "[MARCA4]", "label": "Marca/Modelo 4", "type": "text"},
    "TIPO4": {"placeholder": "[TIPO4]", "label": "Tipo/Aplicación 4", "type": "text"},
    "FECHA4": {"placeholder": "[FECHA4]", "label": "Fecha de calibración 4", "type": "text"},
    "OBSER4": {"placeholder": "[OBSER4]", "label": "Observaciones 4", "type": "text"},
    "EQUIPO5": {"placeholder": "[EQUIPO5]", "label": "Equipo 5", "type": "text"},
    "MARCA5": {"placeholder": "[MARCA5]", "label": "Marca/Modelo 5", "type": "text"},
    "TIPO5": {"placeholder": "[TIPO5]", "label": "Tipo/Aplicación 5", "type": "text"},
    "FECHA5": {"placeholder": "[FECHA5]", "label": "Fecha de calibración 5", "type": "text"},
    "OBSER5": {"placeholder": "[OBSER5]", "label": "Observaciones 5", "type": "text"},
    "EQUIPO6": {"placeholder": "[EQUIPO6]", "label": "Equipo 6", "type": "text"},
    "MARCA6": {"placeholder": "[MARCA6]", "label": "Marca/Modelo 6", "type": "text"},
    "TIPO6": {"placeholder": "[TIPO6]", "label": "Tipo/Aplicación 6", "type": "text"},
    "FECHA6": {"placeholder": "[FECHA6]", "label": "Fecha de calibración 6", "type": "text"},
    "OBSER6": {"placeholder": "[OBSER6]", "label": "Observaciones 6", "type": "text"},
    "EQUIPO7": {"placeholder": "[EQUIPO7]", "label": "Equipo 7", "type": "text"},
    "MARCA7": {"placeholder": "[MARCA7]", "label": "Marca/Modelo 7", "type": "text"},
    "TIPO7": {"placeholder": "[TIPO7]", "label": "Tipo/Aplicación 7", "type": "text"},
    "FECHA7": {"placeholder": "[FECHA7]", "label": "Fecha de calibración 7", "type": "text"},
    "OBSER7": {"placeholder": "[OBSER7]", "label": "Observaciones 7", "type": "text"},
    "TEXT12": {"placeholder": "[TEXT12]", "label": "Método de ensayo", "type": "text"},
    
    # 3. TEMPERATURAS REGISTRADAS
    "PUNTO1": {"placeholder": "[PUNTO1]", "label": "Punto de Medición 1", "type": "text"},
    "UNIDAD1": {"placeholder": "[UNIDAD1]", "label": "Unidad 1", "type": "text"},
    "LIMITE1": {"placeholder": "[LIMITE1]", "label": "Límite Máximo 1", "type": "text"},
    "TEMP1": {"placeholder": "[TEMP1]", "label": "Temperatura Medida 1", "type": "text"},
    "PUNTO2": {"placeholder": "[PUNTO2]", "label": "Punto de Medición 2", "type": "text"},
    "UNIDAD2": {"placeholder": "[UNIDAD2]", "label": "Unidad 2", "type": "text"},
    "LIMITE2": {"placeholder": "[LIMITE2]", "label": "Límite Máximo 2", "type": "text"},
    "TEMP2": {"placeholder": "[TEMP2]", "label": "Temperatura Medida 2", "type": "text"},
    "PUNTO3": {"placeholder": "[PUNTO3]", "label": "Punto de Medición 3", "type": "text"},
    "UNIDAD3": {"placeholder": "[UNIDAD3]", "label": "Unidad 3", "type": "text"},
    "LIMITE3": {"placeholder": "[LIMITE3]", "label": "Límite Máximo 3", "type": "text"},
    "TEMP3": {"placeholder": "[TEMP3]", "label": "Temperatura Medida 3", "type": "text"},
    "PUNTO4": {"placeholder": "[PUNTO4]", "label": "Punto de Medición 4", "type": "text"},
    "UNIDAD4": {"placeholder": "[UNIDAD4]", "label": "Unidad 4", "type": "text"},
    "LIMITE4": {"placeholder": "[LIMITE4]", "label": "Límite Máximo 4", "type": "text"},
    "TEMP4": {"placeholder": "[TEMP4]", "label": "Temperatura Medida 4", "type": "text"},
    "PUNTO5": {"placeholder": "[PUNTO5]", "label": "Punto de Medición 5", "type": "text"},
    "UNIDAD5": {"placeholder": "[UNIDAD5]", "label": "Unidad 5", "type": "text"},
    "LIMITE5": {"placeholder": "[LIMITE5]", "label": "Límite Máximo 5", "type": "text"},
    "TEMP5": {"placeholder": "[TEMP5]", "label": "Temperatura Medida 5", "type": "text"},
    "TEXT13": {"placeholder": "[TEXT13]", "label": "NOTA", "type": "text"},
    
    # 3.1. GRÁFICA GENERADA
    "IMAGE1": {"placeholder": "[IMAGE1]", "label": "Imagen 1", "type": "file"},
    "TITLE1": {"placeholder": "[TITLE1]", "label": "Título 1", "type": "text"},
    "DESC1": {"placeholder": "[DESC1]", "label": "Descripción 1", "type": "text"},
    "IMAGE2": {"placeholder": "[IMAGE2]", "label": "Imagen 2", "type": "file"},
    "DESC2": {"placeholder": "[DESC2]", "label": "Descripción 2", "type": "text"},
    
    # 4. ESTABILIZACIÓN TÉRMICA
    "TEXT14": {"placeholder": "[TEXT14]", "label": "Estabilización térmica", "type": "text"},
    "MEDICI1": {"placeholder": "[MEDICI1]", "label": "Punto de Medición 1", "type": "text"},
    "UNI1": {"placeholder": "[UNI1]", "label": "Unidad 1", "type": "text"},
    "VALMIN1": {"placeholder": "[VALMIN1]", "label": "Valor Mínimo 1", "type": "text"},
    "VALMAX1": {"placeholder": "[VALMAX1]", "label": "Valor Máximo 1", "type": "text"},
    "DESVI1": {"placeholder": "[DESVI1]", "label": "Desviación 1", "type": "text"},
    "MEDICI2": {"placeholder": "[MEDICI2]", "label": "Punto de Medición 2", "type": "text"},
    "UNI2": {"placeholder": "[UNI2]", "label": "Unidad 2", "type": "text"},
    "VALMIN2": {"placeholder": "[VALMIN2]", "label": "Valor Mínimo 2", "type": "text"},
    "VALMAX2": {"placeholder": "[VALMAX2]", "label": "Valor Máximo 2", "type": "text"},
    "DESVI2": {"placeholder": "[DESVI2]", "label": "Desviación 2", "type": "text"},
    "MEDICI3": {"placeholder": "[MEDICI3]", "label": "Punto de Medición 3", "type": "text"},
    "UNI3": {"placeholder": "[UNI3]", "label": "Unidad 3", "type": "text"},
    "VALMIN3": {"placeholder": "[VALMIN3]", "label": "Valor Mínimo 3", "type": "text"},
    "VALMAX3": {"placeholder": "[VALMAX3]", "label": "Valor Máximo 3", "type": "text"},
    "DESVI3": {"placeholder": "[DESVI3]", "label": "Desviación 3", "type": "text"},
    "MEDICI4": {"placeholder": "[MEDICI4]", "label": "Punto de Medición 4", "type": "text"},
    "UNI4": {"placeholder": "[UNI4]", "label": "Unidad 4", "type": "text"},
    "VALMIN4": {"placeholder": "[VALMIN4]", "label": "Valor Mínimo 4", "type": "text"},
    "VALMAX4": {"placeholder": "[VALMAX4]", "label": "Valor Máximo 4", "type": "text"},
    "DESVI4": {"placeholder": "[DESVI4]", "label": "Desviación 4", "type": "text"},
    "MEDICI5": {"placeholder": "[MEDICI5]", "label": "Punto de Medición 5", "type": "text"},
    "UNI5": {"placeholder": "[UNI5]", "label": "Unidad 5", "type": "text"},
    "VALMIN5": {"placeholder": "[VALMIN5]", "label": "Valor Mínimo 5", "type": "text"},
    "VALMAX5": {"placeholder": "[VALMAX5]", "label": "Valor Máximo 5", "type": "text"},
    "DESVI5": {"placeholder": "[DESVI5]", "label": "Desviación 5", "type": "text"},
    
    # 5. RESULTADOS
    "PUNTODE1": {"placeholder": "[PUNTODE1]", "label": "Punto de Medición 1", "type": "text"},
    "UNIC1": {"placeholder": "[UNIC1]", "label": "Unidad 1", "type": "text"},
    "LIMITE1": {"placeholder": "[LIMITE1]", "label": "Límite Máximo 1", "type": "text"},
    "TEMPE1": {"placeholder": "[TEMPE1]", "label": "Temperatura final 1", "type": "text"},
    "RESULT1": {"placeholder": "[RESULT1]", "label": "Resultado 1", "type": "text"},
    "PUNTODE2": {"placeholder": "[PUNTODE2]", "label": "Punto de Medición 2", "type": "text"},
    "UNIC2": {"placeholder": "[UNIC2]", "label": "Unidad 2", "type": "text"},
    "LIMITE2": {"placeholder": "[LIMITE2]", "label": "Límite Máximo 2", "type": "text"},
    "TEMPE2": {"placeholder": "[TEMPE2]", "label": "Temperatura final 2", "type": "text"},
    "RESULT2": {"placeholder": "[RESULT2]", "label": "Resultado 2", "type": "text"},
    "PUNTODE3": {"placeholder": "[PUNTODE3]", "label": "Punto de Medición 3", "type": "text"},
    "UNIC3": {"placeholder": "[UNIC3]", "label": "Unidad 3", "type": "text"},
    "LIMITE3": {"placeholder": "[LIMITE3]", "label": "Límite Máximo 3", "type": "text"},
    "TEMPE3": {"placeholder": "[TEMPE3]", "label": "Temperatura final 3", "type": "text"},
    "RESULT3": {"placeholder": "[RESULT3]", "label": "Resultado 3", "type": "text"},
    "PUNTODE4": {"placeholder": "[PUNTODE4]", "label": "Punto de Medición 4", "type": "text"},
    "UNIC4": {"placeholder": "[UNIC4]", "label": "Unidad 4", "type": "text"},
    "LIMITE4": {"placeholder": "[LIMITE4]", "label": "Límite Máximo 4", "type": "text"},
    "TEMPE4": {"placeholder": "[TEMPE4]", "label": "Temperatura final 4", "type": "text"},
    "RESULT4": {"placeholder": "[RESULT4]", "label": "Resultado 4", "type": "text"},
    "PUNTODE5": {"placeholder": "[PUNTODE5]", "label": "Punto de Medición 5", "type": "text"},
    "UNIC5": {"placeholder": "[UNIC5]", "label": "Unidad 5", "type": "text"},
    "LIMITE5": {"placeholder": "[LIMITE5]", "label": "Límite Máximo 5", "type": "text"},
    "TEMPE5": {"placeholder": "[TEMPE5]", "label": "Temperatura final 5", "type": "text"},
    "RESULT5": {"placeholder": "[RESULT5]", "label": "Resultado 5", "type": "text"},
    
    # 6. CONCLUSIONES DEL LABORATORIO
    "TEXT15": {"placeholder": "[TEXT15]", "label": "Conclusiones del laboratorio", "type": "text"},
    
    # 7. FOTOGRAFIAS
    "TITLE3": {"placeholder": "[TITLE3]", "label": "Título 3", "type": "text"},
    "IMAGE3": {"placeholder": "[IMAGE3]", "label": "Imagen 3", "type": "file"},
    "TITLE4": {"placeholder": "[TITLE4]", "label": "Título 4", "type": "text"},
    "IMAGE4": {"placeholder": "[IMAGE4]", "label": "Imagen 4", "type": "file"},
    "TITLE5": {"placeholder": "[TITLE5]", "label": "Título 5", "type": "text"},
    "IMAGE5": {"placeholder": "[IMAGE5]", "label": "Imagen 5", "type": "file"},
    "TITLE6": {"placeholder": "[TITLE6]", "label": "Título 6", "type": "text"},
    "IMAGE6": {"placeholder": "[IMAGE6]", "label": "Imagen 6", "type": "file"},
    "TITLE7": {"placeholder": "[TITLE7]", "label": "Título 7", "type": "text"},
    "IMAGE7": {"placeholder": "[IMAGE7]", "label": "Imagen 7", "type": "file"},
    "TITLE8": {"placeholder": "[TITLE8]", "label": "Título 8", "type": "text"},
    "IMAGE8": {"placeholder": "[IMAGE8]", "label": "Imagen 8", "type": "file"},
}


class DocumentGeneratorApp(QWidget):
    """Aplikasi untuk menginput data dan menghasilkan dokumen Word dari template."""
    def __init__(self):
        super().__init__()
        self.input_widgets = {} 
        self.setWindowTitle("Generador de Anexo II al Informe")
        self.setStyleSheet("font-size: 14px; font-family: Arial;")
        self.init_ui()

    def init_ui(self):
        """Membangun antarmuka pengguna."""
        main_layout = QVBoxLayout(self)

        # --- Judul ---
        title = QLabel("Ingresar Datos Para el Anexo II")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px; color: #2C3E50;")
        main_layout.addWidget(title)
        
        # --- Scroll Area untuk banyak input ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        form_layout = QVBoxLayout(content_widget)
        form_layout.setSpacing(15)

        # Group 0: INFORMACIÓN DEL SOLICITANTE DEL ENSAYO
        self.create_input_group(form_layout, "0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO", [
            "TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3"
        ])

        # Group 1: INFORMACIÓN GENERAL DEL PRODUCTO
        self.create_input_group(form_layout, "1. INFORMACIÓN GENERAL DEL PRODUCTO", [
            "TEXT6", "TEXT7", "TEXT8"
        ])

        # Group 1.1: CONDICIONES DEL ENSAYO
        self.create_input_group(form_layout, "1.1. CONDICIONES DEL ENSAYO", [
            "TEXT9", "TEXT10", "TEXT11"
        ])

        # Group 2: EQUIPOS Y MÉTODOS UTILIZADOS
        self.create_input_group(form_layout, "2. EQUIPOS Y MÉTODOS UTILIZADOS", [
            "EQUIPO1", "MARCA1", "TIPO1", "FECHA1", "OBSER1",
            "EQUIPO2", "MARCA2", "TIPO2", "FECHA2", "OBSER2",
            "EQUIPO3", "MARCA3", "TIPO3", "FECHA3", "OBSER3",
            "EQUIPO4", "MARCA4", "TIPO4", "FECHA4", "OBSER4",
            "EQUIPO5", "MARCA5", "TIPO5", "FECHA5", "OBSER5",
            "EQUIPO6", "MARCA6", "TIPO6", "FECHA6", "OBSER6",
            "EQUIPO7", "MARCA7", "TIPO7", "FECHA7", "OBSER7",
            "TEXT12"
        ])

        # Group 3: TEMPERATURAS REGISTRADAS
        self.create_input_group(form_layout, "3. TEMPERATURAS REGISTRADAS", [
            "PUNTO1", "UNIDAD1", "LIMITE1", "TEMP1",
            "PUNTO2", "UNIDAD2", "LIMITE2", "TEMP2",
            "PUNTO3", "UNIDAD3", "LIMITE3", "TEMP3",
            "PUNTO4", "UNIDAD4", "LIMITE4", "TEMP4",
            "PUNTO5", "UNIDAD5", "LIMITE5", "TEMP5",
            "TEXT13"
        ])

        # Group 3.1: GRÁFICA GENERADA
        self.create_input_group(form_layout, "3.1. GRÁFICA GENERADA", [
            "IMAGE1", "TITLE1", "DESC1", "IMAGE2", "DESC2"
        ])

        # Group 4: ESTABILIZACIÓN TÉRMICA
        self.create_input_group(form_layout, "4. ESTABILIZACIÓN TÉRMICA", [
            "TEXT14",
            "MEDICI1", "UNI1", "VALMIN1", "VALMAX1", "DESVI1",
            "MEDICI2", "UNI2", "VALMIN2", "VALMAX2", "DESVI2",
            "MEDICI3", "UNI3", "VALMIN3", "VALMAX3", "DESVI3",
            "MEDICI4", "UNI4", "VALMIN4", "VALMAX4", "DESVI4",
            "MEDICI5", "UNI5", "VALMIN5", "VALMAX5", "DESVI5"
        ])

        # Group 5: RESULTADOS
        self.create_input_group(form_layout, "5. RESULTADOS", [
            "PUNTODE1", "UNIC1", "LIMITE1", "TEMPE1", "RESULT1",
            "PUNTODE2", "UNIC2", "LIMITE2", "TEMPE2", "RESULT2",
            "PUNTODE3", "UNIC3", "LIMITE3", "TEMPE3", "RESULT3",
            "PUNTODE4", "UNIC4", "LIMITE4", "TEMPE4", "RESULT4",
            "PUNTODE5", "UNIC5", "LIMITE5", "TEMPE5", "RESULT5"
        ])

        # Group 6: CONCLUSIONES DEL LABORATORIO
        self.create_input_group(form_layout, "6. CONCLUSIONES DEL LABORATORIO", [
            "TEXT15"
        ])

        # Group 7: FOTOGRAFIAS
        self.create_input_group(form_layout, "7. FOTOGRAFIAS", [
            "TITLE3", "IMAGE3", "TITLE4", "IMAGE4", "TITLE5", "IMAGE5",
            "TITLE6", "IMAGE6", "TITLE7", "IMAGE7", "TITLE8", "IMAGE8"
        ])

        scroll.setWidget(content_widget)
        main_layout.addWidget(scroll)

        # --- Tombol Generate ---
        self.generate_button = QPushButton("GENERAR DOCUMENTO DE WORD (.docx)")
        self.generate_button.setStyleSheet(
            "background-color: #3498DB; color: white; padding: 12px; border-radius: 8px; font-weight: bold;"
        )
        self.generate_button.clicked.connect(self.generate_document)
        main_layout.addWidget(self.generate_button)

        # --- Informasi Template ---
        info = QLabel(
            f"**Plantilla utilizada:** '{TEMPLATE_FILENAME}'\n"
            f"Asegúrate de que este archivo esté en la carpeta: '{TEMPLATES_DIR}'"
        )
        info.setStyleSheet("font-size: 10px; color: gray; margin-top: 5px;")
        main_layout.addWidget(info)

        self.setLayout(main_layout)
        self.resize(600, 700)

    def create_input_group(self, parent_layout, title, keys):
        """Membuat group box untuk input yang terorganisir."""
        group_box = QGroupBox(title)
        group_box.setStyleSheet("font-weight: bold; margin-top: 10px;")
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)
        
        row = 0
        col = 0
        
        for key in keys:
            definition = FIELD_DEFINITIONS[key]
            
            label = QLabel(f"{definition['label']}:")
            
            if definition['type'] == "text":
                if key in ["TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3", "TEXT6", "TEXT7", "TEXT8", "TEXT9", "TEXT10", "TEXT11", "TEXT12", "TEXT13", "TEXT14", "TEXT15"]:
                    input_field = QTextEdit()
                    input_field.setMinimumHeight(60)
                    grid_layout.addWidget(label, row, 0, 1, 2)
                    grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                    row += 2
                    col = 0
                else:
                    input_field = QLineEdit()
                    input_field.setMinimumHeight(30)
                    grid_layout.addWidget(label, row, col)
                    grid_layout.addWidget(input_field, row + 1, col)
                    col = 1 - col
                    if col == 0:
                        row += 2
            elif definition['type'] == "file":
                input_field = QLineEdit()
                input_field.setMinimumHeight(30)
                browse_button = QPushButton("Browse")
                browse_button.clicked.connect(lambda _, field=input_field: self.browse_file(field))
                grid_layout.addWidget(label, row, 0)
                grid_layout.addWidget(input_field, row + 1, 0)
                grid_layout.addWidget(browse_button, row + 1, 1)
                row += 2
                col = 0
            
            self.input_widgets[key] = input_field

        group_box.setLayout(grid_layout)
        parent_layout.addWidget(group_box)
        
    def browse_file(self, field):
        """Browse file untuk input gambar."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Image Files (*.png *.jpg *.jpeg *.gif)")
        if file_path:
            field.setText(file_path)

    def replace_in_paragraph(self, paragraph, placeholder, value):
        """Mengganti placeholder di dalam paragraf."""
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, value)
            for run in paragraph.runs:
                run.font.name = "Gordita Light"
                run.font.size = 9 * 12700

    def replace_in_tables(self, document, replacement_data):
        """Mengganti placeholder di dalam sel tabel."""
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, value in replacement_data.items():
                            self.replace_in_paragraph(paragraph, placeholder, value)

    def replace_in_headers(self, document, replacement_data):
        """Mengganti placeholder di dalam header."""
        for section in document.sections:
            header = section.header
            for paragraph in header.paragraphs:
                for placeholder, value in replacement_data.items():
                    self.replace_in_paragraph(paragraph, placeholder, value)
            for table in header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for placeholder, value in replacement_data.items():
                                self.replace_in_paragraph(paragraph, placeholder, value)

    def replace_in_footers(self, document, replacement_data):
        """Mengganti placeholder di dalam footer."""
        for section in document.sections:
            footer = section.footer
            for paragraph in footer.paragraphs:
                for placeholder, value in replacement_data.items():
                    self.replace_in_paragraph(paragraph, placeholder, value)
            for table in footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for placeholder, value in replacement_data.items():
                                self.replace_in_paragraph(paragraph, placeholder, value)

    def replace_images(self, document, replacement_data):
        """Mengganti placeholder gambar dengan gambar yang dipilih."""
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                if run.text in replacement_data:
                    image_path = replacement_data[run.text]
                    if os.path.exists(image_path):
                        run.text = ""
                        paragraph.add_run().add_picture(image_path)

    def generate_document(self):
        """Logika utama untuk membaca input, memuat template, mengganti placeholder, dan menyimpan file."""
        
        replacement_data = {}
        all_required_filled = True
        
        for key, definition in FIELD_DEFINITIONS.items():
            input_widget = self.input_widgets.get(key)
            if isinstance(input_widget, QLineEdit):
                value = input_widget.text().strip()
            elif isinstance(input_widget, QTextEdit):
                value = input_widget.toPlainText().strip()
            else:
                continue

            if key in ["TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3", "TEXT6", "TEXT7", "TEXT8", "TEXT9", "TEXT10", "TEXT11", "TEXT12", "TEXT13", "TEXT14", "TEXT15"] and not value:
                all_required_filled = False
                QMessageBox.warning(self, "Input Kosong", f"Campo obligatorio ('{definition['label']}') no puede estar vacío.")
                return

            replacement_data[definition['placeholder']] = value
        
        if not all_required_filled:
            return

        if not os.path.exists(TEMPLATE_PATH):
            QMessageBox.critical(self, "Error", 
                f"Plantilla no encontrada en: {TEMPLATE_PATH}. "
                f"Por favor coloque el archivo '{TEMPLATE_FILENAME}' que usted proporciona en la carpeta 'templates'."
            )
            return

        try:
            document = Document(TEMPLATE_PATH)
        except Exception as e:
            QMessageBox.critical(self, "Error al leer la plantilla", f"Error al cargar la plantilla: {e}")
            return

        for paragraph in document.paragraphs:
            for placeholder, value in replacement_data.items():
                self.replace_in_paragraph(paragraph, placeholder, value)

        self.replace_in_tables(document, replacement_data)
        self.replace_in_headers(document, replacement_data)
        self.replace_in_footers(document, replacement_data)
        self.replace_images(document, replacement_data)

        try:
            output_filename = f"Generated_Anexo_II_{date.today().strftime('%d_%m_%Y')}.docx"
            file_path, _ = QFileDialog.getSaveFileName(self, "Guardar documento como...", output_filename, "Word Documents (*.docx);;All Files (*)", options=QFileDialog.Options())
            if not file_path:
                QMessageBox.information(self, "Cancelado", "El almacenamiento fue cancelado por el usuario.")
                return
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'

            document.save(file_path)
            QMessageBox.information(
                self,
                "¡Listo!",
                f"El documento de Word se creó y se guardó con éxito como:\n{file_path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar el archivo", f"Error al guardar el documento: {e}")


if __name__ == '__main__':
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
        print(f"La carpeta 'templates' acaba de ser creada. Por favor, coloque el archivo '{TEMPLATE_FILENAME}' Dentro de él, luego vuelve a ejecutar la aplicación.")
        sys.exit()

    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())
