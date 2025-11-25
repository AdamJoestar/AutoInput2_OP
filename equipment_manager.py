import json
import os
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, QHBoxLayout,
    QMessageBox, QDialogButtonBox, QLineEdit, QFormLayout, QLabel, QHeaderView
)
from PyQt5.QtCore import Qt
from fields import EQUIPO_AUTOFILL_DATA, CODIGO_OPTIONS_BY_EQUIPO

EQUIPMENT_CONFIG_FILE = 'equipment_config.json'

def get_default_config():
    """Construye la configuración por defecto desde fields.py"""
    config = []
    
    # Menggabungkan data dari EQUIPO_AUTOFILL_DATA dan CODIGO_OPTIONS_BY_EQUIPO
    all_equipos = set(EQUIPO_AUTOFILL_DATA.keys()) | set(CODIGO_OPTIONS_BY_EQUIPO.keys())

    for equipo_name in sorted(list(all_equipos)):
        data = EQUIPO_AUTOFILL_DATA.get(equipo_name, {})
        codes = CODIGO_OPTIONS_BY_EQUIPO.get(equipo_name, [])
        
        # Jika satu equipo memiliki banyak kode, kita buat satu entri saja
        # dengan semua kode digabungkan.
        config.append({
            "equipo": equipo_name,
            "marca": data.get("marca", ""),
            "tipo": data.get("tipo", ""),
            "fecha": data.get("fecha", ""),
            "codigos": codes
        })
    return config

class EquipmentEditDialog(QDialog):
    """Dialog untuk menambah atau mengedit satu data peralatan."""
    def __init__(self, data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Editar Equipo" if data else "Añadir Nuevo Equipo")

        self.layout = QFormLayout(self)

        self.equipo_input = QLineEdit(data.get('equipo', '') if data else '')
        self.marca_input = QLineEdit(data.get('marca', '') if data else '')
        self.tipo_input = QLineEdit(data.get('tipo', '') if data else '')
        self.fecha_input = QLineEdit(data.get('fecha', '') if data else '')
        self.codigo_input = QLineEdit(', '.join(data.get('codigos', [])) if data else '')

        self.layout.addRow(QLabel("Nombre del Equipo:"), self.equipo_input)
        self.layout.addRow(QLabel("Marca/Modelo:"), self.marca_input)
        self.layout.addRow(QLabel("Tipo/Aplicación:"), self.tipo_input)
        self.layout.addRow(QLabel("Fecha de Calibración (DD/MM/YYYY):"), self.fecha_input)
        self.layout.addRow(QLabel("Códigos (separados por coma):"), self.codigo_input)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self.layout.addWidget(self.button_box)

    def get_data(self):
        """Mengembalikan data yang diinput dalam dialog."""
        if not self.equipo_input.text().strip():
            QMessageBox.warning(self, "Entrada Vacía", "El nombre del equipo no puede estar vacío.")
            return None

        return {
            'equipo': self.equipo_input.text().strip(),
            'marca': self.marca_input.text().strip(),
            'tipo': self.tipo_input.text().strip(),
            'fecha': self.fecha_input.text().strip(),
            'codigos': [c.strip() for c in self.codigo_input.text().split(',') if c.strip()]
        }

class EquipmentManagerDialog(QDialog):
    """Dialog untuk mengelola konfigurasi peralatan."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent
        self.setWindowTitle("Gestión de Equipos")
        self.setMinimumSize(800, 600)

        self.layout = QVBoxLayout(self)

        # Tabel untuk menampilkan peralatan
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Nombre del Equipo", "Marca/Modelo", "Tipo/Aplicación", "Fecha Cal.", "Códigos"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.layout.addWidget(self.table)

        # Tombol-tombol
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Añadir")
        self.edit_button = QPushButton("Editar")
        self.delete_button = QPushButton("Eliminar")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        self.layout.addLayout(button_layout)

        # Tombol Simpan & Batal
        self.dialog_buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.layout.addWidget(self.dialog_buttons)

        # Koneksi sinyal
        self.add_button.clicked.connect(self.add_item)
        self.edit_button.clicked.connect(self.edit_item)
        self.delete_button.clicked.connect(self.delete_item)
        self.dialog_buttons.accepted.connect(self.save_config)
        self.dialog_buttons.rejected.connect(self.reject)

        self.load_config()

    def load_config(self):
        """Memuat konfigurasi dari file JSON dan menampilkannya di tabel."""
        self.config_data = self.parent_app.equipment_config
        self.table.setRowCount(len(self.config_data))

        for row, item in enumerate(self.config_data):
            self.table.setItem(row, 0, QTableWidgetItem(item.get('equipo', '')))
            self.table.setItem(row, 1, QTableWidgetItem(item.get('marca', '')))
            self.table.setItem(row, 2, QTableWidgetItem(item.get('tipo', '')))
            self.table.setItem(row, 3, QTableWidgetItem(item.get('fecha', '')))
            self.table.setItem(row, 4, QTableWidgetItem(', '.join(item.get('codigos', []))))

    def add_item(self):
        """Membuka dialog untuk menambah item baru."""
        dialog = EquipmentEditDialog(parent=self)
        if dialog.exec_() == QDialog.Accepted:
            new_data = dialog.get_data()
            if new_data:
                self.config_data.append(new_data)
                self.refresh_table()

    def edit_item(self):
        """Membuka dialog untuk mengedit item yang dipilih."""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Aviso", "Seleccione una fila para editar.")
            return

        item_data = self.config_data[current_row]
        dialog = EquipmentEditDialog(data=item_data, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            updated_data = dialog.get_data()
            if updated_data:
                self.config_data[current_row] = updated_data
                self.refresh_table()

    def delete_item(self):
        """Menghapus item yang dipilih dari tabel."""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Aviso", "Seleccione una fila para eliminar.")
            return

        reply = QMessageBox.question(self, 'Confirmar Eliminación',
                                     f"¿Está seguro de que desea eliminar '{self.config_data[current_row]['equipo']}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            del self.config_data[current_row]
            self.refresh_table()

    def refresh_table(self):
        """Memperbarui tampilan tabel dengan data saat ini."""
        self.table.setRowCount(0)
        self.table.setRowCount(len(self.config_data))
        for row, item in enumerate(self.config_data):
            self.table.setItem(row, 0, QTableWidgetItem(item.get('equipo', '')))
            self.table.setItem(row, 1, QTableWidgetItem(item.get('marca', '')))
            self.table.setItem(row, 2, QTableWidgetItem(item.get('tipo', '')))
            self.table.setItem(row, 3, QTableWidgetItem(item.get('fecha', '')))
            self.table.setItem(row, 4, QTableWidgetItem(', '.join(item.get('codigos', []))))

    def save_config(self):
        """Menyimpan data konfigurasi saat ini ke file JSON."""
        try:
            with open(EQUIPMENT_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "Éxito", "La configuración de equipos se ha guardado correctamente.")
            self.accept() # Tutup dialog setelah menyimpan
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar la configuración: {e}")