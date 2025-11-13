import pytest
from PyQt5.QtWidgets import QLineEdit, QComboBox, QDateEdit
from PyQt5.QtCore import QDate

# Karena document_processor.py dan fields.py berada di direktori yang sama,
# kita bisa mengimpornya langsung. Pastikan VS Code berjalan dari root proyek.
from document_processor import DocumentProcessor
from fields import FIELD_DEFINITIONS

# `DocumentProcessor` membutuhkan `parent_app` saat diinisialisasi.
# Kita bisa membuat objek palsu (mock) untuk memenuhi kebutuhan ini.
class MockApp:
    pass

def test_collect_replacement_data(qtbot):
    """
    Menguji apakah fungsi `collect_replacement_data` berhasil mengumpulkan
    data dari berbagai jenis widget dan mengubahnya menjadi dictionary yang benar.
    """
    # 1. ARRANGE: Siapkan semua yang dibutuhkan untuk tes

    # Buat instance palsu dari DocumentProcessor
    processor = DocumentProcessor(parent_app=MockApp())

    # Buat widget palsu (mock widgets) seolah-olah dari UI
    mock_widgets = {
        "NO_TEST": QLineEdit("TP-001"),
        "TEXT1": QLineEdit("  John Doe  "), # Uji dengan spasi ekstra untuk memastikan .strip() bekerja
            # Pastikan format tanggal sesuai dengan yang diharapkan oleh collect_replacement_data
        "DATE": QDateEdit(QDate(2025, 11, 13)),
        "EQUIPO1": QComboBox(),
    }
    # Tambahkan item ke ComboBox dan pilih salah satunya
    mock_widgets["EQUIPO1"].addItems(["ALMEMO", "TERMOHIGRÓMETRO"])
    mock_widgets["EQUIPO1"].setCurrentText("ALMEMO")

    # 2. ACT: Jalankan fungsi yang ingin diuji
    replacement_data = processor.collect_replacement_data(mock_widgets)

    # Debugging: Cetak isi replacement_data untuk memeriksa nilai
    print("--- replacement_data ---")
    import json
    print(json.dumps(replacement_data, indent=2))

    # 3. ASSERT: Periksa apakah hasilnya sesuai dengan yang diharapkan
    # Perhatikan kita memeriksa terhadap placeholder, bukan key.
    assert replacement_data["[NO_TEST]"] == "TP-001"
    assert replacement_data["[TEXT1]"] == "John Doe" # Pastikan spasi dihilangkan
    assert replacement_data["[DATE]"] == "13/11/2025"
    assert replacement_data["[EQUIPO1]"] == "ALMEMO"

    # Uji juga apakah placeholder untuk widget yang tidak ada diisi string kosong
    assert replacement_data["[TEXT6]"] == ""