import json
import firebase_admin
from firebase_admin import credentials, firestore

# --- KONFIGURASI ---
# Pastikan kedua file ini ada di folder yang sama dengan skrip ini
SERVICE_ACCOUNT_KEY_FILE = 'serviceAccountKey.json'
EQUIPMENT_JSON_FILE = 'equipment_config.json'
COLLECTION_NAME = 'equipments'
# -------------------

def upload_equipment_data():
    """
    Membaca file equipment_config.json dan mengunggahnya ke Cloud Firestore.
    PERINGATAN: Ini akan menghapus semua data lama di collection 'equipments'
    dan menggantinya dengan data dari file JSON.
    """
    try:
        # 1. Inisialisasi Firebase
        cred = credentials.Certificate(SERVICE_ACCOUNT_KEY_FILE)
        firebase_admin.initialize_app(cred)
        db = firestore.client()
        print("Koneksi ke Firebase berhasil.")

        # 2. Baca file JSON lokal
        with open(EQUIPMENT_JSON_FILE, 'r', encoding='utf-8') as f:
            local_data = json.load(f)
        print(f"Berhasil membaca {len(local_data)} item dari '{EQUIPMENT_JSON_FILE}'.")

        # 3. Dapatkan referensi ke collection
        equip_collection = db.collection(COLLECTION_NAME)

        # 4. Hapus semua dokumen lama (opsional, tapi disarankan untuk sinkronisasi)
        print(f"Menghapus data lama dari collection '{COLLECTION_NAME}'...")
        for doc in equip_collection.stream():
            doc.reference.delete()
        print("Data lama berhasil dihapus.")

        # 5. Unggah data baru
        print("Mengunggah data baru...")
        for item in local_data:
            # Gunakan 'equipo' sebagai ID dokumen
            doc_id = item.get('equipo')
            if doc_id:
                equip_collection.document(doc_id).set(item)
                print(f"  -> Dokumen '{doc_id}' berhasil diunggah.")
            else:
                print("  -> Peringatan: Melewatkan item tanpa nama 'equipo'.")
        
        print("\nProses unggah selesai dengan sukses!")

    except FileNotFoundError as e:
        print(f"ERROR: File tidak ditemukan - {e}. Pastikan '{SERVICE_ACCOUNT_KEY_FILE}' dan '{EQUIPMENT_JSON_FILE}' ada di folder yang benar.")
    except Exception as e:
        print(f"Terjadi error: {e}")

if __name__ == '__main__':
    # Minta konfirmasi sebelum menjalankan
    confirm = input(f"Anda akan MENGGANTI SEMUA data di collection '{COLLECTION_NAME}' dengan data dari '{EQUIPMENT_JSON_FILE}'.\nLanjutkan? (y/n): ")
    if confirm.lower() == 'y':
        upload_equipment_data()
    else:
        print("Operasi dibatalkan.")

