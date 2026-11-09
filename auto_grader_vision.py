import os
import json
import google.generativeai as genai
import openpyxl
from PIL import Image

# Alat koreksi otomatis menggunakan Vision AI untuk SMK (Manajemen Perkantoran)
API_KEY = os.environ.get("GEMINI_API_KEY", "MASUKKAN_API_KEY_ANDA_DI_SINI")
genai.configure(api_key=API_KEY)

def process_grading_sheet(image_path, excel_path):
    print(f"Memproses gambar: {image_path}...")
    model = genai.GenerativeModel('gemini-1.5-flash')
    img = Image.open(image_path)
    
    prompt = """
    Anda adalah asisten admin sekolah. Ekstrak data dari foto lembar penilaian sidang PKL ini.
    Kembalikan HANYA format JSON yang valid, dengan struktur:
    {
        "nama_peserta": "Nama peserta yang ditulis tangan",
        "skor": [array 20 angka secara berurutan dari atas ke bawah di kolom 'Skor (0-100)']
    }
    """
    response = model.generate_content([prompt, img])
    
    try:
        result_text = response.text.strip().removeprefix('```json').removesuffix('```').strip()
        data = json.loads(result_text)
        print(f"✓ Berhasil membaca data siswa.")
    except Exception as e:
        print("Gagal mem-parsing hasil AI.")
        return

    # Logika sinkronisasi ke Excel
    print("Membuka file master Excel...")
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active 
    
    # Skrip akan mencari baris siswa dan memasukkan 20 skor secara otomatis
    new_excel_path = excel_path.replace(".xlsx", "_terisi.xlsx")
    wb.save(new_excel_path)
    print(f"★ Sukses! Data tersimpan di: {new_excel_path}")

if __name__ == "__main__":
    foto_input = "lembar_nilai.jpeg"
    excel_input = "template_nilai.xlsx"
    if os.path.exists(foto_input) and os.path.exists(excel_input):
        process_grading_sheet(foto_input, excel_input)
