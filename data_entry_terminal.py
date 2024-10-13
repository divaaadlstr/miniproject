import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Definisikan scope untuk mengakses Google Sheets dan Google Drive
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Baca file credentials
creds = ServiceAccountCredentials.from_json_keyfile_name(
    "credentials.json", scope)

# Hubungkan ke Google Sheets
client = gspread.authorize(creds)


def input_data():
    data = {
        "Name": [],
        "Age": [],
        "Email": []
    }

    while True:
        name = input("Input nama lengkap: ")
        age = input("Input umur sekarang: ")
        email = input("Input alamat email: ")

        data["Name"].append(name)
        data["Age"].append(age)
        data["Email"].append(email)

        another = input("Tambah data lagi? (y/n): ").lower()
        if another != 'y':
            break

    return data


def export_to_excel(data, filename="dataEntry.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Data berhasil diekspor ke {filename}")


def upload_to_google_sheets(data, sheet_name='Sheet1'):
    df = pd.DataFrame(data)

    # Buat atau buka spreadsheet
    try:
        spreadsheet = client.open(sheet_name)
        print(f"Membuka spreadsheet: {sheet_name}")
    except gspread.SpreadsheetNotFound:
        spreadsheet = client.create(sheet_name)
        print(f"Spreadsheet'{sheet_name}' tidak ditemukan. Membuat yang baru.")

    # Ambil worksheet pertama atau buat yang baru
    try:
        worksheet = spreadsheet.sheet1
        print("Menggunakan worksheet yang sudah ada.")
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title="Sheet1", rows="100", cols="3")
        print("Worksheet baru telah dibuat.")

    # Tanyakan kepada pengguna apakah ingin menghapus data sebelumnya
    choice = input(
        "Ingin menghapus data sebelumnya (h) atau menambahkan data (a)? ").lower()

    if choice == 'h':
        # Menghapus isi worksheet sebelumnya
        try:
            worksheet.clear()
            print("Menghapus isi worksheet sebelumnya.")
        except Exception as e:
            print(f"Gagal menghapus isi worksheet: {e}")

    elif choice == 'a':
        # Ambil data yang sudah ada
        existing_data = worksheet.get_all_records()
        print("Data yang sudah ada di Google Sheets:", existing_data)

        # Tambahkan data baru ke data yang sudah ada
        existing_df = pd.DataFrame(existing_data)
        df = pd.concat([existing_df, df], ignore_index=True)
        print("Data setelah ditambahkan:", df.values.tolist())
    else:
        print("Pilihan tidak valid. Menghapus data sebelumnya secara default.")

    # Menampilkan data yang akan diunggah
    print("Data yang akan diunggah:", df.values.tolist())

    # Menulis data ke worksheet
    try:
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print(f"Data berhasil diunggah ke Google Sheets di {sheet_name}.")
    except Exception as e:
        print(f"Gagal mengunggah data ke Google Sheets: {e}")

    # Cek isi worksheet setelah update
    updated_values = worksheet.get_all_values()
    print("Isi worksheet setelah update:", updated_values)


if __name__ == "__main__":
    user_data = input_data()

    # Tanyakan kepada pengguna apakah ingin mengekspor ke Excel atau Google Sheets
    export_choice = input(
        "Ingin mengekspor ke Excel (e) atau Google Sheets (g)? ").lower()

    if export_choice == 'e':
        export_to_excel(user_data)
    elif export_choice == 'g':
        sheet_name = input("Masukkan nama spreadsheet yang ingin digunakan: ")
        upload_to_google_sheets(user_data, sheet_name)
    else:
        print("Pilihan tidak valid.")
