import tkinter as tk
from tkinter import messagebox
import pandas as pd
import re


# Fungsi untuk validasi email
def validasi_email(email):
    pola = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pola, email) is not None


# Fungsi untuk validasi usia
def validasi_usia(usia):
    return usia.isdigit() and int(usia) > 0


# Fungsi untuk menambahkan data ke tabel
def tambah_data():
    nama = entry_nama.get()
    usia = entry_usia.get()
    email = entry_email.get()

    # Validasi data input
    if not nama:
        messagebox.showerror("Error", "Nama tidak boleh kosong!")
        return
    if not validasi_usia(usia):
        messagebox.showerror("Error", "Usia harus berupa angka positif!")
        return
    if not validasi_email(email):
        messagebox.showerror("Error", "Format email tidak valid!")
        return

    # Tambahkan data ke tabel
    data["Nama"].append(nama)
    data["Usia"].append(int(usia))  # Konversi usia ke integer
    data["Email"].append(email)

    # Update tampilan tabel di GUI
    listbox_nama.insert(tk.END, nama)
    listbox_usia.insert(tk.END, usia)
    listbox_email.insert(tk.END, email)

    # Bersihkan input setelah menambahkan
    entry_nama.delete(0, tk.END)
    entry_usia.delete(0, tk.END)
    entry_email.delete(0, tk.END)


# Fungsi untuk ekspor data ke file Excel
def ekspor_ke_excel():
    if len(data["Nama"]) == 0:
        messagebox.showwarning("Peringatan", "Tidak ada data untuk diekspor!")
        return

    df = pd.DataFrame(data)
    df.to_excel("dataEntry.xlsx", index=False)
    messagebox.showinfo("Sukses", "Data berhasil diekspor ke output.xlsx")


# Inisialisasi data
data = {
    "Nama": [],
    "Usia": [],
    "Email": []
}

# Setup GUI
root = tk.Tk()
root.title("Form Input Data Diri")

# Frame untuk Form Input
frame_input = tk.Frame(root, padx=10, pady=10)
frame_input.pack()

tk.Label(frame_input, text="Nama:").grid(row=0, column=0)
entry_nama = tk.Entry(frame_input, width=30)
entry_nama.grid(row=0, column=1)

tk.Label(frame_input, text="Usia:").grid(row=1, column=0)
entry_usia = tk.Entry(frame_input, width=30)
entry_usia.grid(row=1, column=1)

tk.Label(frame_input, text="Email:").grid(row=2, column=0)
entry_email = tk.Entry(frame_input, width=30)
entry_email.grid(row=2, column=1)

# Tombol Tambah Data
tombol_tambah = tk.Button(frame_input, text="Tambah Data", command=tambah_data)
tombol_tambah.grid(row=3, column=0, columnspan=2, pady=10)

# Frame untuk Tabel Data
frame_tabel = tk.Frame(root, padx=10, pady=10)
frame_tabel.pack()

tk.Label(frame_tabel, text="Nama").grid(row=0, column=0)
tk.Label(frame_tabel, text="Usia").grid(row=0, column=1)
tk.Label(frame_tabel, text="Email").grid(row=0, column=2)

listbox_nama = tk.Listbox(frame_tabel, width=30)
listbox_nama.grid(row=1, column=0)

listbox_usia = tk.Listbox(frame_tabel, width=10)
listbox_usia.grid(row=1, column=1)

listbox_email = tk.Listbox(frame_tabel, width=30)
listbox_email.grid(row=1, column=2)

# Tombol Ekspor ke Excel
tombol_ekspor = tk.Button(
    root, text="Ekspor ke Excel", command=ekspor_ke_excel)
tombol_ekspor.pack(pady=10)

# Jalankan Aplikasi
root.mainloop()
