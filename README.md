# Generator Surat Pernyataan (Hotel & Kendaraan Dinas)

Aplikasi Streamlit untuk membuat:
- **Surat Pernyataan Tidak Menggunakan Hotel**
- **Surat Pernyataan Tidak Menggunakan Kendaraan Dinas**

Alur: pilih jenis surat → isi nomor surat, tanggal surat, tanggal perjalanan, dan data pegawai → sistem otomatis mengisi template Word → hasil bisa **diunduh sebagai .docx** atau **.pdf (untuk print)**.

## Struktur folder

```
surat-generator/
├── app.py                     # Aplikasi Streamlit utama
├── templates/
│   ├── template_hotel.docx        # Template surat tidak menggunakan hotel
│   └── template_kendaraan.docx    # Template surat tidak menggunakan kendaraan dinas
├── requirements.txt
├── packages.txt                # dependency sistem (LibreOffice) untuk Streamlit Cloud
└── README.md
```

## Cara kerja template

Template dibuat dengan **docxtpl** (bukan tanda `<...>` polos), karena placeholder `<...>` mudah "pecah" saat diedit di Word sehingga sulit dideteksi otomatis. Sebagai gantinya, placeholder ditulis dengan format Jinja `{{ nama_variabel }}`, contoh isi template:

```
Nomor: {{ nomor_surat }}
Nama       : {{ nama }}
NIP        : {{ nip }}
...
```

Variabel yang tersedia di kedua template:

| Variabel | Keterangan |
|---|---|
| `nomor_surat` | Nomor surat pernyataan |
| `tanggal_surat` | Tanggal surat (format: 27 Juli 2026) |
| `nomor_surat_tugas` | Nomor Surat Tugas (dasar perjalanan dinas) |
| `tanggal_surat_tugas` | Tanggal Surat Tugas |
| `tanggal_perjalanan` | Tanggal pelaksanaan perjalanan dinas |
| `tujuan` | Lokasi tujuan perjalanan dinas |
| `nama` | Nama pegawai |
| `nip` | NIP pegawai |
| `pangkat_gol` | Pangkat/Golongan |
| `jabatan` | Jabatan |
| `unit_kerja` | Unit kerja |
| `kota_ttd` | Kota untuk tanda tangan |

**Ingin memakai template Anda sendiri?** Buka `template_hotel.docx` / `template_kendaraan.docx` di Word, ubah redaksional sesuka Anda, tapi pertahankan placeholder `{{ ... }}` di posisi yang sesuai (atau tambahkan variabel baru — lalu tambahkan input-nya juga di `app.py`).

## Menjalankan secara lokal

```bash
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Buka `http://localhost:8501` di browser.

> Fitur unduh **PDF** membutuhkan LibreOffice (`soffice`) terpasang di sistem. Kalau tidak tersedia, aplikasi tetap berjalan normal dan hanya menyediakan unduhan **.docx** (bisa langsung diprint dari Word).
>
> Install LibreOffice lokal (Ubuntu/Debian): `sudo apt install libreoffice`

## Deploy ke GitHub + Streamlit Community Cloud

1. **Push ke GitHub**
   ```bash
   cd surat-generator
   git init
   git add .
   git commit -m "Initial commit: generator surat pernyataan"
   git branch -M main
   git remote add origin https://github.com/<username>/<nama-repo>.git
   git push -u origin main
   ```

2. **Deploy di Streamlit Cloud**
   - Buka [share.streamlit.io](https://share.streamlit.io) → **New app**
   - Pilih repo GitHub yang baru dibuat, branch `main`, file utama `app.py`
   - Klik **Deploy**
   - File `packages.txt` akan otomatis membuat Streamlit Cloud memasang LibreOffice, sehingga fitur unduh PDF ikut aktif.

## Menambah jenis surat baru

Tambahkan entri baru di dictionary `JENIS_SURAT` pada `app.py`, lalu siapkan file template `.docx` baru di folder `templates/` dengan placeholder `{{ ... }}` yang sama (atau variabel tambahan sesuai kebutuhan).
