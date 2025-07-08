# Monthly Media Monitoring Report Compiler

Ini adalah skrip Python untuk secara otomatis mengompilasi laporan media monitoring bulanan dari beberapa file Word menjadi satu file Excel yang rapi dengan beberapa kategori sheet.

## Fitur

- Secara otomatis mencari semua file `.docx` di dalam folder dan sub-folder.
- Mengekstrak data berita dari tabel di dalam file Word.
- Menggabungkan beberapa kategori berita menjadi satu kategori utama.
- Menghasilkan satu file Excel dengan sheet terpisah untuk setiap kategori berita.
- Memberi format pada header (warna kuning, rata tengah) dan sel (wrap text).
- Mengatur lebar kolom secara otomatis.

## Cara Penggunaan

1.  Pastikan semua pustaka yang dibutuhkan sudah terpasang: `pip install pandas python-docx xlsxwriter`
2.  Buka file `monthly_compiler.py`.
3.  Ubah path di variabel `root_folder_path` agar menunjuk ke folder utama bulan yang ingin diproses.
4.  Jalankan skrip: `python monthly_compiler.py`

## Acknowledgements

Skrip ini dikembangkan oleh Hizkia Gerald Garibaldi dengan bantuan dan bimbingan dari asisten AI Google, Gemini.