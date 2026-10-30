# Generator Dokumen — CV. Usaha Mandiri Lift

Web statis untuk mengisi template Word dari ponsel. Saat ini: **Invoice**.
Semua berjalan di browser; tidak ada data yang dikirim ke server.

## Pasang di GitHub Pages (gratis)
1. Buat repository baru di github.com (misal `dokumen`), Public.
2. **Add file → Upload files**, unggah 4 file: `index.html`, `docgen.js`, `templates.js`, `jszip.min.js` → **Commit changes**.
3. **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `(root)` → Save.
4. Tunggu 1–2 menit, alamatnya: `https://USERNAME.github.io/dokumen/`. Di ponsel, pilih "Tambahkan ke layar utama".

## Cara pakai
Isi Data umum, lalu satu kartu per tempat (nomor PO, tanggal PO, barang, harga). Nomor invoice otomatis `TTBBHH` + nomor tempat. Tekan **Unduh**.
Template Invoice maksimal 5 baris barang.
