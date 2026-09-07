# Writing Guide — Buku HTML Layout

Panduan ini untuk menulis isi buku ke sistem layout HTML modular ini.

## 1) Tujuan workflow

- Tulis isi buku di file modular kecil.
- Jaga layout di satu CSS engine.
- Preview di browser via `examples/book.html`.
- Final print tetap dari print/PDF context.

## 2) Struktur proyek

### Engine
- `assets/css/buku-layout.css`
  - Source of truth untuk ukuran kertas, margin, font, warna, heading, arabic block, kotak, table, checklist, dan guard.

### Preview
- `examples/book.html`
  - Loader preview berbasis `fetch()`.
  - Buka lewat dev server, bukan `file://`.

### Modul konten
- `examples/modules/cover.html`
- `examples/modules/toc.html`
- `examples/modules/bab-01.html`
- `examples/modules/bab-02.html`
- `examples/modules/bab-03.html`

### Dokumentasi
- `docs/system_map.md`
- `docs/book_map.md`
- `docs/writing_guide.md` ini

## 3) Istilah baku

Pakai istilah ini konsisten:

- **Book / Buku** = keseluruhan proyek.
- **Front matter** = cover, daftar isi, pengantar.
- **Bab / Chapter** = unit utama isi.
- **Subbab / Section** = bagian dalam bab.
- **Pasal / Subsection kecil** = bagian lebih kecil di bawah subbab.
- **Matan / Arabic block** = teks Arab utama.
- **Terjemah / Translation** = penjelasan Indonesia.
- **Kotak / Callout** = faidah, dalil, catatan.
- **Checklist** = tabel ibadah harian.
- **Footnote** = catatan kaki.

## 4) Format penulisan modular

### Aturan dasar
- Satu bab = satu file modul.
- Jangan gabung semua bab ke satu file besar.
- Satu blok isi = satu elemen HTML yang jelas.
- Jangan taruh style inline kalau bisa ditaruh di CSS.

### Struktur file bab
Gunakan struktur seperti ini:

```html
<section class="chapter" id="bab1">
  <div class="bab-header">
    <div class="bab-label">Bab Pertama</div>
    <span class="bab-arab">الطَّهَارَةُ</span>
    <div class="bab-judul">Pengertian Thaharah</div>
  </div>

  <h3 class="pasal" id="s1-1">Definisi dan Dalil</h3>

  <div class="arab">...</div>
  <p class="terjemah">...</p>
  <div class="kotak">...</div>
</section>
```

## 5) Blok yang didukung

### Cover
Pakai `section.cover`.

Contoh isi:
- `.ornament`
- `.judul-arab`
- `.judul-latin`
- `.subtitle`
- `.author`
- `.divider`

### Daftar isi
Pakai `nav.toc`.

Setiap item:
- judul bab
- `span.dots`
- nomor halaman

### Bab
Pakai `section.chapter` atau `section.bab`.

Didalamnya boleh ada:
- `.bab-header`
- `h2` / `h3.pasal`
- `.arab` / `.arabic`
- `.terjemah` / `.translation`
- `ol.daftar` / `ol.pasal`
- `.kotak`
- `.fn` / `.footnote`
- `.checklist-table`

### Arabic block
Pakai `.arab` atau `.arabic`.

Aturan:
- isi Arab berdiri sendiri
- jangan gabung dengan kalimat Indonesia dalam satu baris
- kalau perlu penjelasan, taruh di elemen terpisah sesudahnya

### Terjemah
Pakai `.terjemah` atau `.translation`.

Aturan:
- paragraf biasa
- boleh ada `b`, `i`, `sup`
- jangan dicampur dengan Arab dalam elemen yang sama

### Kotak
Pakai `.kotak`.

Cocok untuk:
- dalil
- faidah
- catatan penting
- warning

### Checklist ibadah
Pakai `table.checklist-table`.

Aturan:
- cocok untuk tabel sederhana A5
- batasi isi agar tidak overflow
- kalau terlalu panjang, pecah jadi beberapa tabel / beberapa halaman

## 6) Urutan kerja menulis

### Workflow standar
1. Tentukan bab baru.
2. Buat file baru di `examples/modules/`.
3. Isi dengan HTML modular sesuai kelas layout.
4. Daftarkan file itu di `examples/book.html`.
5. Update `docs/book_map.md`.
6. Preview via dev server.
7. Revisi per modul, bukan per buku penuh.

### Contoh tambah bab baru
Buat file:
- `examples/modules/bab-04.html`

Isi minimal:
```html
<section class="chapter" id="bab4">
  <div class="bab-header">
    <div class="bab-label">Bab Keempat</div>
    <span class="bab-arab">...</span>
    <div class="bab-judul">Judul Bab</div>
  </div>

  <h3 class="pasal">Subbab</h3>
  <div class="arab">...</div>
  <p class="terjemah">...</p>
</section>
```

Lalu tambahkan urutannya di `examples/book.html`.

## 7) Aturan layout

### Wajib
- Semua halaman pakai satu engine CSS.
- Semua ukuran utama pakai token.
- Teks Arab dan Indonesia tidak satu baris.
- Ganti bab = ganti halaman berikutnya.
- Tabel lebar harus aman di A5.

### Jangan
- Jangan inline CSS untuk styling utama.
- Jangan gabung banyak bab ke satu file HTML besar.
- Jangan paksa tabel besar masuk satu halaman kalau hasilnya rusak.
- Jangan pakai `file://` untuk preview modular yang memakai `fetch()`.

## 8) Preview lokal

Jalankan:

```bash
node dev-server.mjs
```

Buka:

```text
http://localhost:8000/examples/book.html
```

Kenapa:
- `fetch()` modul butuh HTTP server.
- `file://` sering gagal load modul.

## 9) Pola penulisan yang aman

### Baik
- 1 bab = 1 file
- 1 pasal = 1 blok
- 1 ide = 1 elemen HTML
- kalimat Arab berdiri sendiri
- penjelasan Indonesia berdiri sendiri

### Kurang baik
- satu file sangat panjang
- inline style banyak
- tabel terlalu padat
- Arab dan Indonesia digabung dalam satu baris

## 10) Checklist sebelum submit konten

- [ ] File modul sudah benar tempatnya.
- [ ] `id` bab sudah unik.
- [ ] Judul bab ada di `bab-header`.
- [ ] Arab dan Indonesia dipisah elemen.
- [ ] Tabel/checklist tidak overflow.
- [ ] `examples/book.html` sudah daftar modul baru.
- [ ] `docs/book_map.md` sudah diupdate.
- [ ] Preview di browser sudah cek rapi.

## 11) Default template singkat

```html
<section class="chapter" id="babX">
  <div class="bab-header">
    <div class="bab-label">Bab ...</div>
    <span class="bab-arab">...</span>
    <div class="bab-judul">...</div>
  </div>

  <h3 class="pasal">...</h3>
  <div class="arab">...</div>
  <p class="terjemah">...</p>
  <div class="kotak">...</div>
</section>
```

