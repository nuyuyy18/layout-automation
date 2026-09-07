# Buku Layout — Print Engine (A5 + Margin Word-style)

**Tujuan:** Satu CSS engine untuk semua layout buku cetak. Token terpusat, kelas modular, mudah maintained.

## Struktur File

- `buku-layout.css` — engine CSS terpusat (240 baris)
  - Token: margin, spacing scale, color, font, lineheight.
  - @page: A5 + counter halaman otomatis.
  - Modul: cover, toc, bab, arab, terjemah, daftar, kotak, footnote, guard.
  - Media Screen Preview: simulasi tampilan buku di browser layar.
- `examples/book.html` — map/index navigasi modul.
- `examples/modules/` — file pecahan isi buku per bab.

## Token (ubah di `:root`)

### Kertas (mm, Word-style)
```css
--page-w: 148mm;           /* lebar A5 */
--page-h: 210mm;           /* tinggi A5 */
--margin-top: 18mm;        /* atas */
--margin-outer: 16mm;      /* luar (kanan hal. ganjil) */
--margin-inner: 20mm;      /* dalam (dekat jilid) */
--margin-bottom: 20mm;     /* bawah */
```

### Spasi (4pt base, 9 level, named by role)
Pakai token `var(--space-md)` dst, jangan `gap: 16px`. Nanti mudah ubah ritme spacing global.

### Warna (tint, bukan pure)
- `--ink: #1a1a1a` (bukan pure black #000)
- `--paper: #ffffff` (keep white, tapi bg tint)
- `--accent: #1b5e20` (hijau Islami)
- `--accent-tint: #e8f5e9` (kotak info bg)

### Font
- Body: Calibri/Segoe UI (sans)
- Arab: "Arabic Typesetting" (serif, RTL)
- Jangan italic header (global rule)

## Modul

### 1. `.cover`
Halaman pertama. Full bleed, margin:0, gradient hijau. Anak: `.ornament`, `.judul-arab`, `.judul-latin`, `.subtitle`, `.author`, `.divider`.

### 2. `.toc` (table of contents)
Page break sebelum konten. Heading border, list dot-leaders. Anak: `ul > li > a, .dots, .nomor` (TOC item).

### 3. `.bab` dan `.chapter` (alias)
Halaman bab. Page break sebelum. Anak:
- `.bab-header` (label + arab + judul)
- `h2`, `h3.pasal` (heading)
- `.arab` / `.arabic` (matan RTL)
- `.terjemah` / `.translation` (penjelasan)
- `ol.daftar` / `ol.pasal` (list bernomor)
- `.kotak` (kotak faidah)
- `.fn` / `.footnote` (catatan kaki)

### 4. `.arab` / `.arabic` (matan)
RTL, background warm gray, border: 1px solid var(--rule-soft). `font-family: var(--font-arab)`.

### 5. `.terjemah` / `.translation` (penjelasan)
Text-align: justify. Child: `.label` (bold accent).

### 6. `ol.daftar` / `ol.pasal` (list)
Padding-left dari token. Orphan/widow guard.

### 7. `.kotak` (info box)
Green tint bg + border. Anak: `.kotak-label` (uppercase label).

### 8. `.fn` / `.footnote`
Border-top rule. Font-size small. Anak: `sup` (red number).

### 9. Guard (orphan/widow, page-break-inside:avoid)
`p, li { orphans: 2; widows: 2; }`
