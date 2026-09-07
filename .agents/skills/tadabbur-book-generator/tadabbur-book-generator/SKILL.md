---
name: tadabbur-book-generator
description: Generates high-converting, deeply reflective Islamic guided journals, Tadabbur e-books, and personal development workbooks using the 7-Pillar Modular Architecture (Valid Quran Verses, Sahih Supplications, Emotional Storytelling, Neurobiology/Psychology Science, Mindful Dialogue, 1x Action Checklist, Peer-Reviewed Research with active DOIs, Interactive Journaling, and Quote of the Day) in Markdown and print/mobile-ready monochrome A5 DOCX format using python-docx. Use when the user asks to write, draft, or generate Tadabbur books, Islamic psychology chapters, guided journals, or structured reflective workbooks.
---

# Tadabbur Book Generator Skill (7-Pillar Architecture)

Skill ini digunakan untuk menyusun **Buku Tadabbur Interaktif, Guided Journal, & E-Book Pengembangan Diri Berbasis Islam-Sains** dengan standar editorial profesional, validasi riset ilmiah internasional, dan format tata letak A5 hitam-putih (*monochrome clean editorial*) yang dioptimalkan untuk layar HP & Tablet.

---

## 🏛️ 7-Pillar Master Architecture

Setiap bab disusun secara modular dengan urutan psikologis berikut:

```
+-------------------------------------------------------------------+
| 1. [ Ayat Pijakan & Doa Hati ]                                    |
|    - Ayat Al-Qur'an (Teks Arab, Terjemahan, Nomor Surah:Ayat)      |
|    - Doa Ma'tsur / Shahih (Teks Arab, Latin, Terjemahan, Sumber)  |
+-------------------------------------------------------------------+
| 2. [ Catatan Hati ] (Emotional Storytelling & Vulnerability)       |
|    - Scene naratif personal, gejolak emosi, dilema manusiawi       |
|    - Jembatan refleksi menuju ketenangan Ilahi                     |
+-------------------------------------------------------------------+
| 3. [ Sains Jiwa / Sains Ketahanan ] (Neurobiology & Psychology)   |
|    - Penjelasan biologis (Dopamin, Amigdala, Prefrontal Cortex)    |
|    - Mekanisme kognitif (Meaning-Focused Coping, PTG, Grit)       |
+-------------------------------------------------------------------+
| 4. [ Dialog Mindful ] (Somatic Grounding & Meditation)            |
|    - Callout box: Tarik napas, relaksasi bahu, afirmasi batin     |
+-------------------------------------------------------------------+
| 5. [ Amal Tracker ] (1x Daily Action Checklist)                   |
|    - 5 Kotak centang aksi harian tunggal [  ] (Bukan tabel 7 hari)|
+-------------------------------------------------------------------+
| 6. [ Pojok Riset ] (Empirical Peer-Reviewed Citations)            |
|    - Sitasi jurnal primer (Penulis, Tahun, Jurnal, DOI aktif)      |
|    - Metadata diverifikasi & diunduh ke lokal                      |
+-------------------------------------------------------------------+
| 7. [ Kolom Interaktif & Quote of the Day ]                        |
|    - Guided journaling prompts (Ruang kosong/titik-titik isian)   |
|    - Quote of the Day (Kutipan penutup yang berkesan & mendalam)  |
+-------------------------------------------------------------------+
```

---

## 📐 Standar Layout Dokumen A5 (.docx)

* **Ukuran Kertas**: A5 ($14.8\text{ cm} \times 21.0\text{ cm}$).
* **Margin (Binding & Digital View)**:
  * Inside / Gutter (Kiri): $2.2\text{ cm}$
  * Outside (Kanan): $1.8\text{ cm}$
  * Top / Bottom: $2.0\text{ cm}$
* **Palet Warna**: **Murni Hitam (`#000000`)** (*Clean Monochrome / Anti-AI Slop*).
* **Tipografi**:
  * **Font Family**: **Georgia** atau **Garamond** (Serif elegan & mudah dibaca).
  * **H1 (Judul Bab)**: $15\text{--}18\text{ pt}$ Bold, Hitam, **Wajib Halaman Baru (*Page Break Before*)**.
  * **Judul Sub-Bab**: $14\text{--}15\text{ pt}$ Bold, Hitam.
  * **Pillar Tag (`[ Catatan Hati ]`)**: $13\text{ pt}$ Bold, Hitam.
  * **Body Text**: $12\text{ pt}$, *Line Spacing* $1.25\times$, Indent Paragraf Pertama $0.5\text{ cm}$ (Ukuran ramah baca HP/Tablet).
  * **Callout Box**: Background shading netral tipis (`#FAFAFA`), Border kiri $1.5\text{ pt}$ solid hitam (`#000000`), teks $11\text{ pt}$ Italic hitam.
  * **Checklist**: Format tanda kurung siku `[  ]`, label tebal $11.5\text{ pt}$ hitam.

---

## 🔒 Protokol Integritas Ilmiah & Keagamaan (Wajib Dipatuhi)

1. **Verifikasi Ayat & Hadits**:
   * Teks Arab, terjemahan, dan perawi/surah harus shahih dan presisi. Dilarang mengarang teks dalil.
2. **Verifikasi Riset & DOI**:
   * Dilarang membuat sitasi palsu/fiktif.
   * Jurnal harus berasal dari penerbit terindeks (*CellPress, PubMed, APA, ScienceDirect, Nature, Routledge*).
   * Cari DOI resmi, unduh metadata JSON melalui API CrossRef ke folder scratch/lokal sebelum menulis bab.

---

## ⚡ Workflow Penulisan Bab Baru

1. **Tentukan Topik & Tema Bab** (misal: Tokoh Nabi / Fenomena Kejiwaan).
2. **Riset & Unduh Jurnal Primer** (Dapatkan DOI valid, simpan metadata di lokal).
3. **Pilih Ayat Al-Qur'an & Doa Ma'tsur yang Relevan**.
4. **Tulis Naskah Lengkap Berdasarkan 7-Pillar Architecture**.
5. **Jalankan Generator Python (`scripts/build_docx.py`)** untuk membuat file `.docx` siap pakai di folder dokumen.
