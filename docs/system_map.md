# System Map

## Layout system

- `assets/css/buku-layout.css`
  - Source of truth for print tokens, `@page`, and all layout modules.
  - Includes `@media screen` book preview formatting.

## Modules (Modular Book Content)

- `examples/book.html`
  - Web map/index preview page.
- `examples/modules/`
  - `cover.html` — Book Cover.
  - `toc.html` — Table of Contents.
  - `bab-01.html` — Chapter 1: Pengertian Thaharah.
  - `bab-02.html` — Chapter 2: Wudhu.

## Docs

- `docs/buku-layout.md`
  - Implementation notes, token guide, module guide.
- `docs/system_map.md`
  - This file. Fast route map for maintainers.
- `docs/book_map.md`
  - Module order and run flow.
- `docs/writing_guide.md`
  - Main writing guide for authors.
