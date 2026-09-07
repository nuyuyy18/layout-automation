# Book Map

## Source of truth
- `assets/css/buku-layout.css` — layout engine for both screen preview and PDF print.

## Module order
1. `examples/modules/cover.html`
2. `examples/modules/toc.html`
3. `examples/modules/bab-01.html`
4. `examples/modules/bab-02.html`
5. `examples/modules/bab-03.html`

## Current preview entry
- `examples/book.html`
  - Placeholder include markers for module-first workflow.
  - Use this file as navigation map while editing modules one by one.

## Working rule
- Edit one module at a time.
- Keep shared styling in CSS only.
- Keep module files content-only.
- Add new chapters as `examples/modules/bab-04.html`, `bab-05.html`, etc.

## Local dev server

- `node dev-server.mjs`
- Open `http://localhost:8000/examples/book.html`
- Use this route so `fetch()` can load modules normally.

## Author guide

- `docs/writing_guide.md`
  - Full guide for writing new content modules.
