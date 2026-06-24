# GeminiOut

GeminiOut is a browser-only tool for turning Gemini chat exports and shared-link screenshots into readable Word documents.

## What it does

- Import Gemini JSON, Markdown, TXT, or HTML files.
- Preview chats with Markdown and KaTeX math rendering.
- Export selected chats to `.docx`.
- OCR screenshots locally to find `gemini.google.com/share/...` links.
- Run fully in the browser. No backend server is required.

## Privacy model

Files and screenshots are processed in your browser. The static website does not upload your chat files, images, or generated Word document to a server.

## Development

```bash
npm install
npm run dev
```

## Build

```bash
npm run lint
npm run build
```

## GitHub Pages

This repository includes a GitHub Actions workflow that publishes `dist/` to GitHub Pages after every push to `main`.
