# DOCX Template Updater Tool

A fully client-side web tool that updates DOCX document formatting using a template — without changing any text content. No server, no installs, no accounts needed.

**Live at:** `https://flippurs.github.io/DOCX-Template-Updater-Tool/`

## How It Works

1. Open the web page
2. Drag & drop your **template** DOCX (the modern format you want to apply)
3. Drag & drop one or more **target** DOCX files (the documents to update)
4. Click **Update Documents**
5. Updated files download automatically

All processing happens in your browser — files never leave your computer.

## What Gets Updated (from template)

- ✅ Document styles (paragraph, character, table styles)
- ✅ Page layout (margins, page size, orientation)
- ✅ Headers and footers
- ✅ Theme (fonts, colors)
- ✅ Numbering/list definitions
- ✅ Document settings
- ✅ Document properties/metadata

## What Stays Unchanged (in target)

- ✅ All text content
- ✅ Images and media
- ✅ Tables and their data
- ✅ Body structure

## Hosting (GitHub Pages)

This tool is hosted for free on GitHub Pages. To enable it:

1. Go to your repo **Settings** → **Pages**
2. Under "Source", select **Deploy from a branch**
3. Choose **main** branch and **/ (root)** folder
4. Click **Save**
5. Your site will be live at `https://flippurs.github.io/DOCX-Template-Updater-Tool/`

## Running Locally

Just open `index.html` in any browser. No server needed.

## Tech Stack

- Pure HTML/CSS/JavaScript (no frameworks)
- [JSZip](https://stuk.github.io/jszip/) — unzips/rezips DOCX files in browser
- [FileSaver.js](https://github.com/nicolo-ribaudo/FileSaver.js) — triggers file downloads
- GitHub Pages — free static hosting

## Project Structure

```
├── index.html     # Main page
├── style.css      # Styles
├── app.js         # All logic (drag-drop + DOCX processing)
└── README.md
```
