# DOCX Template Updater Tool

A web-based tool that updates DOCX document formatting using a template — without changing any text content.

## How It Works

1. Open the web page
2. Drag & drop your **template** DOCX (the modern format you want to apply)
3. Drag & drop one or more **target** DOCX files (the documents to update)
4. Click **Update Documents**
5. Download the updated files

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

## Setup & Deployment

### Prerequisites

- Node.js 18+
- Firebase CLI: `npm install -g firebase-tools`
- A Firebase project (free Spark plan works)

### Initial Setup

```bash
# Login to Firebase
firebase login

# Update .firebaserc with your project ID
# Replace "your-project-id" with your actual Firebase project ID

# Install function dependencies
cd functions
npm install
cd ..
```

### Local Testing

```bash
firebase emulators:start
```

Then open `http://localhost:5000` in your browser.

### Deploy

```bash
firebase deploy
```

This deploys both the frontend (Firebase Hosting) and backend (Cloud Functions).

## Project Structure

```
├── firebase.json          # Firebase configuration
├── .firebaserc            # Firebase project link
├── public/                # Frontend (Firebase Hosting)
│   ├── index.html         # Main page
│   ├── style.css          # Styles
│   └── app.js             # Client-side logic
└── functions/             # Backend (Cloud Functions)
    ├── package.json       # Dependencies
    ├── index.js           # Express API setup
    └── docxUpdater.js     # Core DOCX update logic
```

## Limits

- Max file size: 50MB per file
- Max batch: 50 files at once
- Firebase Cloud Functions free tier: 2 million invocations/month

## Future Enhancements

- Support for XLSX and PPTX templates
- Network/SharePoint path support
- User authentication
- Processing history/logs
