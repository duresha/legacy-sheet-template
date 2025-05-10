# Legacies - Genealogy Document Generator

A modern web application for extracting genealogical data from PDFs and converting it into beautifully formatted Legacy Sheets.

## Overview

This project provides a precise US Letter (8.5×11") template for genealogical documents with professional styling and formatting, designed to streamline the conversion of raw PDF data into visually appealing documents.

## Technology Stack

- **Vite**: Fast, modern frontend build tool
- **Vanilla JavaScript**: Clean, no-framework approach
- **CSS3**: Advanced styling with custom properties and media queries
- **Google Fonts**: Public Sans, Nunito Sans, and custom Roca Two Bold for professional typography
- **PDF.js**: For PDF text extraction and parsing
- **html2canvas**: For capturing the rendered template as an image
- **jsPDF**: For generating PDFs from the captured template
- **FileSaver.js**: For triggering browser download animations

## Features

- ✅ Two-panel interface with PDF upload and template preview
- ✅ PDF upload via drag & drop or file picker
- ✅ PDF text extraction and structured data parsing
- ✅ Real-time extraction progress visualization
- ✅ US Letter (8.5×11") document format with proper margins
- ✅ Professional typography with custom fonts and consistent styling
- ✅ Circular numbered bullets in brown (#a85733) with proper formatting
- ✅ Support for person images with rounded corners
- ✅ Family structure formatting with proper indentation
- ✅ Person-to-person hyperlinking
- ✅ Apply/Revert template functionality
- ✅ One-click PDF export with download animation
- ✅ Responsive design for different screen sizes
- ✅ Print-ready styling with proper page breaks
- **OCR Integration**: Extracts text from uploaded genealogy images
- **Smart Parsing**: Identifies person entries, names, and relationships
- **Interactive Editor**: Edit extracted content before generating the final document
- **PDF Generation**: Save the finalized legacy sheet as a PDF file
- **Responsive Design**: Works on both desktop and mobile devices
- **Google Fonts**: Public Sans, Nunito Sans, and custom Roca Two Bold for professional typography
- **Keyboard Shortcuts**: Use the Enter key to quickly move to the next person entry

## Getting Started

### Prerequisites

- Node.js 16.0 or higher

### Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/legacies.git
   cd legacies
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev
   ```

4. Open your browser to the URL shown in your terminal (typically http://localhost:3000)

## How to Use

1. **Upload PDF**: Use the left panel to upload a genealogy PDF via drag & drop or file picker
2. **Review Extracted Data**: The app will extract structured genealogy data and display it in a preview
3. **Apply to Template**: Click "Apply to Template" to populate the right panel with the extracted data
4. **Customize if Needed**: The template will update with the extracted data, maintaining proper formatting
5. **Export as PDF**: Click "Save PDF" to export the final document with all formatting preserved

## Building for Production

```bash
npm run build
```

The built files will be in the `dist` directory, ready for deployment.

## Repository Structure

```
legacies/
├── public/               # Static assets
│   ├── fonts/            # Custom fonts
│   │   └── FontsFree-Net-Roca-Two-Bold.ttf
│   └── images/           # Sample images and icons
│       ├── upload-icon.svg
│       ├── pdf-icon.svg
│       └── ehlers.jpg
├── src/
│   ├── index.html        # Main HTML template
│   ├── style.css         # CSS styling
│   ├── main.js           # JavaScript functionality
│   └── pdf.worker.min.mjs # PDF.js worker script
├── package.json          # Project dependencies
└── vite.config.js        # Vite configuration
```

## License

ISC
