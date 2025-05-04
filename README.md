# Legacy Sheet Template

A modern web application for converting genealogical PDF data into beautifully formatted Legacy Sheets.

## Overview

This project provides a precise US Letter (8.5×11") template for genealogical documents with professional styling and formatting, designed to streamline the conversion of raw PDF data into visually appealing documents.

## Technology Stack

- **Vite**: Fast, modern frontend build tool
- **Vanilla JavaScript**: Clean, no-framework approach
- **CSS3**: Advanced styling with custom properties and media queries
- **Google Fonts**: Merriweather and Source Sans Pro for professional typography

## Features

- ✅ US Letter (8.5×11") document format with proper margins
- ✅ Professional typography with serif headers and consistent text styling
- ✅ Circular numbered bullets in brown (#8B4513) with support for large numbers
- ✅ Support for person images with 20% rounded corners
- ✅ Family structure formatting with proper indentation
- ✅ Person-to-person hyperlinking
- ✅ Print-ready styling with proper page breaks

## Getting Started

### Prerequisites

- Node.js 16.0 or higher

### Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/duresha/legacy-sheet-template.git
   cd legacy-sheet-template
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm run dev
   ```

4. Open your browser to the URL shown in your terminal (typically http://localhost:5173)

## Building for Production

```bash
npm run build
```

The built files will be in the `dist` directory, ready for deployment.

## Repository Structure

```
legacy-sheet-template/
├── public/             # Static assets
│   └── images/         # Sample images
├── src/
│   ├── index.html      # Main HTML template
│   ├── style.css       # CSS styling
│   └── main.js         # JavaScript functionality
├── package.json        # Project dependencies
└── vite.config.js      # Vite configuration
```

## License

ISC