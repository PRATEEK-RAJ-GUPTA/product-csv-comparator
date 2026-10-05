# Product CSV Comparator - Browser Only

## Architecture

This application has NO backend.

Browser:
    |
    |-- reads Yesterday CSV using File API
    |
    |-- reads Today CSV using File API
    |
    |-- parses CSV using JavaScript
    |
    |-- compares products locally
    |
    |-- displays TXT report
    |
    |-- generates XLSX locally using SheetJS
    |
    `-- downloads reports

The CSV contents are not uploaded to the hosting server.

## Files

- index.html
- style.css
- app.js

## Features

- Select two local CSV files
- Detects sku, SKU, or #SKU
- Finds new SKUs
- Finds removed SKUs
- Finds field-level changes
- Counts changed SKUs
- Detects duplicate SKUs
- Displays the TXT report in the browser
- Downloads TXT report
- Downloads Excel report
- Excel contains:
  - Summary
  - New SKUs
  - Removed SKUs
  - Column Changes
  - Detailed Changes
  - Duplicate SKUs
  - Common SKU Summary

## Hosting

Because this is a static website, it can be hosted on:

- GitHub Pages
- Netlify
- Vercel
- Cloudflare Pages
- Any normal static web server

No Python installation or server is required.

## Excel dependency

SheetJS is loaded from its CDN:

https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js

The comparison itself remains local. The CDN is only used to provide the browser-side Excel generation library.

If you require completely offline operation, download/bundle the SheetJS library locally and replace the CDN script in index.html.
