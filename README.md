# ppt-to-pdf

A Node.js tool that converts PowerPoint presentations (`.pptx` / `.ppt`) into compact, multi‑slide PDFs using LibreOffice in headless mode. Slides are arranged in a newspaper‑style grid (2, 3, 4, 6, or 9 per page).

## What it does
- Upload a `.pptx` or `.ppt` file via the browser
- Pick how many slides you want per page (1‑9)
- Get back a PDF with slides packed tightly on each page
- Works on Windows, macOS, and Linux

## Requirements
- [Node.js](https://nodejs.org/) (16 or later)
- [LibreOffice](https://www.libreoffice.org/download/) (24.8 or later recommended)

## Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/your-username/ppt-to-pdf.git
cd ppt-to-pdf
