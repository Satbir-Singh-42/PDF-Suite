# PDF Converter & Merger

A professional desktop application for converting various file formats to PDF and merging existing PDF documents. Built with PyQt5 for a clean, cross‑platform GUI.

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Python: 3.8+](https://img.shields.io/badge/Python-3.8%2B-green.svg)

## Features

- **Convert to PDF** from:
  - Images: JPG, JPEG, PNG, BMP, GIF, TIFF
  - Microsoft Word documents: DOC, DOCX, RTF *(Windows only)*
  - Plain text files: TXT and other text-based files
- **Merge multiple PDFs** into a single document (with drag‑and‑drop style reordering)
- **Batch conversion** – add many files at once
- **Lossless image handling** – preserve original image quality with `img2pdf`
- **JPEG quality control** – adjust compression for image‑based PDFs
- **Optional merged output** – directly combine all converted files into one PDF
- **Overwrite protection** – configurable warning before replacing existing files
- **Automatic open** – optionally open the resulting PDF after operation
- **Clean, responsive UI** with progress feedback and cancellation support

## Screenshots

*(Add screenshots of the three tabs here: File Conversion, PDF Merge, Settings & About)*

## Installation

### Prerequisites
- Python 3.8 or newer
- **Windows only:** Microsoft Word (for DOC/DOCX conversion)

- pip install -r requirements.txt