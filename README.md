# PDF Generator with Customizable Text Overlay

## Overview
This Python application generates personalized PDF files based on a template PDF and entries from a CSV file. The generated PDFs include custom text that can be formatted and positioned dynamically.

## Features
- Select a template PDF, CSV file, and output directory via a simple GUI.
- Customize the output file names with a user-defined tag.
- Format text dynamically using placeholders: `{name}`, `{surnames}`, `{amount}`.
- Specify font type, size, text position (X and Y in %), and maximum line width.
- Generate one PDF per entry in the CSV file.

## Requirements
- Python 3.7+
- Dependencies listed in `requirements.txt`.

## Installation
1. Clone this repository:
   ```bash
   git clone <repository_url>
   cd pdf_generator_project

