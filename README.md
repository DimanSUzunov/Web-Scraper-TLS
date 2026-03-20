# TUDelftwebtoword

A Python tool to scrape main content from a TU Delft web page and export it to a formatted Word document. Includes a simple GUI for URL input.

## Features
- Scrapes and cleans main content from a given URL
- Exports content to a Word (.docx) file
- Handles images, tables, lists, and notices
- Simple Tkinter-based GUI

## Requirements
- Python 3.7+
- See `requirements.txt` for dependencies

## Installation
1. Clone or download this repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Run the main script:
   ```bash
   python main.py
   ```
2. Enter the URL of the TU Delft web page you want to scrape when prompted.
3. The resulting Word file will be saved in the current directory.

## Project Structure
- `main.py` — GUI and entry point
- `fetcher.py` — Fetches and cleans HTML
- `word_writer.py` — Converts HTML to Word
- `html_utils.py` — HTML-related helpers (e.g., image handling)
- `config.py` — Centralized configuration/constants

## Customization
Edit `config.py` to change selectors, styles, or other constants.

## License
MIT License

