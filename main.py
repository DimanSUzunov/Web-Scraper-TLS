"""
UPDATED
"""
import tkinter as tk
from tkinter import simpledialog, messagebox
from urllib.parse import urlparse
import os
from typing import Optional

from docx import Document
from fetcher import fetch_and_clean_main
from word_writer import extract_text_elements
import traceback

def fetch_main_to_word(url: str) -> None:
    """
    Fetches the main content from the given URL, processes it, and saves it as a Word document.
    """
    parsed = urlparse(url)
    base_name = os.path.basename(parsed.path.rstrip("/")) or "webpage"
    output_file = f"{base_name}.docx"
    try:
        t3ce_blocks = fetch_and_clean_main(url)
        doc = Document()
        doc.add_heading("Webpagina Inhoud", level=1)
        num_id_counter = 1
        for block in t3ce_blocks:
            num_id_counter = extract_text_elements(doc, block, num_id_counter)
        doc.save(output_file)
        messagebox.showinfo("Voltooid", f"Het Word-bestand is opgeslagen als:\n{os.path.abspath(output_file)}")
    except ValueError as ve:
        messagebox.showerror("Fout", f"HTML-fout: {ve}")
    except RuntimeError as re:
        messagebox.showerror("Fout", f"Netwerkfout: {re}")
    except Exception as e:
        messagebox.showerror("Fout", f"Onbekende fout: {e}\n\n{traceback.format_exc()}")

def get_url_from_user() -> Optional[str]:
    """Shows a dialog to get a URL from the user."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    url = simpledialog.askstring("Invoer", "Voer de URL in om te scrapen:")
    root.destroy()
    return url

def main() -> None:
    """
    Main function to handle GUI input and trigger the scraping and Word export process.
    """
    url = get_url_from_user()
    if not url:
        print("Geen URL opgegeven. Afsluiten.")
        return
    fetch_main_to_word(url)

if __name__ == "__main__":
    main()
