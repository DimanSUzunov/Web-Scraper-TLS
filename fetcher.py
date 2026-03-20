"""
UPDATED
"""

from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
from config import UNWANTED_SELECTORS, T3CE_BLOCK_SELECTOR

class FAQ:
    def __init__(self, title, url):
        self.title = title
        self.url = url


class Button:
    def __init__(self, text, url):
        self.text = text
        self.url = url

class Image:
    def __init__(self, url):
        self.url = url


def fetch_and_clean_main(url: str):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Network error while fetching URL: {e}")

    soup = BeautifulSoup(response.content, "html.parser")
    main_elem = soup.select_one('main#main')
    if not main_elem:
        raise ValueError("Geen <main> gevonden in de pagina")

    # Remove unwanted layout sections
    for selector in UNWANTED_SELECTORS:
        for tag in main_elem.select(selector):
            tag.decompose()

    t3ce_blocks = main_elem.select(T3CE_BLOCK_SELECTOR)

    # Extract FAQ links and attach to blocks
    faq_list = []
    
    for block in t3ce_blocks:
        faq_cards = block.select("a.card")  # select the FAQ links in this block
        faqs = []
        for card in faq_cards:
            href = card.get("href")
            title_elem = card.find("h3")
            if href and title_elem:
                full_url = urljoin(url, href)
                title = title_elem.get_text(strip=True)
                faqs.append(FAQ(title, full_url))
        block.faqs = faqs  # attach FAQ list to block

    # Extract buttons and attach them to their blocks
    for block in t3ce_blocks:

        buttons = []

        # Find <a> buttons
        for btn in block.select("a.button, a.btn, a.elementor-button"):
            text = btn.get_text(strip=True)
            href = btn.get("href")

            if text and href:
                full_url = urljoin(url, href)
                buttons.append(Button(text, full_url))

        # Find <button onclick="">
        for btn in block.select("button"):
            text = btn.get_text(strip=True)
            onclick = btn.get("onclick")

            if onclick and "location.href" in onclick:
                try:
                    link = onclick.split("'")[1]
                    full_url = urljoin(url, link)
                    buttons.append(Button(text, full_url))
                except IndexError:
                    pass

        # Attach buttons to block
        block.buttons = buttons
    
    for block in t3ce_blocks:
        images = []

        # Normal images (including svg)
        for img in block.select("img"):
            src = img.get("src") or img.get("data-src") or img.get("data-lazy-src")
            if src:
                full_url = urljoin(url, src)
                images.append(Image(full_url))

        # Inline SVGs
        for svg in block.select("svg"):
            images.append(Image(str(svg)))

        block.images = images


    return t3ce_blocks
