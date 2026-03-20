"""
UPDATED
Configuration constants for the TUDelftwebtoword project.
"""

# CSS selectors for elements to remove from the main content
UNWANTED_SELECTORS = ['.nav-aside', '.hiddenSidebarContent', 'nav', 'footer', 'aside']

# Selector for top-level content blocks
T3CE_BLOCK_SELECTOR = 'div.t3ce:not(div.t3ce div.t3ce)'

# Default document heading
DEFAULT_HEADING = "Webpagina Inhoud"

# Default image width (in inches)
DEFAULT_IMAGE_WIDTH = 4

# Notice colors
NOTICE_COLORS = {
    'notice--blue_lighter': '0070C0',
    'notice--yellow': 'FFC000',
    'notice--red': 'FF0000',
}

