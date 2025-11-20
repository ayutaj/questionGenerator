import re
def to_camel_case(text: str) -> str:
    # Remove all non-alphanumeric characters and split into words
    words = re.split(r'[^A-Za-z0-9]+', text.strip())
    words = [w for w in words if w] 

    if not words:
        return ""

    # First word lowercase + remaining words capitalized
    first = words[0].lower()
    rest = [w.capitalize() for w in words[1:]]

    return first + "".join(rest)


# Famous Font Names (Reference)
# --------------------------------
# Serif (Traditional / Book Style):
#   - Times New Roman
#   - Georgia
#   - Garamond
#   - Cambria
#   - Palatino Linotype
#   - Baskerville
#   - Didot
#   - Book Antiqua
#   - Constantia
#
# Sans-Serif (Modern / Clean):
#   - Arial
#   - Helvetica
#   - Calibri
#   - Segoe UI
#   - Verdana
#   - Tahoma
#   - Roboto
#   - Open Sans
#   - Lato
#   - Futura
#   - Montserrat
#
# Display / Heading Fonts (Bold / Stylish):
#   - Impact
#   - Arial Black
#   - Franklin Gothic Heavy
#   - Rockwell
#   - Bebas Neue
#   - Oswald
#   - Anton
#   - Playfair Display
#
# Monospace (Coding / Technical):
#   - Courier New
#   - Consolas
#   - Fira Code
#   - JetBrains Mono
#   - Source Code Pro
#
# Script / Handwriting:
#   - Brush Script MT
#   - Pacifico
#   - Dancing Script
#   - Great Vibes
#   - Lobster
