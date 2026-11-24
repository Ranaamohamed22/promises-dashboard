"""Brand-aware palettes for the isolated Streamlit design preview.

Presentation only: never used to select a report data source or change
business rules. The soft charcoal mode keeps the active bank's brand accent.
"""

BANK_PALETTES = {
    "SNB": {
        "accent": "#07883E",
        "accent_hover": "#056C32",
        "accent_soft": "#E8F7EE",
        "accent_text": "#056D35",
        "background": "#F3F7FC",
        "surface": "#FFFFFF",
        "surface_alt": "#F2F7FC",
        "text": "#102B45",
        "muted": "#51667C",
        "border": "#CEDCEB",
        "neutral_chart": "#BBC9D7",
        "grid": "#DFE9F2",
        "sidebar": "#092541",
    },
    "STC": {
        "accent": "#6A35B7",
        "accent_hover": "#54299A",
        "accent_soft": "#F2EAFD",
        "accent_text": "#56279E",
        "background": "#F7F4FC",
        "surface": "#FFFFFF",
        "surface_alt": "#F8F4FF",
        "text": "#281C43",
        "muted": "#605779",
        "border": "#DCD0EC",
        "neutral_chart": "#C9BDD9",
        "grid": "#EBE3F5",
        "sidebar": "#2A1A4F",
    },
}

THEME_OPTIONS = ("حسب البنك", "أسود هادي")


def resolve_theme(bank, mode="حسب البنك"):
    """Get a fresh palette for the selected bank and optional charcoal mode."""
    if bank not in BANK_PALETTES:
        raise ValueError("Unknown bank; no business logic or palette guessing.")
    if mode not in THEME_OPTIONS:
        raise ValueError("Unknown theme option.")
    palette = dict(BANK_PALETTES[bank])
    if mode == "أسود هادي":
        palette.update({
            "background": "#303741",
            "surface": "#3C4651",
            "surface_alt": "#36404B",
            "text": "#F4F8FC",
            "muted": "#D4DFE9",
            "border": "#657385",
            "neutral_chart": "#9DADBE",
            "grid": "#596677",
            "sidebar": "#272D35",
            "accent_soft": "#3E4A54",
            "accent_text": "#FFFFFF",
        })
    return palette
