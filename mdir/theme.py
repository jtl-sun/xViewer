from __future__ import annotations

import os
from typing import Mapping

THEME_CHOICES = ("Bright", "Dark", "System")

LIGHT_PALETTE: dict[str, str] = {
    "name": "Bright",
    "window": "#f0f0f0",
    "panel": "#f6f6f6",
    "field": "#ffffff",
    "foreground": "#202020",
    "muted": "#666666",
    "button": "#f4f4f4",
    "button_active": "#e7e7e7",
    "border": "#cfcfcf",
    "divider": "#dedede",
    "selection": "#0a64ad",
    "selection_foreground": "#ffffff",
    "header": "#f1f1f1",
    "header_foreground": "#222222",
    "header_border": "#bdbdbd",
    "canvas_margin": "#ffffff",
    "path_bg": "#0b6f3c",
    "path_border": "#174b34",
    "path_focus": "#36c9c6",
    "path_foreground": "#ffffff",
    "path_hover": "#d7fff0",
}

DARK_PALETTE: dict[str, str] = {
    "name": "Dark",
    "window": "#1e1e1e",
    "panel": "#252526",
    "field": "#2d2d30",
    "foreground": "#f2f2f2",
    "muted": "#aaaaaa",
    "button": "#333337",
    "button_active": "#414146",
    "border": "#484848",
    "divider": "#383838",
    "selection": "#0e639c",
    "selection_foreground": "#ffffff",
    "header": "#2b2b2b",
    "header_foreground": "#f1f1f1",
    "header_border": "#4a4a4a",
    "canvas_margin": "#1e1e1e",
    "path_bg": "#0d5a35",
    "path_border": "#25704a",
    "path_focus": "#36c986",
    "path_foreground": "#ffffff",
    "path_hover": "#bfffdc",
}


def normalize_theme_mode(value: object) -> str:
    text = str(value or "System").strip().lower()
    if text in {"bright", "light"}:
        return "Bright"
    if text == "dark":
        return "Dark"
    return "System"


def system_prefers_dark() -> bool:
    """Return the OS app-theme preference when it can be read reliably."""
    if os.name == "nt":
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                value, _kind = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return int(value) == 0
        except Exception:
            return False
    # Tk does not expose a portable Linux/macOS dark-mode flag.  Fall back to
    # the bright palette there rather than guessing from terminal settings.
    return False


def effective_theme_name(mode: object) -> str:
    normalized = normalize_theme_mode(mode)
    if normalized == "System":
        return "Dark" if system_prefers_dark() else "Bright"
    return normalized


def palette_for(mode: object) -> Mapping[str, str]:
    return DARK_PALETTE if effective_theme_name(mode) == "Dark" else LIGHT_PALETTE
