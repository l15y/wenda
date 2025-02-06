"""
colour
"""

from pptx.dml.color import RGBColor, MSO_THEME_COLOR
import re

RGBRegex = re.compile("#([0-9a-fA-F]{6})")

def setColour(x, colour):
    colourType, colourValue = colour
    if colourType == "Theme":
        x.theme_color = colourValue
    else:
        x.rgb = RGBColor.from_string(colourValue[1:])


def parseThemeColour(value):
    value2 = value.upper()
    if value2 == "NONE":
        return MSO_THEME_COLOR.NOT_THEME_COLOR
    elif value2 == "ACCENT 1":
        return MSO_THEME_COLOR.ACCENT_1
    elif value2 == "ACCENT 2":
        return MSO_THEME_COLOR.ACCENT_2
    elif value2 == "ACCENT 3":
        return MSO_THEME_COLOR.ACCENT_3
    elif value2 == "ACCENT 4":
        return MSO_THEME_COLOR.ACCENT_4
    elif value2 == "ACCENT 5":
        return MSO_THEME_COLOR.ACCENT_5
    elif value2 == "ACCENT 6":
        return MSO_THEME_COLOR.ACCENT_6
    elif value2 == "BACKGROUND 1":
        return MSO_THEME_COLOR.BACKGROUND_1
    elif value2 == "BACKGROUND 2":
        return MSO_THEME_COLOR.BACKGROUND_2
    elif value2 == "DARK 1":
        return MSO_THEME_COLOR.DARK_1
    elif value2 == "DARK 2":
        return MSO_THEME_COLOR.DARK_2
    elif value2 == "FOLLOWED HYPERLINK":
        return MSO_THEME_COLOR.FOLLOWED_HYPERLINK
    elif value2 == "HYPERLINK":
        return MSO_THEME_COLOR.HYPERLINK
    elif value2 == "LIGHT 1":
        return MSO_THEME_COLOR.LIGHT_1
    elif value2 == "LIGHT 2":
        return MSO_THEME_COLOR.LIGHT_2
    elif value2 == "TEXT 1":
        return MSO_THEME_COLOR.TEXT_1
    elif value2 == "TEXT 2":
        return MSO_THEME_COLOR.TEXT_2
    elif value2 == "MIXED":
        return MSO_THEME_COLOR.MIXED

def parseColour(value):
    if value[0] == "#":
        return ("RGB", value)
    else:
        return ("Theme", parseThemeColour(value))

def parseRGB(str):
    if RGBmatch := RGBRegex.match(str):
        # Matches
        return (True, RGBmatch.group(1))
    else:
        return (False, "")

