# -*- coding: UTF-8 -*-
# buildVars.py - variables used by SCons when building the addon.

# Add-on information variables
addon_info = {
    "addon_name": "ThaiTypeTest",

    "addon_version": "2025.08.28",

    "addon_author": "NVDA_TH <nvdainth@gmail.com>, assisted by A.I.",

    "addon_summary": "An add-on to test and develop Thai typing speed and accuracy.",

    "addon_description": (
        "A complete toolkit for practicing and measuring Thai typing skills for NVDA users. "
        "It features various test modes, including Random Words (General/Hard), Sentences, Lyrics, and Literature. "
        "The add-on provides detailed metrics like Net WPM, Gross WPM, and Accuracy. "
        "It also includes dynamic dataset management, allowing users to add new lyrics via URL from supported websites (Kapook, Siamzone, Meemodel) "
        "or edit the datasets directly."
    ),

    "addon_url": "https://nvda.in.th",

    "addon_docFileName": "readme.html",

    "addon_minimumNVDAVersion": "2025.1",
    "addon_lastTestedNVDAVersion": "2025.2",

    "addon_updateChannel": None,
}

# Define the folder that contains the Python source code for the add-on
pythonSources = [
    "addon/globalPlugins",
]

# Files that are not Python source code but should be included in the add-on
# We don't need this as our .txt files are inside the globalPlugins/ThaiTypeTest/lib folder
# and will be included automatically.
i18nSources = []
docFiles = ["readme.html"]

# Files to be excluded from the build
tests = []
excludedFiles = []