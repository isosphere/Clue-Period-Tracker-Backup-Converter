# Clue-Period-Tracker-Backup-Converter

Converts a Clue (Android-based period tracker) backup to a Microsoft Excel file.

This is not an advanced script; a `.cluedata` file is a JSON file. This script simply:
- loads that JSON
- enumerates the columns that should be used in the output
- outputs a representation of the JSON in Excel format
