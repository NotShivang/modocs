# MoDocs

**Mo**bile **Doc**ument**s** — a lightweight Android app for viewing and working with documents on your phone.

If you're like me, you sometimes use your phone to look at docs, fill out forms, or open up a quick Excel file — but you don't do any heavy document editing on your phone. There was no single lightweight app on the Play Store that supported opening PDF, DOCX, XLSX, and PPTX without showing you a bunch of ads, so I decided to vibe code one.

## Features

- **PDF** — View, search, fill forms (tap to place text, signatures, checkmarks, dates), and save filled copies
- **DOCX** — View Word documents with formatting, images, and basic layout
- **XLSX** — View and edit spreadsheet cells, save changes
- **PPTX** — View presentations with text, shapes, images, and backgrounds

## Installation

Grab the latest APK from [Releases](https://github.com/NotShivang/modocs/releases) and sideload it on your Android device.

**Requires Android 8.0 (API 26) or higher.**

## Known Limitations

### DOCX
- Documents won't show up exactly how they look in Word, especially if they use complicated features like nested tables or unusual spacing/positioning. For heavy Word editing, use the real thing.

### XLSX
- No formula support — formulas are not evaluated, only their last cached values are shown
- Editing is limited to viewing and modifying cell values (no formatting changes, no inserting rows/columns)

### PPTX
- This is the least baked feature (because it's very rare I open a PPT on my phone)
- Complex graphics, SmartArt, charts, and animations will not render
- No editing support for presentations — just viewing
- If you need to do a lot of presentation editing on mobile, just get Microsoft Office

### PDF
- Fill & sign flattens annotations onto a re-rendered copy of the PDF (not a direct edit of the original file structure)
- Very large PDFs may be slow to save

## Building from Source

```bash
# Clone the repo
git clone https://github.com/NotShivang/modocs.git
cd modocs

# Build debug APK
./gradlew assembleDebug

# APK will be at app/build/outputs/apk/debug/app-debug.apk
```

## License

[MIT](LICENSE)
