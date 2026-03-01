# Macedonian Font Converter / Македонски конвертер на фонтови

This tool converts old Macedonian font encoding (Latin characters) to proper Macedonian Cyrillic letters.

## Supported File Formats
- Microsoft Word: `.docx`, `.doc`
- Microsoft Excel: `.xlsx`, `.xls`

## Two Ways to Use

### Option 1: Web-Based Converter (Easiest)
1. Open `macedonian_converter.html` in any web browser
2. Drag and drop your file or click to browse
3. Click "Конвертирај / Convert"
4. Download the converted file (saved with `_converted` suffix)

**Features:**
- Works directly in your browser
- No installation needed
- Text converter for quick testing
- Visual conversion map reference

**Limitations:**
- Works best with `.docx` and `.xlsx` files
- For `.doc` and `.xls` files, convert them to newer formats first

### Option 2: Python Script (More Powerful)
For batch processing or `.doc`/`.xls` files:

1. Install Python 3.7 or higher
2. Run the converter:
```bash
python macedonian_converter.py file1.docx file2.xlsx file3.doc
```

The script will automatically install required packages and process all files.

## Conversion Map

The converter uses the following character mapping:

| Old Font | Cyrillic | Old Font | Cyrillic |
|----------|----------|----------|----------|
| A, a     | А, а     | N, n     | Н, н     |
| B, b     | Б, б     | O, o     | О, о     |
| C, c     | Ц, ц     | P, p     | П, п     |
| D, d     | Д, д     | Q, q     | Љ, љ     |
| E, e     | Е, е     | R, r     | Р, р     |
| F, f     | Ф, ф     | S, s     | С, с     |
| G, g     | Г, г     | T, t     | Т, т     |
| H, h     | Х, х     | U, u     | У, у     |
| I, i     | И, и     | V, v     | В, в     |
| J, j     | Ј, ј     | W, w     | Њ, њ     |
| K, k     | К, к     | X, x     | Џ, џ     |
| L, l     | Л, л     | Y, y     | Ѕ, ѕ     |
| M, m     | М, м     | Z, z     | З, з     |

**Special Characters:**
- ` → Ж
- ~ → ж
- @ → Ќ
- ^ → ч
- [ → Ш
- ] → Ќ
- { → ш
- } → Ќ
- | → Ѓ

## Output Files

Converted files are saved with `_converted` added to the filename:
- `document.docx` → `document_converted.docx`
- `spreadsheet.xlsx` → `spreadsheet_converted.xlsx`

## Troubleshooting

**Web version doesn't work:**
- Make sure you're using a modern browser (Chrome, Firefox, Edge)
- Check that JavaScript is enabled
- Try converting to `.docx` or `.xlsx` format first if using older formats

**Python script errors:**
- Ensure Python 3.7+ is installed: `python --version`
- The script will automatically install required packages
- For `.doc` files on Linux/Mac, convert to `.docx` first

## Examples

**Convert a single file:**
```bash
python macedonian_converter.py my_document.docx
```

**Convert multiple files:**
```bash
python macedonian_converter.py file1.docx file2.xlsx file3.docx
```

**Test conversion with text:**
Open the HTML file and use the text converter section to test the conversion on sample text.

## Notes

- Original files are never modified
- Converted files are saved as new files with `_converted` suffix
- The conversion preserves all formatting, styles, and structure
- Numbers, punctuation, and other characters remain unchanged

---

За прашања или проблеми / For questions or issues, please check the conversion map to verify the character mappings match your old font.
