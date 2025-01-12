# CSV2PDF

A quick&dirty GUI application that generates multiple PDFs from a PDF template by combining it with text content and data from a CSV file (using tags as {name} defined as CSV headers). The application generates 1 PDF for each CSV entry.

- [Download Linux binary](https://github.com/emprius/csv2pdf/raw/refs/heads/main/dist/csv2pdf)

![screenshot](https://github.com/emprius/csv2pdf/blob/main/dist/screenshot.png?raw=true)

## Installation

### Option 1: Download Portable Binary (Recommended)
1. Download the latest `csv2pdf` binary from the releases page
2. Make it executable: `chmod +x csv2pdf`
3. Double-click to run, or run from terminal: `./csv2pdf`

### Option 2: Build from Source
1. Clone the repository
2. Install make: `sudo apt install make` (on Debian/Ubuntu) or `sudo pacman -S make` (on Arch)
3. Run `make install` and `make run` to run the program.
4. Run `make build` to create the portable binary. The binary will be created in the `dist` directory

## Features

- Load a PDF template
- Insert text with formatting (bold, italic, underline)
- Use tags that get replaced with values from CSV file
- Support for special characters (\n for new line, \t for tab)
- Preserve text formatting and indentation
- Save and restore settings between sessions
- Customizable font, size, and position

## How to Use

1. **Select Template PDF**: Click "Browse" to select your PDF template file
2. **Select CSV File**: Click "Browse" to select your CSV file with data
3. **Select Output Directory**: Choose where to save the generated PDFs
4. **Enter Text**: Type or paste your text in the text area
   - Use `{column_name}` to insert values from CSV
   - Use `\n` for new line
   - Use `\t` for tab (4 spaces)
   - Select text and use B/I/U buttons for formatting
5. **Configure Settings**:
   - Font: Choose between Helvetica, Times-Roman, or Courier
   - Font Size: Set the text size
   - Text Position: Set X and Y positions (in %)
   - Chars per Line: Set maximum characters per line
   - Filename: Set output filename pattern using tags
6. **Generate PDFs**: Click "Generate PDFs" button

## Example

### CSV File (data.csv):
```csv
name,age,city
John Doe,30,New York
Jane Smith,25,Los Angeles
```

### Text Content Example:
```
\tDear {name},

I hope this letter finds you well. I understand you are {age} years old
and living in {city}.

\tBest regards,
```

### Output
This will generate a PDF for each row in the CSV file, replacing the tags with the corresponding values:

For first row:
```
    Dear John Doe,

I hope this letter finds you well. I understand you are 30 years old
and living in New York.

    Best regards,
```

For second row:
```
    Dear Jane Smith,

I hope this letter finds you well. I understand you are 25 years old
and living in Los Angeles.

    Best regards,
```

### Filename Pattern Example:
```
letter_{name}_{city}
```
This will generate files like:
- letter_John_Doe_New_York.pdf
- letter_Jane_Smith_Los_Angeles.pdf

## Tips

1. **Tags**:
   - Must match CSV column names exactly: `{name}`, `{age}`, `{city}`
   - Are case-sensitive
   - Will be replaced with the corresponding value from the CSV

2. **Special Characters**:
   - `\n`: Creates a new line
   - `\t`: Adds 4 spaces of indentation
   - Can be used anywhere in the text

3. **Text Formatting**:
   - Select text and click B for bold
   - Select text and click I for italic
   - Select text and click U for underline
   - Can combine multiple formats
   - Formatting is preserved in the PDF

4. **Settings**:
   - All settings are automatically saved
   - Will be restored next time you open the program
   - Includes file paths, text content, and formatting preferences
