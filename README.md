# Obfusel - Excel Obfuscator

An Excel obfuscator tool that helps you remove sensitive data from Excel files while preserving their structure for analysis.

## Features

- **Read Excel files** (.xlsx format)
- **Output to CSV** format (due to library limitations)
- **Preserve headers** optionally
- **String obfuscation** with multiple modes:
  - Random: Replace with random strings
  - Consistent: Same original value gets same replacement
  - None: Don't replace strings
- **Number obfuscation** with multiple modes:
  - Jitter: Add random noise to numbers (default)
  - Random: Replace with completely random numbers
  - Consistent: Same original value gets same replacement
  - None: Don't replace numbers
- **Row and column shuffling**

## Installation

Make sure you have Nim and Nimble installed, then:

```bash
nimble build
```

## Usage

```bash
./bin/obfusel [options] input_file.xlsx output_file.csv
```

### Options

- `--help, -h`: Show help message
- `--preserve-headers, -ph`: Preserve header values in the first row
- `--preserve-formulas, -pf`: Preserve formulas (don't replace them with values)
- `--preserve-numbers, -pn`: Keep numeric values unchanged
- `--shuffle-rows, -sr`: Shuffle the order of rows (except headers if preserved)
- `--shuffle-columns, -sc`: Shuffle the order of columns
- `--string-replacement=TYPE`: How to replace string values (random, consistent, none)
- `--number-replacement=TYPE`: How to replace numeric values (jitter, random, consistent, none)

### Example

```bash
./bin/obfusel --preserve-headers --string-replacement=consistent input.xlsx output.csv
```

## Dependencies

- Nim >= 1.6.0
- xlsx library (for reading Excel files)

## Limitations

- Output is limited to CSV format due to the xlsx library being read-only
- Only processes the first sheet in multi-sheet Excel files when outputting to CSV

## Use Case

This tool is particularly useful when you need to:
1. Share Excel data with AI systems or external parties
2. Remove sensitive information while preserving data structure
3. Maintain referential integrity with consistent replacements
4. Test systems with realistic but non-sensitive data
