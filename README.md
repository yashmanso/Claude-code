# Word to Markdown Reference Validator

A Python tool that converts Word documents (.docx) to Markdown format and validates that all inline references are included in the document's reference list.

## Features

- **Word to Markdown Conversion**: Converts .docx files to clean Markdown format using mammoth
- **Inline Reference Extraction**: Detects various citation formats:
  - Numeric citations: `[1]`, `[2]`, etc.
  - Author-year citations: `(Author, 2020)`, `(Smith & Jones, 2019)`
  - Alternative formats: `[Author, 2020]`, `(Author 2020)`
- **Reference List Extraction**: Automatically finds and parses the References/Bibliography section
- **Validation**: Checks that every inline reference has a corresponding entry in the reference list
- **Detailed Reporting**: Generates comprehensive validation reports showing:
  - Total inline references found
  - Total reference list entries
  - Missing references (if any)
  - Full listing of inline references and reference entries

## Installation

### Prerequisites

- Python 3.6 or higher
- pip

### Install Dependencies

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install python-docx mammoth
```

## Usage

### Basic Usage

```bash
python word_to_markdown_validator.py document.docx
```

This will:
1. Convert `document.docx` to Markdown
2. Save the output as `document.md`
3. Validate all references
4. Display a detailed report in the console
5. Save the validation report as `document.validation_report.txt`

### Specify Output File

```bash
python word_to_markdown_validator.py document.docx output.md
```

### Verbose Mode (Debugging)

Use `--verbose` to see detailed debug information about what references are being detected:

```bash
python word_to_markdown_validator.py document.docx --verbose
```

This shows:
- Which citation patterns matched
- Sample inline references found
- Sample reference list entries
- All document headers (useful for troubleshooting)

### Example Output

```
Converting document.docx to Markdown...
✓ Conversion completed successfully

Extracting inline references...
✓ Found 15 unique inline references

Extracting reference list...
✓ Found 14 entries in reference list

Validating references...
✗ Found 1 inline references not in reference list

======================================================================
WORD TO MARKDOWN REFERENCE VALIDATION REPORT
======================================================================

Document: document.docx
Total inline references: 15
Total reference list entries: 14
Missing references: 1

----------------------------------------------------------------------
INLINE REFERENCES FOUND:
----------------------------------------------------------------------
  • Smith, 2020
  • Jones et al., 2019
  • Brown & Davis, 2021
  ...

----------------------------------------------------------------------
⚠ MISSING REFERENCES (not found in reference list):
----------------------------------------------------------------------
  ✗ Wilson, 2022

----------------------------------------------------------------------
REFERENCE LIST ENTRIES:
----------------------------------------------------------------------
  1. Smith, J. (2020). Example Article. Journal Name, 15(2), 123-145.
  2. Jones, A., et al. (2019). Another Study. Conference Proceedings.
  ...

======================================================================
✗ VALIDATION FAILED: Some inline references are missing
======================================================================

✓ Markdown saved to: document.md
✓ Validation report saved to: document.validation_report.txt
```

## Supported Citation Formats

The tool recognizes many common academic citation formats:

### Numeric Citations
- `[1]`, `[2]`, `[3]` - Bracketed numbers
- `^1`, `^2` - Superscript numbers (markdown format)
- `(1)`, `(2)` - Parenthetical numbers

### Author-Year Citations
- `(Smith, 2020)` - Single author with comma
- `(Smith 2020)` - Single author without comma
- `[Smith, 2020]` - Bracketed citation
- `(Jones & Brown, 2019)` - Two authors with ampersand
- `(Jones and Brown, 2019)` - Two authors with "and"
- `(Davis et al., 2021)` - Multiple authors with et al.
- `(Davis et al. 2021)` - Et al. without comma
- `(Smith, Jones, and Brown, 2020)` - Three+ authors listed
- `(Smith J., 2020)` - Author with initials
- Works with author names containing apostrophes and hyphens

## Reference Section Detection

The tool automatically detects reference sections with these common headers:
- References
- Reference
- Bibliography
- Works Cited
- Literature Cited

Headers can use any markdown heading level (# or ##, etc.).

## Exit Codes

- `0`: Validation passed - all inline references are in the reference list
- `1`: Validation failed - some references are missing

## Use Cases

- **Academic Writing**: Ensure all citations are properly referenced
- **Document Review**: Validate reference completeness before submission
- **Format Conversion**: Convert Word documents to Markdown while checking references
- **Automation**: Integrate into CI/CD pipelines for document validation

## Programmatic Usage

You can also use the tool as a library in your Python code:

```python
from word_to_markdown_validator import ReferenceValidator

# Create validator instance
validator = ReferenceValidator("document.docx")

# Convert to markdown
validator.convert_to_markdown()

# Extract references
validator.extract_inline_references()
validator.extract_reference_list()

# Validate
missing_refs, matched_refs = validator.validate_references()

# Generate report
report = validator.generate_report()
print(report)

# Save markdown
validator.save_markdown("output.md")

# Check results
if missing_refs:
    print(f"Warning: {len(missing_refs)} references are missing")
    for ref in missing_refs:
        print(f"  - {ref}")
```

## Limitations

- The tool works best with standard academic citation formats
- Custom or non-standard citation styles may not be fully detected
- Reference list must be in a clearly marked section (References, Bibliography, etc.)
- Complex footnote or endnote citations may require additional pattern configuration

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## License

MIT License - feel free to use this tool in your projects.