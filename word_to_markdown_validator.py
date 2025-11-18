#!/usr/bin/env python3
"""
Word to Markdown Converter with Reference Validation

This script converts Word documents to Markdown format and validates that all
inline references are included in the document's reference list.

Requirements:
    pip install python-docx mammoth
"""

import re
import sys
from pathlib import Path
from typing import Set, List, Tuple, Dict
from collections import defaultdict


try:
    import mammoth
    from docx import Document
except ImportError:
    print("Error: Required libraries not installed.")
    print("Please run: pip install python-docx mammoth")
    sys.exit(1)


class ReferenceValidator:
    """Validates inline references against a reference list."""

    # Comprehensive inline citation patterns
    CITATION_PATTERNS = [
        # Numeric citations
        r'\[(\d+)\]',  # [1], [2], etc.
        r'\^(\d+)',    # ^1, ^2 (superscript in markdown)
        r'\((\d+)\)',  # (1), (2)

        # Author-year patterns (various formats)
        r"\(([A-Z][A-Za-z''\-]+(?:\s+et\s+al\.?)?[,\s]+\d{4}[a-z]?)\)",  # (Author et al., 2020)
        r"\[([A-Z][A-Za-z''\-]+(?:\s+et\s+al\.?)?[,\s]+\d{4}[a-z]?)\]",  # [Author et al., 2020]
        r"\(([A-Z][A-Za-z''\-]+\s+and\s+[A-Z][A-Za-z''\-]+[,\s]+\d{4}[a-z]?)\)",  # (Author and Author, 2020)
        r"\(([A-Z][A-Za-z''\-]+\s+&\s+[A-Z][A-Za-z''\-]+[,\s]+\d{4}[a-z]?)\)",  # (Author & Author, 2020)
        r"\(([A-Z][A-Za-z''\-]+\s+\d{4}[a-z]?)\)",  # (Author 2020)
        r"\[([A-Z][A-Za-z''\-]+\s+\d{4}[a-z]?)\]",  # [Author 2020]

        # Multiple authors variations
        r"\(([A-Z][A-Za-z''\-]+,\s+[A-Z][A-Za-z''\-]+,?\s+(?:and|&)\s+[A-Z][A-Za-z''\-]+[,\s]+\d{4}[a-z]?)\)",  # (A, B, and C, 2020)

        # Author with initials
        r"\(([A-Z][A-Za-z''\-]+\s+[A-Z]\.(?:\s+[A-Z]\.)?[,\s]+\d{4}[a-z]?)\)",  # (Smith J., 2020)

        # Common patterns with "et al"
        r"\(([A-Z][A-Za-z''\-]+\s+et\s+al\.\s+\d{4}[a-z]?)\)",  # (Smith et al. 2020)
    ]

    def __init__(self, docx_path: str, verbose: bool = False):
        """Initialize validator with a Word document path."""
        self.docx_path = Path(docx_path)
        self.markdown_content = ""
        self.inline_refs = []
        self.reference_list = []
        self.missing_refs = []
        self.verbose = verbose
        self.pattern_matches = defaultdict(list)  # Track which patterns matched what

    def convert_to_markdown(self) -> str:
        """Convert Word document to Markdown format."""
        print(f"Converting {self.docx_path} to Markdown...")

        try:
            with open(self.docx_path, "rb") as docx_file:
                result = mammoth.convert_to_markdown(docx_file)
                self.markdown_content = result.value

                if result.messages:
                    print("Conversion warnings:")
                    for message in result.messages:
                        print(f"  - {message}")

            print("✓ Conversion completed successfully")
            return self.markdown_content

        except Exception as e:
            print(f"Error converting document: {e}")
            sys.exit(1)

    def extract_inline_references(self) -> List[str]:
        """Extract all inline references from the markdown content."""
        print("\nExtracting inline references...")

        inline_refs_set = set()
        self.pattern_matches = defaultdict(list)

        for i, pattern in enumerate(self.CITATION_PATTERNS):
            matches = re.findall(pattern, self.markdown_content)
            if matches:
                if self.verbose:
                    print(f"  Pattern {i+1} matched: {len(matches)} citations")
                for match in matches:
                    inline_refs_set.add(match)
                    self.pattern_matches[pattern].append(match)

        self.inline_refs = sorted(list(inline_refs_set))
        print(f"✓ Found {len(self.inline_refs)} unique inline references")

        if self.verbose and self.inline_refs:
            print("\n  Sample inline references detected:")
            for ref in list(self.inline_refs)[:10]:
                print(f"    - {ref}")

        return self.inline_refs

    def extract_reference_list(self) -> List[str]:
        """Extract the reference list from the document."""
        print("\nExtracting reference list...")

        # Common section headers for references
        ref_headers = [
            r'#+\s*References?\s*$',
            r'#+\s*Bibliography\s*$',
            r'#+\s*Works?\s+Cited\s*$',
            r'#+\s*Literature\s+Cited\s*$',
        ]

        # Find the references section
        ref_section_start = -1
        lines = self.markdown_content.split('\n')

        for i, line in enumerate(lines):
            for header_pattern in ref_headers:
                if re.match(header_pattern, line, re.IGNORECASE):
                    ref_section_start = i
                    if self.verbose:
                        print(f"  Found reference section at line {i+1}: '{line.strip()}'")
                    break
            if ref_section_start != -1:
                break

        if ref_section_start == -1:
            print("⚠ Warning: Could not find reference section")
            print("  Looking for sections with headers: References, Bibliography, Works Cited, etc.")
            if self.verbose:
                print("\n  Showing all lines with '#' (headers) in the document:")
                for i, line in enumerate(lines):
                    if line.strip().startswith('#'):
                        print(f"    Line {i+1}: {line.strip()}")
            return []

        # Extract references from that section onwards
        reference_lines = lines[ref_section_start + 1:]

        # Filter out empty lines and extract references
        self.reference_list = [
            line.strip() for line in reference_lines
            if line.strip() and not line.strip().startswith('#')
        ]

        print(f"✓ Found {len(self.reference_list)} entries in reference list")

        if self.verbose and self.reference_list:
            print("\n  Sample reference list entries:")
            for ref in self.reference_list[:5]:
                print(f"    - {ref[:100]}{'...' if len(ref) > 100 else ''}")

        return self.reference_list

    def validate_references(self) -> Tuple[List[str], Dict[str, List[str]]]:
        """
        Validate that all inline references exist in the reference list.

        Returns:
            Tuple of (missing_refs, matched_refs_dict)
        """
        print("\nValidating references...")

        missing = []
        matched = defaultdict(list)

        for inline_ref in self.inline_refs:
            found = False

            # Check if inline reference appears in any reference list entry
            for ref_entry in self.reference_list:
                # For numeric citations like [1], check if the number matches
                if inline_ref.isdigit():
                    # Look for the number at the start of the reference
                    if re.match(rf'^\[?{inline_ref}\]?\.?\s+', ref_entry):
                        matched[inline_ref].append(ref_entry)
                        found = True
                        break
                else:
                    # For author-year citations, check if it appears in the reference
                    if inline_ref in ref_entry or self._normalize_citation(inline_ref) in self._normalize_citation(ref_entry):
                        matched[inline_ref].append(ref_entry)
                        found = True
                        break

            if not found:
                missing.append(inline_ref)

        self.missing_refs = missing

        if missing:
            print(f"✗ Found {len(missing)} inline references not in reference list")
        else:
            print("✓ All inline references are present in the reference list")

        return missing, dict(matched)

    @staticmethod
    def _normalize_citation(text: str) -> str:
        """Normalize citation text for comparison."""
        # Remove punctuation, extra spaces, and convert to lowercase
        normalized = re.sub(r'[^\w\s]', '', text.lower())
        normalized = re.sub(r'\s+', ' ', normalized).strip()
        return normalized

    def generate_report(self) -> str:
        """Generate a detailed validation report."""
        report = []
        report.append("=" * 70)
        report.append("WORD TO MARKDOWN REFERENCE VALIDATION REPORT")
        report.append("=" * 70)
        report.append(f"\nDocument: {self.docx_path}")
        report.append(f"Total inline references: {len(self.inline_refs)}")
        report.append(f"Total reference list entries: {len(self.reference_list)}")
        report.append(f"Missing references: {len(self.missing_refs)}")

        if self.inline_refs:
            report.append("\n" + "-" * 70)
            report.append("INLINE REFERENCES FOUND:")
            report.append("-" * 70)
            for ref in self.inline_refs[:20]:  # Show first 20
                report.append(f"  • {ref}")
            if len(self.inline_refs) > 20:
                report.append(f"  ... and {len(self.inline_refs) - 20} more")

        if self.missing_refs:
            report.append("\n" + "-" * 70)
            report.append("⚠ MISSING REFERENCES (not found in reference list):")
            report.append("-" * 70)
            for ref in self.missing_refs:
                report.append(f"  ✗ {ref}")

        if self.reference_list:
            report.append("\n" + "-" * 70)
            report.append("REFERENCE LIST ENTRIES:")
            report.append("-" * 70)
            for i, ref in enumerate(self.reference_list[:10], 1):  # Show first 10
                report.append(f"  {i}. {ref[:100]}{'...' if len(ref) > 100 else ''}")
            if len(self.reference_list) > 10:
                report.append(f"  ... and {len(self.reference_list) - 10} more")

        report.append("\n" + "=" * 70)
        if not self.missing_refs:
            report.append("✓ VALIDATION PASSED: All inline references are accounted for")
        else:
            report.append("✗ VALIDATION FAILED: Some inline references are missing")
        report.append("=" * 70)

        return "\n".join(report)

    def save_markdown(self, output_path: str = None) -> Path:
        """Save the converted markdown to a file."""
        if output_path is None:
            output_path = self.docx_path.with_suffix('.md')
        else:
            output_path = Path(output_path)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self.markdown_content)

        print(f"✓ Markdown saved to: {output_path}")
        return output_path

    def save_report(self, report_path: str = None) -> Path:
        """Save the validation report to a file."""
        if report_path is None:
            report_path = self.docx_path.with_suffix('.validation_report.txt')
        else:
            report_path = Path(report_path)

        report_content = self.generate_report()

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)

        print(f"✓ Validation report saved to: {report_path}")
        return report_path


def main():
    """Main entry point for the script."""
    if len(sys.argv) < 2:
        print("Usage: python word_to_markdown_validator.py <word_document.docx> [output.md] [--verbose]")
        print("\nThis script will:")
        print("  1. Convert the Word document to Markdown")
        print("  2. Extract all inline references")
        print("  3. Extract the reference list")
        print("  4. Validate that all inline references exist in the reference list")
        print("  5. Generate a detailed validation report")
        print("  6. Save the report to a .validation_report.txt file")
        print("\nOptions:")
        print("  --verbose    Show detailed debug information during processing")
        sys.exit(1)

    # Parse arguments
    verbose = '--verbose' in sys.argv or '-v' in sys.argv
    args = [arg for arg in sys.argv[1:] if not arg.startswith('-')]

    docx_path = args[0]
    output_path = args[1] if len(args) > 1 else None

    # Check if input file exists
    if not Path(docx_path).exists():
        print(f"Error: File '{docx_path}' not found")
        sys.exit(1)

    # Create validator and run the process
    validator = ReferenceValidator(docx_path, verbose=verbose)

    # Step 1: Convert to markdown
    validator.convert_to_markdown()

    # Step 2: Extract inline references
    validator.extract_inline_references()

    # Step 3: Extract reference list
    validator.extract_reference_list()

    # Step 4: Validate references
    missing, matched = validator.validate_references()

    # Step 5: Generate and display report
    report = validator.generate_report()
    print("\n" + report)

    # Step 6: Save files
    validator.save_markdown(output_path)
    validator.save_report()

    # Exit with appropriate code
    sys.exit(0 if not missing else 1)


if __name__ == "__main__":
    main()
