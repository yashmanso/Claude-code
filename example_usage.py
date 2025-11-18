#!/usr/bin/env python3
"""
Example usage of the Word to Markdown Reference Validator.

This script demonstrates how to use the ReferenceValidator class
programmatically in your own Python code.
"""

from word_to_markdown_validator import ReferenceValidator


def example_basic_usage():
    """Basic usage example."""
    print("=" * 70)
    print("EXAMPLE: Basic Usage")
    print("=" * 70)

    # Create validator instance
    validator = ReferenceValidator("your_document.docx")

    # Convert to markdown
    markdown_content = validator.convert_to_markdown()

    # Extract references
    inline_refs = validator.extract_inline_references()
    ref_list = validator.extract_reference_list()

    # Validate
    missing_refs, matched_refs = validator.validate_references()

    # Generate and print report
    report = validator.generate_report()
    print(report)

    # Save markdown
    validator.save_markdown("output.md")

    return len(missing_refs) == 0


def example_detailed_analysis():
    """Example showing detailed analysis of references."""
    print("\n" + "=" * 70)
    print("EXAMPLE: Detailed Analysis")
    print("=" * 70)

    validator = ReferenceValidator("your_document.docx")

    # Process document
    validator.convert_to_markdown()
    validator.extract_inline_references()
    validator.extract_reference_list()
    missing_refs, matched_refs = validator.validate_references()

    # Analyze results
    print("\nDetailed Analysis:")
    print(f"  Total words in document: {len(validator.markdown_content.split())}")
    print(f"  Inline references: {len(validator.inline_refs)}")
    print(f"  Reference list entries: {len(validator.reference_list)}")
    print(f"  Missing references: {len(missing_refs)}")

    if matched_refs:
        print("\nMatched References (sample):")
        for i, (inline_ref, ref_entries) in enumerate(list(matched_refs.items())[:5]):
            print(f"  {i+1}. '{inline_ref}' matched with:")
            for entry in ref_entries:
                print(f"     - {entry[:80]}...")

    if missing_refs:
        print("\nMissing References:")
        for ref in missing_refs:
            print(f"  ✗ {ref}")

    return validator


def example_validation_only():
    """Example focusing only on validation, not conversion."""
    print("\n" + "=" * 70)
    print("EXAMPLE: Validation Only")
    print("=" * 70)

    validator = ReferenceValidator("your_document.docx")
    validator.convert_to_markdown()
    validator.extract_inline_references()
    validator.extract_reference_list()
    missing_refs, matched_refs = validator.validate_references()

    # Simple validation result
    if not missing_refs:
        print("✓ VALIDATION PASSED: All references are accounted for!")
        return True
    else:
        print(f"✗ VALIDATION FAILED: {len(missing_refs)} references are missing:")
        for ref in missing_refs:
            print(f"  - {ref}")
        return False


def example_custom_patterns():
    """
    Example showing how you might extend the validator
    with custom citation patterns.
    """
    print("\n" + "=" * 70)
    print("EXAMPLE: Custom Citation Patterns")
    print("=" * 70)

    validator = ReferenceValidator("your_document.docx")

    # Add custom patterns (extend the class in practice)
    # This is just demonstrative
    custom_patterns = [
        r'\{([A-Z][a-z]+\s+\d{4})\}',  # {Author 2020}
        r'<([A-Z][a-z]+,\s+\d{4})>',   # <Author, 2020>
    ]

    # In a real implementation, you'd extend the class
    # validator.CITATION_PATTERNS.extend(custom_patterns)

    print("Custom patterns that could be added:")
    for pattern in custom_patterns:
        print(f"  - {pattern}")

    print("\nNote: To use custom patterns, extend the ReferenceValidator class")
    print("and modify the CITATION_PATTERNS class variable.")


if __name__ == "__main__":
    print("\nWord to Markdown Reference Validator - Example Usage")
    print("=" * 70)
    print("\nThis file demonstrates various ways to use the validator.")
    print("Replace 'your_document.docx' with an actual Word document path.\n")

    # Note: These examples will fail without an actual .docx file
    # Uncomment to run with a real document:

    # example_basic_usage()
    # example_detailed_analysis()
    # example_validation_only()
    # example_custom_patterns()

    print("\n" + "=" * 70)
    print("To run these examples:")
    print("  1. Uncomment the function calls above")
    print("  2. Replace 'your_document.docx' with your actual file path")
    print("  3. Run: python example_usage.py")
    print("=" * 70)
