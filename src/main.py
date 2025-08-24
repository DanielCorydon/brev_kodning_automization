import argparse
import os
from components.convert_text_fields import convert_document_fields


def main():
    """Main entry point for the field conversion tool."""
    parser = argparse.ArgumentParser(
        description="Convert text-based field representations in Word documents to actual Word fields"
    )
    parser.add_argument("input", help="Input Word document path")
    parser.add_argument(
        "--output",
        help="Output Word document path. If not specified, will add '_converted' to the input filename.",
    )

    args = parser.parse_args()

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        input_base, input_ext = os.path.splitext(args.input)
        output_path = f"{input_base}_converted{input_ext}"

    # Convert the document
    convert_document_fields(args.input, output_path)
    print(f"Conversion complete. Saved to: {output_path}")


if __name__ == "__main__":
    main()
