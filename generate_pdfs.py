#!/usr/bin/env python3
"""
Simple PDF Generator Wrapper
Easy-to-use script for generating PDFs from Teams call flow HTML files
"""

import os
import sys
import argparse
from pathlib import Path

# Add modules to path
sys.path.insert(0, str(Path(__file__).parent / "Modules"))

try:
    from pdf_generator import convert_html_to_pdf, batch_convert_directory, check_playwright
except ImportError:
    print("✗ Could not import PDF generator module")
    print("Ensure the Modules/pdf_generator.py file exists")
    sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description="Generate PDFs from Teams Call Flow HTML files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Generate PDFs for all HTML files in a directory
  python3 generate_pdfs.py CallFlowMaps_20250926_161304/Individual
  
  # Generate PDF for a single HTML file
  python3 generate_pdfs.py CallFlowMaps_20250926_161304/Summary/Dashboard.html
  
  # Specify custom output directory
  python3 generate_pdfs.py CallFlowMaps_20250926_161304/Individual -o MyPDFs
  
  # Check if prerequisites are installed
  python3 generate_pdfs.py --check
        """
    )
    
    parser.add_argument(
        "input",
        nargs="?",
        help="HTML file or directory containing HTML files to convert"
    )
    
    parser.add_argument(
        "-o", "--output",
        help="Output directory for PDF files (default: same location as HTML with PDF suffix)"
    )
    
    parser.add_argument(
        "--check",
        action="store_true",
        help="Check if Playwright and other prerequisites are available"
    )
    
    args = parser.parse_args()
    
    # Check prerequisites
    if args.check or not args.input:
        print("=== PDF Generation Prerequisites ===")
        if check_playwright():
            print("✓ All prerequisites are available")
            print("✓ Ready to generate PDFs")
            if not args.input:
                sys.exit(0)
        else:
            print("✗ Prerequisites not met")
            print("\nTo install Playwright:")
            print("  pip install playwright")
            print("  playwright install chromium")
            sys.exit(1)
    
    # Validate input
    if not args.input:
        parser.print_help()
        sys.exit(1)
    
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"✗ Input path does not exist: {args.input}")
        sys.exit(1)
    
    print(f"=== PDF Generation ===")
    print(f"Input: {input_path}")
    
    # Determine output directory
    if args.output:
        output_dir = Path(args.output)
        print(f"Output: {output_dir}")
    else:
        if input_path.is_file():
            output_dir = input_path.parent / "PDF"
        else:
            output_dir = input_path.parent / "PDF"
        print(f"Output: {output_dir} (default)")
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate PDFs
    try:
        if input_path.is_file() and input_path.suffix.lower() == '.html':
            # Single file
            pdf_file = convert_html_to_pdf(str(input_path), str(output_dir))
            if pdf_file:
                print(f"✓ Generated: {Path(pdf_file).name}")
                print(f"✓ PDF generation successful")
            else:
                print("✗ PDF generation failed")
                sys.exit(1)
        
        elif input_path.is_dir():
            # Directory
            generated_pdfs = batch_convert_directory(str(input_path), str(output_dir))
            
            if generated_pdfs:
                print(f"\n=== Generation Complete ===")
                print(f"✓ Generated {len(generated_pdfs)} PDF files")
                print(f"Location: {output_dir}")
                
                # List the files
                print("\nGenerated PDFs:")
                for pdf_file in generated_pdfs:
                    print(f"  • {Path(pdf_file).name}")
            else:
                print("✗ No PDFs were generated")
                print("Check that the input directory contains HTML files")
                sys.exit(1)
        
        else:
            print(f"✗ Invalid input: {args.input}")
            print("Input must be an HTML file or directory containing HTML files")
            sys.exit(1)
    
    except KeyboardInterrupt:
        print("\n⚠ PDF generation interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Error during PDF generation: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()