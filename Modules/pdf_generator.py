#!/usr/bin/env python3
"""
PDF Generator for Teams Call Flow Maps
Uses Playwright for reliable PDF generation from HTML
"""

import os
import sys
import asyncio
import argparse
from pathlib import Path
import platform

try:
    from playwright.async_api import async_playwright  # type: ignore
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

def check_playwright():
    """Check if Playwright is available and install instructions if not"""
    if PLAYWRIGHT_AVAILABLE:
        print("✓ Playwright is available")
        return True
    
    print("⚠ Playwright not found. Installation instructions:")
    print("  pip install playwright")
    print("  playwright install chromium")
    print("\nAlternatively, install all dependencies with:")
    print("  pip install playwright && playwright install")
    
    return False

async def convert_html_to_pdf_async(html_file, output_dir=None):
    """Convert HTML file to PDF using Playwright"""
    if not os.path.exists(html_file):
        print(f"✗ HTML file not found: {html_file}")
        return None
    
    # Determine output path
    html_path = Path(html_file)
    if output_dir:
        pdf_path = Path(output_dir) / f"{html_path.stem}.pdf"
    else:
        pdf_path = html_path.with_suffix('.pdf')
    
    # Create output directory if needed
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        async with async_playwright() as p:  # type: ignore
            # Launch browser
            browser = await p.chromium.launch()
            page = await browser.new_page()
            
            # Convert file path to file:// URL
            file_url = html_path.resolve().as_uri()
            
            # Load the HTML file
            await page.goto(file_url, wait_until='networkidle')
            
            # Wait a bit for any dynamic content
            await page.wait_for_timeout(2000)
            
            # Generate PDF with good settings
            await page.pdf(
                path=str(pdf_path),
                format='Letter',
                margin={
                    'top': '0.75in',
                    'right': '0.75in', 
                    'bottom': '0.75in',
                    'left': '0.75in'
                },
                print_background=True,
                prefer_css_page_size=True
            )
            
            await browser.close()
            
        if pdf_path.exists():
            print(f"✓ Generated PDF: {pdf_path.name}")
            return str(pdf_path)
        else:
            print(f"✗ PDF generation failed for {html_file}")
            return None
            
    except Exception as e:
        print(f"✗ PDF generation error: {e}")
        return None

def convert_html_to_pdf(html_file, output_dir=None):
    """Synchronous wrapper for async PDF conversion"""
    return asyncio.run(convert_html_to_pdf_async(html_file, output_dir))

async def batch_convert_directory_async(html_dir, pdf_dir=None):
    """Convert all HTML files in a directory to PDF"""
    html_path = Path(html_dir)
    if not html_path.exists():
        print(f"✗ HTML directory not found: {html_dir}")
        return []
    
    # Set up PDF output directory
    if pdf_dir is None:
        pdf_dir = html_path.parent / "PDF"
    
    pdf_path = Path(pdf_dir)
    pdf_path.mkdir(parents=True, exist_ok=True)
    
    # Find all HTML files
    html_files = list(html_path.glob("*.html"))
    if not html_files:
        print(f"✗ No HTML files found in: {html_dir}")
        return []
    
    print(f"Converting {len(html_files)} HTML files to PDF...")
    
    generated_pdfs = []
    
    async with async_playwright() as p:  # type: ignore
        # Launch browser once for all conversions
        browser = await p.chromium.launch()
        
        for html_file in html_files:
            try:
                page = await browser.new_page()
                
                # Determine output path
                pdf_output_path = pdf_path / f"{html_file.stem}.pdf"
                
                # Convert file path to file:// URL
                file_url = html_file.resolve().as_uri()
                
                # Load the HTML file
                await page.goto(file_url, wait_until='networkidle')
                
                # Wait a bit for any dynamic content
                await page.wait_for_timeout(1000)
                
                # Generate PDF
                await page.pdf(
                    path=str(pdf_output_path),
                    format='Letter',
                    margin={
                        'top': '0.75in',
                        'right': '0.75in', 
                        'bottom': '0.75in',
                        'left': '0.75in'
                    },
                    print_background=True,
                    prefer_css_page_size=True
                )
                
                await page.close()
                
                if pdf_output_path.exists():
                    print(f"✓ Generated PDF: {pdf_output_path.name}")
                    generated_pdfs.append(str(pdf_output_path))
                else:
                    print(f"✗ Failed to generate PDF for: {html_file.name}")
                    
            except Exception as e:
                print(f"✗ Error converting {html_file.name}: {e}")
        
        await browser.close()
    
    return generated_pdfs

def batch_convert_directory(html_dir, pdf_dir=None):
    """Synchronous wrapper for async batch conversion"""
    return asyncio.run(batch_convert_directory_async(html_dir, pdf_dir))

def main():
    parser = argparse.ArgumentParser(description="Convert Teams Call Flow HTML files to PDF")
    parser.add_argument("input", nargs='?', help="HTML file or directory to convert")
    parser.add_argument("-o", "--output", help="Output directory for PDFs")
    parser.add_argument("--check-deps", action="store_true", help="Check for required dependencies")
    
    args = parser.parse_args()
    
    if args.check_deps:
        if check_playwright():
            print("✓ All dependencies are available")
            sys.exit(0)
        else:
            print("✗ Missing dependencies")
            sys.exit(1)
    
    if not args.input:
        print("✗ Input file or directory is required")
        parser.print_help()
        sys.exit(1)
    
    if not check_playwright():
        print("✗ Playwright is required but not found")
        print("Install with: pip install playwright && playwright install chromium")
        sys.exit(1)
    
    input_path = Path(args.input)
    
    if input_path.is_file() and input_path.suffix.lower() == '.html':
        # Convert single file
        pdf_file = convert_html_to_pdf(str(input_path), args.output)
        if pdf_file:
            print(f"✓ Conversion complete: {pdf_file}")
        else:
            print("✗ Conversion failed")
            sys.exit(1)
    
    elif input_path.is_dir():
        # Convert all HTML files in directory
        generated_pdfs = batch_convert_directory(str(input_path), args.output)
        
        print(f"\n=== Conversion Complete ===")
        print(f"Generated {len(generated_pdfs)} PDF files")
        
        if generated_pdfs and args.output:
            print(f"Output directory: {args.output}")
    
    else:
        print(f"✗ Invalid input: {args.input}")
        print("Input must be an HTML file or directory containing HTML files")
        sys.exit(1)

if __name__ == "__main__":
    main()