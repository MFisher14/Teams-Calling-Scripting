#!/usr/bin/env python3
"""
Setup script for Teams Call Flow PDF Generation
Installs and configures Playwright for PDF generation
"""

import os
import sys
import subprocess
import platform

def check_python_version():
    """Check if Python version is adequate"""
    if sys.version_info < (3, 7):
        print("✗ Python 3.7 or higher is required for Playwright")
        print(f"  Current version: {sys.version}")
        return False
    
    print(f"✓ Python version: {sys.version.split()[0]}")
    return True

def install_playwright():
    """Install Playwright and browser dependencies"""
    print("Installing Playwright...")
    
    try:
        # Install Playwright package
        print("Installing Playwright package...")
        result = subprocess.run([sys.executable, "-m", "pip", "install", "playwright"], 
                              capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"✗ Playwright installation failed: {result.stderr}")
            return False
        
        print("✓ Playwright package installed")
        
        # Install browser binaries
        print("Installing browser binaries...")
        result = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], 
                              capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"✗ Browser installation failed: {result.stderr}")
            return False
        
        print("✓ Chromium browser installed")
        
        # Install system dependencies (Linux)
        if platform.system().lower() == "linux":
            print("Installing system dependencies...")
            result = subprocess.run([sys.executable, "-m", "playwright", "install-deps"], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                print("✓ System dependencies installed")
            else:
                print("⚠ System dependencies installation had issues (may still work)")
        
        return True
        
    except Exception as e:
        print(f"✗ Installation error: {e}")
        return False

def test_installation():
    """Test if Playwright is working"""
    try:
        import subprocess
        test_script = """
import asyncio
from playwright.async_api import async_playwright

async def test():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto('data:text/html,<h1>Test</h1>')
        await browser.close()
    print('✓ Playwright test successful')

asyncio.run(test())
"""
        
        result = subprocess.run([sys.executable, "-c", test_script], 
                              capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            print("✓ Playwright is working correctly")
            return True
        else:
            print(f"✗ Playwright test failed: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ Playwright test timed out")
        return False
    except Exception as e:
        print(f"✗ Test error: {e}")
        return False

def main():
    print("=== Teams Call Flow PDF Generator Setup (Playwright) ===\n")
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install Playwright
    if not install_playwright():
        print("\n✗ Playwright installation failed")
        print("You can try manual installation:")
        print("  pip install playwright")
        print("  playwright install chromium")
        sys.exit(1)
    
    # Test installation
    print("\nTesting installation...")
    if not test_installation():
        print("\n⚠ Installation test failed")
        print("Playwright may still work, but there might be issues")
        print("Try running: python pdf_generator.py --check-deps")
    
    print("\n✓ Setup completed successfully!")
    print("You can now run: python pdf_generator.py --help")
    print("\nTo test PDF generation:")
    print("  python pdf_generator.py --check-deps")

if __name__ == "__main__":
    main()

import os
import sys
import subprocess
import platform
import shutil
from pathlib import Path

def run_command(command, shell=False):
    """Run a command and return success status"""
    try:
        result = subprocess.run(command, shell=shell, capture_output=True, text=True)
        return result.returncode == 0, result.stdout, result.stderr
    except Exception as e:
        return False, "", str(e)

def check_python_version():
    """Check if Python version is adequate"""
    version = sys.version_info
    if version.major >= 3 and version.minor >= 6:
        print(f"✓ Python {version.major}.{version.minor}.{version.micro} is adequate")
        return True
    else:
        print(f"✗ Python {version.major}.{version.minor}.{version.micro} is too old (need 3.6+)")
        return False

def check_wkhtmltopdf():
    """Check if wkhtmltopdf is installed"""
    if shutil.which('wkhtmltopdf'):
        success, stdout, stderr = run_command(['wkhtmltopdf', '--version'])
        if success:
            version = stdout.split('\n')[0] if stdout else "unknown version"
            print(f"✓ wkhtmltopdf is installed: {version}")
            return True
    
    print("✗ wkhtmltopdf is not installed")
    return False

def install_wkhtmltopdf_macos():
    """Install wkhtmltopdf on macOS using Homebrew"""
    print("Attempting to install wkhtmltopdf on macOS...")
    
    # Check if Homebrew is installed
    if not shutil.which('brew'):
        print("✗ Homebrew is not installed. Please install Homebrew first:")
        print("  /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"")
        return False
    
    # Try to install wkhtmltopdf
    success, stdout, stderr = run_command(['brew', 'install', 'wkhtmltopdf'])
    if success:
        print("✓ wkhtmltopdf installed successfully via Homebrew")
        return True
    else:
        print(f"✗ Failed to install wkhtmltopdf via Homebrew: {stderr}")
        return False

def install_wkhtmltopdf_linux():
    """Install wkhtmltopdf on Linux"""
    print("Attempting to install wkhtmltopdf on Linux...")
    
    # Try different package managers
    package_managers = [
        (['apt-get', 'update'], ['apt-get', 'install', '-y', 'wkhtmltopdf']),
        (['yum', 'install', '-y', 'wkhtmltopdf'],),
        (['dnf', 'install', '-y', 'wkhtmltopdf'],),
        (['pacman', '-S', '--noconfirm', 'wkhtmltopdf'],),
    ]
    
    for commands in package_managers:
        pm_available = shutil.which(commands[0][0])
        if pm_available:
            print(f"Using {commands[0][0]} package manager...")
            
            # Run update command if provided
            if len(commands) > 1:
                run_command(commands[0], shell=False)
                success, stdout, stderr = run_command(commands[1], shell=False)
            else:
                success, stdout, stderr = run_command(commands[0], shell=False)
            
            if success:
                print("✓ wkhtmltopdf installed successfully")
                return True
            else:
                print(f"Failed with {commands[0][0]}: {stderr}")
    
    print("✗ Could not install wkhtmltopdf automatically")
    print("Please install manually:")
    print("  Ubuntu/Debian: sudo apt-get install wkhtmltopdf")
    print("  CentOS/RHEL: sudo yum install wkhtmltopdf")
    print("  Fedora: sudo dnf install wkhtmltopdf")
    print("  Arch: sudo pacman -S wkhtmltopdf")
    return False

def install_wkhtmltopdf_windows():
    """Install wkhtmltopdf on Windows"""
    print("Attempting to install wkhtmltopdf on Windows...")
    
    # Check if Chocolatey is available
    if shutil.which('choco'):
        print("Using Chocolatey package manager...")
        success, stdout, stderr = run_command(['choco', 'install', 'wkhtmltopdf', '-y'])
        if success:
            print("✓ wkhtmltopdf installed successfully via Chocolatey")
            return True
        else:
            print(f"Failed with Chocolatey: {stderr}")
    
    # Check if winget is available
    if shutil.which('winget'):
        print("Using winget package manager...")
        success, stdout, stderr = run_command(['winget', 'install', 'wkhtmltopdf'])
        if success:
            print("✓ wkhtmltopdf installed successfully via winget")
            return True
        else:
            print(f"Failed with winget: {stderr}")
    
    print("✗ Could not install wkhtmltopdf automatically")
    print("Please install manually:")
    print("  1. Download installer from: https://wkhtmltopdf.org/downloads.html")
    print("  2. Or install Chocolatey and run: choco install wkhtmltopdf")
    return False

def install_wkhtmltopdf():
    """Install wkhtmltopdf based on the current platform"""
    system = platform.system().lower()
    
    if system == 'darwin':
        return install_wkhtmltopdf_macos()
    elif system == 'linux':
        return install_wkhtmltopdf_linux()
    elif system == 'windows':
        return install_wkhtmltopdf_windows()
    else:
        print(f"✗ Unsupported platform: {system}")
        print("Please install wkhtmltopdf manually from: https://wkhtmltopdf.org/downloads.html")
        return False

def test_pdf_generation():
    """Test PDF generation with a simple HTML file"""
    print("\nTesting PDF generation...")
    
    # Create a simple test HTML file
    test_html = """<!DOCTYPE html>
<html>
<head><title>Test</title></head>
<body>
    <h1>PDF Generation Test</h1>
    <p>If you can see this in a PDF, the setup is working correctly!</p>
</body>
</html>"""
    
    test_dir = Path("test_pdf_output")
    test_dir.mkdir(exist_ok=True)
    
    html_file = test_dir / "test.html"
    pdf_file = test_dir / "test.pdf"
    
    try:
        # Write test HTML
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(test_html)
        
        # Generate PDF
        cmd = [
            'wkhtmltopdf',
            '--page-size', 'Letter',
            '--margin-top', '0.75in',
            '--margin-right', '0.75in',
            '--margin-bottom', '0.75in',
            '--margin-left', '0.75in',
            str(html_file),
            str(pdf_file)
        ]
        
        success, stdout, stderr = run_command(cmd)
        
        if success and pdf_file.exists():
            print("✓ PDF generation test successful")
            print(f"  Test PDF created: {pdf_file}")
            
            # Clean up test files
            try:
                html_file.unlink()
                pdf_file.unlink()
                test_dir.rmdir()
            except:
                pass
            
            return True
        else:
            print("✗ PDF generation test failed")
            if stderr:
                print(f"  Error: {stderr}")
            return False
    
    except Exception as e:
        print(f"✗ PDF generation test failed: {e}")
        return False

def main():
    print("=== Teams Call Flow PDF Setup ===")
    print("Checking system dependencies...\n")
    
    # Check Python version
    if not check_python_version():
        print("\nPlease upgrade to Python 3.6 or newer")
        sys.exit(1)
    
    # Check if wkhtmltopdf is already installed
    if check_wkhtmltopdf():
        if test_pdf_generation():
            print("\n✓ Setup complete! PDF generation is ready to use.")
            sys.exit(0)
    
    # Try to install wkhtmltopdf
    print("\nInstalling wkhtmltopdf...")
    if install_wkhtmltopdf():
        # Verify installation
        if check_wkhtmltopdf() and test_pdf_generation():
            print("\n✓ Setup complete! PDF generation is ready to use.")
            sys.exit(0)
    
    print("\n⚠ Setup incomplete - manual installation required")
    print("\nNext steps:")
    print("1. Install wkhtmltopdf manually from: https://wkhtmltopdf.org/downloads.html")
    print("2. Ensure wkhtmltopdf is in your system PATH")
    print("3. Run this script again to verify installation")
    
    system = platform.system().lower()
    if system == 'darwin':
        print("\nmacOS users can also try:")
        print("  brew install wkhtmltopdf")
    elif system == 'linux':
        print("\nLinux users can also try:")
        print("  sudo apt-get install wkhtmltopdf  # Ubuntu/Debian")
        print("  sudo yum install wkhtmltopdf      # CentOS/RHEL")
    elif system == 'windows':
        print("\nWindows users can also try:")
        print("  choco install wkhtmltopdf")
        print("  winget install wkhtmltopdf")
    
    sys.exit(1)

if __name__ == "__main__":
    main()