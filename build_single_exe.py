# -*- coding: utf-8 -*-
"""
Script for building a single .exe file
Output: install/IRealJade.exe
"""
import os
import sys
import subprocess
import shutil

# Fix console encoding for Thai Windows
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def main():
    print("=" * 60)
    print("  Build IRealJade.exe (Single File) -> install folder")
    print("=" * 60)

    project_dir = os.path.dirname(os.path.abspath(__file__))
    install_dir = os.path.join(project_dir, "install")
    
    # ========== Step 1: Check PyInstaller ==========
    print("\n[1/4] Checking PyInstaller...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--version"],
            capture_output=True, text=True, check=True
        )
        print(f"      OK - PyInstaller version {result.stdout.strip()}")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("      Not found. Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("      OK - PyInstaller installed!")

    # ========== Step 2: Create install folder ==========
    print("\n[2/4] Creating install folder...")
    os.makedirs(install_dir, exist_ok=True)
    print(f"      OK - Folder: {install_dir}")

    # ========== Step 3: Build with PyInstaller ==========
    print("\n[3/4] Building .exe with PyInstaller... (may take 2-5 minutes)")
    
    # Collect --add-data args
    add_data_args = []
    img_template_dir = os.path.join(project_dir, "img_template")
    if os.path.isdir(img_template_dir):
        add_data_args += ["--add-data", f"{img_template_dir};img_template"]
        print(f"      Including img_template/")

    config_file = os.path.join(project_dir, "config.json")
    if os.path.isfile(config_file):
        add_data_args += ["--add-data", f"{config_file};."]
        print(f"      Including config.json")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--noconsole",
        "--name", "IRealJade",
        "--distpath", install_dir,
        "--workpath", os.path.join(project_dir, "build", "_build_temp"),
        "--specpath", os.path.join(project_dir, "build"),
        "--hidden-import", "customtkinter",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL._tkinter_finder",
        "--collect-all", "customtkinter",
    ] + add_data_args + [
        os.path.join(project_dir, "app.pyw")
    ]

    print(f"\n      Command: {' '.join(cmd)}\n")
    
    try:
        subprocess.run(cmd, check=True, cwd=project_dir)
    except subprocess.CalledProcessError as e:
        print(f"\n      FAILED: PyInstaller error: {e}")
        sys.exit(1)

    # ========== Step 4: Verify result ==========
    print("\n[4/4] Verifying result...")
    exe_path = os.path.join(install_dir, "IRealJade.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"\n{'=' * 60}")
        print(f"  SUCCESS! .exe file created!")
        print(f"  File: {exe_path}")
        print(f"  Size: {size_mb:.1f} MB")
        print(f"{'=' * 60}")
    else:
        print(f"\n      FAILED: File not found: {exe_path}")
        sys.exit(1)

if __name__ == "__main__":
    main()
