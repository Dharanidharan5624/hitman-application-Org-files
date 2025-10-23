import os
import sys
import subprocess
from unittest import result
import urllib.request
import platform
import time
import ctypes
import shutil
import zipfile
import configparser
import socket
import threading
import stat
import json
import ast
import requests
import importlib
import ctypes
import psutil
from pathlib import Path
from collections import defaultdict

EXE_PATH = r"C:\exe_files"
if not os.path.exists(EXE_PATH):
            os.makedirs(EXE_PATH)
LOG_FILE = os.path.join(EXE_PATH,"xampp_install_log1.txt")

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

_original_print = print

def print(*args, **kwargs):
    """Override print to also log into a file"""
    message = " ".join(str(a) for a in args)
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')

    # Log to file
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        _original_print(f"Logging failed: {e}")

    # Still print to console
    _original_print(*args, **kwargs)


try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False
    if not getattr(sys, 'frozen', False):
        print("Warning: tkinter not available, attempting to install...")
        subprocess.run([sys.executable, "-m", "pip", "install", "tk"], capture_output=True)
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox, ttk
            GUI_AVAILABLE = True
        except ImportError:
            GUI_AVAILABLE = False
            print("Warning: Failed to install tkinter, using default directory")
    else:
        print("Error: tkinter not available in bundled executable")
try:
    import mysql.connector
    from mysql.connector import errorcode
except ImportError:
    if not getattr(sys, 'frozen', False):
        print("Installing mysql-connector-python...")
        subprocess.run([sys.executable, "-m", "pip", "install", "mysql-connector-python"], capture_output=True)
        try:
            import mysql.connector
            from mysql.connector import errorcode
        except ImportError:
            print("Error: Failed to install mysql-connector-python")
            sys.exit(1)
    else:
        print("Error: mysql-connector-python not available in bundled executable")
        sys.exit(1)
if platform.system().lower() == "linux":
    try:
        import pyodbc
    except ImportError:
        if not getattr(sys, 'frozen', False):
            print("Installing pyodbc for Linux...")
            subprocess.run([sys.executable, "-m", "pip", "install", "pyodbc"], capture_output=True)
            try:
                import pyodbc
            except ImportError:
                print("Error: Failed to install pyodbc")
                sys.exit(1)
        else:
            print("Error: pyodbc not available in bundled executable")
            sys.exit(1)
CURRENT_OS = platform.system().lower()
IS_WINDOWS = CURRENT_OS == "windows"
IS_MACOS = CURRENT_OS == "darwin"
IS_LINUX = CURRENT_OS == "linux"

#LOG_FILE = None

winreg = None
db_username, db_password = None, None
print("Detected OS:", IS_WINDOWS)
if IS_WINDOWS:
    import ctypes
    try:
        import winreg as _winreg
        winreg = _winreg
    except ImportError:
        winreg = None
else:
    winreg = None
print("winreg available:", winreg)

db_folder_path = None
REPO_URL = "https://github.com/Dharanidharan5624/Hitman_Testing.git"
CONFIG_TARGET_DIR = None
#LOG_FILE = None

BG_COLOR = "white"
BTN_COLOR = "#0078D7"
BTN_TEXT = "white"
LABEL_COLOR = "#0078D7"

REQUIRED_FOLDERS = [ 'database', 'powerbuilder']
if IS_WINDOWS:
   
    PYTHON_INSTALLER_FILE = "python-3.13.5-amd64.exe"
    PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.13.5/python-3.13.5-amd64.exe"
elif IS_MACOS:
  
    PYTHON_INSTALLER_FILE = "python-3.13.5-macos11.pkg"
    PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.13.5/python-3.13.5-macos11.pkg"
else: 
    
    PYTHON_INSTALLER_FILE = "Python-3.13.5.tgz"
    PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.13.5/Python-3.13.5.tgz"

if IS_WINDOWS:
    DEFAULT_TARGET_DIR = r"C:\HitmanEdge"
    WEB_SERVER_PATH = r"C:\xampp"
    VCREDIST_INSTALL_PATH = r"C:\HitmanEdge\vcredist"
    PYTHON_INSTALL_PATHS = [
        r"C:\Program Files\Python313",
        os.path.expanduser(r"~\AppData\Local\Programs\Python"),
    ]
    PYTHON_DEFAULT_INSTALL = r"C:\Program Files\Python313"
    XAMPP_INSTALLER_URL = "https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/8.2.12/xampp-windows-x64-8.2.12-0-VS16-installer.exe"
    INSTALLER_FILE_XAMPP = os.path.join(EXE_PATH, "xampp_installer.exe")
    MYSQL_ODBC_URL = "https://downloads.mysql.com/archives/get/p/10/file/mysql-connector-odbc-8.0.42-win32.msi"
    INSTALLER_FILE_ODBC = os.path.join(EXE_PATH, "odbc_installer.msi")
    # New: AutoIt constants
    AUTOIT_INSTALL_PATH = r"C:\Program Files (x86)\AutoIt3"
    AUTOIT_INSTALLER_URL = "https://www.autoitscript.com/files/autoit3/autoit-v3-setup.exe"
    AUTOIT_INSTALLER_FILE = os.path.join(EXE_PATH, "autoit-v3-setup.exe")
 
elif IS_MACOS:
    DEFAULT_TARGET_DIR = os.path.expanduser("~/HitmanEdge")
    WEB_SERVER_PATH = "/Applications/XAMPP"
    PYTHON_INSTALL_PATHS = [
        "/usr/local/bin",
        "/opt/homebrew/bin",
        "/Library/Frameworks/Python.framework/Versions/Current/bin",
        "/usr/bin",
    ]
    PYTHON_DEFAULT_INSTALL = "/usr/local/bin"
    XAMPP_INSTALLER_URL = "https://sourceforge.net/projects/xampp/files/XAMPP%20Mac%20OS%20X/8.2.12/xampp-osx-8.2.12-0-installer.dmg"
    INSTALLER_FILE = "xampp_installer.dmg"
    MYSQL_ODBC_URL = "https://archive.mariadb.org/connector-odbc-3.1.22/mariadb-connector-odbc-3.1.22-macos-x86-64.pkg"
else:  # Linux
    DEFAULT_TARGET_DIR = os.path.expanduser("~/HitmanEdge")
    WEB_SERVER_PATH = "/opt/lampp"
    PYTHON_INSTALL_PATHS = [
        "/usr/bin",
        "/usr/local/bin",
        "/opt/python/bin",
        "/home/{}/.local/bin".format(os.environ.get('USER', 'user')),
    ]
    PYTHON_DEFAULT_INSTALL = "/usr/local/bin"
    XAMPP_INSTALLER_URL = "https://sourceforge.net/projects/xampp/files/XAMPP%20Linux/8.2.12/xampp-linux-x64-8.2.12-0-installer.run"
    INSTALLER_FILE = "xampp_installer.run"
    MYSQL_ODBC_URL = "https://archive.mariadb.org/connector-odbc-3.1.22/mariadb-connector-odbc-3.1.22-linux-x86-64.rpm"


VCREDIST_PACKAGES = {
    "2015-2022_x86": {
        "url": "https://aka.ms/vs/17/release/vc_redist.x86.exe",
        "name": "Microsoft Visual C++ 2015-2022 Redistributable (x86)",
        "file": "vc_redist_2015_2022_x86.exe",
        "registry_keys": [
            r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X86",
            r"SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{*Microsoft Visual C++ 2015-2022*x64*}",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{*Microsoft Visual C++ 2015-2022*x64*}"
        ],
        "min_version": "14.42.34438"
    }
}


def detect_system_info():
    system_info = {
        "os": CURRENT_OS,
        "version": platform.release(),
        "machine": platform.machine(),
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "architecture": platform.architecture()[0],
        "processor": platform.processor()
    }
    
    if IS_WINDOWS:
        system_info["windows_version"] = platform.win32_ver()[0]
        try:
            system_info["windows_edition"] = platform.win32_edition()
        except:
            system_info["windows_edition"] = "Unknown"
        system_info["is_64bit"] = platform.machine().endswith('64')
    elif IS_LINUX:
        try:
            with open('/etc/os-release', 'r') as f:
                lines = f.readlines()
                for line in lines:
                    if line.startswith('PRETTY_NAME'):
                        system_info["distro"] = line.split('=')[1].strip().replace('"', '')
                        break
        except:
            system_info["distro"] = "Unknown Linux"
    elif IS_MACOS:
        system_info["macos_version"] = platform.mac_ver()[0]
    
    return system_info

def get_windows_version_info():
    if not IS_WINDOWS:
        return None
    try:
        version = platform.win32_ver()[0]
        version_mapping = {
            "10": "Windows 10/11",
            "6.3": "Windows 8.1",
            "6.2": "Windows 8",
            "6.1": "Windows 7",
            "6.0": "Windows Vista",
            "5.2": "Windows XP 64-bit",
            "5.1": "Windows XP"
        }
        return {
            "version": version,
            "friendly_name": version_mapping.get(version, f"Windows {version}"),
            "is_64bit": platform.machine().endswith('64')
        }
    except:
        return {
            "version": "Unknown",
            "friendly_name": "Unknown Windows",
            "is_64bit": platform.machine().endswith('64')
        }
    
def get_windows_architecture():
    """Force only x86 package requirement"""
    print_status("Forcing installation of x86 (32-bit) package only")
    return ['2015-2022_x86']

def parse_version_string(version_str):
    """Parse version string into comparable tuple"""
    if not version_str:
        return (0, 0, 0, 0)
    try:
        parts = version_str.replace('v', '').split('.')
        return tuple(int(part) for part in parts[:4])
    except (ValueError, AttributeError):
        return (0, 0, 0, 0)
    

def check_vcredist_version_registry(package_key):
    global winreg
    """Check if VC++ redistributable is installed with proper version"""
    if not IS_WINDOWS or not winreg:
        return False, None
    
    package_info = VCREDIST_PACKAGES[package_key]
    min_version = parse_version_string(package_info['min_version'])
    
    print_status(f"Checking {package_info['name']}...")

    # Check runtime registry entries
    runtime_keys = [key for key in package_info['registry_keys'] if 'Runtimes' in key]
    for registry_path in runtime_keys:
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
            try:
                version_value, _ = winreg.QueryValueEx(registry_key, "Version")
                installed_value, _ = winreg.QueryValueEx(registry_key, "Installed")
                
                if installed_value == 1:
                    installed_version = parse_version_string(version_value)
                    version_str = '.'.join(map(str, installed_version))
                    
                    if installed_version >= min_version:
                        print_status(f"Found {package_info['name']} v{version_str} (Compatible)", "SUCCESS")
                        winreg.CloseKey(registry_key)
                        return True, version_str
                    else:
                        print_status(f"Found {package_info['name']} v{version_str} (Outdated)", "WARNING")
                        winreg.CloseKey(registry_key)
                        return False, version_str
            except FileNotFoundError:
                pass
            finally:
                winreg.CloseKey(registry_key)
        except (OSError, FileNotFoundError):
            continue
    
    print_status(f"{package_info['name']} not found or incompatible", "WARNING")
    return False, None

def download_vcredist_package(package_key, download_dir):
    """Download VC++ redistributable package"""
    package_info = VCREDIST_PACKAGES[package_key]
    file_path = os.path.join(download_dir, package_info['file'])
    
    if os.path.exists(file_path):
        print_status(f"Using existing installer: {package_info['file']}", "INFO")
        return file_path
    
    print_status(f"Downloading {package_info['name']}...")
    try:
        response = requests.get(package_info['url'], stream=True)
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        progress = (downloaded / total_size) * 100
                        print_status(f"Download progress: {progress:.1f}%", "INFO")
        
        print_status(f"Downloaded successfully: {package_info['file']}", "SUCCESS")
        return file_path
    except Exception as e:
        print_status(f"Download failed for {package_info['name']}: {e}", "ERROR")
        return None
def install_vcredist_package(installer_path, package_key):
    """Install VC++ redistributable package"""
    if not installer_path or not os.path.exists(installer_path):
        return False
    
    package_info = VCREDIST_PACKAGES[package_key]
    print_status(f"Installing {package_info['name']}...")
    try:
        cmd = [installer_path, '/quiet', '/norestart']
        print_status("Running installer (this may take a few minutes)...", "INFO")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print_status(f"Installation successful: {package_info['name']}", "SUCCESS")
            return True
        elif result.returncode == 1638:
            print_status(f"Already installed (newer version): {package_info['name']}", "SUCCESS")
            return True
        elif result.returncode == 3010:
            print_status(f"Installation successful (reboot required): {package_info['name']}", "SUCCESS")
            return True
        else:
            error_msg = result.stderr.strip() if result.stderr else f"Exit code: {result.returncode}"
            print_status(f"Installation failed: {error_msg}", "ERROR")
            return False
    except subprocess.TimeoutExpired:
        print_status(f"Installation timed out: {package_info['name']}", "ERROR")
        return False
    except Exception as e:
        print_status(f"Installation error: {e}", "ERROR")
        return False

def verify_vcredist_post_install(package_key):
    """Verify installation after install attempt"""
    print_status(f"Verifying installation: {VCREDIST_PACKAGES[package_key]['name']}")
    import time
    time.sleep(2)
    is_installed, version = check_vcredist_version_registry(package_key)
    if is_installed:
        print_status(f"Verification successful: v{version}", "SUCCESS")
        return True
    else:
        print_status("Verification failed - package may not be properly installed", "WARNING")
        return False
    

def determine_required_vcredist_packages():
    if not IS_WINDOWS:
        return []
    
    win_info = get_windows_version_info()
    required_packages = []
    
    if win_info["is_64bit"]:
        required_packages.extend([
            "2015-2022_x86"
        ])
    else:
        required_packages.append("2015-2022_x86")
    
    return required_packages

def check_vcredist_installed(package_key):
    global winreg
    if not IS_WINDOWS or not winreg:
        return False
    
    package_year = package_key.split('_')[0]
    
    registry_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    for registry_path in registry_paths:
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
            for i in range(winreg.QueryInfoKey(registry_key)[0]):
                try:
                    subkey_name = winreg.EnumKey(registry_key, i)
                    subkey = winreg.OpenKey(registry_key, subkey_name)
                    try:
                        display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                        if "Microsoft Visual C++" in display_name and package_year in display_name:
                            winreg.CloseKey(subkey)
                            winreg.CloseKey(registry_key)
                            return True
                    except FileNotFoundError:
                        pass
                    winreg.CloseKey(subkey)
                except (OSError, FileNotFoundError):
                    continue
            winreg.CloseKey(registry_key)
        except (OSError, FileNotFoundError):
            continue
    return False

def install_all_vcredist_packages():
    """Main function to check and install x86 VC++ redistributable"""
    if not IS_WINDOWS:
        print_status("Visual C++ Redistributables are only required on Windows", "INFO")
        return True
    
    print_status("=== Microsoft Visual C++ Redistributables Setup (x86 only) ===")
    
    required_packages = get_windows_architecture()
    download_dir = os.path.join(EXE_PATH, "vcredist")
    os.makedirs(download_dir, exist_ok=True)
    
    installation_results = {}
    for package_key in required_packages:
        is_installed, version = check_vcredist_version_registry(package_key)
        installation_results[package_key] = {
            'initially_installed': is_installed,
            'initial_version': version,
            'download_success': False,
            'install_success': False,
            'final_verification': False
        }
    
    for package_key in required_packages:
        if installation_results[package_key]['initially_installed']:
            installation_results[package_key]['install_success'] = True
            installation_results[package_key]['final_verification'] = True
            continue
        
        installer_path = download_vcredist_package(package_key, download_dir)
        if installer_path:
            installation_results[package_key]['download_success'] = True
            install_success = install_vcredist_package(installer_path, package_key)
            installation_results[package_key]['install_success'] = install_success
            
            if install_success:
                verification_success = verify_vcredist_post_install(package_key)
                installation_results[package_key]['final_verification'] = verification_success
            
            try:
                os.remove(installer_path)
                print_status(f"Cleaned up installer: {os.path.basename(installer_path)}", "INFO")
            except:
                pass
    
    print_status("=== Installation Summary ===")
    successful_installs = 0
    total_packages = len(required_packages)
    
    for package_key, results in installation_results.items():
        package_name = VCREDIST_PACKAGES[package_key]['name']
        if results['final_verification'] or results['initially_installed']:
            print_status(f" {package_name} - Ready", "SUCCESS")
            successful_installs += 1
        else:
            print_status(f" {package_name} - Failed", "ERROR")
    
    success_rate = (successful_installs / total_packages) * 100
    print_status(f"Overall success rate: {successful_installs}/{total_packages} ({success_rate:.1f}%)")
    
    return successful_installs == total_packages

def install_autoit():
    """Download and install/reinstall AutoIt on Windows."""
    if not IS_WINDOWS:
        print_status("AutoIt is only required on Windows", "INFO")
        return True
    
    print_status("=== AutoIt Setup ===")
    
    # Check if AutoIt is installed
    installed = os.path.exists(AUTOIT_INSTALL_PATH)
    action = "Reinstalling" if installed else "Installing"
    print_status(f"{action} AutoIt...", "INFO")
    
    # Download installer if not exists
    if os.path.exists(AUTOIT_INSTALLER_FILE):
        print_status(f"Using existing installer: {AUTOIT_INSTALLER_FILE}", "INFO")
    else:
        print_status(f"Downloading AutoIt from {AUTOIT_INSTALLER_URL}...")
        try:
            response = requests.get(AUTOIT_INSTALLER_URL, stream=True)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(AUTOIT_INSTALLER_FILE, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            progress = (downloaded / total_size) * 100
                            print_status(f"Download progress: {progress:.1f}%", "INFO")
            
            print_status("AutoIt installer downloaded successfully", "SUCCESS")
        except Exception as e:
            print_status(f"Download failed: {e}", "ERROR")
            return False
    
    # Install silently
    try:
        print_status("Running AutoIt installer (this may take a few minutes)...", "INFO")
        result = subprocess.run([AUTOIT_INSTALLER_FILE, '/S'], capture_output=True, text=True)
        
        if result.returncode == 0:
            print_status("AutoIt installation/reinstallation successful", "SUCCESS")
        else:
            error_msg = result.stderr.strip() if result.stderr else f"Exit code: {result.returncode}"
            print_status(f"Installation failed: {error_msg}", "ERROR")
            return False
    except Exception as e:
        print_status(f"AutoIt installation error: {e}", "ERROR")
        return False
    
    # Verify installation
    time.sleep(2)  # Wait for installation to complete
    if os.path.exists(AUTOIT_INSTALL_PATH):
        print_status("AutoIt verification successful", "SUCCESS")
    else:
        print_status("AutoIt verification failed", "ERROR")
        return False
    
    # Clean up installer
    try:
        os.remove(AUTOIT_INSTALLER_FILE)
        print_status("Cleaned up AutoIt installer", "INFO")
    except Exception as e:
        print_status(f"Failed to clean up installer: {e}", "WARNING")
    
    return True

def install_platform_dependencies():
    if IS_WINDOWS:
        return install_all_vcredist_packages()
    elif IS_MACOS:
        return install_macos_dependencies()
    elif IS_LINUX:
        return install_linux_dependencies()
    else:
        print_status("Unsupported platform detected, no dependencies installed", "WARNING")
        return True

def install_macos_dependencies():
    print_status("Installing dependencies for macOS...")
   
    try:
        # Install MySQL ODBC driver prerequisites
        print_status("Please install unixODBC manually using brew: brew install unixodbc", "WARNING")
        # Assume user does it
       
        # Install MySQL ODBC driver
        print_status("Please download and install MariaDB ODBC driver from: " + MYSQL_ODBC_URL, "WARNING")
        # Assume user does it
       
        # Install pyodbc
        try:
            import pyodbc
        except ImportError:
            if not getattr(sys, 'frozen', False):
                print_status("pyodbc not found, installing...", "INFO")
                subprocess.run([sys.executable, "-m", "pip", "install", "pyodbc"], check=True)
                print_status("pyodbc installed successfully", "SUCCESS")
            else:
                print_status("Error: pyodbc not available in bundled executable", "ERROR")
                return False
       
        return True
    except Exception as e:
        print_status(f"macOS dependency installation error: {e}", "ERROR")
        return False

def install_linux_dependencies():
    print_status("Installing dependencies for Linux...")
   
    try:
        
        print_status("Please install unixODBC using your package manager (e.g., sudo apt install unixodbc unixodbc-dev)", "WARNING")
        # Assume user does it
       
        # Install MySQL ODBC driver
        print_status("Please download and install MariaDB ODBC driver from: " + MYSQL_ODBC_URL, "WARNING")
        # Assume user does it
       
        # Install pyodbc
        try:
            import pyodbc
        except ImportError:
            if not getattr(sys, 'frozen', False):
                print_status("pyodbc not found, installing...", "INFO")
                subprocess.run([sys.executable, "-m", "pip", "install", "pyodbc"], check=True)
                print_status("pyodbc installed successfully", "SUCCESS")
            else:
                print_status("Error: pyodbc not available in bundled executable", "ERROR")
                return False
       
        return True
    except Exception as e:
        print_status(f"Linux dependency installation error: {e}", "ERROR")
        return False

def log_to_file(message):
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except:
        pass

def create_setup_lock():
    global CONFIG_TARGET_DIR
    if CONFIG_TARGET_DIR:
        lock_file_path = os.path.join(CONFIG_TARGET_DIR, "setup.lock")
        try:
            os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
            with open(lock_file_path, 'w') as f:
                f.write(f"Setup completed at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            os.chmod(lock_file_path, 0o644)
            return True
        except Exception as e:
            print_status(f"Could not create setup lock: {e}", "WARNING")
            return False
    return False

def check_setup_lock():
    global CONFIG_TARGET_DIR
    if CONFIG_TARGET_DIR:
        lock_file_path = os.path.join(CONFIG_TARGET_DIR, "setup.lock")
        return os.path.exists(lock_file_path)
    return False

def get_python_executable_path():
    return sys.executable

def print_status(message, status="INFO"):
    log_to_file(f"[{status}] {message}")
    if hasattr(InstallerApp, 'instance') and InstallerApp.instance:
        InstallerApp.instance.update_progress(f"[{status}] {message}")
        print(f"[{status}] {message}")
    else:
        print(f"[{status}] {message}")


def update_config_paths(config_path):
    global db_folder_path, CONFIG_TARGET_DIR
    if not CONFIG_TARGET_DIR:
        print_status("Target directory not set for config creation", "ERROR")
        return False

    print_status("Updating config.ini file...", "INFO")
    try:
        config_dir = os.path.dirname(config_path)
        os.makedirs(config_dir, exist_ok=True)
        config = configparser.ConfigParser()
        config.optionxform = str

        # Read existing config.ini if it exists
        if os.path.exists(config_path):
            config.read(config_path, encoding='utf-8')

        # Ensure sections exist
        if "database" not in config:
            config.add_section("database")
        if "app" not in config:
            config.add_section("app")
        if "paths" not in config:
            config.add_section("paths")

        # Check for existing Python installations
        print_status("Searching for Python installations...", "INFO")
        valid_pythons, _ = check_existing_python_installations()
        print_status(f"Valid Python installations found: {[p['path'] for p in valid_pythons]}", "INFO")
        python_exe = None

        # Handle bundled executable case
        is_bundled = getattr(sys, 'frozen', False)
        if is_bundled:
            print_status("Running as bundled executable, ignoring sys.executable", "INFO")
        else:
            # Check if sys.executable is a valid Python interpreter
            current_exe = os.path.normpath(sys.executable)
            if (os.path.isfile(current_exe) and 
                current_exe.lower().endswith('python.exe' if IS_WINDOWS else 'python') and 
                current_exe not in [p['path'] for p in valid_pythons]):
                try:
                    import subprocess
                    result = subprocess.run([current_exe, "--version"], capture_output=True, text=True, timeout=5)
                    if result.returncode == 0:
                        version_str = result.stdout.strip().replace("Python ", "")
                        version_parts = version_str.split(".")
                        if len(version_parts) >= 2 and tuple(map(int, version_parts[:2])) >= (3, 7):
                            valid_pythons.append({
                                'path': current_exe,
                                'version': version_str,
                                'valid': True,
                                'type': 'Current Runtime'
                            })
                            print_status(f"Added sys.executable to valid pythons: {current_exe}", "INFO")
                except Exception as e:
                    print_status(f"Failed to validate sys.executable {current_exe}: {e}", "WARNING")

        # Select the first valid Python executable
        if valid_pythons:
            for python in valid_pythons:
                candidate_path = os.path.normpath(python['path'])
                # Ensure the path is a Python executable (not the bundled .exe)
                if (os.path.isfile(candidate_path) and 
                    candidate_path.lower().endswith('python.exe' if IS_WINDOWS else 'python')):
                    python_exe = candidate_path
                    print_status(f"Selected Python executable: {python_exe}", "SUCCESS")
                    break
            if not python_exe:
                print_status("No valid Python executable found in valid_pythons", "ERROR")
        else:
            print_status("No valid Python executable found, attempting installation...", "INFO")
            success, python_path = ensure_python_compatibility()
            if success and os.path.isfile(python_path):
                python_exe = os.path.normpath(python_path)
                if python_exe.lower().endswith('python.exe' if IS_WINDOWS else 'python'):
                    print_status(f"Installed Python executable: {python_exe}", "SUCCESS")
                else:
                    print_status(f"Installed path is not a Python executable: {python_path}", "ERROR")
                    return False
            else:
                print_status(f"Failed to install Python or invalid path: {python_path}", "ERROR")
                return False

        # Ensure python_exe is set and valid
        if not python_exe:
            print_status("No valid Python executable available", "ERROR")
            return False

        # Update database section - preserve existing if present, else set defaults
        config["database"].setdefault("HE_HOSTNAME", "127.0.0.1")
        config["database"].setdefault("HE_PORT", "3306")
        config["database"].setdefault("HE_DB_USERNAME", "Hitman")
        config["database"].setdefault("HE_DB_PASSWORD", "Hitman@123")
        config["database"].setdefault("HE_DB_DEV", "hitman_edge_dev")
        config["database"].setdefault("HE_DB_TEST", "hitman_edge_test")
        config["database"].setdefault("HE_DB_PROD", "hitman_edge_prod")
        config["database"].setdefault("HE_ODBC_NAME", "hitman_edge_dev")
        config["database"].setdefault("HE_DBPARM", "ConnectString='DRIVER={MySQL ODBC 8.0 Unicode Driver};SERVER=127.0.0.1;PORT=3306;DATABASE=hitman_edge_dev;UID=Hitman;PWD=Hitman@123;OPTION=3;'")
        config["database"].setdefault("HE_AUTOCOMMIT", "false")
        config["database"].setdefault("HE_DBMS", "ODBC")

        # Update app section - preserve existing if present, else set defaults
        config["app"].setdefault("APP_NAME", "HitmanEdge")
        config["app"].setdefault("VERSION", "1.0.0")
        config["app"].setdefault("LANGUAGE", "en")
        config["app"].setdefault("TIMEZONE", "Asia/Kolkata")

        # Update paths section - these should be updated always
        config["paths"]["HE_ROOT_PATH"] = CONFIG_TARGET_DIR
        config["paths"]["HE_PYTHON_EXE"] = python_exe
        config["paths"]["HE_PYTHON_PATH"] = r"\script"
        config["paths"]["HE_IMAGE_PATH"] = r"\assets"
        config["paths"]["HE_PATH"] = config_path
        config["paths"]["HE_POWERBUILDER_EXE"] = os.path.join(CONFIG_TARGET_DIR, "powerbuilder")
        config["paths"]["HE_DATABASE"] = os.path.join(CONFIG_TARGET_DIR, "database")
        db_folder_path = config["paths"]["HE_DATABASE"]

        # Write the config file
        with open(config_path, 'w', encoding='utf-8') as f:
            config.write(f)
        os.chmod(config_path, 0o644)
        print_status(f"Config file updated successfully with HE_PYTHON_EXE={python_exe}", "SUCCESS")
        return True
    except Exception as e:
        print_status(f"Failed to update config file: {e}", "ERROR")
        return False
   
def create_default_config_if_missing():
    if not CONFIG_TARGET_DIR:
        return None
    config_folder = os.path.join(CONFIG_TARGET_DIR, "config")
    # config_folder = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config")
    config_path = os.path.join(config_folder, "config.ini")
    if os.path.exists(config_path):
        print_status("Existing config.ini found", "SUCCESS")
        return config_path
    print_status("Creating default config.ini...", "INFO")
    try:
        os.makedirs(config_folder, exist_ok=True)
        os.chmod(config_folder, 0o755)
        if update_config_paths(config_path):
            print_status("Default config.ini created", "SUCCESS")
            return config_path
        else:
            print_status("Failed to create default config.ini", "ERROR")
            return None
    except Exception as e:
        print_status(f"Error creating default config: {e}", "ERROR")
        return None

def find_config_ini(repo_path):
    """Find existing config.ini in the repository structure"""
    if not repo_path or not os.path.exists(repo_path):
        return None
    for root, _, files in os.walk(repo_path):
        for file in files:
            if file.lower() == 'config.ini':
                print_status(f"Found existing config.ini at: {os.path.relpath(os.path.join(root, file), repo_path)}", "SUCCESS")
                return os.path.join(root, file)
    print_status("No existing config.ini found in repository", "INFO")
    return None

def find_sql_file():
    global db_folder_path
    print(db_folder_path)
    if not db_folder_path or not os.path.exists(db_folder_path):
        print("Configured database folder does not exist.")
        return None

    for file in os.listdir(db_folder_path):
        if file.endswith('.sql'):
            sql_file = os.path.join(db_folder_path, file)
            print(f"Found SQL file: {sql_file}")
            return sql_file

    print("No SQL file found in database folder.")
    return None


def remove_readme_files():
    if not CONFIG_TARGET_DIR:
        return
    removed = 0
    for root, _, files in os.walk(CONFIG_TARGET_DIR):
        for file in files:
            if file.lower() == 'readme.md':
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    print_status(f"Removed README.md: {os.path.relpath(file_path, CONFIG_TARGET_DIR)}", "SUCCESS")
                    removed += 1
                except Exception as e:
                    print_status(f"Failed to remove {file_path}: {e}", "WARNING")
    if removed > 0:
        print_status(f"Removed {removed} README.md files", "SUCCESS")
    else:
        print_status("No README.md files found", "INFO")

def get_platform_info():
    info = {
        "system": platform.system(),
        "release": platform.release(),
        "machine": platform.machine(),
        "python_version": sys.version.split()[0]
    }
    return info

def install_missing_packages_from_requirements(requirements_path=None, python_exe=sys.executable):
    if not requirements_path:
        requirements_path = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "script", "requirements.txt")
   
    if not os.path.exists(requirements_path):
        print_status("requirements.txt not found", "WARNING")
        return False
   
    print_status(f"Installing packages from {requirements_path}...")
    try:
        with open(requirements_path, 'r') as f:
            requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        
        # Prepare subprocess kwargs to hide console windows on Windows
        subprocess_kwargs = {
            'capture_output': True,
            'text': True,
        }
        if IS_WINDOWS:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = subprocess.SW_HIDE
            subprocess_kwargs['startupinfo'] = si
            subprocess_kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW

        # Install packages from requirements.txt
        for requirement in requirements:
            for cmd in [
                [python_exe, "-m", "pip", "install", requirement],
                [python_exe, "-m", "pip", "install", requirement, "--user"]
            ]:
                try:
                    result = subprocess.run(cmd, **subprocess_kwargs, check=True)
                    if result.returncode == 0:
                        print_status(f"Installed: {requirement}", "SUCCESS")
                        break
                    else:
                        print_status(f"Retrying {requirement} with alternative method...", "INFO")
                except Exception as e:
                    print_status(f"Error during installation attempt for {requirement}: {e}", "WARNING")
                    continue
            else:
                print_status(f"Failed to install {requirement}", "ERROR")

        # Final step: Install or upgrade websockets package
        print_status("Upgrading websockets package as final step...", "INFO")
        cmd_websockets = [python_exe, "-m", "pip", "install", "--upgrade", "websockets"]
        try:
            result = subprocess.run(cmd_websockets, **subprocess_kwargs, check=True)
            if result.returncode == 0:
                print_status("Upgraded websockets successfully", "SUCCESS")
            else:
                print_status(f"Failed to upgrade websockets: {result.stderr}", "ERROR")
                return False
        except Exception as e:
            print_status(f"Error upgrading websockets: {e}", "ERROR")
            return False

        return True
    except Exception as e:
        print_status(f"Package installation failed: {e}", "ERROR")
        return False

def check_existing_python_installations():
    """
    Check for valid Python installations in allowed paths.
    Returns a list of dicts with valid Python installations and all paths checked.
    """
    valid_pythons = []
    all_paths = []

    for path in PYTHON_INSTALL_PATHS:
        candidate = os.path.join(path, "python.exe" if IS_WINDOWS else "python3")
        all_paths.append(candidate)
        if os.path.isfile(candidate):
            try:
                result = subprocess.run([candidate, "--version"], capture_output=True, text=True, timeout=5,check=True)
                if result.returncode == 0:
                    version_str = result.stdout.strip().replace("Python ", "")
                    version_parts = version_str.split(".")
                    if len(version_parts) >= 2 and tuple(map(int, version_parts[:2])) >= (3, 7):
                        valid_pythons.append({
                            'path': candidate,
                            'version': version_str,
                            'valid': True,
                            'type': 'System'
                        })
            except Exception as e:
                print_status(f"Failed to validate Python at {candidate}: {e}", "WARNING")

    return valid_pythons, all_paths


def ensure_python_compatibility():
    """
    Check for compatible Python installation, install if needed.
    Returns a tuple (success: bool, python_path: str).
    """
    print_status("Checking Python compatibility...", "INFO")

    # Step 1: Check for existing Python installations
    python_exe = None
    for path in PYTHON_INSTALL_PATHS:
        candidate = os.path.join(path, "python.exe" if IS_WINDOWS else "python3")
        if os.path.isfile(candidate):
            try:
                result = subprocess.run([candidate, "--version"], capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    version_str = result.stdout.strip().replace("Python ", "")
                    version_parts = version_str.split(".")
                    if len(version_parts) >= 2 and tuple(map(int, version_parts[:2])) >= (3, 7):
                        print_status(f"Found compatible Python: {candidate} (v{version_str})", "SUCCESS")
                        return True, candidate
            except Exception as e:
                print_status(f"Failed to validate Python at {candidate}: {e}", "WARNING")
                continue

    # Step 2: No compatible Python found, proceed to install
    print_status("No compatible Python found. Installing...", "WARNING")

    # Create download directory
    installer_dir = EXE_PATH
    os.makedirs(installer_dir, exist_ok=True)
    installer_path = os.path.join(installer_dir, PYTHON_INSTALLER_FILE)

    # Step 3: Download Python installer
    if not os.path.exists(installer_path):
        print_status(f"Downloading Python from {PYTHON_INSTALLER_URL}...", "INFO")
        try:
            response = requests.get(PYTHON_INSTALLER_URL, stream=True, verify=True)
            response.raise_for_status()
            with open(installer_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print_status("Python installer downloaded successfully", "SUCCESS")
        except Exception as e:
            print_status(f"Failed to download Python installer: {e}", "ERROR")
            return False, None

    # Step 4: Install Python
    print_status(f"Installing Python to {PYTHON_DEFAULT_INSTALL}...", "INFO")
    try:
        if IS_WINDOWS:
            cmd = [
                installer_path,
                "/quiet",
                "InstallAllUsers=1",
                f"TargetDir={PYTHON_DEFAULT_INSTALL}",
                "PrependPath=1",
                "Include_test=0",
                "Include_pip=1"
            ]
        elif IS_MACOS:
            cmd = ["sudo", "installer", "-pkg", installer_path, "-target", "/"]
        else:  # Linux
            cmd = ["sudo", "bash", installer_path, "--no-download"]

        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        if result.returncode == 0:
            print_status("Python installation completed", "SUCCESS")
        else:
            print_status(f"Python installation failed: {result.stderr}", "ERROR")
            return False, None

        # Step 5: Verify installation
        python_exe = os.path.join(PYTHON_DEFAULT_INSTALL, "python.exe" if IS_WINDOWS else "python3")
        if os.path.isfile(python_exe):
            result = subprocess.run([python_exe, "--version"], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print_status(f"Python verified: {python_exe} (v{result.stdout.strip()})", "SUCCESS")
                try:
                    os.remove(installer_path)
                    print_status("Cleaned up installer", "INFO")
                except Exception as e:
                    print_status(f"Failed to clean up installer: {e}", "WARNING")
                return True, python_exe
        print_status("Python installation verification failed", "ERROR")
        return False, None
    except Exception as e:
        print_status(f"Python installation error: {e}", "ERROR")
        return False, None

    
       

def check_system_compatibility():
    system_info = detect_system_info()
    print_status(f"Detected OS: {system_info['os']} {system_info['version']}")
    print_status(f"Architecture: {system_info['architecture']}")
    print_status(f"Current Python: {system_info['python_version']}")
    
    compatible = True
    
    if IS_WINDOWS:
        if system_info.get("windows_version"):
            print_status(f"Windows Version: {system_info['windows_version']}")
        if system_info.get("windows_edition"):
            print_status(f"Windows Edition: {system_info['windows_edition']}")
        compatible = True
    elif IS_LINUX:
        print_status(f"Linux Distribution: {system_info.get('distro', 'Unknown')}")
        compatible = True
    elif IS_MACOS:
        if system_info.get("macos_version"):
            print_status(f"macOS Version: {system_info['macos_version']}")
        compatible = True
    else:
        print_status("Unsupported operating system", "ERROR")
        compatible = False
    
    return compatible, system_info


def check_port_availability(port, service_name="Service"):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(2)
            result = s.connect_ex(('localhost', port))
            if result == 0:
                print_status(f"Port {port} is occupied ({service_name} may be running)", "WARNING")
                return False
            else:
                print_status(f"Port {port} is available for {service_name}", "SUCCESS")
                return True
    except Exception as e:
        print_status(f"Port {port} check failed: {e}", "WARNING")
        return False



def configure_apache_port():
    print_status("Configuring Apache port settings...")

    if IS_WINDOWS:
        httpd_conf_path = os.path.join(WEB_SERVER_PATH, "apache", "conf", "httpd.conf")
        apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")
    else:
        httpd_conf_path = os.path.join(WEB_SERVER_PATH, "etc", "httpd.conf")
        apache_exe = os.path.join(WEB_SERVER_PATH, "bin", "httpd")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "bin", "mysqld")

    if not os.path.exists(httpd_conf_path):
        print_status("Apache configuration file not found", "ERROR")
        return False, 80

    port_80_available = check_port_availability(80, "Web Server")
    if port_80_available:
        print_status("Port 80 available - Apache will use default port", "SUCCESS")
        apache_port = 80
    else:
        print_status("Configuring Apache to use port 8080...")
        try:
            with open(httpd_conf_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            backup_path = httpd_conf_path + f".backup_{int(time.time())}"
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(content)
            if not IS_WINDOWS:
                os.chmod(backup_path, 0o644)
            print_status(f"Configuration backup created: {os.path.basename(backup_path)}", "SUCCESS")

            if "Listen 8080" in content and "ServerName localhost:8080" in content:
                print_status("Apache already configured for port 8080", "SUCCESS")
                apache_port = 8080
            else:
                lines = content.split('\n')
                modified_lines = []
                changes_made = 0

                for line in lines:
                    if line.strip().startswith("Listen 80") and not line.strip().startswith("Listen 8080"):
                        line = line.replace("Listen 80", "Listen 8080")
                        changes_made += 1
                        print_status("Updated: Listen 80 to Listen 8080", "SUCCESS")
                    elif "ServerName localhost:80" in line:
                        line = line.replace("ServerName localhost:80", "ServerName localhost:8080")
                        changes_made += 1
                        print_status("Updated ServerName to port 8080", "SUCCESS")
                    elif ":80>" in line:
                        line = line.replace(":80>", ":8080>")
                        changes_made += 1
                        print_status("Updated virtual host to port 8080", "SUCCESS")
                    modified_lines.append(line)

                if changes_made > 0:
                    with open(httpd_conf_path, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(modified_lines))
                    if not IS_WINDOWS:
                        os.chmod(httpd_conf_path, 0o644)
                    print_status(f"Apache configuration updated ({changes_made} changes)", "SUCCESS")
                    print_status("Apache will run on http://localhost:8080")
                    apache_port = 8080
                else:
                    print_status("No port configurations found to modify", "WARNING")
                    apache_port = 80

        except Exception as e:
            print_status(f"Apache configuration failed: {e}", "ERROR")
            return False, 80


    if IS_WINDOWS:
        try:
            
            apache_service_name = "Apache-XAMPP"
            subprocess.run([apache_exe, "-k", "install", "-n", apache_service_name], check=True, shell=True)
            subprocess.run(["sc", "config", apache_service_name, "start=", "auto"], check=True, shell=True)
            subprocess.run(["net", "start", apache_service_name], check=True, shell=True)
            print_status(f"Apache service '{apache_service_name}' installed and started", "SUCCESS")
        except subprocess.CalledProcessError as e:
            print_status(f"Error installing/starting Apache service: {e}", "ERROR")

        try:
           
            mysql_service_name = "MySQL-XAMPP"
            subprocess.run([mysql_exe, "--install", mysql_service_name], check=True,shell=True)
            subprocess.run(["sc", "config", mysql_service_name, "start=", "auto"], check=True,shell=True)
            subprocess.run(["net", "start", mysql_service_name], check=True,shell=True)
            print_status(f"MySQL service '{mysql_service_name}' installed and started", "SUCCESS")
        except subprocess.CalledProcessError as e:
            print_status(f"Error installing/starting MySQL service: {e}", "ERROR")

    return True, apache_port


def install_xampp_if_needed():
    installed, running = check_xampp_installation_and_running()
    if installed:
        print("XAMPP already installed.")
        if running:
            print("XAMPP services are running.")
        else:
            print("XAMPP installed but services not running. Starting them...")
            start_services_enhanced(db_port=3306, db_host='localhost')
        return True

    # else download and install
    r = requests.get(XAMPP_INSTALLER_URL, verify=True)
    with open(INSTALLER_FILE_XAMPP, "wb") as f:
        f.write(r.content)
    subprocess.run(INSTALLER_FILE_XAMPP, check=True, shell=True)
    return True

def check_xampp_installation():
    if IS_WINDOWS:
        apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
        print_status (f"Xamp check if apache_exe:  {apache_exe}")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")
        print_status (f"Xamp check if mysql_exe: {mysql_exe}")

    else:
        apache_exe = os.path.join(WEB_SERVER_PATH, "bin", "httpd")
        print_status (f"Xamp check else apache_exe: {apache_exe}")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "bin", "mysqld")
        print_status (f"Xamp check else mysql_exe: {mysql_exe}")
   
    xampp_installed = os.path.exists(WEB_SERVER_PATH) and os.path.exists(apache_exe) and os.path.exists(mysql_exe)
    print_status (f"Xamp check: {xampp_installed} {WEB_SERVER_PATH} {apache_exe} {mysql_exe}")
   
    return xampp_installed


def is_xampp_running():
    
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and ('httpd' in proc.info['name'].lower() or 'mysqld' in proc.info['name'].lower()):
            return True
    return False

def check_xampp_installation_and_running():
    installed = check_xampp_installation()
    running = is_xampp_running()
    return installed, running


def start_services_enhanced(db_port, db_host):
    print_status("Starting XAMPP services...")
   
    apache_success, apache_port = configure_apache_port()
    if not apache_success:
        print_status("Apache configuration failed", "ERROR")
        return False
   
    if IS_WINDOWS:
        apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")
        mysql_ini = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "my.ini")
    else:
        apache_exe = os.path.join(WEB_SERVER_PATH, "bin", "httpd")
        mysql_exe = os.path.join(WEB_SERVER_PATH, "bin", "mysqld")
        mysql_ini = os.path.join(WEB_SERVER_PATH, "etc", "my.cnf")
   
    if IS_LINUX:
        subprocess.run(["sudo", "systemctl", "stop", "apache2"], capture_output=True)
        subprocess.run(["sudo", "systemctl", "stop", "mysql"], capture_output=True)
        subprocess.run(["sudo", os.path.join(WEB_SERVER_PATH, "lampp"), "start"], capture_output=True)
    elif IS_MACOS:
        subprocess.run(["sudo", "launchctl", "unload", "/Library/LaunchDaemons/org.apache.httpd.plist"], capture_output=True)
        subprocess.run(["sudo", "launchctl", "unload", "/Library/LaunchDaemons/com.mysql.mysql.plist"], capture_output=True)
        subprocess.run(["sudo", os.path.join(WEB_SERVER_PATH, "xampp"), "start"], capture_output=True)
   
    print_status(f"Starting Apache on port {apache_port}...")
    max_retries = 3
    for attempt in range(max_retries):
        if not is_port_open(apache_port):
            try:
                if IS_WINDOWS:
                    subprocess.Popen([apache_exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    subprocess.Popen(["sudo", apache_exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
               
                print_status("Waiting for Apache startup...")
                time.sleep(10)
               
                if is_port_open(apache_port):
                    print_status(f"Apache started successfully on port {apache_port}", "SUCCESS")
                    break
                else:
                    print_status(f"Apache startup attempt {attempt + 1}/{max_retries} failed", "WARNING")
            except Exception as e:
                print_status(f"Apache startup error: {e}", "ERROR")
        else:
            print_status(f"Apache already running on port {apache_port}", "SUCCESS")
            break
    else:
        print_status("Apache startup failed after all retries", "ERROR")
   
    print_status(f"Starting MySQL on port {db_port}...")
    for attempt in range(max_retries):
        if not is_port_open(db_port, db_host):
            try:
                if IS_WINDOWS:
                    subprocess.Popen([mysql_exe, f"--defaults-file={mysql_ini}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    subprocess.Popen(["sudo", mysql_exe, f"--defaults-file={mysql_ini}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
               
                print_status("Waiting for MySQL startup...")
                time.sleep(30)
               
                if is_port_open(db_port, db_host):
                    print_status(f"MySQL started successfully on port {db_port}", "SUCCESS")
                    return True
                else:
                    print_status(f"MySQL startup attempt {attempt + 1}/{max_retries} failed", "WARNING")
            except Exception as e:
                print_status(f"MySQL startup error: {e}", "ERROR")
        else:
            print_status(f"MySQL already running on port {db_port}", "SUCCESS")
            return True
    else:
        print_status("MySQL startup failed after all retries", "ERROR")
        return False

def is_port_open(port, host='localhost'):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            return s.connect_ex((host, port)) == 0
    except:
        return False

def download_and_extract_repo():
    try:
        print_status("Preparing repository download...")
        zip_url = REPO_URL.replace(".git", "") + "/archive/refs/heads/main.zip"
        temp_dir = os.path.join(CONFIG_TARGET_DIR, "temp_extract")
        zip_path = os.path.join(temp_dir, "repo.zip")
       
        os.makedirs(temp_dir, exist_ok=True)
       
        print_status("Downloading HitmanEdge repository...")
        urllib.request.urlretrieve(zip_url, zip_path)
       
        print_status("Extracting repository files...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
       
        os.remove(zip_path)
        extracted_folders = [d for d in os.listdir(temp_dir)
                           if os.path.isdir(os.path.join(temp_dir, d)) and 'main' in d.lower()]
       
        if not extracted_folders:
            raise Exception("Repository extraction failed: main folder not found")
        inner_folder = os.path.join(temp_dir, extracted_folders[0])
        
        # Ensure target directory exists
        os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
        
        # Extract only required folders
        for folder in REQUIRED_FOLDERS:
            src_folder = os.path.join(inner_folder, folder)
            dest_folder = os.path.join(CONFIG_TARGET_DIR, folder)
            if os.path.exists(src_folder):
                if os.path.exists(dest_folder):
                    shutil.rmtree(dest_folder)
                shutil.move(src_folder, dest_folder)
                print_status(f"Moved folder: {folder}", "SUCCESS")
            else:
                print_status(f"Folder not found in repository: {folder}", "WARNING")
        
        # Move config.ini from config/config.ini to powerbuilder/config.ini
        src_config = os.path.join(inner_folder, "config", "config.ini")
        dest_config_dir = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config")
        dest_config = os.path.join(dest_config_dir, "config.ini")
        dest_config_dir = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "assets")
        dest_config_dir = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "script")
       
        if os.path.exists(src_config):
            os.makedirs(dest_config_dir, exist_ok=True)
            shutil.move(src_config, dest_config)
            print_status(f"Moved config.ini to {dest_config}", "SUCCESS")
            # Update config.ini with correct paths
            update_config_paths(dest_config)
        else:
            print_status("config.ini not found in repository at config/config.ini", "ERROR")
            # Create a default config.ini if not found
            os.makedirs(dest_config_dir, exist_ok=True)
            update_config_paths(dest_config)
        
        # Clean up
        shutil.rmtree(temp_dir)
        print_status("Repository extracted and structure preserved successfully", "SUCCESS")
        return CONFIG_TARGET_DIR
    except Exception as e:
        print_status(f"Repository download/extraction failed: {e}", "ERROR")
        return None

def clean_config_folder():
    # Find config folder in the preserved structure and clean it
    config_folder = None
    for root, dirs, _ in os.walk(CONFIG_TARGET_DIR):
        if 'config' in [d.lower() for d in dirs]:
            config_folder = os.path.join(root, 'config')
            break
    
    if not config_folder or not os.path.exists(config_folder):
        print_status("Config folder not found in preserved structure", "WARNING")
        return
   
    config_ini_path = os.path.join(config_folder, "config.ini")
   
    print_status("Cleaning config folder - keeping only config.ini...")
   
    try:
        all_items = os.listdir(config_folder)
        config_ini_exists = os.path.exists(config_ini_path)
       
        if not config_ini_exists:
            print_status("config.ini not found in config folder", "ERROR")
            return
       
        removed_count = 0
        for item_name in all_items:
            item_path = os.path.join(config_folder, item_name)
           
            if item_name.lower() != "config.ini":
                try:
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                        removed_count += 1
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                        removed_count += 1
                except Exception as e:
                    print_status(f"Could not remove {item_name}: {e}", "WARNING")
       
        if removed_count > 0:
            print_status(f"Config folder cleaned - removed {removed_count} items", "SUCCESS")
        else:
            print_status("Config folder already clean", "SUCCESS")
    except Exception as e:
        print_status(f"Error cleaning config folder: {e}", "ERROR")

def load_config(config_path):
    if not config_path or not os.path.exists(config_path):
        raise FileNotFoundError("config.ini not found. Please ensure it exists in your GitHub repository.")
       
    print_status(f"Loading configuration: {os.path.basename(config_path)}")
   
    try:
        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
        if "database" not in config:
            raise KeyError("Missing [database] section in config.ini")
        db = config["database"]
        required_keys = [
            "HE_HOSTNAME", "HE_PORT", "HE_DB_USERNAME", "HE_DB_PASSWORD",
            "HE_DB_DEV", "HE_DB_TEST", "HE_DB_PROD", "HE_ODBC_NAME"
        ]
        print_status("Validating configuration keys...")
        for key in required_keys:
            if key not in db:
                raise KeyError(f"Missing key '{key}' in [database] section of config.ini")
        db_info = {
            "host": db["HE_HOSTNAME"],
            "port": int(db["HE_PORT"]),
            "user": db["HE_DB_USERNAME"],
            "pass": db["HE_DB_PASSWORD"],
            "names": [db["HE_DB_DEV"], db["HE_DB_TEST"], db["HE_DB_PROD"]],
            "dsn": db["HE_ODBC_NAME"]
        }
        print_status("Configuration Summary:")
        print_status(f" Host: {db_info['host']}:{db_info['port']}")
        print_status(f" User: {db_info['user']}")
        print_status(f" Databases: {', '.join(db_info['names'])}")
        print_status(f" DSN Name: {db_info['dsn']}")
       
        print_status("Configuration loaded successfully", "SUCCESS")
        return db_info
    except Exception as e:
        print_status(f"Configuration loading failed: {e}", "ERROR")
        raise

def create_user_and_dbs(info):
    print_status("Setting up MySQL databases and user...")
   
    max_retries = 3
    for attempt in range(max_retries):
        try:
            conn = mysql.connector.connect(
                host=info["host"],
                port=info["port"],
                user="root",
                password=""
            )
            print_status("MySQL connection established", "SUCCESS")
            break
        except Exception as e:
            print_status(f"MySQL connection attempt {attempt + 1}/{max_retries} failed: {e}", "WARNING")
            if attempt < max_retries - 1:
                print_status("Waiting 10 seconds before retry...", "INFO")
                time.sleep(10)
            else:
                print_status("MySQL connection failed after all retries", "ERROR")
                return False
   
    try:
        cur = conn.cursor()
       
        try:
            cur.execute(f"CREATE USER '{info['user']}'@'localhost' IDENTIFIED BY '{info['pass']}';")
            print_status(f"Created MySQL user: {info['user']}", "SUCCESS")
        except mysql.connector.Error as e:
            if e.errno == errorcode.ER_CANNOT_USER:
                print_status(f"User {info['user']} already exists", "SUCCESS")
            else:
                raise
       
        cur.execute(f"GRANT ALL PRIVILEGES ON *.* TO '{info['user']}'@'localhost';")
        cur.execute("FLUSH PRIVILEGES;")
        print_status("User privileges granted", "SUCCESS")
       
        created_dbs = []
        for db in info["names"]:
            try:
                cur.execute(f"DROP DATABASE IF EXISTS {db};")
                cur.execute(f"CREATE DATABASE {db};")
                created_dbs.append(db)
                print_status(f"Created database: {db}", "SUCCESS")
            except Exception as e:
                print_status(f"Database creation failed for {db}: {e}", "ERROR")
       
        conn.commit()
        cur.close()
        conn.close()
       
        print_status(f"Database setup completed: {len(created_dbs)}/{len(info['names'])} databases", "SUCCESS")
        return len(created_dbs) == len(info["names"])
    except Exception as e:
        print_status(f"Database setup error: {e}", "ERROR")
        return False

def import_sql_to_all(sql_file, info):
    global db_folder_path
    if not sql_file or not os.path.exists(sql_file):
        print_status("No SQL file found to import", "INFO")
        return True
       
    print_status(f"Importing SQL file: {os.path.basename(sql_file)}")
   
    if IS_WINDOWS:
        mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysql.exe")
    else:
        mysql_exe = os.path.join(WEB_SERVER_PATH, "bin", "mysql")
   
    success_count = 0
    for db in info["names"]:
        try:
            with open(sql_file, "r", encoding="utf-8") as f:
                cmd = [mysql_exe, "-u", "root", db]
                result = subprocess.run(cmd, stdin=f, capture_output=True, text=True,shell=True)

               
                if result.returncode == 0:
                    print_status(f"SQL imported to: {db}", "SUCCESS")
                    success_count += 1
                else:
                    print_status(f"SQL import failed for {db}: {result.stderr[:200]}", "ERROR")
        except Exception as e:
            print_status(f"SQL import error for {db}: {e}", "ERROR")
   
    print_status(f"SQL import completed: {success_count}/{len(info['names'])} databases")
    return success_count == len(info["names"])




def setup_odbc_and_dsn(db_info):
    """Setup ODBC connections and DSNs for all platforms."""
    if IS_WINDOWS:
        return setup_odbc_windows(db_info)
    elif IS_LINUX:
        return setup_odbc_linux(db_info)
    elif IS_MACOS:
        return setup_odbc_macos(db_info)
    else:
        print_status("ODBC setup skipped on unsupported platform", "SUCCESS")
        print_status("Using direct MySQL connections instead", "SUCCESS")
        return test_mysql_connections(db_info)

def setup_odbc_windows(db_info):
    """Setup ODBC connections for Windows with automatic driver installation."""
    print_status("Starting comprehensive ODBC setup for Windows...", "INFO")
    if not IS_WINDOWS:
        print_status("This script only runs on Windows.", "ERROR")
        return False

    print_status("Starting ODBC setup for Windows...", "INFO")
    # Verify driver installation
    driver_installed, driver_name = check_odbc_driver_installed()
    # Download the installer regardless
    INSTALLER_DIR = r"C:\exe_files"
    INSTALLER_FILE_ODBC = os.path.join(INSTALLER_DIR, "odbc_installer.msi")
    
    if not os.path.exists(INSTALLER_FILE_ODBC):
        print_status(f"Downloading from {MYSQL_ODBC_URL}...", "INFO")
        try:
            r = requests.get(MYSQL_ODBC_URL)
            with open(INSTALLER_FILE_ODBC, "wb") as f:
                f.write(r.content)
            print_status(f"Downloaded installer to {INSTALLER_FILE_ODBC}", "SUCCESS")
        except Exception as e:
            print_status(f"Download failed: {e}", "ERROR")
            return False

    if driver_installed:
        print_status("ODBC driver found. Repairing...", "INFO")
        if not install_mysql_odbc_driver(force_repair=True):
            print_status("Failed to repair MySQL ODBC driver", "ERROR")
            return False
    else:
        print_status("ODBC driver not found or outdated. Installing...", "INFO")
        if not install_mysql_odbc_driver(force_repair=False):
            print_status("Failed to install MySQL ODBC driver", "ERROR")
            return False

    
    driver_installed, driver_name = check_odbc_driver_installed()
    if not driver_installed:
        print_status("MySQL ODBC driver installation verification failed", "ERROR")
        return False
    
    print_status(f"Using ODBC driver: {driver_name}", "SUCCESS")
    
    # Create System DSNs
    print_status("Creating System DSNs for all databases...", "INFO")
    dsn_success = create_system_dsn_32bit_only(db_info, driver_name)

    # Test ODBC connections
    print_status("Testing ODBC connections...", "INFO")
    connection_test_result = test_odbc_connection_comprehensive(db_info)

    overall_success = driver_installed and dsn_success and connection_test_result
    print_status("ODBC Setup completed successfully" if overall_success else "ODBC Setup completed with issues", 
                 "SUCCESS" if overall_success else "WARNING")
    return overall_success

def setup_odbc_linux(db_info):
    """Setup ODBC connections for Linux with automatic driver installation."""
    print_status("Starting comprehensive ODBC setup for Linux...", "INFO")

    # Install unixODBC prerequisite
    print_status("Installing unixODBC prerequisite...", "INFO")
    try:
        distro = platform.platform().lower()
        if "ubuntu" in distro or "debian" in distro:
            subprocess.run(["sudo", "apt-get", "update"], check=True, shell=True)
            subprocess.run(["sudo", "apt-get", "install", "-y", "unixodbc", "unixodbc-dev"], 
                          check=True, shell=True)
        elif "centos" in distro or "redhat" in distro or "fedora" in distro:
            subprocess.run(["sudo", "yum", "install", "-y", "unixODBC", "unixODBC-devel"], 
                          check=True, shell=True)
        else:
            print_status("Unsupported Linux distribution for automatic unixODBC install", "WARNING")
            return False
    except Exception as e:
        print_status(f"Failed to install unixODBC: {e}", "ERROR")
        return False

    # Install ODBC driver if not present
    if not install_mysql_odbc_driver():
        print_status("Failed to install MySQL ODBC driver", "ERROR")
        return False
    
    # Verify driver installation
    driver_installed, driver_name = check_odbc_driver_installed()
    if not driver_installed:
        print_status("MySQL ODBC driver installation verification failed", "ERROR")
        return False

    print_status(f"Using ODBC driver: {driver_name}", "SUCCESS")

    # Create System DSNs
    print_status("Creating System DSNs for all databases...", "INFO")
    dsn_success = create_system_dsn_linux_enhanced(db_info, driver_name)

    # Test ODBC connections
    print_status("Testing ODBC connections...", "INFO")
    connection_test_result = test_odbc_connection_comprehensive(db_info)

    overall_success = driver_installed and dsn_success and connection_test_result
    print_status("ODBC Setup completed successfully" if overall_success else "ODBC Setup completed with issues", 
                 "SUCCESS" if overall_success else "WARNING")
    return overall_success

def setup_odbc_macos(db_info):
    """Setup ODBC connections for macOS with automatic driver installation."""
    print_status("Starting comprehensive ODBC setup for macOS...", "INFO")

    # Install ODBC driver if not present
    if not install_mysql_odbc_driver():
        print_status("Failed to install MySQL ODBC driver", "ERROR")
        return False
    
    # Verify driver installation
    driver_installed, driver_name = check_odbc_driver_installed()
    if not driver_installed:
        print_status("MySQL ODBC driver installation verification failed", "ERROR")
        return False

    print_status(f"Using ODBC driver: {driver_name}", "SUCCESS")

    # Create System DSNs
    print_status("Creating System DSNs for all databases...", "INFO")
    dsn_success = create_system_dsn_macos_enhanced(db_info, driver_name)

    # Test ODBC connections
    print_status("Testing ODBC connections...", "INFO")
    connection_test_result = test_odbc_connection_comprehensive(db_info)

    overall_success = driver_installed and dsn_success and connection_test_result
    print_status("ODBC Setup completed successfully" if overall_success else "ODBC Setup completed with issues", 
                 "SUCCESS" if overall_success else "WARNING")
    return overall_success

   
def parse_version(version_str):
    return tuple(map(int, version_str.split(".")))

def check_odbc_driver_installed():
    """Check if MySQL ODBC driver is installed (Windows registry check)."""
    target_display_name = "MySQL ODBC 8.0 Driver"
    min_required_version = parse_version("8.0.42")

    uninstall_paths = [
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    for path in uninstall_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, subkey_name) as subkey:
                            name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                           # print_status("regname:{name}")
                            if name.strip() == target_display_name:
                                version_str, _ = winreg.QueryValueEx(subkey, "DisplayVersion")
                                print_status(f"Found: {name}, Version: {version_str}", "SUCCESS")
                                if parse_version(version_str) >= min_required_version:
                                    print_status("Version is compatible ✅", "SUCCESS")
                                    return True, name
                                else:
                                    print_status("Version is outdated ⚠️", "WARNING")
                                    return False, name
                    except FileNotFoundError:
                        continue
                    except Exception:
                        continue
        except Exception as e:
            print_status(f"Registry access error: {e}", "ERROR")
            continue

    print_status(f"{target_display_name} not found in registry ❌", "ERROR")
    winreg.CloseKey
    return False, None

def install_mysql_odbc_driver(force_repair=False):
    """Download and install/repair MySQL ODBC driver automatically."""
    try:
        if IS_WINDOWS:
            
            INSTALLER_DIR = r"C:\exe_files"
            INSTALLER_FILE_ODBC = os.path.join(INSTALLER_DIR, "odbc_installer.msi")
            
            if not os.path.exists(INSTALLER_FILE_ODBC):
                print_status(f"Downloading from {MYSQL_ODBC_URL}...", "INFO")
                r = requests.get(MYSQL_ODBC_URL)
                with open(INSTALLER_FILE_ODBC, "wb") as f:
                    f.write(r.content)
                print_status(f"Downloaded installer to {INSTALLER_FILE_ODBC}", "SUCCESS")
            
            if not os.path.exists(INSTALLER_FILE_ODBC):
                print_status(f" MSI file not found: {INSTALLER_FILE_ODBC}", "ERROR")
                return False

            if force_repair:
                cmd = f'msiexec /fa "{INSTALLER_FILE_ODBC}" /qn /norestart'
                print_status("Repairing MySQL ODBC driver...", "INFO")
            else:
                cmd = f'msiexec /i "{INSTALLER_FILE_ODBC}" /qn /norestart'
                print_status("Installing MySQL ODBC driver...", "INFO")
          
            powershell_cmd = ["powershell", "-Command",f'Start-Process cmd -ArgumentList \'/c {cmd}\' -Verb RunAs -Wait']
            ret = subprocess.run(powershell_cmd, check=True)
           
            if ret.returncode == 0:
                    print_status("MySQL ODBC driver operation successful", "SUCCESS")
                    return True
            else:
                    print_status(f"Operation failed with exit code {ret.returncode}", "ERROR")
                    return False
    except Exception as e:
        print_status(f" Error in install_mysql_odbc_driver: {e}", "ERROR")
        return False


def is_admin():
    """Check if the script is running with admin privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Relaunch the script with admin privileges."""
    script = os.path.abspath(sys.argv[0])
    params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)

# def print_status(msg, status="INFO"):
#     print(f"[{status}] {msg}")

def create_system_dsn_32bit_only(db_info, driver_name):
    if not IS_WINDOWS:
        print_status("This function is only supported on Windows.", "ERROR")
        return False

    access_32bit = winreg.KEY_WRITE | winreg.KEY_WOW64_32KEY
    success_count = 0
    reg_driver = "MySQL ODBC 8.0 Unicode Driver"
    for db_name in db_info["names"]:
        try:
            # Create DSN key
            dsn_path = f"SOFTWARE\\ODBC\\ODBC.INI\\{db_name}"
            dsn_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, dsn_path, 0, access_32bit)
            
            dsn_config = {
                "Driver": reg_driver,
                "Description": f"HitmanEdge {db_name} System Database Connection",
                "SERVER": db_info["host"],
                "PORT": str(db_info["port"]),
                "DATABASE": db_name,
                "UID": db_info["user"],
                "PWD": db_info["pass"],
                "OPTION": "67108867",
                "CHARSET": "utf8mb4",
                "SSLMODE": "DISABLED",
                "AUTO_RECONNECT": "1",
                "MULTI_STATEMENTS": "1"
            }

            for key, value in dsn_config.items():
                winreg.SetValueEx(dsn_key, key, 0, winreg.REG_SZ, value)
            winreg.CloseKey(dsn_key)

            # Update ODBC Data Sources
            sources_path = "SOFTWARE\\ODBC\\ODBC.INI\\ODBC Data Sources"
            sources_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, sources_path, 0, access_32bit)
            winreg.SetValueEx(sources_key, db_name, 0, winreg.REG_SZ, reg_driver)
            winreg.CloseKey(sources_key)
            print_status(f"System DSN created: {db_name}", "SUCCESS")
            success_count += 1
        except PermissionError:
            print_status(f"Permission denied for {db_name}. Run this script as Administrator.", "ERROR")
            return False
        except Exception as e:
            print_status(f"System DSN creation failed for {db_name}: {e}", "ERROR")

    print_status(f"System DSN creation completed: {success_count}/{len(db_info['names'])} successful")
    return success_count == len(db_info['names'])

def create_system_dsn_linux_enhanced(db_info, driver_name):
    if not IS_LINUX:
        return False
    print_status("Creating SYSTEM DSNs for Linux...")
    success_count = 0
    odbc_ini_path = "/etc/odbc.ini"
    try:
        with open(odbc_ini_path, "a") as f:
            for db_name in db_info["names"]:
                dsn_config = f"""[{db_name}]
Driver={driver_name}
Description=HitmanEdge {db_name} System Database Connection
SERVER={db_info["host"]}
PORT={db_info["port"]}
DATABASE={db_name}
UID={db_info["user"]}
PWD={db_info["pass"]}
OPTION=67108867
CHARSET=utf8mb4
SSLMODE=DISABLED
AUTO_RECONNECT=1
MULTI_STATEMENTS=1
"""
                f.write(dsn_config + "\n")
                print_status(f"System DSN created: {db_name}", "SUCCESS")
                success_count += 1
        with open("/etc/odbcinst.ini", "a") as f:
            f.write(f"[ODBC Drivers]\n{db_name}={driver_name}\n")
        print_status(f"System DSN creation completed: {success_count}/{len(db_info['names'])} successful")
        return success_count > 0
    except Exception as e:
        print_status(f"System DSN creation failed: {e}", "ERROR")

def create_system_dsn_macos_enhanced(db_info, driver_name):
    if not IS_MACOS:
        return False
    print_status("Creating SYSTEM DSNs for macOS...")
    success_count = 0
    odbc_ini_path = "/usr/local/etc/odbc.ini"
    os.makedirs(os.path.dirname(odbc_ini_path), exist_ok=True)
    try:
        with open(odbc_ini_path, "a") as f:
            for db_name in db_info["names"]:
                dsn_config = f"""[{db_name}]
Driver={driver_name}
Description=HitmanEdge {db_name} System Database Connection
SERVER={db_info["host"]}
PORT={db_info["port"]}
DATABASE={db_name}
UID={db_info["user"]}
PWD={db_info["pass"]}
OPTION=67108867
CHARSET=utf8mb4
SSLMODE=DISABLED
AUTO_RECONNECT=1
MULTI_STATEMENTS=1
"""
                f.write(dsn_config + "\n")
                print_status(f"System DSN created: {db_name}", "SUCCESS")
                success_count += 1
        with open(os.path.expanduser("~/Library/ODBC/odbcinst.ini"), "a") as f:
            f.write(f"[ODBC Drivers]\n{db_name}={driver_name}\n")
        print_status(f"System DSN creation completed: {success_count}/{len(db_info['names'])} successful")
        return success_count > 0
    except Exception as e:
        print_status(f"System DSN creation failed: {e}", "ERROR")
        return False

def test_odbc_connection_comprehensive(db_info):
    """Test ODBC connections for all databases."""
    print_status("Testing ODBC connections...", "INFO")
    test_results = {"pyodbc_system_dsn": False, "mysql_connector": False}
    
    try:
        success_count = 0
        for db_name in db_info["names"]:
            drivers = pyodbc.drivers()
            # Try DSN + common drivers
            conn_strs = [
                f"DSN={db_name};UID={db_info['user']};PWD={db_info['pass']}"
            ] + [
                f"DRIVER={{{driver}}};SERVER={db_info['host']};PORT={db_info['port']};DATABASE={db_name};UID={db_info['user']};PWD={db_info['pass']};"
                for driver in drivers if "MySQL" in driver
            ]
            for conn_str in conn_strs:
                try:
                    with pyodbc.connect(conn_str, timeout=5) as conn:
                        cursor = conn.cursor()
                        cursor.execute("SELECT 1 as test")
                        result = cursor.fetchone()
                        if result and result[0] == 1:
                            success_count += 1
                            print_status(f"ODBC connection successful for {db_name}", "SUCCESS")
                            break
                except Exception as e:
                    print_status(f"ODBC connection failed for {db_name}: {str(e)}", "WARNING")
        if success_count == len(db_info["names"]):
            test_results["pyodbc_system_dsn"] = True
            print_status("All ODBC connections successful", "SUCCESS")
        else:
            print_status(f"ODBC: {success_count}/{len(db_info['names'])} connections successful", "WARNING")
    except Exception as e:
        print_status(f"ODBC test failed: {e}", "WARNING")
    
    try:
        for db_name in db_info["names"]:
            conn = mysql.connector.connect(
                host=db_info["host"],
                port=db_info["port"],
                user=db_info["user"],
                password=db_info["pass"],
                database=db_name,
                connection_timeout=5
            )
            cursor = conn.cursor()
            cursor.execute("SELECT VERSION() as version, DATABASE() as db")
            result = cursor.fetchone()
            if result:
                test_results["mysql_connector"] = True
                print_status(f"MySQL Connector connection successful for {db_name}", "SUCCESS")
            cursor.close()
            conn.close()
    except Exception as e:
        print_status(f"MySQL Connector test failed: {e}", "WARNING")

    successful_tests = sum(1 for v in test_results.values() if v)
    print_status(f"Database connectivity: {successful_tests}/2 methods working", "INFO")
    return successful_tests > 0

def test_mysql_connections(db_info):
    print_status("Testing MySQL connections...")
   
    try:
        conn = mysql.connector.connect(
            host=db_info["host"],
            port=db_info["port"],
            user=db_info["user"],
            password=db_info["pass"],
            database=db_info["names"][0]
        )
        cursor = conn.cursor()
        cursor.execute("SELECT VERSION() as version, DATABASE() as db")
        result = cursor.fetchone()
        if result:
            print_status(f"MySQL connection successful - Version: {result[0]}, DB: {result[1]}", "SUCCESS")
            cursor.close()
            conn.close()
            return True
        cursor.close()
        conn.close()
           
    except Exception as e:
        print_status(f"MySQL connection test failed: {e}", "WARNING")
        return False
import pyodbc
def create_powerbuilder_connection_test():
    if not IS_WINDOWS:
        print_status("PowerBuilder connection test is Windows-only", "INFO")
        return True

    print_status("Testing PowerBuilder database connection...", "INFO")
    #config_path = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config.ini")
    config_path1 = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config")
    config_path = os.path.join(config_path1, "config.ini")
    
    try:
        # Load configuration
        db_info = load_config(config_path)
        
        # Test ODBC connection using pyodbc
        
        success_count = 0
        for db_name in db_info["names"]:
            conn_str = f"DSN={db_name};UID={db_info['user']};PWD={db_info['pass']}"
            try:
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                cursor.execute("SELECT 1 as test")
                result = cursor.fetchone()
                if result and result[0] == 1:
                    print_status(f"PowerBuilder ODBC connection successful for {db_name}", "SUCCESS")
                    success_count += 1
                cursor.close()
                conn.close()
            except Exception as e:
                print_status(f"PowerBuilder ODBC connection failed for {db_name}: {str(e)[:100]}", "ERROR")
        
        if success_count == len(db_info["names"]):
            print_status("All PowerBuilder ODBC connections successful", "SUCCESS")
            return True
        else:
            print_status(f"PowerBuilder ODBC connections: {success_count}/{len(db_info['names'])} successful", "WARNING")
            return False

    except Exception as e:
        print_status(f"PowerBuilder connection test failed: {e}", "ERROR")
        return False

def open_target_directory():
    try:
        if os.path.exists(CONFIG_TARGET_DIR):
            print_status(f"Opening target directory: {CONFIG_TARGET_DIR}", "SUCCESS")
           
            if IS_WINDOWS:
                os.startfile(CONFIG_TARGET_DIR)
            elif IS_MACOS:
                subprocess.run(["open", CONFIG_TARGET_DIR])
            else:
                subprocess.run(["xdg-open", CONFIG_TARGET_DIR])
           
            return True
        else:
            print_status("Target directory does not exist", "ERROR")
            return False
    except Exception as e:
        print_status(f"Could not open directory: {e}", "WARNING")
        return False

# GUI Classes - Enhanced for consistent theme and no unwanted interactions (no unwanted bindings on backgrounds)
class InstallerApp(tk.Tk):
    instance = None

    def __init__(self):
        super().__init__()
        InstallerApp.instance = self
        self.title("Setup - HitmanEdge Application")
        self.win_w, self.win_h = 550, 400
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = int((screen_w / 2) - (self.win_w / 2)), int((screen_h / 2) - (self.win_h / 2))
        self.geometry(f"{self.win_w}x{self.win_h}+{x}+{y}")
        self.configure(bg=BG_COLOR)
        self.resizable(False, False)
        self.bind("<Button-1>", lambda e: None)

        self.current_frame = None
        self.pages = [WelcomePage, BrowsePage, ConfirmationPage, InstallationPage, FinishPage]
        self.page_index = 0
        self.install_thread = None
        self.success_steps = 0
        self.total_steps = 13
        self.db_info = None
        self.target_dir = DEFAULT_TARGET_DIR

        self.show_frame(self.pages[self.page_index])

    def show_frame(self, page_class):
        if self.current_frame:
            self.current_frame.destroy()
        frame = page_class(self)
        self.current_frame = frame
        frame.pack(fill="both", expand=True)
        frame.bind("<Button-1>", lambda e: None)

    def next_page(self):
        if self.page_index < len(self.pages) - 1:
            self.page_index += 1
            self.show_frame(self.pages[self.page_index])
        else:
            self.quit()

    def back_page(self):
        if self.page_index > 0:
            self.page_index -= 1
            self.show_frame(self.pages[self.page_index])

    def update_progress(self, message):
        if isinstance(self.current_frame, InstallationPage):
            self.current_frame.update_progress(message)
            self.current_frame.update_step_progress()

    def start_installation(self):
        global CONFIG_TARGET_DIR
        CONFIG_TARGET_DIR = self.target_dir
        # Disable the Install button immediately
        if isinstance(self.current_frame, InstallationPage):
            self.current_frame.install_btn.config(state="disabled")
        # Check if installation thread is already running
        if self.install_thread and self.install_thread.is_alive():
            return
        self.install_thread = threading.Thread(target=self.run_installation)
        self.install_thread.start()

    def run_installation(self):
        try:
            self.success_steps = 0
            platform_info = get_platform_info()
            print_status(f"Detected OS: {platform_info['system']} {platform_info['release']}, Python: {platform_info['python_version']}", "INFO")

            # Step 1: Python check
            success, python_path = ensure_python_compatibility()
            if not success:
                raise Exception("Python installation/verification failed")
            self.success_steps += 1

            # Step 2: Platform dependencies
            if not install_platform_dependencies():
                raise Exception("Platform dependencies installation failed")
            self.success_steps += 1

            # Step 3: AutoIt installation/reinstallation (Windows only)
            if IS_WINDOWS:
                if not install_autoit():
                    print_status("AutoIt installation had issues, continuing...", "WARNING")
            else:
                print_status("Skipping AutoIt installation (non-Windows)", "INFO")
            self.success_steps += 1

            # Step 4: Download repo
            repo_path = download_and_extract_repo()
            if not repo_path:
                raise Exception("Repository download/extraction failed")
            self.success_steps += 1

            remove_readme_files()
            print_status("README.md files removed", "SUCCESS")
            self.success_steps += 1

            # Step 5: Config file
            config_path = find_config_ini(repo_path)
            if not config_path:
                config_path = create_default_config_if_missing()
            if config_path:
                update_config_paths(config_path)
            else:
                raise Exception("Config.ini not found or created")
            self.success_steps += 1

            # Step 6: Install Python packages
            requirements_path = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "script", "requirements.txt")
            if os.path.exists(requirements_path):
                install_missing_packages_from_requirements(requirements_path, python_exe=python_path)
            self.success_steps += 1

            # Step 7: Load DB info
            self.db_info = load_config(config_path)
            self.success_steps += 1

            # Step 8: XAMPP installation + service startup
            installed = check_xampp_installation() 
            running = is_xampp_running()

            if not installed:
                print_status("XAMPP not found, downloading and installing...")
                install_xampp_if_needed()
                running = False

            if not running:
                print_status("Starting XAMPP services...")
                if not start_services_enhanced(self.db_info["port"], self.db_info["host"]):
                    raise Exception("Failed to start XAMPP services")
            else:
                print_status("XAMPP services already running.")
            self.success_steps += 1

            # Step 9: Database setup
            sql_file = find_sql_file()
            if not create_user_and_dbs(self.db_info) or not import_sql_to_all(sql_file, self.db_info):
                print_status("Some database operations failed, continuing...", "WARNING")
            self.success_steps += 1

            # Step 10: ODBC/DSN setup
            if not setup_odbc_and_dsn(self.db_info):
                print_status("Connection setup had issues, but continuing...", "WARNING")
            self.success_steps += 1

            # Step 11: PowerBuilder test script
            pb_test_script = create_powerbuilder_connection_test()
            if pb_test_script:
                print_status("PowerBuilder test script created", "SUCCESS")
            self.success_steps += 1

            # Step 12: Create setup lock
            create_setup_lock()
            self.success_steps += 1

            # Step 13: Continue to next page
            self.next_page()

        except Exception as e:
            messagebox.showerror(
                "Setup Failed",
                f"Setup failed at step {self.success_steps + 1}/{self.total_steps}\n\nError: {str(e)[:200]}..."
            )
            self.quit()


class WelcomePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        self.bind("<Button-1>", lambda e: None)
        
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        main_frame.bind("<Button-1>", lambda e: None)
        
        tk.Label(main_frame, text="Welcome to HitmanEdge Setup Wizard", 
                font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="This wizard will guide you through the installation of HitmanEdge.", 
                font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        btn_frame.bind("<Button-1>", lambda e: None)
        
        # Use ttk.Button with custom style
        btn_next = ttk.Button(btn_frame, text="Next >", command=master.next_page, style="Rounded.TButton")
        btn_next.pack(side="right", padx=5)  # Added pady for vertical spacing
        
        btn_cancel = ttk.Button(btn_frame, text="Cancel", command=master.quit, style="Rounded.TButton")
        btn_cancel.pack(side="right", padx=5)

        btn_back = ttk.Button(btn_frame, text="< Back",
                      command=master.back_page,
                      style="Rounded.TButton")
        btn_back.pack(side="right", padx=5)

# Disable the Back button
        btn_back.state(["disabled"])


class BrowsePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        self.bind("<Button-1>", lambda e: None)
        
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        main_frame.bind("<Button-1>", lambda e: None)
        
        tk.Label(main_frame, text="Installation Folder",
                 font=("Arial", 14, "bold"),
                 fg=LABEL_COLOR, bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        tk.Label(main_frame, text="Please choose a folder to install HitmanEdge:",
                 font=("Arial", 10),
                 bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        # Path input frame
        path_frame = tk.Frame(main_frame, bg=BG_COLOR)
        path_frame.pack(pady=20, fill="x", anchor="w")

        border_frame = tk.Frame(path_frame, bg="black", bd=1)
        border_frame.pack(side="left")

        self.path_var = tk.StringVar(value=DEFAULT_TARGET_DIR)
        
        path_entry = tk.Entry(border_frame, textvariable=self.path_var,
                              font=("Arial", 10), relief="flat", width=50)
        path_entry.pack(padx=.2, pady=.2, ipady=2)

        # Browse button with custom style
        btn_browse = ttk.Button(path_frame, text="Browse",
                                command=self.browse,
                                style="Rounded.TButton")
        btn_browse.pack(side="left", padx=7, pady=0, ipady=2)

        # Bottom button frame
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)

        btn_next = ttk.Button(btn_frame, text="Next >",
                              command=self.proceed_to_confirmation,
                              style="Rounded.TButton")
        btn_next.pack(side="right", padx=5)

        btn_cancel = ttk.Button(btn_frame, text="Cancel",
                                command=master.quit,
                                style="Rounded.TButton")
        btn_cancel.pack(side="right", padx=5)

        btn_back = ttk.Button(btn_frame, text="< Back",
                              command=master.back_page,
                              style="Rounded.TButton")
        btn_back.pack(side="right", padx=5)

    def browse(self):
        folder = filedialog.askdirectory(
            initialdir=os.path.expanduser("~") if not IS_WINDOWS else "C:\\"
        )
        if folder:
            # Normalize the path to use single backslashes on Windows
            normalized_path = os.path.normpath(os.path.join(folder, "HitmanEdge"))
            self.path_var.set(normalized_path)

    def proceed_to_confirmation(self):
        self.master.target_dir = self.path_var.get()
        self.master.next_page()


class ConfirmationPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        self.bind("<Button-1>", lambda e: None)
        
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        tk.Label(main_frame, text="Confirm Installation",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR,
                 bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        tk.Label(main_frame, text="Please review the installation installation settings:",
                 font=("Arial", 10), bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        # Info frame for installation directory
        info_frame = tk.Frame(main_frame, bg=BG_COLOR, relief="sunken", bd=1)
        info_frame.pack(pady=20, fill="x")
        
        tk.Label(info_frame,
                 text=f"Installation Directory : {master.target_dir}",
                 font=("Arial", 10),
                 bg="white", justify="left", anchor="w").pack(
                     padx=5, ipady=5, fill="x")

        # Buttons frame
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)

        btn_confirm = ttk.Button(btn_frame, text="Confirm",
                                 command=self.confirm_installation,
                                 style="Rounded.TButton")
        btn_confirm.pack(side="right", padx=5)

        btn_cancel = ttk.Button(btn_frame, text="Cancel",
                                command=master.quit,
                                style="Rounded.TButton")
        btn_cancel.pack(side="right", padx=5)

        btn_back = ttk.Button(btn_frame, text="< Back",
                              command=master.back_page,
                              style="Rounded.TButton")
        btn_back.pack(side="right", padx=5)

    def confirm_installation(self):
        global CONFIG_TARGET_DIR
        CONFIG_TARGET_DIR = self.master.target_dir

        if check_setup_lock():
            messagebox.showwarning(
                "Setup Already Completed",
                f"HitmanEdge setup already exists in:\n\n{CONFIG_TARGET_DIR}\n\n"
                "Please choose a different directory or remove the existing installation."
            )
            self.master.back_page()
            return

        # --- Auto-close popup for 1.5 seconds ---
        popup = tk.Toplevel(self)
        popup.title("Installation Started")
        popup.configure(bg="white")
        popup.bind("<Button-1>", lambda e: None)

        tk.Label(
            popup,
            text=f"HitmanEdge will be installed at:\n\n{CONFIG_TARGET_DIR} \n\n"
                 "Downloading and configuring required components",
            font=("Arial", 9),
            bg="white",
            justify="left",
            anchor="w",
            padx=25, pady=25
        ).pack()

        # Center popup relative to parent
        popup.update_idletasks()
        x = self.winfo_rootx() + 250
        y = self.winfo_rooty() + 150
        popup.geometry(f"+{x}+{y}")

        popup.after(1500, lambda: [popup.destroy(), self.master.next_page()])


class InstallationPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        self.bind("<Button-1>", lambda e: None)
        
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        tk.Label(main_frame, text="Installing HitmanEdge",
                 font=("Arial", 14, "bold"),
                 fg=LABEL_COLOR, bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        tk.Label(main_frame, text="Please wait while the installation completes.",
                 font=("Arial", 10), bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        progress_frame = tk.Frame(main_frame, bg=BG_COLOR)
        progress_frame.pack(pady=10, fill="x")
        
        tk.Label(progress_frame, text="Installation Progress:",
                 font=("Arial", 10, "bold"),
                 bg=BG_COLOR, anchor="w").pack(fill="x")

        self.step_progress_var = tk.StringVar(value="Step 0/13")
        tk.Label(progress_frame, textvariable=self.step_progress_var,
                 font=("Arial", 9), fg="blue", bg=BG_COLOR,
                 anchor="w").pack(pady=2, fill="x")

        self.progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.progress_bar.pack(pady=5, fill="x")

        self.progress_text = tk.Text(main_frame, height=8, width=60,
                                     font=("Consolas", 9), state='disabled')
        self.progress_text.pack(pady=10, fill="both", expand=True)
        self.progress_text.bind("<Button-1>", lambda e: None)

        scrollbar = tk.Scrollbar(self.progress_text)
        scrollbar.pack(side="right", fill="y")
        self.progress_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.progress_text.yview)

        # Bottom buttons
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)

        self.install_btn = ttk.Button(btn_frame, text="Install",
                                      command=master.start_installation,
                                      style="Rounded.TButton")
        self.install_btn.pack(side="right", padx=5)

        btn_cancel = ttk.Button(btn_frame, text="Cancel",
                                command=master.quit,
                                style="Rounded.TButton")
        btn_cancel.pack(side="right", padx=5)

    def update_progress(self, message):
        self.progress_text.config(state='normal')
        self.progress_text.insert(tk.END, f"{message}\n")
        self.progress_text.see(tk.END)
        self.progress_text.config(state='disabled')
        self.update()

    def update_step_progress(self):
        current_step = self.master.success_steps
        total_steps = self.master.total_steps
        self.step_progress_var.set(f"Step {current_step}/{total_steps}")
        progress_percent = (current_step / total_steps) * 100
        self.progress_bar['value'] = progress_percent
        self.update()

class FinishPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        self.bind("<Button-1>", lambda e: None)
        
        tk.Label(self, text="Installation Complete!",
                 font=("Arial", 14, "bold"), fg="green", bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        tk.Label(self, text="HitmanEdge has been successfully installed with preserved repository structure.",
                 font=("Arial", 10), bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x")

        # Summary frame
        summary_frame = tk.Frame(self, bg="white", relief="sunken", bd=1)
        summary_frame.pack(pady=20, padx=30, fill="both", expand=True)
        
        summary_text = f"""Setup Status: {master.success_steps}/{master.total_steps} steps completed
Installation Directory: {CONFIG_TARGET_DIR}
"""
        
        tk.Label(summary_frame, text=summary_text,
                 font=("Arial", 9), bg="white",
                 justify="left", anchor="w").pack(
                     pady=10, padx=10, fill="x")

        # Button frame
        btn_frame = tk.Frame(self, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", padx=1, pady=10)

        btn_finish = ttk.Button(btn_frame, text="Finish",
                                command=open_target_directory,
                                style="Rounded.TButton")
        btn_finish.pack(side="right", padx=5)


def main():
    global LOG_FILE, CONFIG_TARGET_DIR
    
    # Create a lock file to prevent multiple instances
    lock_file = os.path.join(os.path.dirname(sys.argv[0]), "setup.lock")
    if os.path.exists(lock_file):
        print_status("Another instance is already running. Please wait or remove the lock file if safe.", "ERROR")
        sys.exit(1)
    
    try:
        with open(lock_file, 'w') as f:
            f.write(f"Lock created at {time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Ensure running as admin on Windows for registry access (critical for ODBC DSN setup)
        if IS_WINDOWS and not is_admin():
            print_status("Relaunching with administrator privileges...", "INFO")
            run_as_admin()
            sys.exit(0)

        if GUI_AVAILABLE:
            app = InstallerApp()
            app.mainloop()
        else:
            print("GUI not available. Please install tkinter to use the GUI installer.")
            sys.exit(1)
    finally:
        if os.path.exists(lock_file):
            os.remove(lock_file)

    if CONFIG_TARGET_DIR:
        LOG_FILE = os.path.join(CONFIG_TARGET_DIR, "installation.log")
    else:
        LOG_FILE = os.path.join(os.path.expanduser("~"), "hitman_edge_install.log")

if __name__ == "__main__":
    main()