import os
import sys
import subprocess
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
import requests
import psutil
import json
from pathlib import Path
from collections import defaultdict
from datetime import datetime

BUILD_ENV = "test"

EXE_PATH = r"C:\exe_files"
LOG_FILE = os.path.join(EXE_PATH, "installation_log.txt")

os.makedirs(EXE_PATH, exist_ok=True)

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

try:
    import pyodbc
except ImportError:
    if not getattr(sys, 'frozen', False):
        print("Installing pyodbc...")
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

if not IS_WINDOWS:
    print("This script is designed for Windows only.")
    sys.exit(1)

winreg = None
db_username, db_password = None, None
try:
    import winreg as _winreg
    winreg = _winreg
except ImportError:
    winreg = None

db_folder_path = None
REPO_URL = "https://github.com/Dharanidharan5624/Hitman_Testing.git"
CONFIG_TARGET_DIR = None
SELECTED_ENVIRONMENT = None
SELECTED_VERSION = None

BG_COLOR = "white"
BTN_COLOR = "#0078D7"
BTN_TEXT = "white"
LABEL_COLOR = "#0078D7"

REQUIRED_FOLDERS = ['database', 'powerbuilder']
PYTHON_INSTALLER_FILE = "python-3.13.5-amd64.exe"
PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.13.5/python-3.13.5-amd64.exe"
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
INSTALLER_FILE_ODBC = os.path.join(EXE_PATH, "odbc_installer_x86.msi")
AUTOIT_INSTALL_PATH = r"C:\Program Files (x86)\AutoIt3"
AUTOIT_INSTALLER_URL = "https://www.autoitscript.com/files/autoit3/autoit-v3-setup.exe"
AUTOIT_INSTALLER_FILE = os.path.join(EXE_PATH, "autoit-v3-setup.exe")
POWERBUILDER_INSTALLER_FILE = os.path.join(EXE_PATH, "PowerBuilderInstaller_bootstrapper.exe")

VCREDIST_PACKAGES = {
    "2015-2022_x86": {
        "url": "https://aka.ms/vs/17/release/vc_redist.x86.exe",
        "name": "Microsoft Visual C++ 2015-2022 Redistributable (x86)",
        "file": "vc_redist_2015_2022_x86.exe",
        "registry_keys": [
            r"SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X86",
            r"SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86"
        ],
        "min_version": "14.42.34438"
    }
}

def save_version_file(version_data=None):
    if not CONFIG_TARGET_DIR:
        return False
    version_file_path = os.path.join(CONFIG_TARGET_DIR,"Hitman.json")
    try:
        os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
        if version_data is None:
            version_data = {"build_env": BUILD_ENV}
        version_data["installation_date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        version_data["installation_path"] = CONFIG_TARGET_DIR
        version_data["selected_environment"] = SELECTED_ENVIRONMENT
        version_data["selected_version"] = SELECTED_VERSION
        with open(version_file_path, 'w', encoding='utf-8') as f:
            json.dump(version_data, f, indent=4)
        print_status(f"Version file created ({BUILD_ENV})", "SUCCESS")
        return True
    except Exception as e:
        print_status(f"Failed to save version file: {e}", "ERROR")
        return False

def load_version_file():
    if not CONFIG_TARGET_DIR:
        return None
    version_file_path = os.path.join(CONFIG_TARGET_DIR,"Hitman.json")
    try:
        if os.path.exists(version_file_path):
            with open(version_file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None
    except Exception as e:
        print_status(f"Failed to load version file: {e}", "WARNING")
        return None

def compare_versions(v1, v2):
    v1_parts = [int(x) for x in v1.split('.')]
    v2_parts = [int(x) for x in v2.split('.')]
    for i in range(max(len(v1_parts), len(v2_parts))):
        val1 = v1_parts[i] if i < len(v1_parts) else 0
        val2 = v2_parts[i] if i < len(v2_parts) else 0
        if val1 > val2:
            return 1
        elif val1 < val2:
            return -1
    return 0

def check_existing_installation():
    if not CONFIG_TARGET_DIR:
        return False, None
    version_data = load_version_file()
    if version_data:
        existing_env = version_data.get("build_env", "unknown")
        print_status(f"Found existing installation ({existing_env})", "INFO")
        return True, version_data
    return False, None

def check_upgrade_needed(existing_version_data):
    if not existing_version_data:
        return False, "fresh"
    
    existing_env = existing_version_data.get("build_env", "dev")

    if BUILD_ENV != existing_env:
        return True, "env_change"
    return False, "same"

def backup_existing_installation():
    if not CONFIG_TARGET_DIR or not os.path.exists(CONFIG_TARGET_DIR):
        return True
    backup_dir = f"{CONFIG_TARGET_DIR}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    try:
        print_status(f"Creating backup: {backup_dir}", "INFO")
        shutil.copytree(CONFIG_TARGET_DIR, backup_dir)
        print_status("Backup created successfully", "SUCCESS")
        return True
    except Exception as e:
        print_status(f"Backup failed: {e}", "ERROR")
        return False

def detect_system_info():
    system_info = {
        "os": CURRENT_OS,
        "version": platform.release(),
        "machine": platform.machine(),
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "architecture": platform.architecture()[0],
        "processor": platform.processor(),
        "windows_version": platform.win32_ver()[0],
        "windows_edition": platform.win32_edition() or "Unknown",
        "is_64bit": platform.machine().endswith('64')
    }
    return system_info

def get_windows_version_info():
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
    print_status("Forcing installation of x86 (32-bit) packages only")
    return ["2015-2022_x86"]

def parse_version_string(version_str):
    if not version_str:
        return (0, 0, 0, 0)
    try:
        parts = version_str.replace('v', '').split('.')
        return tuple(int(part) for part in parts[:4])
    except (ValueError, AttributeError):
        return (0, 0, 0, 0)

def check_vcredist_version_registry(package_key):
    global winreg
    if not winreg:
        return False, None
    
    package_info = VCREDIST_PACKAGES[package_key]
    min_version = parse_version_string(package_info['min_version'])
    
    print_status(f"Checking {package_info['name']}...")
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
    package_info = VCREDIST_PACKAGES[package_key]
    file_path = os.path.join(download_dir, package_info['file'])

    if os.path.exists(file_path):
        print_status(f"Using existing installer: {package_info['file']}", "INFO")
        return file_path

    print_status(f"Downloading {package_info['name']} ( progress)...")
    try:
        response = requests.get(package_info['url'], stream=True, timeout=30)
        response.raise_for_status()

        total = int(response.headers.get('content-length', 0))
        downloaded = 0
        start = time.time()

        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)

                    now = time.time()
                    if now - start >= 0.1 or downloaded == total:
                        start = now
                        if total:
                            pct = (downloaded / total) * 100
                            print_status(f"Progress: {pct:6.2f}%  ({downloaded/1_048_576:.2f} MiB)", "INFO")
                        else:
                            print_status(f"Downloaded {downloaded/1_048_576:.2f} MiB", "INFO")

        f.flush()
        os.fsync(f.fileno())

        print_status(f"Downloaded: {package_info['file']}", "SUCCESS")
        return file_path
    except Exception as e:
        print_status(f"Download failed for {package_info['name']}: {e}", "ERROR")
        return None

def install_vcredist_package(installer_path, package_key):
    if not installer_path or not os.path.exists(installer_path):
        return False
    
    package_info = VCREDIST_PACKAGES[package_key]
    print_status(f"Installing {package_info['name']}...")
    try:
        cmd = [installer_path, '/quiet', '/norestart']
        print_status("Running installer (this may take a few minutes)...", "INFO")
        result = subprocess.run(cmd, capture_output=True, text=True,check=True)
        
        if result.returncode in [0, 1638, 3010]:
            print_status(f"Installation successful: {package_info['name']}", "SUCCESS")
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
    print_status(f"Verifying installation: {VCREDIST_PACKAGES[package_key]['name']}")
    time.sleep(2)
    is_installed, version = check_vcredist_version_registry(package_key)
    if is_installed:
        print_status(f"Verification successful: v{version}", "SUCCESS")
        return True
    else:
        print_status("Verification failed - package may not be properly installed", "WARNING")
        return False

def determine_required_vcredist_packages():
    win_info = get_windows_version_info()
    required_packages = ["2015-2022_x86"]
    return required_packages

def check_vcredist_installed(package_key):
    global winreg
    if not winreg:
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
    print_status("Visual C++ (x86) Setup")
    required = get_windows_architecture()
    cache_dir = os.path.join(EXE_PATH, "vcredist")
    os.makedirs(cache_dir, exist_ok=True)

    success_count = 0
    for pkg in required:
        info = VCREDIST_PACKAGES[pkg]
        installed, version = check_vcredist_version_registry(pkg)

        if installed:
            print_status(f"{info['name']} v{version} – Already OK", "SUCCESS")
            success_count += 1
            time.sleep(2)
            continue

        installer_path = os.path.join(cache_dir, info["file"])
        if not os.path.exists(installer_path):
            print_status(f"Downloading {info['name']}...")
            try:
                with requests.get(info["url"], stream=True, timeout=60) as r:
                    r.raise_for_status()
                    total = int(r.headers.get("content-length", 0))
                    downloaded = 0
                    with open(installer_path, "wb") as f:
                        for chunk in r.iter_content(chunk_size=6000):  
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total > 0:
                                    pct = (downloaded / total) * 100
                                    print_status(f" Downloading: {pct:.1f}%", "INFO")
                                    time.sleep(0.2)
                print_status(" Download Complete!", "SUCCESS")
                time.sleep(3)
            except Exception as e:
                print_status(f"Download Failed: {e}", "ERROR")
                continue
        else:
            print_status(f"Using Cached (but still  install)", "INFO")
            time.sleep(2)

        print_status(f"Installing {info['name']} ...")
        time.sleep(2)
        try:
            subprocess.run(
                [installer_path, "/quiet", "/norestart"],
                check=True, capture_output=True, text=True, timeout=300
            )
            print_status("Install Done – Verifying ...", "INFO")
            time.sleep(3)
            if verify_vcredist_post_install(pkg):
                print_status("Verified!", "SUCCESS")
                success_count += 1
            else:
                print_status("Verify failed", "WARNING")
        except Exception as e:
            print_status(f"Install Failed: {e}", "ERROR")

    print_status(f"Visual C++ Done: {success_count}/{len(required)}", "SUCCESS")
    time.sleep(2)
    return success_count == len(required)

def install_autoit():
    if BUILD_ENV != "prod":
        print_status("AutoIt Setup", "INFO")
        installer = AUTOIT_INSTALLER_FILE

        if not os.path.exists(installer):
            print_status("Downloading AutoIt ...", "INFO")
            try:
                with requests.get(AUTOIT_INSTALLER_URL, stream=True, timeout=60) as r:
                    r.raise_for_status()
                    total = int(r.headers.get("content-length", 0))
                    downloaded = 0
                    with open(installer, "wb") as f:
                        for chunk in r.iter_content(6000):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total > 0:
                                    pct = (downloaded / total) * 100
                                    print_status(f"Downloading: {pct:.1f}%", "INFO")
                                    time.sleep(0.2)
                print_status("Download Complete!", "SUCCESS")
                time.sleep(4)
            except Exception as e:
                print_status(f"Download Failed: {e}", "ERROR")
                return False
        else:
            print_status("Using Cached but install", "INFO")
            time.sleep(3)

        print_status("Installing AutoIt ...", "INFO")
        time.sleep(3)
        try:
            subprocess.run([installer, "/S"], check=True, capture_output=True, text=True, timeout=300)
            print_status("AutoIt Installed!", "SUCCESS")
            time.sleep(3)
        except Exception as e:
            print_status(f"Install Failed: {e}", "ERROR")
            return False

        if os.path.isdir(AUTOIT_INSTALL_PATH):
            print_status("AutoIt Ready!", "SUCCESS")
            time.sleep(2)
            return True
        else:
            print_status("AutoIt Folder Not Found!", "ERROR")
            return False
    else:
        print_status("Skipping AutoIt Setup installation", "INFO")
        return True

def install_platform_dependencies():
    return install_all_vcredist_packages()

def log_to_file(message):
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except Exception as e:
        print(f"Logging failed: {e}")

def create_setup_lock():
    global CONFIG_TARGET_DIR
    if CONFIG_TARGET_DIR:
        lock_file_path = os.path.join(CONFIG_TARGET_DIR, "setup.lock")
        try:
            os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
            lock_data = {
                "build_env": BUILD_ENV,
                "setup_completed": time.strftime('%Y-%m-%d %H:%M:%S'),
                "selected_environment": SELECTED_ENVIRONMENT,
                "selected_version": SELECTED_VERSION
            }
            with open(lock_file_path, 'w') as f:
                json.dump(lock_data, f, indent=4)
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
    message = f"[{status}] {message}"
    log_to_file(message)
    if hasattr(InstallerApp, 'instance') and InstallerApp.instance:
        InstallerApp.instance.after(0, lambda: InstallerApp.instance.update_progress(message))
    print(message)

def get_db_name_by_environment(env, version=None):
    base_map = {
        "dev": "hitman_edge_dev",
        "test": "hitman_edge_test",
        "prod": "hitman_edge_prod"
    }
    base = base_map.get(env.lower(), "hitman_edge_dev")
    if version:
        clean_version = version.lstrip('v')
        return f"{base}_{clean_version}"
    return base

def update_config_paths(config_path):
    global db_folder_path, CONFIG_TARGET_DIR, SELECTED_ENVIRONMENT, SELECTED_VERSION
    if not CONFIG_TARGET_DIR:
        print_status("Target directory not set for config creation", "ERROR")
        return False
    print_status("Updating config.ini file...", "INFO")
    try:
        config_dir = os.path.dirname(config_path)
        os.makedirs(config_dir, exist_ok=True)
        config = configparser.ConfigParser()
        config.optionxform = str
        if os.path.exists(config_path):
            config.read(config_path, encoding='utf-8')
        if "database" not in config:
            config.add_section("database")
        if "app" not in config:
            config.add_section("app")
        if "paths" not in config:
            config.add_section("paths")
        if "version" not in config:
            config.add_section("version")
        print_status("Searching for Python installations...", "INFO")
        valid_pythons, _ = check_existing_python_installations()
        print_status(f"Valid Python installations found: {[p['path'] for p in valid_pythons]}", "INFO")
        python_exe = None
        is_bundled = getattr(sys, 'frozen', False)
        if is_bundled:
            print_status("Running as bundled executable, ignoring sys.executable", "INFO")
        else:
            current_exe = os.path.normpath(sys.executable)
            if (os.path.isfile(current_exe) and
                current_exe.lower().endswith('python.exe') and
                current_exe not in [p['path'] for p in valid_pythons]):
                try:
                    result = subprocess.run([current_exe, "--version"], capture_output=True, text=True, check=True)
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
        if valid_pythons:
            for python in valid_pythons:
                candidate_path = os.path.normpath(python['path'])
                if os.path.isfile(candidate_path) and candidate_path.lower().endswith('python.exe'):
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
                if python_exe.lower().endswith('python.exe'):
                    print_status(f"Installed Python executable: {python_exe}", "SUCCESS")
                else:
                    print_status(f"Installed path is not a Python executable: {python_path}", "ERROR")
                    return False
            else:
                print_status(f"Failed to install Python or invalid path: {python_path}", "ERROR")
                return False
        if not python_exe:
            print_status("No valid Python executable available", "ERROR")
            return False

        env_suffix = SELECTED_ENVIRONMENT if SELECTED_ENVIRONMENT else BUILD_ENV
        db_name = get_db_name_by_environment(env_suffix, SELECTED_VERSION)
        odbc_name = db_name

        config["database"]["HE_HOSTNAME"] = "127.0.0.1"
        config["database"]["HE_PORT"] = "3306"
        config["database"]["HE_DB_USERNAME"] = "Hitman"
        config["database"]["HE_DB_PASSWORD"] = "Hitman@123"
        config["database"]["HE_DB_DEV"] = db_name
        config["database"]["HE_DB_TEST"] = db_name
        config["database"]["HE_DB_PROD"] = db_name
        config["database"]["HE_ODBC_NAME"] = odbc_name
        config["database"]["HE_DBPARM"] = f"ConnectString='DRIVER={{MySQL ODBC 8.0 Unicode Driver}};SERVER=127.0.0.1;PORT=3306;DATABASE={db_name};UID=Hitman;PWD=Hitman@123;OPTION=3;'"
        config["database"]["HE_AUTOCOMMIT"] = "false"
        config["database"]["HE_DBMS"] = "ODBC"
        config["app"]["APP_NAME"] = "HitmanEdge"
        config["app"]["LANGUAGE"] = "en"
        config["app"]["TIMEZONE"] = "Asia/Kolkata"
        config["version"]["BUILD_ENV"] = BUILD_ENV
        config["version"]["SELECTED_ENVIRONMENT"] = env_suffix
        config["version"]["SELECTED_VERSION"] = SELECTED_VERSION if SELECTED_VERSION else ""
        config["version"]["LAST_UPDATE"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        config["paths"]["HE_ROOT_PATH"] = CONFIG_TARGET_DIR
        config["paths"]["HE_PYTHON_EXE"] = python_exe
        config["paths"]["HE_PYTHON_PATH"] = r"\script"
        config["paths"]["HE_IMAGE_PATH"] = r"\assets"
        config["paths"]["HE_PATH"] = config_path
        config["paths"]["HE_POWERBUILDER_EXE"] = os.path.join(CONFIG_TARGET_DIR, "powerbuilder")
        config["paths"]["HE_DATABASE"] = os.path.join(CONFIG_TARGET_DIR, "database")
        db_folder_path = config["paths"]["HE_DATABASE"]
        with open(config_path, 'w', encoding='utf-8') as f:
            config.write(f)
        os.chmod(config_path, 0o644)
        print_status(f"Config file updated successfully with HE_PYTHON_EXE={python_exe}", "SUCCESS")
        print_status(f"Environment: {env_suffix} | DB: {db_name} | ODBC: {odbc_name}", "INFO")
        return True
    except Exception as e:
        print_status(f"Failed to update config file: {e}", "ERROR")
        return False

def create_default_config_if_missing():
    if not CONFIG_TARGET_DIR:
        return None
    config_folder = os.path.join(CONFIG_TARGET_DIR, "config")
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
    if not db_folder_path or not os.path.exists(db_folder_path):
        print_status("Configured database folder does not exist.", "ERROR")
        return None

    for file in os.listdir(db_folder_path):
        if file.endswith('.sql'):
            sql_file = os.path.join(db_folder_path, file)
            print_status(f"Found SQL file: {sql_file}", "SUCCESS")
            return sql_file

    print_status("No SQL file found in database folder.", "INFO")
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
                    os.chmod(file_path, stat.S_IWRITE)
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

# def install_missing_packages_from_requirements(requirements_path=None, python_exe=sys.executable):
#     if not requirements_path:
#         requirements_path = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "script", "requirements.txt")
    
#     if not os.path.exists(requirements_path):
#         print_status("requirements.txt not found", "WARNING")
#         return False
    
#     print_status(f"Installing packages from {requirements_path}...")
#     try:
#         with open(requirements_path, 'r') as f:
#             requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        
#         subprocess_kwargs = {
#             'capture_output': True,
#             'text': True,
#             'startupinfo': subprocess.STARTUPINFO(),
#             'creationflags': subprocess.CREATE_NO_WINDOW
#         }
#         subprocess_kwargs['startupinfo'].dwFlags |= subprocess.STARTF_USESHOWWINDOW
#         subprocess_kwargs['startupinfo'].wShowWindow = subprocess.SW_HIDE

#         for requirement in requirements:
#             for cmd in [
#                 [python_exe, "-m", "pip", "install", requirement],
#                 [python_exe, "-m", "pip", "install", requirement, "--user"]
#             ]:
#                 try:
#                     result = subprocess.run(cmd, **subprocess_kwargs, check=True, timeout=300)
#                     print_status(f"Installed: {requirement}", "SUCCESS")
#                     break
#                 except Exception as e:
#                     print_status(f"Error during installation attempt for {requirement}: {e}", "WARNING")
#                     continue
#             else:
#                 print_status(f"Failed to install {requirement}", "ERROR")

#         print_status("Upgrading websockets package as final step...", "INFO")
#         cmd_websockets = [python_exe, "-m", "pip", "install", "--upgrade", "websockets"]
#         try:
#             result = subprocess.run(cmd_websockets, **subprocess_kwargs, check=True, timeout=300)
#             print_status("Upgraded websockets successfully", "SUCCESS")
#         except Exception as e:
#             print_status(f"Error upgrading websockets: {e}", "ERROR")
#             return False

#         return True
#     except Exception as e:
#         print_status(f"Package installation failed: {e}", "ERROR")
#         return False

def check_existing_python_installations():
    valid_pythons = []
    all_paths = []
    for path in PYTHON_INSTALL_PATHS:
        candidate = os.path.join(path, "python.exe")
        all_paths.append(candidate)
        if os.path.isfile(candidate):
            try:
                result = subprocess.run([candidate, "--version"], capture_output=True, text=True, check=True)
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
    print_status("Checking Python compatibility...", "INFO")
    python_exe = None
    for path in PYTHON_INSTALL_PATHS:
        candidate = os.path.join(path, "python.exe")
        if os.path.isfile(candidate):
            try:
                result = subprocess.run([candidate, "--version"], capture_output=True, text=True,check=True)
                version_str = result.stdout.strip().replace("Python ", "")
                version_parts = version_str.split(".")
                if len(version_parts) >= 2 and tuple(map(int, version_parts[:2])) >= (3, 7):
                    print_status(f"Found compatible Python: {candidate} (v{version_str})", "SUCCESS")
                    return True, candidate
            except Exception as e:
                print_status(f"Failed to validate Python at {candidate}: {e}", "WARNING")
                continue

    print_status("No compatible Python found. Installing...", "WARNING")
    installer_dir = EXE_PATH
    os.makedirs(installer_dir, exist_ok=True)
    installer_path = os.path.join(installer_dir, PYTHON_INSTALLER_FILE)

    if not os.path.exists(installer_path):
        print_status(f"Downloading Python from {PYTHON_INSTALLER_URL}...", "INFO")
        try:
            response = requests.get(PYTHON_INSTALLER_URL, stream=True, verify=True, timeout=30)
            response.raise_for_status()
            with open(installer_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print_status("Python installer downloaded successfully", "SUCCESS")
        except Exception as e:
            print_status(f"Failed to download Python installer: {e}", "ERROR")
            return False, None

    print_status(f"Installing Python to {PYTHON_DEFAULT_INSTALL}...", "INFO")
    try:
        cmd = [
            installer_path,
            "/quiet",
            "InstallAllUsers=1",
            f"TargetDir={PYTHON_DEFAULT_INSTALL}",
            "PrependPath=1",
            "Include_test=0",
            "Include_pip=1"
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, timeout=600)
        print_status("Python installation completed", "SUCCESS")
        
        python_exe = os.path.join(PYTHON_DEFAULT_INSTALL, "python.exe")
        if os.path.isfile(python_exe):
            result = subprocess.run([python_exe, "--version"], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print_status(f"Python verified: {python_exe} (v{result.stdout.strip()})", "SUCCESS")
                try:
                    os.chmod(installer_path, stat.S_IWRITE)
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
    print_status(f"Windows Version: {system_info['windows_version']}")
    print_status(f"Windows Edition: {system_info['windows_edition']}")
    return True, system_info

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
    httpd_conf_path = os.path.join(WEB_SERVER_PATH, "apache", "conf", "httpd.conf")
    apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
    mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")

    if not os.path.exists(httpd_conf_path):
        print_status("Apache configuration file not found", "ERROR")
        return False, 80

    possible_ports = [80, 8080, 8081, 8082, 8083]
    apache_port = None

    for port in possible_ports:
        if check_port_availability(port, "Web Server"):
            apache_port = port
            break

    if apache_port is None:
        print_status("No available ports found, defaulting to 8083", "WARNING")
        apache_port = 8083

    print_status(f"Configuring Apache to use port {apache_port}...")
    try:
        with open(httpd_conf_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        backup_path = httpd_conf_path + f".backup_{int(time.time())}"
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print_status(f"Configuration backup created: {os.path.basename(backup_path)}", "SUCCESS")

        if f"Listen {apache_port}" in content and f"ServerName localhost:{apache_port}" in content:
            print_status(f"Apache already configured for port {apache_port}", "SUCCESS")
        else:
            lines = content.split('\n')
            modified_lines = []
            changes_made = 0

            for line in lines:
                if line.strip().startswith("Listen ") and not line.strip().startswith(f"Listen {apache_port}"):
                    line = line.replace(line.strip(), f"Listen {apache_port}")
                    changes_made += 1
                    print_status(f"Updated: Listen to port {apache_port}", "SUCCESS")
                elif f"ServerName localhost:" in line:
                    line = line.replace(f"ServerName localhost:{line.split(':')[1].split()[0]}", f"ServerName localhost:{apache_port}")
                    changes_made += 1
                    print_status(f"Updated ServerName to port {apache_port}", "SUCCESS")
                elif any(f":{old_port}>" in line for old_port in possible_ports):
                    for old_port in possible_ports:
                        if f":{old_port}>" in line:
                            line = line.replace(f":{old_port}>", f":{apache_port}>")
                            changes_made += 1
                            print_status(f"Updated virtual host to port {apache_port}", "SUCCESS")
                            break
                modified_lines.append(line)

            if changes_made > 0:
                with open(httpd_conf_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(modified_lines))
                print_status(f"Apache configuration updated ({changes_made} changes)", "SUCCESS")
                print_status(f"Apache will run on http://localhost:{apache_port}", "INFO")
            else:
                print_status("No port configurations found to modify", "WARNING")

    except Exception as e:
        print_status(f"Apache configuration failed: {e}", "ERROR")
        return False, apache_port

    try:
        apache_service_name = "Apache-XAMPP"
        subprocess.run([apache_exe, "-k", "install", "-n", apache_service_name], check=True, capture_output=True, text=True)
        subprocess.run(["sc", "config", apache_service_name, "start=", "auto"], check=True, capture_output=True, text=True)
        subprocess.run(["net", "start", apache_service_name], check=True, capture_output=True, text=True)
        print_status(f"Apache service '{apache_service_name}' installed and started", "SUCCESS")
    except subprocess.CalledProcessError as e:
        print_status(f"Error installing/starting Apache service: {e.stderr}", "ERROR")

    try:
        mysql_service_name = "MySQL-XAMPP"
        subprocess.run([mysql_exe, "--install", mysql_service_name], check=True, capture_output=True, text=True)
        subprocess.run(["sc", "config", mysql_service_name, "start=", "auto"], check=True, capture_output=True, text=True)
        subprocess.run(["net", "start", mysql_service_name], check=True, capture_output=True, text=True)
        print_status(f"MySQL service '{mysql_service_name}' installed and started", "SUCCESS")
    except subprocess.CalledProcessError as e:
        print_status(f"Error installing/starting MySQL service: {e.stderr}", "ERROR")

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

    r = requests.get(XAMPP_INSTALLER_URL, verify=True)
    with open(INSTALLER_FILE_XAMPP, "wb") as f:
        f.write(r.content)
    subprocess.run(INSTALLER_FILE_XAMPP, check=True, shell=True)
    return True

def check_xampp_installation():
    apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
    mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")
    xampp_installed = os.path.exists(WEB_SERVER_PATH) and os.path.exists(apache_exe) and os.path.exists(mysql_exe)
    print_status(f"XAMPP check: {'Installed' if xampp_installed else 'Not installed'}", "INFO")
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
    
    apache_exe = os.path.join(WEB_SERVER_PATH, "apache", "bin", "httpd.exe")
    mysql_exe = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "mysqld.exe")
    mysql_ini = os.path.join(WEB_SERVER_PATH, "mysql", "bin", "my.ini")

    print_status(f"Starting Apache on port {apache_port}...")
    max_retries = 5
    for attempt in range(max_retries):
        if not is_port_open(apache_port):
            try:
                subprocess.Popen([apache_exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print_status("Waiting for Apache startup...")
                time.sleep(15)
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
                subprocess.Popen([mysql_exe, f"--defaults-file={mysql_ini}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print_status("Waiting for MySQL startup...")
                time.sleep(45)
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
            s.settimeout(2)
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
        response = requests.get(zip_url, stream=True, timeout=30)
        response.raise_for_status()
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
       
        print_status("Extracting repository files...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
       
        os.chmod(zip_path, stat.S_IWRITE)
        os.remove(zip_path)
        extracted_folders = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d)) and 'main' in d.lower()]
       
        if not extracted_folders:
            raise Exception("Repository extraction failed: main folder not found")
        inner_folder = os.path.join(temp_dir, extracted_folders[0])
       
        os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
       
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
       
        powerbuilder_exe_src = None
        for root, _, files in os.walk(inner_folder):
            if 'PowerBuilderInstaller_bootstrapper.exe' in files:
                powerbuilder_exe_src = os.path.join(root, 'PowerBuilderInstaller_bootstrapper.exe')
                break
       
        if powerbuilder_exe_src:
            shutil.move(powerbuilder_exe_src, POWERBUILDER_INSTALLER_FILE)
            print_status(f"Moved PowerBuilderInstaller_bootstrapper.exe to {POWERBUILDER_INSTALLER_FILE}", "SUCCESS")
        else:
            print_status("PowerBuilderInstaller_bootstrapper.exe not found in repository", "WARNING")
       
        src_config = os.path.join(inner_folder, "config", "config.ini")
        dest_config_dir = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config")
        dest_config = os.path.join(dest_config_dir, "config.ini")
       
        if os.path.exists(src_config):
            os.makedirs(dest_config_dir, exist_ok=True)
            shutil.move(src_config, dest_config)
            print_status(f"Moved config.ini to {dest_config}", "SUCCESS")
        else:
            print_status("config.ini not found in repository, creating default...", "INFO")
            os.makedirs(dest_config_dir, exist_ok=True)
            with open(dest_config, 'w') as f:
                f.write("[database]\n[app]\n[paths]\n[version]\n")
       
        user_config_path = dest_config
        update_config_paths(user_config_path)
        print_status(f"Updated config.ini in user path: {user_config_path}", "SUCCESS")
       
        program_files_dest = r"C:\Program Files\HitmanEdge"
        os.makedirs(program_files_dest, exist_ok=True)
       
        for folder_name in ['database', 'powerbuilder']:
            src = os.path.join(CONFIG_TARGET_DIR, folder_name)
            dst = os.path.join(program_files_dest, folder_name)
            if os.path.exists(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
                print_status(f"Copied {folder_name} to {program_files_dest}", "SUCCESS")
            else:
                print_status(f"Source folder {folder_name} not found in user path", "WARNING")
       
        program_config_path = os.path.join(program_files_dest, "powerbuilder", "config", "config.ini")
        if os.path.exists(program_config_path):
            original_target = CONFIG_TARGET_DIR
            try:
                globals()['CONFIG_TARGET_DIR'] = program_files_dest
                update_config_paths(program_config_path)
                print_status(f"Updated config.ini in Program Files: {program_config_path}", "SUCCESS")
            finally:
                globals()['CONFIG_TARGET_DIR'] = original_target
        else:
            print_status("config.ini not found in Program Files after copy", "ERROR")
       
        print_status(f"Both config.ini files updated with correct paths", "SUCCESS")
       
        shutil.rmtree(temp_dir)
        print_status("Repository extracted and dual-path setup completed", "SUCCESS")
        return CONFIG_TARGET_DIR
    except Exception as e:
        print_status(f"Repository download/extraction or config setup failed: {e}", "ERROR")
        return None

def install_powerbuilder_if_needed():
    if BUILD_ENV != "prod":
        possible_paths = [
            r"C:\Program Files (x86)\Appeon\PowerBuilder",
        ]
        installed = any(os.path.exists(path) for path in possible_paths)
        
        if installed:
            print_status("PowerBuilder already installed.", "INFO")
            return True

        if not os.path.exists(POWERBUILDER_INSTALLER_FILE):
            print_status("PowerBuilder installer not found in EXE_PATH. Ensure it's extracted from repo.", "ERROR")
            return False
        
        try:
            print_status("Running PowerBuilder installer...", "INFO")
            cmd = [POWERBUILDER_INSTALLER_FILE, "/s", "/v/qn"]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            if result.returncode == 0:
                print_status("PowerBuilder installation completed", "SUCCESS")
                try:
                    os.chmod(POWERBUILDER_INSTALLER_FILE, stat.S_IWRITE)
                except Exception as e:
                    print_status(f"Failed to clean up PowerBuilder installer: {e}", "WARNING")
                return True
            else:
                print_status(f"PowerBuilder installation failed: {result.stderr}", "ERROR")
                return False
        except Exception as e:
            print_status(f"PowerBuilder installation failed: {e}", "ERROR")
            return False
    else:
        print_status("Skipping PowerBuilder Setup installation ", "INFO")
        return True

def clean_config_folder():
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
                        os.chmod(item_path, stat.S_IWRITE)
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
        raise FileNotFoundError("config.ini not found.")
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
        for key in required_keys:
            if key not in db:
                raise KeyError(f"Missing key '{key}' in [database] section")
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
    print_status("Setting up MySQL databases and user...", "INFO")
    max_retries = 5
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
                print_status("Waiting 15 seconds before retry...", "INFO")
                time.sleep(15)
            else:
                print_status("MySQL connection failed after all retries", "ERROR")
                return False
    
    try:
        cur = conn.cursor()
        try:
            cur.execute(f"CREATE USER IF NOT EXISTS '{info['user']}'@'localhost' IDENTIFIED BY '{info['pass']}';")
            print_status(f"Created/Updated MySQL user: {info['user']}", "SUCCESS")
        except mysql.connector.Error as e:
            if e.errno == errorcode.ER_CANNOT_USER:
                print_status(f"User {info['user']} already exists", "SUCCESS")
            else:
                raise
        
        cur.execute(f"GRANT ALL PRIVILEGES ON *.* TO '{info['user']}'@'localhost';")
        cur.execute("FLUSH PRIVILEGES;")
        print_status("User privileges granted", "SUCCESS")
        
        created_dbs = []
        db_name = info["names"][0]
        try:
            cur.execute(f"DROP DATABASE IF EXISTS {db_name};")
            cur.execute(f"CREATE DATABASE {db_name} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;")
            created_dbs.append(db_name)
            print_status(f"Created database: {db_name}", "SUCCESS")
        except Exception as e:
            print_status(f"Database creation failed for {db_name}: {e}", "ERROR")
        
        conn.commit()
        cur.close()
        conn.close()
        print_status(f"Database setup completed: {len(created_dbs)}/1 database", "SUCCESS")
        return len(created_dbs) == 1
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
    target_db = info["names"][0]  # Only import to the single active DB (based on BUILD_ENV)

    try:
        with open(rf"{db_folder_path}\hitman_edge_dev.sql", "r", encoding="utf-8") as f:
            cmd = [mysql_exe, "-u", "root", target_db]
            result = subprocess.run(cmd, stdin=f, capture_output=True, text=True, shell=True, check=True)
            
            if result.returncode == 0:
                print_status(f"SQL imported to: {target_db}", "SUCCESS")
                success_count += 1
            else:
                print_status(f"SQL import failed for {target_db}: {result.stderr[:200]}", "ERROR")
    except Exception as e:
        print_status(f"SQL import error for {target_db}: {e}", "ERROR")
    
    print_status(f"SQL import completed: {success_count}/1 database")
    return success_count == 1

def setup_odbc_and_dsn(db_info):
    return setup_odbc_windows(db_info)

def setup_odbc_windows(db_info):
    print_status("Starting ODBC setup for 32-bit...", "INFO")

    driver_installed, driver_name = check_odbc_driver_installed()
    if not driver_installed:
        print_status("MySQL ODBC 8.0.42 (32-bit) driver not found. Downloading and installing...", "INFO")
        if not download_odbc_installer(MYSQL_ODBC_URL, INSTALLER_FILE_ODBC):
            print_status("Failed to download ODBC installer", "ERROR")
            return False
        if not install_mysql_odbc_driver(INSTALLER_FILE_ODBC, force_repair=False):
            print_status("Failed to install MySQL ODBC driver", "ERROR")
            return False
    else:
        print_status("MySQL ODBC 8.0.42 (32-bit) driver found. Repairing...", "INFO")
        if not download_odbc_installer(MYSQL_ODBC_URL, INSTALLER_FILE_ODBC):
            print_status("Failed to download ODBC installer for repair", "ERROR")
            return False
        if not install_mysql_odbc_driver(INSTALLER_FILE_ODBC, force_repair=True):
            print_status("Failed to repair MySQL ODBC driver", "ERROR")
            return False

    driver_name = "MySQL ODBC 8.0 Unicode Driver"
    print_status(f"Using ODBC driver: {driver_name}", "SUCCESS")

    dsn_success = create_system_dsn(db_info, driver_name)
    
    print_status("Testing ODBC connections...", "INFO")
    connection_test_result = create_system_dsn(db_info, driver_name)
    
    overall_success = dsn_success and connection_test_result
    print_status("ODBC Setup completed successfully" if overall_success else "ODBC Setup completed with issues", 
                 "SUCCESS" if overall_success else "WARNING")
    return overall_success

def download_odbc_installer(url, file_path):
    if os.path.exists(file_path):
        print_status(f"Using existing installer: {file_path}", "INFO")
        return True
    
    print_status(f"Downloading from {url}...", "INFO")
    try:
        response = requests.get(url, stream=True, verify=True, timeout=30)
        response.raise_for_status()
        expected_size = int(response.headers.get('content-length', 0))
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        actual_size = os.path.getsize(file_path)
        if actual_size == expected_size:
            print_status(f"Downloaded successfully (Size: {actual_size} bytes)", "SUCCESS")
            return True
        else:
            print_status(f"Download size mismatch: Expected {expected_size}, got {actual_size}", "ERROR")
            os.remove(file_path)
            return False
    except Exception as e:
        print_status(f"Download failed: {e}", "ERROR")
        return False

def parse_version(version_str):
    return tuple(map(int, (version_str or "0.0.0").split(".")))

def check_odbc_driver_installed():
    target_display_name = "MySQL ODBC 8.0 Unicode Driver"
    min_required_version = parse_version("8.0.42")
    uninstall_paths = [r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"]

    for path in uninstall_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, subkey_name) as subkey:
                            name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                            if name.strip() == target_display_name:
                                version_str, _ = winreg.QueryValueEx(subkey, "DisplayVersion")
                                print_status(f"Found (32-bit): {name}, Version: {version_str}", "SUCCESS")
                                if parse_version(version_str) >= min_required_version:
                                    print_status("Version is compatible (32-bit) ✅", "SUCCESS")
                                    return True, name
                                else:
                                    print_status("Version is outdated (32-bit) ⚠️", "WARNING")
                                    return False, name
                    except FileNotFoundError:
                        continue
                    except Exception:
                        continue
        except Exception as e:
            print_status(f"Registry access error (32-bit): {e}", "ERROR")
            continue

    print_status(f"{target_display_name} not found in registry (32-bit) ❌", "ERROR")
    return False, None

def install_mysql_odbc_driver(INSTALLER_FILE_ODBC, force_repair=False):
    try:
        if not os.path.exists(INSTALLER_FILE_ODBC):
            print_status(f"MSI file not found: {INSTALLER_FILE_ODBC}", "ERROR")
            return False

        if force_repair:
            cmd = f'msiexec /fa "{INSTALLER_FILE_ODBC}" /qn /norestart'
            print_status("Repairing MySQL ODBC driver...", "INFO")
        else:
            cmd = f'msiexec /i "{INSTALLER_FILE_ODBC}" /qn /norestart'
            print_status("Installing MySQL ODBC driver...", "INFO")
        
        powershell_cmd = ["powershell", "-Command", f'Start-Process cmd -ArgumentList \'/c {cmd}\' -Verb RunAs -Wait']
        result = subprocess.run(powershell_cmd, capture_output=True, text=True, check=True)
        
        if result.returncode == 0:
            print_status("MySQL ODBC driver operation successful", "SUCCESS")
            try:
                os.chmod(INSTALLER_FILE_ODBC, stat.S_IWRITE)
            except Exception as e:
                print_status(f"Failed to clean up ODBC installer: {e}", "WARNING")
            return True
        else:
            print_status(f"Operation failed with exit code {result.returncode}: {result.stderr}", "ERROR")
            return False
    except Exception as e:
        print_status(f"Error in install_mysql_odbc_driver: {e}", "ERROR")
        return False

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    script = os.path.abspath(sys.argv[0])
    params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)

def create_system_dsn(db_info, driver_name):
    print_status("Creating System DSN (32-bit)...", "INFO")
    db_name = db_info["names"][0]
    reg_driver = driver_name
    access = winreg.KEY_WRITE | winreg.KEY_WOW64_32KEY

    if is_dsn_exists(db_name):
        print_status(f"System DSN already exists: {db_name}", "INFO")
        return True

    try:
        dsn_path = f"SOFTWARE\\ODBC\\ODBC.INI\\{db_name}"
        dsn_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, dsn_path, 0, access)

        cfg = {
            "Driver": reg_driver,
            "Description": f"HitmanEdge {db_name} System Database Connection",
            "SERVER": db_info["host"],
            "PORT": str(db_info["port"]),
            "DATABASE": db_name,
            "UID": db_info["user"],
            "PWD": db_info["pass"],
            "OPTION": "3",
            "CHARSET": "utf8mb4",
            "SSLMODE": "DISABLED",
            "AUTO_RECONNECT": "1",
            "MULTI_STATEMENTS": "1"
        }
        for k, v in cfg.items():
            winreg.SetValueEx(dsn_key, k, 0, winreg.REG_SZ, v)
        winreg.CloseKey(dsn_key)

        src_path = r"SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
        src_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, src_path, 0, access)
        winreg.SetValueEx(src_key, db_name, 0, winreg.REG_SZ, reg_driver)
        winreg.CloseKey(src_key)

        print_status(f"System DSN created (32-bit): {db_name}", "SUCCESS")
        return True

    except PermissionError:
        print_status(f"Permission denied – run as Administrator: {db_name}", "ERROR")
        return False
    except Exception as e:
        print_status(f"DSN creation failed: {e}", "ERROR")
        return False

def is_dsn_exists(dsn_name: str) -> bool:
    try:
        path = r"SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path,
                             0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        winreg.QueryValueEx(key, dsn_name)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False
    except Exception:
        return False

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

def create_powerbuilder_connection_test():
    print_status("Testing PowerBuilder database connection...", "INFO")
    config_path1 = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "config")
    config_path = os.path.join(config_path1, "config.ini")
    
    try:
        db_info = load_config(config_path)
        success_count = 0
        for db_name in db_info["names"]:
            conn_str = f"DSN={db_name};UID={db_info['user']};PWD={db_info['pass']}"
            try:
                conn = pyodbc.connect(conn_str, timeout=5)
                cursor = conn.cursor()
                cursor.execute("SELECT 1 as test")
                result = cursor.fetchone()
                if result and result[0] == 1:
                    print_status(f"PowerBuilder ODBC connection successful for {db_name}", "SUCCESS")
                    success_count += 1
                cursor.close()
                conn.close()
            except Exception as e:
                print_status(f"PowerBuilder ODBC connection failed for {db_name}: {e}", "WARNING")
        if success_count == len(db_info["names"]):
            print_status("All PowerBuilder ODBC connections successful", "SUCCESS")
            return True
        else:
            print_status(f"PowerBuilder ODBC: {success_count}/{len(db_info['names'])} connections successful", "WARNING")
            return False
    except Exception as e:
        print_status(f"PowerBuilder connection test failed: {e}", "ERROR")
        return False
    
def open_target_directory():
    try:
        if os.path.exists(CONFIG_TARGET_DIR):
            print_status(f"Opening target directory: {CONFIG_TARGET_DIR}", "SUCCESS")
            os.startfile(CONFIG_TARGET_DIR)
            return True
        else:
            print_status("Target directory does not exist", "ERROR")
            return False
    except Exception as e:
        print_status(f"Could not open directory: {e}", "WARNING")
        return False

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
        self.full_pages = [
            WelcomePage, BrowsePage, EnvironmentPage,
            VersionPage, ConfirmationPage, InstallationPage, FinishPage
        ]

        if BUILD_ENV.lower() == "prod":
            self.pages = [WelcomePage, BrowsePage, VersionPage, ConfirmationPage, InstallationPage, FinishPage]
            self.selected_env = "prod"
            self.selected_version = None
        else:
            self.pages = self.full_pages[:]
            self.selected_env = "dev"
            self.selected_version = None

        self.page_index = 0
        self.install_thread = None
        self.success_steps = 0
        self.failed_steps = []
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
        global CONFIG_TARGET_DIR, SELECTED_ENVIRONMENT, SELECTED_VERSION
        CONFIG_TARGET_DIR = self.target_dir
        SELECTED_ENVIRONMENT = self.selected_env
        SELECTED_VERSION = self.selected_version
        if isinstance(self.current_frame, InstallationPage):
            self.current_frame.install_btn.config(state="disabled")
        if self.install_thread and self.install_thread.is_alive():
            return
        self.install_thread = threading.Thread(target=self.run_installation)
        self.install_thread.start()

    def run_installation(self):
        self.success_steps = 0
        self.failed_steps = []
        python_path = None
        config_path = None

        try:
            platform_info = get_platform_info()
            print_status(f"Detected OS: {platform_info['system']} {platform_info['release']}", "INFO")
        except Exception as e:
            print_status(f"Failed to get platform info: {e}", "ERROR")
            
        step_name = "Python compatibility check and installation"
        try:
            success, python_path = ensure_python_compatibility()
            if success:
                self.success_steps += 1
            else:
                print_status("Python installation/verification failed, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in Python step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Platform dependencies installation"
        try:
            if install_platform_dependencies():
                self.success_steps += 1
            else:
                print_status("Platform dependencies installation failed, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in platform dependencies step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "AutoIt installation"
        try:
            if IS_WINDOWS:
                if install_autoit():
                    self.success_steps += 1
                else:
                    print_status("AutoIt installation had issues, skipping...", "WARNING")
                    self.failed_steps.append(step_name)
            else:
                print_status("Skipping AutoIt installation (non-Windows)", "INFO")
                self.success_steps += 1
        except Exception as e:
            print_status(f"Error in AutoIt step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Repository download and extraction"
        try:
            repo_path = download_and_extract_repo()
            if repo_path:
                self.success_steps += 1
            else:
                print_status("Repository download/extraction failed, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in repo download step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "PowerBuilder installation"
        try:
            if install_powerbuilder_if_needed():
                self.success_steps += 1
            else:
                print_status("PowerBuilder installation failed, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in PowerBuilder installation step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Remove README files"
        try:
            remove_readme_files()
            print_status("README.md files removed", "SUCCESS")
            self.success_steps += 1
        except Exception as e:
            print_status(f"Error in removing READMEs: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Config file handling"
        try:
            config_path = find_config_ini(repo_path if 'repo_path' in locals() else CONFIG_TARGET_DIR)
            if not config_path:
                config_path = create_default_config_if_missing()
            if config_path:
                update_config_paths(config_path)
                self.success_steps += 1
            else:
                print_status("Config.ini not found or created, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in config step: {e}", "ERROR")
            self.failed_steps.append(step_name)


        # step_name = "Python packages installation"
        # try:
        #     requirements_path = os.path.join(CONFIG_TARGET_DIR, "powerbuilder", "script", "requirements.txt")
        #     if os.path.exists(requirements_path):
        #         if install_missing_packages_from_requirements(requirements_path, python_exe=python_path if python_path else sys.executable):
        #             self.success_steps += 1
        #         else:
        #             print_status("Python packages installation failed, skipping...", "WARNING")
        #             self.failed_steps.append(step_name)
        #     else:
        #         print_status("requirements.txt not found, skipping...", "WARNING")
        #         self.failed_steps.append(step_name)
        # except Exception as e:
        #     print_status(f"Error in Python packages step: {e}", "ERROR")
        #     self.failed_steps.append(step_name)

        step_name = "Load database info"
        try:
            if config_path:
                self.db_info = load_config(config_path)
                self.success_steps += 1
            else:
                print_status("No config path available for DB info, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in loading DB info: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "XAMPP installation and services startup"
        try:
            installed = check_xampp_installation() 
            running = is_xampp_running()

            if not installed:
                print_status("XAMPP not found, downloading and installing...")
                install_xampp_if_needed()
                running = False

            if not running:
                print_status("Starting XAMPP services...")
                if start_services_enhanced(self.db_info["port"] if self.db_info else 3306, self.db_info["host"] if self.db_info else 'localhost'):
                    self.success_steps += 1
                else:
                    print_status("Failed to start XAMPP services, skipping...", "WARNING")
                    self.failed_steps.append(step_name)
            else:
                print_status("XAMPP services already running.")
                self.success_steps += 1
        except Exception as e:
            print_status(f"Error in XAMPP step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Database setup"
        try:
            sql_file = find_sql_file()
            db_success = create_user_and_dbs(self.db_info) if self.db_info else False
            import_success = import_sql_to_all(sql_file, self.db_info) if self.db_info and sql_file else False
            if db_success and import_success:
                self.success_steps += 1
            else:
                print_status("Some database operations failed, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in database setup: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "ODBC/DSN setup"
        try:
            if setup_odbc_and_dsn(self.db_info if self.db_info else {}):
                self.success_steps += 1
            else:
                print_status("Connection setup had issues, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in ODBC/DSN step: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Create setup lock"
        try:
            if create_setup_lock():
                self.success_steps += 1
            else:
                print_status("Failed to create setup lock, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in setup lock: {e}", "ERROR")
            self.failed_steps.append(step_name)

        step_name = "Save version file"
        try:
            if save_version_file():
                self.success_steps += 1
            else:
                print_status("Failed to save version file, skipping...", "WARNING")
                self.failed_steps.append(step_name)
        except Exception as e:
            print_status(f"Error in save version file: {e}", "ERROR")
            self.failed_steps.append(step_name)

        self.next_page()


class WelcomePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Welcome to HitmanEdge Setup Wizard",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="This wizard will guide you through the installation of HitmanEdge.",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="Next >", command=master.next_page).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)
        btn_back = ttk.Button(btn_frame, text="< Back", command=master.back_page)
        btn_back.pack(side="right", padx=5)
        btn_back.state(["disabled"])


class BrowsePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Installation Folder",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="Please choose a folder to install HitmanEdge:",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        path_frame = tk.Frame(main_frame, bg=BG_COLOR)
        path_frame.pack(pady=20, fill="x", anchor="w")
        border_frame = tk.Frame(path_frame, bg="black", bd=1)
        border_frame.pack(side="left")
        self.path_var = tk.StringVar(value=DEFAULT_TARGET_DIR)
        tk.Entry(border_frame, textvariable=self.path_var,
                 font=("Arial", 10), relief="flat", width=50).pack(padx=.2, pady=.2, ipady=2)
        ttk.Button(path_frame, text="Browse", command=self.browse).pack(side="left", padx=7, pady=0, ipady=2)

        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="Next >", command=self.proceed).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="< Back", command=master.back_page).pack(side="right", padx=5)

    def browse(self):
        folder = filedialog.askdirectory(initialdir="C:\\")
        if folder:
            self.path_var.set(os.path.normpath(os.path.join(folder, "HitmanEdge")))

    def proceed(self):
        self.master.target_dir = self.path_var.get()
        self.master.next_page()


class EnvironmentPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Select Environment",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="Choose the environment for your installation:",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        env_frame = tk.Frame(main_frame, bg=BG_COLOR)
        env_frame.pack(pady=30, fill="both", expand=True)

        self.env_var = tk.StringVar(value="dev")
        environments = [
            ("Development", "dev", "For development and testing purposes."),
            ("Testing", "test", "For quality assurance and testing.")
        ]
        for label, value, desc in environments:
            radio_frame = tk.Frame(env_frame, bg=BG_COLOR)
            radio_frame.pack(pady=8, fill="x", anchor="w")
            tk.Radiobutton(radio_frame, text=label, variable=self.env_var, value=value,
                           font=("Arial", 11, "bold"), bg=BG_COLOR, fg=LABEL_COLOR,
                           activebackground=BG_COLOR, selectcolor="white", cursor="hand2").pack(anchor="w")
            tk.Label(radio_frame, text=desc, font=("Arial", 9), bg=BG_COLOR,
                     fg="gray", anchor="w").pack(anchor="w", padx=25)

        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="Next >", command=self.proceed).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="< Back", command=master.back_page).pack(side="right", padx=5)

    def proceed(self):
        self.master.selected_env = self.env_var.get()
        self.master.next_page()


class VersionPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Select Version",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="Choose the version to install:",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        version_frame = tk.Frame(main_frame, bg=BG_COLOR)
        version_frame.pack(pady=30, fill="x")

        self.version_var = tk.StringVar()
        versions = self.get_available_versions(master.selected_env)
        if versions:
            self.version_var.set(versions[0])

        tk.Label(version_frame, text="Available Versions:",
                 font=("Arial", 10, "bold"), bg=BG_COLOR).pack(anchor="w")
        dropdown = ttk.Combobox(version_frame, textvariable=self.version_var,
                                values=versions, state="readonly", width=30)
        dropdown.pack(pady=5, anchor="w")

        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="Next >", command=self.proceed).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="< Back", command=master.back_page).pack(side="right", padx=5)

    def get_available_versions(self, env: str):
        if env == "dev":
            return [f"v1.0.{i}" for i in range(5)]
        if env == "test":
            return [f"v1.0.{i}" for i in range(3)]
        if env == "prod":
            return [f"v1.0.{i}" for i in range(4)]
        return []

    def proceed(self):
        self.master.selected_version = self.version_var.get()
        self.master.next_page()


class ConfirmationPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        tk.Label(main_frame, text="Confirm Installation",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="Please review the installation settings:",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        info_frame = tk.Frame(main_frame, bg=BG_COLOR, relief="sunken", bd=1)
        info_frame.pack(pady=20, fill="both", expand=True)
        env_names = {"dev": "Development", "test": "Testing", "prod": "Production"}
        env_name = env_names.get(master.selected_env, "Development")
        version_suffix = f"_{master.selected_version}" if master.selected_version else ""
        db_name = f"hitman_edge_{master.selected_env}{version_suffix}"
        target_dir = master.target_dir
        if master.selected_version:
            target_dir = target_dir.rstrip(os.sep) + f"_{master.selected_version}"
        version_line = f"Version: {master.selected_version}" if master.selected_version else ""
        info_text = (
            f"Installation Directory: {target_dir}\n\n"
            f"Environment: {env_name}\n\n"
            f"Database: {db_name}\n\n"
            f"{version_line}"
        )
        tk.Label(info_frame, text=info_text, font=("Arial", 10),
                 bg="white", justify="left", anchor="w").pack(
                     padx=15, pady=15, fill="both", expand=True)
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        ttk.Button(btn_frame, text="Confirm", command=self.confirm_installation).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="< Back", command=master.back_page).pack(side="right", padx=5)

    def confirm_installation(self):
        global CONFIG_TARGET_DIR, SELECTED_ENVIRONMENT, SELECTED_VERSION
        version_suffix = f"_{self.master.selected_version}" if self.master.selected_version else ""
        base_dir = self.master.target_dir.rstrip(os.sep)
        CONFIG_TARGET_DIR = f"{base_dir}{version_suffix}"
        SELECTED_ENVIRONMENT = self.master.selected_env
        SELECTED_VERSION = self.master.selected_version

        if check_setup_lock():
            messagebox.showwarning(
                "Setup Already Completed",
                f"HitmanEdge setup already exists in:\n\n{CONFIG_TARGET_DIR}\n\n"
                "Please choose a different directory or remove the existing installation."
            )
            self.master.page_index = 1
            self.master.show_frame(BrowsePage)
            return

        os.makedirs(CONFIG_TARGET_DIR, exist_ok=True)
        print_status(f"Created installation directory: {CONFIG_TARGET_DIR}", "SUCCESS")

        env_names = {"dev": "Development", "test": "Testing", "prod": "Production"}
        env_name = env_names.get(SELECTED_ENVIRONMENT, "Development")
        db_name = f"hitman_edge_{SELECTED_ENVIRONMENT}{version_suffix}"
        popup = tk.Toplevel(self)
        popup.title("Installation Started")
        popup.configure(bg="white")
        version_line = f"Version: {SELECTED_VERSION}\n" if SELECTED_VERSION else ""
        tk.Label(
            popup,
            text=(
                f"HitmanEdge will be installed at:\n\n{CONFIG_TARGET_DIR}\n\n"
                f"Environment: {env_name}\n"
                f"Database: {db_name}\n"
                f"{version_line}\n"
                "Downloading and configuring required components..."
            ),
            font=("Arial", 9), bg="white", justify="left", anchor="w",
            padx=25, pady=25
        ).pack()
        popup.update_idletasks()
        x = self.winfo_rootx() + 250
        y = self.winfo_rooty() + 150
        popup.geometry(f"+{x}+{y}")

        def proceed():
            popup.destroy()
            repo_path = download_and_extract_repo()
            if repo_path:
                self.master.next_page()
            else:
                messagebox.showerror("Installation Failed", "Failed to download and extract repository.")
                self.master.page_index = 1
                self.master.show_frame(BrowsePage)

        popup.after(1500, lambda: [popup.destroy(), self.master.next_page()])


class InstallationPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Installing HitmanEdge",
                 font=("Arial", 14, "bold"), fg=LABEL_COLOR, bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")
        tk.Label(main_frame, text="Please wait while the installation completes.",
                 font=("Arial", 10), bg=BG_COLOR, anchor="w").pack(pady=1, fill="x")

        progress_frame = tk.Frame(main_frame, bg=BG_COLOR)
        progress_frame.pack(pady=10, fill="x")
        tk.Label(progress_frame, text="Installation Progress:",
                 font=("Arial", 10, "bold"), bg=BG_COLOR, anchor="w").pack(fill="x")
        self.step_progress_var = tk.StringVar(value="Step 0/13")
        tk.Label(progress_frame, textvariable=self.step_progress_var,
                 font=("Arial", 9), fg="blue", bg=BG_COLOR, anchor="w").pack(pady=2, fill="x")
        self.progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.progress_bar.pack(pady=5, fill="x")

        self.progress_text = tk.Text(main_frame, height=8, width=60,
                                     font=("Consolas", 9), state='disabled')
        self.progress_text.pack(pady=10, fill="both", expand=True)
        scrollbar = tk.Scrollbar(self.progress_text)
        scrollbar.pack(side="right", fill="y")
        self.progress_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.progress_text.yview)

        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        self.install_btn = ttk.Button(btn_frame, text="Install",
                                      command=master.start_installation)
        self.install_btn.pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=master.quit).pack(side="right", padx=5)

    def update_progress(self, message):
        self.progress_text.config(state='normal')
        self.progress_text.insert(tk.END, f"{message}\n")
        self.progress_text.see(tk.END)
        self.progress_text.config(state='disabled')
        self.update()

    def update_step_progress(self):
        cur = self.master.success_steps
        tot = self.master.total_steps
        self.step_progress_var.set(f"Step {cur}/{tot}")
        self.progress_bar['value'] = (cur / tot) * 100
        self.update()


class FinishPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG_COLOR)
        tk.Label(self, text="Installation Complete!",
                 font=("Arial", 14, "bold"), fg="green", bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x", padx=10)
        tk.Label(self, text="HitmanEdge has been successfully installed.",
                 font=("Arial", 10), bg=BG_COLOR,
                 anchor="w").pack(pady=1, fill="x", padx=10)
        summary_frame = tk.Frame(self, bg="white", relief="sunken", bd=1)
        summary_frame.pack(pady=20, padx=30, fill="both", expand=True)
        env_names = {"dev": "Development", "test": "Testing", "prod": "Production"}
        env_name = env_names.get(master.selected_env, "Development")
        version_suffix = f"_{master.selected_version}" if master.selected_version else ""
        db_name = f"hitman_edge_{master.selected_env}{version_suffix}"
        target_dir = master.target_dir
        if master.selected_version:
            target_dir = target_dir.rstrip(os.sep) + f"_{master.selected_version}"
        version_line = f"Version: {master.selected_version}\n" if master.selected_version else ""
        summary_text = (
            f"Setup Status: {master.success_steps}/{master.total_steps} steps completed\n"
            f"Installation Directory: {target_dir}\n"
            f"Environment: {env_name}\n"
            f"Database: {db_name}\n"
            f"{version_line}\n"
        )
        if master.failed_steps:
            summary_text += f"Failed Steps: {', '.join(master.failed_steps)}"
        else:
            summary_text += "All steps completed successfully."
        tk.Label(summary_frame, text=summary_text,
                 font=("Arial", 9), bg="white",
                 justify="left", anchor="w").pack(pady=10, padx=10, fill="both", expand=True)
        btn_frame = tk.Frame(self, bg=BG_COLOR)
        btn_frame.pack(side="bottom", fill="x", padx=1, pady=10)
        ttk.Button(btn_frame, text="Finish", command=self.finish_and_close).pack(side="right", padx=5)
    def finish_and_close(self):
        open_target_directory()
        self.master.quit()


def main():
    global LOG_FILE, CONFIG_TARGET_DIR
    lock_file = os.path.join(os.path.dirname(sys.argv[0]), "setup.lock")
    if os.path.exists(lock_file):
        sys.exit(1)

    try:
        with open(lock_file, 'w') as f:
            f.write(f"Lock created at {time.strftime('%Y-%m-%d %H:%M:%S')}")

        if not is_admin():
            run_as_admin()
            sys.exit()

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