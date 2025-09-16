from cx_Freeze import Executable, setup
import sys

APP_NAME = "PCAD Batch Export"
APP_VERSION = "0.3.4"
APP_DESCRIPTION = "Batch export tool for PCAD files."
TARGET_EXE = "PCAD_Batch_Export.exe"
BUILD_DIR = "dist/PCAD_Batch_Export"
UPGRADE_CODE = "{3F582720-2F2C-4F29-95F9-50EF86BE5671}"
TARGET_DIR = r"[ProgramFilesFolder]\\PCAD Batch Export"
MSI_TARGET_NAME = "PCADBatchExport"

PACKAGES = [
    "os",
    "sys",
    "time",
    "re",
    "warnings",
    "PyPDF2",
    "pdf2image",
    "PIL",
    "io",
    "docx",
    "docxcompose",
    "PySide6",
    "PySide6.QtWidgets",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtNetwork",
    "PySide6.QtPrintSupport",
    "multiprocessing",
    "collections",
    "qt_material",
    "pythoncom",
    "pywinauto",
    "keyboard",
    "comtypes",
]

INCLUDES = [
    "qt_material",
    "PySide6.QtWidgets",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtNetwork",
    "PySide6.QtPrintSupport",
]

INCLUDE_FILES = [
    ("default.docx", "default.docx"),
    ("color-theme.xml", "color-theme.xml"),
    ("pdfinfo.exe", "pdfinfo.exe"),
    ("pdftoppm.exe", "pdftoppm.exe"),
    ("template.docx", "template.docx"),
]


BIN_EXCLUDES = ["Qt6WebEngineCore.dll", "Qt6WebEngine.dll", "Qt6WebEngineWidgets.dll", "QtPdf.dll", "QtPdfQuick.dll"]

EXCLUDES_MODULES = ["tkinter", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineCore", "PySide6.QtWebEngine", "PySide6.QtWebEngineQuick", "PySide6.QtWebEngineProcess", "PySide6.QtWebEngineCore.DebugTools"]

build_exe_options = {
    "packages": PACKAGES,
    "includes": INCLUDES,
    "include_files": INCLUDE_FILES,
    "include_msvcr": True,
    "excludes": EXCLUDES_MODULES,
    "build_exe": BUILD_DIR,
    "bin_excludes": BIN_EXCLUDES,
}

shortcut_table = [
    (
        "PCADBatchExportShortcut",
        "ProgramMenuFolder",
        "PCAD Batch Export",
        "TARGETDIR",
        f"[TARGETDIR]{TARGET_EXE}",
        None,
        "Launch PCAD Batch Export",
        None,
        None,
        None,
        None,
        "TARGETDIR",
    ),
]

bdist_msi_options = {
    "upgrade_code": UPGRADE_CODE,
    "add_to_path": False,
    "all_users": False,
    "initial_target_dir": TARGET_DIR,
    "target_name": MSI_TARGET_NAME,
    "data": {"Shortcut": shortcut_table},
}

base = "Win32GUI" if sys.platform == "win32" else None

executables = [
    Executable(
        "batch_export.py",
        base=base,
        target_name=TARGET_EXE,
    )
]

setup(
    name=APP_NAME,
    version=APP_VERSION,
    description=APP_DESCRIPTION,
    options={"build_exe": build_exe_options, "bdist_msi": bdist_msi_options},
    executables=executables,
)
