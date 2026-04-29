import img2pdf
import os
import sys
import tempfile
import shutil
import subprocess
from pathlib import Path
from typing import List, Optional, Dict, Tuple, Any

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QFileDialog, QComboBox,
    QCheckBox, QMessageBox, QTabWidget, QProgressBar, QSpinBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from fpdf import FPDF
from PIL import Image
import pythoncom
from win32com import client as wc
from PyPDF2 import PdfMerger

# ----------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------
APP_NAME = "PDF Converter & Merger"
VERSION = "2.1"

# ----------------------------------------------------------------------
# Conversion Worker (runs in a separate thread)
# ----------------------------------------------------------------------
class ConversionThread(QThread):
    """Thread that performs file-to-PDF conversions and optional merging."""
    progress_updated = pyqtSignal(int)
    status_message = pyqtSignal(str)       # current file progress text
    conversion_finished = pyqtSignal(dict) # final result dictionary

    def __init__(self, files: List[str], merge_to_single: bool,
                 output_folder_or_file: str, quality: int, parent=None):
        super().__init__(parent)
        self.files = files
        self.merge_to_single = merge_to_single
        self.output_target = output_folder_or_file  # folder or merged PDF path
        self.quality = quality
        self._canceled = False

    def run(self) -> None:
        """Main conversion entry point."""
        results = {
            'success_files': [],   # list of (original_file, output_path)
            'failed_files': [],    # list of (original_file, error_message)
            'merged_output': None, # string if merge_to_single and successful
            'cancelled': False
        }
        temp_files_to_clean = []   # list of temporary PDFs created for merging

        try:
            total = len(self.files)
            for idx, file_path in enumerate(self.files):
                if self._canceled:
                    results['cancelled'] = True
                    break

                self.status_message.emit(f"Processing {os.path.basename(file_path)}")
                progress = int((idx / total) * 100) if total > 0 else 0
                self.progress_updated.emit(progress)

                # Convert single file
                try:
                    out_pdf = self._convert_to_pdf(file_path)
                    if out_pdf:
                        results['success_files'].append((file_path, out_pdf))
                        if self.merge_to_single:
                            temp_files_to_clean.append(out_pdf)  # will be merged later
                    else:
                        # _convert_to_pdf returned None internally after printing error
                        results['failed_files'].append((file_path, "Unknown conversion error"))
                except Exception as e:
                    results['failed_files'].append((file_path, str(e)))

            # If not cancelled and merge requested, merge all PDFs
            if not results['cancelled'] and self.merge_to_single and results['success_files']:
                try:
                    merged_pdf = self.output_target
                    merger = PdfMerger()
                    for _, pdf_path in results['success_files']:
                        merger.append(pdf_path)
                    merger.write(merged_pdf)
                    merger.close()
                    results['merged_output'] = merged_pdf
                    # Remove individual temp PDFs after merging
                    for pdf_path in temp_files_to_clean:
                        try:
                            os.remove(pdf_path)
                        except OSError:
                            pass
                except Exception as e:
                    # If merging fails, report and keep individual files
                    results['failed_files'].append(("Merging", f"Merge failed: {e}"))
                    results['merged_output'] = None

            self.progress_updated.emit(100)
            self.status_message.emit("Conversion finished.")
        except Exception as e:
            results['failed_files'].append(("Global", f"Unexpected thread error: {e}"))
        finally:
            self.conversion_finished.emit(results)

    def cancel(self) -> None:
        """Request cancellation of the ongoing operation."""
        self._canceled = True

    def _convert_to_pdf(self, file_path: str) -> Optional[str]:
        """
        Convert a single file to PDF. Returns the output PDF path on success,
        or raises an exception on failure.
        """
        file_ext = os.path.splitext(file_path)[1].lower()
        # Determine output path: if not merging, each file goes to output folder
        # with the same base name. Else, a temp file is used (path comes from
        # the thread's logic).
        if not self.merge_to_single:
            # output_target is a folder
            base = os.path.splitext(os.path.basename(file_path))[0]
            out_path = os.path.join(self.output_target, f"{base}.pdf")
        else:
            # create a temporary PDF for later merging
            fd, out_path = tempfile.mkstemp(suffix='.pdf')
            os.close(fd)  # we'll truncate when we write, or use NamedTemporaryFile
            # but we need the path to return; using mkstemp is safe.
            # Note: we must delete it if conversion fails. We'll handle that.
            # Actually safer to use NamedTemporaryFile but we need the path.
            # Using tempfile.NamedTemporaryFile(delete=False) is simpler.
            # I'll switch to that.
            try:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                tmp.close()
                out_path = tmp.name
            except Exception:
                raise RuntimeError("Could not create temporary file.")

        # Perform format-specific conversion
        if file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'):
            self._image_to_pdf(file_path, out_path)
        elif file_ext in ('.doc', '.docx', '.rtf'):
            self._word_to_pdf(file_path, out_path)
        elif file_ext == '.pdf':
            # just copy
            shutil.copy2(file_path, out_path)
        else:
            # assume text file
            self._text_to_pdf(file_path, out_path)

        return out_path

    def _image_to_pdf(self, src: str, dst: str) -> None:
        """Convert image to PDF, using lossless or JPEG quality settings."""
        if self.quality > 0:
            with Image.open(src) as img:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(dst, "PDF", quality=self.quality)
        else:
            # lossless using img2pdf
            pdf_bytes = img2pdf.convert(src)
            with open(dst, 'wb') as f:
                f.write(pdf_bytes)

    def _word_to_pdf(self, src: str, dst: str) -> None:
        """Convert Word/RTF document to PDF using COM automation on Windows."""
        if sys.platform != 'win32':
            raise RuntimeError("Word conversion is supported only on Windows.")

        pythoncom.CoInitialize()
        word = None
        try:
            word = wc.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone
            doc = word.Documents.Open(src)
            # 17 = wdFormatPDF
            doc.SaveAs(dst, FileFormat=17)
            doc.Close()
        except Exception as e:
            raise RuntimeError(f"Word conversion error: {e}")
        finally:
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    def _text_to_pdf(self, src: str, dst: str) -> None:
        """Convert plain text file to PDF using FPDF."""
        try:
            with open(src, 'r', encoding='utf-8') as f:
                text = f.read()
        except UnicodeDecodeError:
            # try with a common Windows encoding
            with open(src, 'r', encoding='cp1252', errors='replace') as f:
                text = f.read()

        pdf = FPDF()
        pdf.add_page()
        # Use a built-in core font (Helvetica) to guarantee availability
        pdf.set_font("Helvetica", size=12)
        pdf.multi_cell(0, 10, txt=text)
        pdf.output(dst)


# ----------------------------------------------------------------------
# Main Application Window
# ----------------------------------------------------------------------
class PDFConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setGeometry(100, 100, 850, 650)
        self.setWindowIcon(self._load_app_icon())

        # State variables
        self.current_files: List[str] = []      # conversion tab file list
        self.merge_files: List[str] = []        # merge tab PDF list
        self.output_folder: Optional[str] = None  # selected output folder
        self.merge_output_path: Optional[str] = None

        self.conversion_thread: Optional[ConversionThread] = None

        # Central widget and main layout
        central = QWidget()
        self.setCentralWidget(central)
        self.main_layout = QVBoxLayout(central)

        self._create_tabs()
        self._create_status_bar()

    # ---------- Icon loading ----------
    def _load_app_icon(self) -> QIcon:
        """Try to load an application icon from known locations."""
        # Search in script directory and common icon paths
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(script_dir, "pdf_icon.png"),
            "pdf_icon.png",
            "/usr/share/icons/hicolor/48x48/apps/pdf.png",
        ]
        for path in candidates:
            if os.path.exists(path):
                return QIcon(path)
        # Fallback to Qt standard icon (a generic document)
        return QApplication.style().standardIcon(QApplication.style().SP_FileIcon)

    # ---------- UI Construction ----------
    def _create_tabs(self) -> None:
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        self._create_conversion_tab()
        self._create_merge_tab()
        self._create_settings_tab()

    def _create_conversion_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # ---- File list with management buttons ----
        file_widget = QWidget()
        file_hbox = QHBoxLayout(file_widget)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)

        add_btn = QPushButton("Add Files")
        add_btn.clicked.connect(self._add_files)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_files)
        clear_btn = QPushButton("Clear All")
        clear_btn.clicked.connect(self._clear_files)

        btn_vbox = QVBoxLayout()
        btn_vbox.addWidget(add_btn)
        btn_vbox.addWidget(remove_btn)
        btn_vbox.addWidget(clear_btn)
        btn_vbox.addStretch()

        file_hbox.addWidget(self.file_list)
        file_hbox.addLayout(btn_vbox)

        # ---- Output options ----
        out_widget = QWidget()
        out_hbox = QHBoxLayout(out_widget)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF", "Image (PNG)", "Image (JPEG)"])
        self.format_combo.setCurrentText("PDF")

        self.out_folder_btn = QPushButton("Select Output Folder")
        self.out_folder_btn.clicked.connect(self._select_output_folder)
        self.out_folder_lbl = QLabel("Output will be saved to the same folder as input")
        self.out_folder_lbl.setWordWrap(True)

        out_hbox.addWidget(QLabel("Output Format:"))
        out_hbox.addWidget(self.format_combo)
        out_hbox.addWidget(self.out_folder_btn)
        out_hbox.addWidget(self.out_folder_lbl)
        out_hbox.addStretch()

        # ---- Conversion options ----
        opt_widget = QWidget()
        opt_hbox = QHBoxLayout(opt_widget)
        self.merge_cb = QCheckBox("Merge into single PDF")
        self.merge_cb.setChecked(False)
        self.quality_spin = QSpinBox()
        self.quality_spin.setRange(0, 100)
        self.quality_spin.setValue(95)
        self.quality_spin.setToolTip(
            "Image quality (0 = lossless via img2pdf, 1-100 = JPEG quality in PDF)")

        opt_hbox.addWidget(self.merge_cb)
        opt_hbox.addWidget(QLabel("Quality:"))
        opt_hbox.addWidget(self.quality_spin)
        opt_hbox.addStretch()

        # ---- Progress and control ----
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_lbl = QLabel("")
        self.status_lbl.setVisible(False)

        control_hbox = QHBoxLayout()
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.clicked.connect(self._start_conversion)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self._cancel_conversion)
        self.cancel_btn.setEnabled(False)

        control_hbox.addStretch()
        control_hbox.addWidget(self.convert_btn)
        control_hbox.addWidget(self.cancel_btn)

        # ---- Assemble tab layout ----
        layout.addWidget(QLabel("<b>Files to Convert:</b>"))
        layout.addWidget(file_widget)
        layout.addWidget(QLabel("<b>Output Options:</b>"))
        layout.addWidget(out_widget)
        layout.addWidget(QLabel("<b>Conversion Options:</b>"))
        layout.addWidget(opt_widget)
        layout.addWidget(self.status_lbl)
        layout.addWidget(self.progress_bar)
        layout.addLayout(control_hbox)
        layout.addStretch()

        self.tabs.addTab(tab, "File Conversion")

    def _create_merge_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # PDF list
        pdf_widget = QWidget()
        pdf_hbox = QHBoxLayout(pdf_widget)

        self.pdf_list = QListWidget()
        self.pdf_list.setSelectionMode(QListWidget.ExtendedSelection)

        add_pdf_btn = QPushButton("Add PDFs")
        add_pdf_btn.clicked.connect(self._add_pdfs)
        remove_pdf_btn = QPushButton("Remove Selected")
        remove_pdf_btn.clicked.connect(self._remove_pdfs)
        clear_pdf_btn = QPushButton("Clear All")
        clear_pdf_btn.clicked.connect(self._clear_pdfs)
        up_btn = QPushButton("Move Up")
        up_btn.clicked.connect(self._move_pdf_up)
        down_btn = QPushButton("Move Down")
        down_btn.clicked.connect(self._move_pdf_down)

        pdf_btn_vbox = QVBoxLayout()
        pdf_btn_vbox.addWidget(add_pdf_btn)
        pdf_btn_vbox.addWidget(remove_pdf_btn)
        pdf_btn_vbox.addWidget(clear_pdf_btn)
        pdf_btn_vbox.addWidget(up_btn)
        pdf_btn_vbox.addWidget(down_btn)
        pdf_btn_vbox.addStretch()

        pdf_hbox.addWidget(self.pdf_list)
        pdf_hbox.addLayout(pdf_btn_vbox)

        # Output
        merge_out_widget = QWidget()
        merge_out_hbox = QHBoxLayout(merge_out_widget)
        self.merge_out_btn = QPushButton("Select Output File")
        self.merge_out_btn.clicked.connect(self._select_merge_output)
        self.merge_out_lbl = QLabel("Default: merged.pdf in Documents")
        self.merge_out_lbl.setWordWrap(True)

        merge_out_hbox.addWidget(self.merge_out_btn)
        merge_out_hbox.addWidget(self.merge_out_lbl)
        merge_out_hbox.addStretch()

        self.merge_btn = QPushButton("Merge PDFs")
        self.merge_btn.clicked.connect(self._merge_pdfs)

        layout.addWidget(QLabel("<b>PDFs to Merge:</b>"))
        layout.addWidget(pdf_widget)
        layout.addWidget(QLabel("<b>Output Options:</b>"))
        layout.addWidget(merge_out_widget)
        layout.addWidget(self.merge_btn)
        layout.addStretch()

        self.tabs.addTab(tab, "PDF Merge")

    def _create_settings_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Settings
        self.open_cb = QCheckBox("Open output file(s) after operation")
        self.open_cb.setChecked(True)
        self.overwrite_cb = QCheckBox("Overwrite existing files without asking")
        self.overwrite_cb.setChecked(False)

        settings_widget = QWidget()
        settings_vbox = QVBoxLayout(settings_widget)
        settings_vbox.addWidget(QLabel("<b>Application Settings:</b>"))
        settings_vbox.addWidget(self.open_cb)
        settings_vbox.addWidget(self.overwrite_cb)
        settings_vbox.addStretch()

        # About
        about_text = f"""
        <b>{APP_NAME}</b><br><br>
        Version {VERSION}<br>
        Supported formats:<br>
        - Images (JPG, PNG, BMP, GIF, TIFF)<br>
        - Microsoft Word (DOC, DOCX, RTF)<br>
        - Text files (TXT and other plain text)<br>
        - PDF merging<br><br>
        © 2024 PDF Tools
        """
        about_label = QLabel(about_text)
        about_label.setWordWrap(True)

        layout.addWidget(settings_widget)
        layout.addWidget(about_label)
        layout.addStretch()

        self.tabs.addTab(tab, "Settings & About")

    def _create_status_bar(self) -> None:
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready")

    # ---------- File list management ----------
    def _add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "",
            "All Supported Files (*.pdf *.doc *.docx *.rtf *.txt *.jpg *.jpeg *.png *.bmp *.gif *.tiff);;"
            "PDF (*.pdf);;Word Documents (*.doc *.docx);;"
            "Images (*.jpg *.jpeg *.png *.bmp *.gif *.tiff);;"
            "Text Files (*.txt *.rtf)")
        for f in files:
            if f not in self.current_files:
                self.current_files.append(f)
                self.file_list.addItem(os.path.basename(f))

    def _remove_files(self) -> None:
        for item in self.file_list.selectedItems():
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            del self.current_files[row]

    def _clear_files(self) -> None:
        self.file_list.clear()
        self.current_files.clear()

    def _select_output_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder = folder
            self.out_folder_lbl.setText(f"Output folder: {folder}")
        else:
            self.output_folder = None
            self.out_folder_lbl.setText("Same as input files")

    # ---------- PDF merge list management ----------
    def _add_pdfs(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)")
        for f in files:
            self.merge_files.append(f)
            self.pdf_list.addItem(os.path.basename(f))

    def _remove_pdfs(self) -> None:
        indices = [self.pdf_list.row(item) for item in self.pdf_list.selectedItems()]
        for idx in sorted(indices, reverse=True):
            self.pdf_list.takeItem(idx)
            del self.merge_files[idx]

    def _clear_pdfs(self) -> None:
        self.pdf_list.clear()
        self.merge_files.clear()

    def _move_pdf_up(self) -> None:
        row = self.pdf_list.currentRow()
        if row > 0:
            item = self.pdf_list.takeItem(row)
            self.pdf_list.insertItem(row - 1, item)
            self.pdf_list.setCurrentRow(row - 1)
            self.merge_files.insert(row - 1, self.merge_files.pop(row))

    def _move_pdf_down(self) -> None:
        row = self.pdf_list.currentRow()
        if row < self.pdf_list.count() - 1:
            item = self.pdf_list.takeItem(row)
            self.pdf_list.insertItem(row + 1, item)
            self.pdf_list.setCurrentRow(row + 1)
            self.merge_files.insert(row + 1, self.merge_files.pop(row))

    def _select_merge_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged PDF As", "", "PDF Files (*.pdf)")
        if path:
            self.merge_output_path = path
            self.merge_out_lbl.setText(f"Output: {os.path.basename(path)}")

    # ---------- Conversion workflow ----------
    def _start_conversion(self) -> None:
        if not self.current_files:
            QMessageBox.warning(self, "No Files", "Add files to convert first.")
            return

        if self.format_combo.currentText() != "PDF":
            QMessageBox.information(self, "Info", "Currently only PDF output is supported.")
            return

        # Determine output target
        if self.merge_cb.isChecked():
            # Merged: need a file path
            if self.output_folder:
                out_path = os.path.join(self.output_folder, "merged.pdf")
            else:
                out_path = os.path.join(os.path.dirname(self.current_files[0]), "merged.pdf")
            # Check overwrite policy for merged output
            if not self.overwrite_cb.isChecked() and os.path.exists(out_path):
                reply = QMessageBox.question(
                    self, "Overwrite?", f"{out_path}\nalready exists. Replace?",
                    QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    return
        else:
            # Single files: output goes to selected folder or to same directory as input
            folder = self.output_folder if self.output_folder else os.path.dirname(self.current_files[0])
            self.output_folder = folder   # update state
            out_path = folder

        # Disable UI during conversion
        self._set_conversion_controls_enabled(False)

        self.status_lbl.setText("Starting conversion...")
        self.status_lbl.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # Launch thread
        self.conversion_thread = ConversionThread(
            files=self.current_files,
            merge_to_single=self.merge_cb.isChecked(),
            output_folder_or_file=out_path,
            quality=self.quality_spin.value()
        )
        self.conversion_thread.progress_updated.connect(self.progress_bar.setValue)
        self.conversion_thread.status_message.connect(self.status_lbl.setText)
        self.conversion_thread.conversion_finished.connect(self._on_conversion_finished)
        self.conversion_thread.start()

    def _cancel_conversion(self) -> None:
        if self.conversion_thread and self.conversion_thread.isRunning():
            self.conversion_thread.cancel()
            self.status_lbl.setText("Cancelling...")
            self.cancel_btn.setEnabled(False)

    def _on_conversion_finished(self, results: Dict[str, Any]) -> None:
        """Called when the conversion thread completes."""
        self._set_conversion_controls_enabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setVisible(False)

        success = results['success_files']
        failed = results['failed_files']
        merged = results['merged_output']
        cancelled = results['cancelled']

        # Build detailed message
        msg = ""
        if cancelled:
            msg += "Conversion was cancelled.\n"
        if success:
            if merged:
                msg += f"Merged PDF created:\n{merged}\n\n"
            else:
                msg += "Successfully converted files:\n"
                for src, out in success:
                    msg += f"  {os.path.basename(src)} → {os.path.basename(out)}\n"
        if failed:
            msg += "\nFailed conversions:\n"
            for src, err in failed:
                msg += f"  {os.path.basename(src)}: {err}\n"

        if cancelled and not success and not failed:
            QMessageBox.information(self, "Cancelled", "Operation cancelled.")
        elif failed:
            QMessageBox.critical(self, "Conversion Errors", msg)
        else:
            QMessageBox.information(self, "Success", msg)

        # Open output if requested and at least one success
        if success and self.open_cb.isChecked():
            if merged and os.path.exists(merged):
                self._open_file(merged)
            elif not merged:
                # open the first converted file (or all? safer to open the folder)
                # we open the parent folder of the first output
                if success:
                    first_out = success[0][1]
                    self._open_file(first_out)

    def _set_conversion_controls_enabled(self, enabled: bool) -> None:
        """Enable/disable UI elements during conversion."""
        self.convert_btn.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)
        self.file_list.setEnabled(enabled)
        self.out_folder_btn.setEnabled(enabled)
        self.format_combo.setEnabled(enabled)
        self.merge_cb.setEnabled(enabled)
        self.quality_spin.setEnabled(enabled)

    # ---------- PDF merge (explicit tab) ----------
    def _merge_pdfs(self) -> None:
        if len(self.merge_files) < 2:
            QMessageBox.warning(self, "Error", "Add at least two PDFs to merge.")
            return

        if not self.merge_output_path:
            docs = os.path.expanduser("~/Documents")
            os.makedirs(docs, exist_ok=True)
            self.merge_output_path = os.path.join(docs, "merged.pdf")

        # Overwrite check
        if not self.overwrite_cb.isChecked() and os.path.exists(self.merge_output_path):
            reply = QMessageBox.question(
                self, "Overwrite?", f"Replace existing {os.path.basename(self.merge_output_path)}?",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return

        try:
            merger = PdfMerger()
            for pdf in self.merge_files:
                merger.append(pdf)
            merger.write(self.merge_output_path)
            merger.close()
            QMessageBox.information(self, "Success", f"Merged PDF saved to:\n{self.merge_output_path}")
            if self.open_cb.isChecked():
                self._open_file(self.merge_output_path)
        except Exception as e:
            QMessageBox.critical(self, "Merge Error", str(e))

    # ---------- Utility ----------
    @staticmethod
    def _open_file(path: str) -> None:
        """Open a file with the default system application securely."""
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=True)
            else:
                subprocess.run(["xdg-open", path], check=True)
        except Exception:
            pass

# ----------------------------------------------------------------------
# Application entry point
# ----------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PDFConverterApp()
    window.show()
    sys.exit(app.exec_())