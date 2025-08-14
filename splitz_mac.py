# Persyaratan: pip install PyQt5 qfluentwidgets pymupdf
import sys
import os
import fitz  # PyMuPDF
import webbrowser
import subprocess
import json
import re

from PyQt5.QtCore import Qt, QPoint, QSize, QStandardPaths
from PyQt5.QtWidgets import (QApplication, QWidget, QFileDialog, QVBoxLayout,
                             QHBoxLayout, QGridLayout, QInputDialog, QSplitter,
                             QTableWidgetItem, QAbstractItemView, QHeaderView,
                             QFormLayout, QLabel, QFrame, QToolButton,
                             QPushButton, QLineEdit, QMessageBox, QSpinBox, QMenu, QTableWidget, QComboBox,
                             QMainWindow, QSizePolicy)
from PyQt5.QtGui import QIcon

# Pustaka yang hanya kita gunakan untuk mendapatkan ikonnya yang bagus
from qfluentwidgets import FluentIcon as FIF

# --- DATA & FUNGSI BANTU (Tidak berubah) ---
TELU_TEMPLATE_DATA = [
    {"name": "Cover", "filename": "cover.pdf", "start_page": 1, "end_page": 1}, {"name": "Disclaimer", "filename": "disclaimer.pdf", "start_page": 2, "end_page": 2}, {"name": "Lembar Pengesahan", "filename": "lembarpersetujuan.pdf", "start_page": 3, "end_page": 3}, {"name": "Abstrak (ID)", "filename": "abstraksi.pdf", "start_page": 6, "end_page": 6}, {"name": "Abstract (EN)", "filename": "abstract.pdf", "start_page": 7, "end_page": 7}, {"name": "Kata Pengantar", "filename": "kpdi.pdf", "start_page": 8, "end_page": 8}, {"name": "Daftar Isi", "filename": "daftarisi.pdf", "start_page": 9, "end_page": 10}, {"name": "Daftar Gambar", "filename": "daftargambar.pdf", "start_page": 11, "end_page": 11}, {"name": "Daftar Tabel", "filename": "daftartabel.pdf", "start_page": 12, "end_page": 12}, {"name": "Daftar Lampiran", "filename": "daftarlampiran.pdf", "start_page": 13, "end_page": 13}, {"name": "BAB I", "filename": "bab1.pdf", "start_page": 14, "end_page": 15}, {"name": "BAB II", "filename": "bab2.pdf", "start_page": 16, "end_page": 19}, {"name": "BAB III", "filename": "bab3.pdf", "start_page": 20, "end_page": 31}, {"name": "BAB IV", "filename": "bab4.pdf", "start_page": 32, "end_page": 46}, {"name": "BAB V", "filename": "bab5.pdf", "start_page": 47, "end_page": 48}, {"name": "Daftar Pustaka", "filename": "dp.pdf", "start_page": 49, "end_page": 49}, {"name": "Lampiran", "filename": "lampiran.pdf", "start_page": 50, "end_page": 50},
]
def _get_template_filepath():
    data_dir = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    if not os.path.exists(data_dir): os.makedirs(data_dir)
    return os.path.join(data_dir, "templates.json")
def _load_all_templates():
    filepath = _get_template_filepath()
    if not os.path.exists(filepath):
        default_templates = {"Tel-U": TELU_TEMPLATE_DATA}
        with open(filepath, 'w', encoding='utf-8') as f: json.dump(default_templates, f, indent=4)
        return default_templates
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            templates = json.load(f); templates.setdefault("Tel-U", TELU_TEMPLATE_DATA); return templates
    except (json.JSONDecodeError, IOError): return {"Tel-U": TELU_TEMPLATE_DATA}
def _save_all_templates(templates_data):
    filepath = _get_template_filepath()
    with open(filepath, 'w', encoding='utf-8') as f: json.dump(templates_data, f, indent=4)

# --- JENDELA UTAMA (QMainWindow) ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Super Splitter PDF")
        self.resize(1000, 750)

        self.sections, self.all_templates, self.current_selected_row = [], {}, -1

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        self._init_file_folder_section(main_layout)
        self._init_splitter_section(main_layout)
        self._init_tools_and_actions_section(main_layout)

        self.retranslateUi()
        self.connect_signals()
        self.all_templates = _load_all_templates()
        self._populate_template_combo()
        self._update_detail_panel_state()
        
        self.setStyleSheet("") # Memastikan tampilan 100% native

    def _init_file_folder_section(self, parent_layout):
        layout = QGridLayout(); layout.setSpacing(10)
        self.pdf_path_edit, self.pdf_path_button = QLineEdit(self), QPushButton(self.tr("Choose..."), self)
        self.output_dir_edit, self.output_dir_button = QLineEdit(self), QPushButton(self.tr("Choose..."), self)
        layout.addWidget(QLabel(self.tr("Input PDF:")), 0, 0); layout.addWidget(self.pdf_path_edit, 0, 1); layout.addWidget(self.pdf_path_button, 0, 2)
        layout.addWidget(QLabel(self.tr("Output Folder:")), 1, 0); layout.addWidget(self.output_dir_edit, 1, 1); layout.addWidget(self.output_dir_button, 1, 2)
        parent_layout.addLayout(layout)

    def _init_splitter_section(self, parent_layout):
        self.splitter = QSplitter(Qt.Horizontal); parent_layout.addWidget(self.splitter, 1)

        # Kiri: Panel Tabel dengan tombol + dan -
        left_pane = QWidget()
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        self.table_widget = QTableWidget(self)
        self.table_widget.setColumnCount(5); self.table_widget.verticalHeader().hide()
        self.table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.table_widget)

        # Baris kontrol di bawah tabel
        table_controls_layout = QHBoxLayout()
        self.add_section_button = QToolButton(); self.add_section_button.setIcon(FIF.ADD.icon()); self.add_section_button.setToolTip(self.tr("Add a new section"))
        self.delete_section_button = QToolButton(); self.delete_section_button.setIcon(FIF.DELETE.icon()); self.delete_section_button.setToolTip(self.tr("Delete selected section"))
        table_controls_layout.addWidget(self.add_section_button)
        table_controls_layout.addWidget(self.delete_section_button)
        table_controls_layout.addStretch(1)
        left_layout.addLayout(table_controls_layout)
        
        self.splitter.addWidget(left_pane)
        
        # Kanan: Panel Detail
        self.detail_panel = QFrame(self); self.detail_panel.setFrameShape(QFrame.StyledPanel)
        detail_layout = QFormLayout(self.detail_panel); detail_layout.setContentsMargins(15, 15, 15, 15); detail_layout.setSpacing(10)
        self.detail_section_name_edit, self.detail_start_page_spin, self.detail_end_page_spin, self.detail_filename_edit = QLineEdit(self), QSpinBox(self), QSpinBox(self), QLineEdit(self)
        self.detail_start_page_spin.setRange(1, 9999); self.detail_end_page_spin.setRange(1, 9999)
        detail_layout.addRow(QLabel(f"<b>{self.tr('Section Details')}</b>")); detail_layout.addRow(self.tr("Section Name:"), self.detail_section_name_edit)
        detail_layout.addRow(self.tr("Start Page:"), self.detail_start_page_spin); detail_layout.addRow(self.tr("End Page:"), self.detail_end_page_spin)
        detail_layout.addRow(self.tr("File Name:"), self.detail_filename_edit)
        self.splitter.addWidget(self.detail_panel); self.splitter.setSizes([550, 450])

    def _init_tools_and_actions_section(self, parent_layout):
        tools_layout = QHBoxLayout(); tools_layout.setSpacing(10)
        tools_layout.addWidget(QLabel(self.tr("Template:")))
        self.template_combo = QComboBox(self); tools_layout.addWidget(self.template_combo)
        self.load_template_button = QPushButton(self.tr("Load"), self); tools_layout.addWidget(self.load_template_button)
        self.delete_template_button = QToolButton(self); self.delete_template_button.setIcon(FIF.DELETE.icon()); tools_layout.addWidget(self.delete_template_button)
        tools_layout.addStretch(1)
        self.save_template_button = QPushButton(self.tr("Save as Template..."), self); tools_layout.addWidget(self.save_template_button)
        self.split_button = QPushButton(self.tr("Start Splitting"), self); self.split_button.setDefault(True)
        tools_layout.addWidget(self.split_button)
        parent_layout.addLayout(tools_layout)

    def retranslateUi(self):
        self.setWindowTitle(self.tr("Super Splitter PDF"))
        self.table_widget.setHorizontalHeaderLabels([self.tr('No.'), self.tr('Section Name'), self.tr('Start'), self.tr('End'), self.tr('File Name')])

    def connect_signals(self):
        self.table_widget.itemSelectionChanged.connect(self._on_row_selection_changed)
        self.pdf_path_button.clicked.connect(self._browse_pdf)
        self.output_dir_button.clicked.connect(self._browse_output)
        self.add_section_button.clicked.connect(lambda: self._add_section())
        self.delete_section_button.clicked.connect(self._delete_selected_section)
        self.load_template_button.clicked.connect(self._load_selected_template)
        self.delete_template_button.clicked.connect(self._delete_selected_template)
        self.save_template_button.clicked.connect(self._save_as_template)
        self.split_button.clicked.connect(self._split_pdf)
        self.detail_section_name_edit.textChanged.connect(self._on_detail_changed)
        self.detail_start_page_spin.valueChanged.connect(self._on_detail_changed)
        self.detail_end_page_spin.valueChanged.connect(self._on_detail_changed)
        self.detail_filename_edit.textChanged.connect(self._on_detail_changed)
        
    def _populate_table(self):
        self.table_widget.blockSignals(True); self.table_widget.setRowCount(0)
        for i, section in enumerate(self.sections):
            self.table_widget.insertRow(i)
            for j, key in enumerate(['', 'name', 'start_page', 'end_page', 'filename']):
                item = QTableWidgetItem(str(section.get(key, '')) if j > 0 else str(i+1))
                if j in [0, 2, 3]: item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(i, j, item)
        self.table_widget.blockSignals(False); self._update_detail_panel_state()

    def _on_row_selection_changed(self):
        self.current_selected_row = self.table_widget.currentRow()
        self._update_detail_panel_state()
        if self.current_selected_row >= 0: self._populate_detail_panel(self.sections[self.current_selected_row])

    def _update_detail_panel_state(self):
        is_row_selected = self.current_selected_row >= 0
        self.detail_panel.setEnabled(is_row_selected)
        self.delete_section_button.setEnabled(is_row_selected) # Kontrol tombol hapus
        if not is_row_selected:
            for w in [self.detail_section_name_edit, self.detail_filename_edit]: w.clear()
            for w in [self.detail_start_page_spin, self.detail_end_page_spin]: w.setValue(0)

    def _populate_detail_panel(self, data):
        for w in [self.detail_section_name_edit, self.detail_start_page_spin, self.detail_end_page_spin, self.detail_filename_edit]: w.blockSignals(True)
        self.detail_section_name_edit.setText(data.get('name', '')); self.detail_start_page_spin.setValue(int(data.get('start_page', 0)))
        self.detail_end_page_spin.setValue(int(data.get('end_page', 0))); self.detail_filename_edit.setText(data.get('filename', ''))
        for w in [self.detail_section_name_edit, self.detail_start_page_spin, self.detail_end_page_spin, self.detail_filename_edit]: w.blockSignals(False)

    def _on_detail_changed(self):
        row = self.current_selected_row
        if not (0 <= row < len(self.sections)): return
        data = {"name": self.detail_section_name_edit.text(), "start_page": self.detail_start_page_spin.value(), "end_page": self.detail_end_page_spin.value(), "filename": self.detail_filename_edit.text()}
        self.sections[row] = data
        self.table_widget.blockSignals(True)
        for i, key in enumerate(['name', 'start_page', 'end_page', 'filename']): self.table_widget.item(row, i + 1).setText(str(data[key]))
        self.table_widget.blockSignals(False)

    def _add_section(self):
        self.sections.append({"name": self.tr("New Section"), "filename": "new_section.pdf", "start_page": 1, "end_page": 1})
        self._populate_table(); self.table_widget.selectRow(len(self.sections) - 1)

    def _delete_selected_section(self):
        row = self.table_widget.currentRow()
        if row >= 0:
            del self.sections[row]; self._populate_table()
            if self.sections: self.table_widget.selectRow(min(row, len(self.sections) - 1))

    def _browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, self.tr("Select PDF File"), "", "PDF Files (*.pdf)");
        if path: self.pdf_path_edit.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, self.tr("Select Output Folder"), "");
        if path: self.output_dir_edit.setText(path)
        
    def _populate_template_combo(self):
        current = self.template_combo.currentText()
        self.template_combo.blockSignals(True); self.template_combo.clear()
        self.template_combo.addItem(self.tr("Load a template..."))
        self.template_combo.addItems(sorted(self.all_templates.keys()))
        self.template_combo.blockSignals(False)
        if current in self.all_templates: self.template_combo.setCurrentText(current)

    def _save_as_template(self):
        if not self.sections: return QMessageBox.warning(self, "List is Empty", "There is no data to save.")
        name, ok = QInputDialog.getText(self, "Save Template", "Enter a name for the new template:")
        if ok and name:
            if name in self.all_templates: return QMessageBox.warning(self, "Name Exists", f"A template named '{name}' already exists.")
            self.all_templates[name] = [s.copy() for s in self.sections]
            _save_all_templates(self.all_templates); self._populate_template_combo()

    def _load_selected_template(self):
        name = self.template_combo.currentText()
        if not name or name == self.tr("Load a template..."): return
        self.sections = [item.copy() for item in self.all_templates.get(name, [])]
        self._populate_table()

    def _delete_selected_template(self):
        name = self.template_combo.currentText()
        if not name or name == self.tr("Load a template..."): return
        if name == "Tel-U": return QMessageBox.warning(self, "Cannot Delete", "The default 'Tel-U' template cannot be deleted.")
        if QMessageBox.question(self, "Confirm Deletion", f"Are you sure you want to delete the template '{name}'?") == QMessageBox.Yes:
            del self.all_templates[name]; _save_all_templates(self.all_templates); self._populate_template_combo()

    def _split_pdf(self):
        pdf_path, out_dir = self.pdf_path_edit.text(), self.output_dir_edit.text()
        if not pdf_path or not out_dir or not self.sections:
            return QMessageBox.warning(self, self.tr("Incomplete Input"), self.tr("Please provide an input PDF, an output folder, and at least one section."))
        try:
            doc = fitz.open(pdf_path); os.makedirs(out_dir, exist_ok=True)
            for section in self.sections:
                start, end, filename = section['start_page'], section['end_page'], section['filename']
                if not (0 < start <= end <= len(doc)): raise ValueError(f"Invalid page range for '{section['name']}' ({start}-{end}). Total pages: {len(doc)}")
                new_doc = fitz.open(); new_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
                if len(new_doc) > 0: new_doc.save(os.path.join(out_dir, filename))
                new_doc.close()
            doc.close()
            QMessageBox.information(self, self.tr("Success"), self.tr("{0} PDF sections were successfully split.").format(len(self.sections)))
        except Exception as e:
            QMessageBox.critical(self, self.tr("An Error Occurred"), self.tr("Failed to process the PDF: {0}").format(e))
        
if __name__ == '__main__':
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling); QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    window = MainWindow(); window.show(); sys.exit(app.exec_())