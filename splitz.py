# Persyaratan: pip install PyQt5 PyQt-Fluent-Widgets==1.8.6 pymupdf
import sys
import os
import fitz  # PyMuPDF
import webbrowser
from typing import List, Dict
import subprocess
import json
import warnings
import re

warnings.filterwarnings("ignore", category=DeprecationWarning)

from PyQt5.QtCore import Qt, QEvent, QSettings, pyqtSignal, QStandardPaths, QTranslator, QLocale, QSize
from PyQt5.QtWidgets import (QApplication, QWidget, QFileDialog,
                             QVBoxLayout, QHBoxLayout, QGridLayout,
                             QInputDialog, QSplitter, QTableWidgetItem,
                             QAbstractItemView, QHeaderView,
                             QFormLayout)
from PyQt5.QtGui import QIcon

from qfluentwidgets import (
    FluentWindow, setTheme, Theme, isDarkTheme,
    PushButton, PrimaryPushButton, ToolButton, LineEdit, SpinBox,
    TableWidget,
    MessageBox, InfoBar, InfoBarPosition, FluentIcon, BodyLabel,
    SubtitleLabel, ComboBox, NavigationItemPosition,
    CardWidget
)

if sys.platform == 'win32':
    import ctypes

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Data template
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
            templates = json.load(f)
            if "Tel-U" not in templates: templates["Tel-U"] = TELU_TEMPLATE_DATA
            return templates
    except (json.JSONDecodeError, IOError): return {"Tel-U": TELU_TEMPLATE_DATA}
def _save_all_templates(templates_data):
    filepath = _get_template_filepath()
    with open(filepath, 'w', encoding='utf-8') as f: json.dump(templates_data, f, indent=4)


class MainInterface(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainInterface")
        self.sections: List[Dict] = []
        self.all_templates = _load_all_templates()
        self.current_selected_row: int = -1

        # --- PERBAIKAN: Muat ikon PNG untuk tema terang dan gelap ---
        self.link_icon_light = QIcon(resource_path("icons/link_light.png"))
        self.link_icon_dark = QIcon(resource_path("icons/link_dark.png"))
        self.unlink_icon_light = QIcon(resource_path("icons/unlink_light.png"))
        self.unlink_icon_dark = QIcon(resource_path("icons/unlink_dark.png"))

        self.file_settings_card = CardWidget(self)
        self.detail_card = CardWidget(self)
        self.tools_template_card = CardWidget(self)
        self.main_layout = QVBoxLayout(self)
        self.file_settings_layout = QGridLayout(self.file_settings_card)
        self.pdf_path_edit = LineEdit(self)
        self.pdf_path_button = PushButton(self)
        self.open_pdf_button = ToolButton(FluentIcon.DOCUMENT, self)
        self.output_dir_edit = LineEdit(self)
        self.output_dir_button = PushButton(self)
        self.open_output_dir_button = ToolButton(FluentIcon.FOLDER, self)
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setStyleSheet("QSplitter::handle { background: transparent; image: none; } QSplitter { border: none; }")
        self.table_widget = TableWidget(self)
        self.detail_panel = QWidget()
        self.detail_layout = QFormLayout(self.detail_panel)
        self.detail_section_name_edit = LineEdit(self)
        self.detail_start_page_spin = SpinBox(self)
        self.detail_end_page_spin = SpinBox(self)
        self.detail_filename_edit = LineEdit(self)
        self.delete_section_button = PushButton(self)
        self.link_button = ToolButton(self)
        self.link_button.setCheckable(True)
        self.link_button.setFixedSize(32, 32)
        self.link_button.setIconSize(QSize(20, 20))
        self.template_layout = QHBoxLayout(self.tools_template_card)
        self.template_combo = ComboBox(self)
        self.load_template_button = PushButton(self)
        self.delete_template_button = ToolButton(FluentIcon.DELETE, self)
        self.save_template_button = PushButton(self)
        self.add_section_button = PushButton(self)
        self.clear_all_button = PushButton(self)
        self.split_button = PrimaryPushButton(self)
        
        self.init_layout()
        self.retranslateUi()
        self.connect_signals()
        self._populate_template_combo()
        self._update_detail_panel_state()

    def init_layout(self):
        self.main_layout.setContentsMargins(30, 20, 30, 20)
        self.main_layout.setSpacing(20)
        self.file_settings_label = SubtitleLabel(self)
        self.main_layout.addWidget(self.file_settings_label)
        self.main_layout.addWidget(self.file_settings_card)
        self.file_settings_layout.setContentsMargins(20, 15, 20, 20)
        self.file_settings_layout.setVerticalSpacing(15)
        self.pdf_input_label = BodyLabel(self)
        self.file_settings_layout.addWidget(self.pdf_input_label, 0, 0)
        self.file_settings_layout.addWidget(self.pdf_path_edit, 0, 1)
        self.file_settings_layout.addWidget(self.pdf_path_button, 0, 2)
        self.file_settings_layout.addWidget(self.open_pdf_button, 0, 3)
        self.output_folder_label = BodyLabel(self)
        self.file_settings_layout.addWidget(self.output_folder_label, 1, 0)
        self.file_settings_layout.addWidget(self.output_dir_edit, 1, 1)
        self.file_settings_layout.addWidget(self.output_dir_button, 1, 2)
        self.file_settings_layout.addWidget(self.open_output_dir_button, 1, 3)
        self.sections_config_label = SubtitleLabel(self)
        self.main_layout.addWidget(self.sections_config_label)
        self.main_layout.addWidget(self.splitter, 1)
        self.table_widget.setColumnCount(5)
        self.table_widget.verticalHeader().hide()
        self.table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.setWordWrap(False)
        self.table_widget.setColumnWidth(0, 50)
        self.table_widget.setColumnWidth(1, 200)
        self.table_widget.setColumnWidth(2, 60)
        self.table_widget.setColumnWidth(3, 60)
        self.splitter.addWidget(self.table_widget)
        self.detail_card_layout = QVBoxLayout(self.detail_card)
        self.detail_card_layout.setContentsMargins(15, 15, 15, 15)
        self.detail_card_layout.addWidget(self.detail_panel)
        self.detail_layout.setVerticalSpacing(15)
        self.detail_section_details_label = SubtitleLabel(self)
        self.detail_layout.addRow(self.detail_section_details_label)
        self.detail_section_name_label = BodyLabel(self)
        self.detail_layout.addRow(self.detail_section_name_label, self.detail_section_name_edit)
        self.detail_start_page_label = BodyLabel(self)
        self.detail_layout.addRow(self.detail_start_page_label, self.detail_start_page_spin)
        self.detail_end_page_label = BodyLabel(self)
        self.detail_layout.addRow(self.detail_end_page_label, self.detail_end_page_spin)
        self.detail_filename_label = BodyLabel(self)
        filename_layout = QHBoxLayout()
        filename_layout.setContentsMargins(0,0,0,0)
        filename_layout.setSpacing(5)
        filename_layout.addWidget(self.detail_filename_edit)
        filename_layout.addWidget(self.link_button)
        self.detail_layout.addRow(self.detail_filename_label, filename_layout)
        self.detail_layout.addRow(self.delete_section_button)
        self.splitter.addWidget(self.detail_card)
        self.splitter.setSizes([600, 400])
        self.tools_template_label = SubtitleLabel(self)
        self.main_layout.addWidget(self.tools_template_label)
        self.main_layout.addWidget(self.tools_template_card)
        self.template_layout.setContentsMargins(15, 15, 15, 15)
        self.template_label = BodyLabel(self)
        self.template_layout.addWidget(self.template_label)
        self.template_layout.addWidget(self.template_combo, 1)
        self.template_layout.addWidget(self.load_template_button)
        self.template_layout.addWidget(self.delete_template_button)
        self.template_layout.addStretch(1)
        self.template_layout.addWidget(self.clear_all_button)
        self.template_layout.addWidget(self.add_section_button)
        self.template_layout.addWidget(self.save_template_button)
        self.action_layout = QHBoxLayout()
        self.action_layout.addStretch(1)
        self.action_layout.addWidget(self.split_button)
        self.main_layout.addLayout(self.action_layout)

    def retranslateUi(self):
        self.table_widget.setHorizontalHeaderLabels([
            self.tr('No.'), self.tr('Section Name'), self.tr('Start'), 
            self.tr('End'), self.tr('File Name')
        ])
        self.file_settings_label.setText(self.tr("File & Folder"))
        self.sections_config_label.setText(self.tr("Sections Configuration"))
        self.tools_template_label.setText(self.tr("Tools & Templates"))
        self.pdf_input_label.setText(self.tr("Input PDF File:"))
        self.output_folder_label.setText(self.tr("Output Folder:"))
        self.pdf_path_edit.setPlaceholderText(self.tr("Select a PDF file to split..."))
        self.output_dir_edit.setPlaceholderText(self.tr("Select a folder to save the results..."))
        self.pdf_path_button.setText(self.tr("Choose..."))
        self.output_dir_button.setText(self.tr("Choose..."))
        self.open_pdf_button.setToolTip(self.tr("Open PDF in default application"))
        self.open_output_dir_button.setToolTip(self.tr("Open output folder in File Explorer"))
        self.detail_section_details_label.setText(self.tr("Section Details"))
        self.detail_section_name_label.setText(self.tr("Section Name:"))
        self.detail_start_page_label.setText(self.tr("Start Page:"))
        self.detail_end_page_label.setText(self.tr("End Page:"))
        self.detail_filename_label.setText(self.tr("File Name:"))
        self.delete_section_button.setText(self.tr("Delete This Section"))
        self.delete_section_button.setIcon(FluentIcon.DELETE)
        self.template_label.setText(self.tr("Template:"))
        self.load_template_button.setText(self.tr("Load"))
        self.delete_template_button.setToolTip(self.tr("Delete selected template"))
        self.link_button.setToolTip(self.tr("Link file name to section name"))
        self.add_section_button.setText(self.tr("Add Section"))
        self.add_section_button.setIcon(FluentIcon.ADD)
        self.clear_all_button.setText(self.tr("Clear All"))
        self.clear_all_button.setIcon(FluentIcon.DELETE)
        self.save_template_button.setText(self.tr("Save as Template"))
        self.save_template_button.setIcon(FluentIcon.SAVE)
        self.split_button.setText(self.tr("START SPLITTING PROCESS"))
        self.split_button.setIcon(FluentIcon.PLAY)
        self._update_link_button_state(self.link_button.isChecked())

    def connect_signals(self):
        self.table_widget.itemSelectionChanged.connect(self._on_row_selection_changed)
        self.pdf_path_button.clicked.connect(self._browse_pdf)
        self.output_dir_button.clicked.connect(self._browse_output)
        self.open_pdf_button.clicked.connect(self._open_pdf)
        self.open_output_dir_button.clicked.connect(self._open_output_dir)
        self.load_template_button.clicked.connect(self._load_selected_template)
        self.delete_template_button.clicked.connect(self._delete_selected_template)
        self.save_template_button.clicked.connect(self._save_as_template)
        self.add_section_button.clicked.connect(self._add_section)
        self.clear_all_button.clicked.connect(self._clear_all_sections)
        self.split_button.clicked.connect(self._split_pdf)
        self.detail_section_name_edit.textChanged.connect(self._on_detail_changed)
        self.detail_start_page_spin.valueChanged.connect(self._on_detail_changed)
        self.detail_end_page_spin.valueChanged.connect(self._on_detail_changed)
        self.detail_filename_edit.textChanged.connect(self._on_filename_edited_manually)
        self.link_button.clicked.connect(self._on_link_button_toggled)
        self.delete_section_button.clicked.connect(self._delete_selected_section)

    def _populate_table(self):
        self.table_widget.blockSignals(True)
        self.table_widget.setRowCount(0)
        for i, section in enumerate(self.sections):
            self.table_widget.insertRow(i)
            no_item = QTableWidgetItem(str(i + 1))
            name_item = QTableWidgetItem(section.get('name', ''))
            start_item = QTableWidgetItem(str(section.get('start_page', '')))
            end_item = QTableWidgetItem(str(section.get('end_page', '')))
            file_item = QTableWidgetItem(section.get('filename', ''))
            no_item.setTextAlignment(Qt.AlignCenter)
            start_item.setTextAlignment(Qt.AlignCenter)
            end_item.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(i, 0, no_item)
            self.table_widget.setItem(i, 1, name_item)
            self.table_widget.setItem(i, 2, start_item)
            self.table_widget.setItem(i, 3, end_item)
            self.table_widget.setItem(i, 4, file_item)
        self.table_widget.blockSignals(False)
        self._update_detail_panel_state()

    def _on_row_selection_changed(self):
        row = self.table_widget.currentRow()
        if row == self.current_selected_row: return
        self.current_selected_row = row
        self._update_detail_panel_state()
        if 0 <= row < len(self.sections):
            section = self.sections[row]
            self._populate_detail_panel(section)
            is_linked = section.get("is_linked", True)
            self._update_link_button_state(is_linked)

    def _sanitize_filename(self, text: str) -> str:
        if not text: return ""
        text = text.lower().replace(" ", "_")
        text = re.sub(r'[^a-z0-9_-]', '', text)
        return re.sub(r'__+', '_', text)

    def _update_link_button_state(self, is_linked: bool):
        # --- PERBAIKAN: Pilih ikon PNG berdasarkan tema ---
        if is_linked:
            icon = self.link_icon_dark if isDarkTheme() else self.link_icon_light
            self.link_button.setToolTip(self.tr("File name is linked. Click to unlink."))
        else:
            icon = self.unlink_icon_dark if isDarkTheme() else self.unlink_icon_light
            self.link_button.setToolTip(self.tr("File name is unlinked. Click to re-link."))
        self.link_button.setIcon(icon)
        self.link_button.setChecked(is_linked)

    def _on_filename_edited_manually(self):
        if self.detail_filename_edit.signalsBlocked() or not self.detail_card.isEnabled(): return
        row = self.current_selected_row
        if 0 <= row < len(self.sections) and self.sections[row].get("is_linked", False):
            self.sections[row]["is_linked"] = False
            self._update_link_button_state(False)
        self._on_detail_changed()

    def _on_link_button_toggled(self):
        row = self.current_selected_row
        if not (0 <= row < len(self.sections)): return
        is_now_linked = self.link_button.isChecked()
        self.sections[row]["is_linked"] = is_now_linked
        self._update_link_button_state(is_now_linked)
        if is_now_linked: self._on_detail_changed()

    def _on_detail_changed(self):
        row = self.current_selected_row
        if not (0 <= row < len(self.sections)): return
        
        name = self.detail_section_name_edit.text()
        start_page = self.detail_start_page_spin.value()
        end_page = self.detail_end_page_spin.value()
        is_linked = self.sections[row].get("is_linked", True)
        
        if is_linked:
            sanitized_name = self._sanitize_filename(name)
            filename = f"{sanitized_name}.pdf" if sanitized_name else "new_section.pdf"
            self.detail_filename_edit.blockSignals(True)
            self.detail_filename_edit.setText(filename)
            self.detail_filename_edit.blockSignals(False)
        else:
            filename = self.detail_filename_edit.text()
            
        self.sections[row] = { "name": name, "start_page": start_page, "end_page": end_page, "filename": filename, "is_linked": is_linked }
        
        self.table_widget.blockSignals(True)
        self.table_widget.item(row, 1).setText(name)
        self.table_widget.item(row, 2).setText(str(start_page))
        self.table_widget.item(row, 3).setText(str(end_page))
        self.table_widget.item(row, 4).setText(filename)
        self.table_widget.blockSignals(False)

    def _add_section(self):
        default_name = self.tr("New Section")
        sanitized_name = self._sanitize_filename(default_name)
        new_section = { "name": default_name, "filename": f"{sanitized_name}.pdf", "start_page": 0, "end_page": 0, "is_linked": True }
        self.sections.append(new_section)
        self._populate_table()
        self.table_widget.selectRow(len(self.sections) - 1)

    def _delete_selected_section(self):
        row = self.table_widget.currentRow()
        if row >= 0:
            del self.sections[row]
            self._populate_table()
            if len(self.sections) > 0:
                self.table_widget.selectRow(min(row, len(self.sections) - 1))
            else:
                self._update_detail_panel_state()
    
    def _clear_all_sections(self):
        if not self.sections: return
        title = self.tr("Confirm Clear")
        content = self.tr("Are you sure you want to remove all sections from the list?")
        m = MessageBox(title, content, self.window())
        if m.exec():
            self.sections.clear()
            self._populate_table()
            InfoBar.success(self.tr("List Cleared"), self.tr("All sections have been removed."), parent=self, duration=3000)

    def _load_selected_template(self):
        template_name = self.template_combo.currentText()
        if not template_name: return
        loaded_sections = self.all_templates.get(template_name, [])
        self.sections = []
        for item in loaded_sections:
            new_item = item.copy()
            new_item.setdefault("is_linked", True)
            self.sections.append(new_item)
        self._populate_table()
        InfoBar.success(self.tr("Template Loaded"), self.tr("Template '{0}' was loaded successfully.").format(template_name), parent=self, duration=3000)

    def _delete_selected_template(self):
        template_name = self.template_combo.currentText()
        if not template_name:
            InfoBar.warning(self.tr("No Template Selected"), self.tr("Please select a template to delete."), parent=self, duration=3000)
            return
        if template_name == "Tel-U":
            InfoBar.error(self.tr("Deletion Denied"), self.tr("The default 'Tel-U' template cannot be deleted."), parent=self, duration=4000)
            return
        title = self.tr("Confirm Deletion")
        content = self.tr("Are you sure you want to delete the template '{0}'? This action cannot be undone.").format(template_name)
        if MessageBox(title, content, self.window()).exec():
            del self.all_templates[template_name]
            _save_all_templates(self.all_templates)
            self._populate_template_combo()
            InfoBar.success(self.tr("Success"), self.tr("Template '{0}' was deleted successfully.").format(template_name), parent=self, duration=3000)

    def _populate_template_combo(self):
        self.template_combo.clear()
        self.template_combo.addItems(sorted(self.all_templates.keys()))
        
    def _save_as_template(self):
        if not self.sections:
            MessageBox(self.tr("List is Empty"), self.tr("There is no data to save as a template."), self.window()).exec()
            return
        name, ok = QInputDialog.getText(self, self.tr("Save Template"), self.tr("Enter a name for the new template:"))
        if ok and name:
            if name in self.all_templates:
                MessageBox(self.tr("Name Exists"), self.tr("A template with the name '{0}' already exists.").format(name), self.window()).exec()
                return
            template_data = [item.copy() for item in self.sections]
            self.all_templates[name] = template_data
            _save_all_templates(self.all_templates)
            self._populate_template_combo()
            self.template_combo.setCurrentText(name)
            InfoBar.success(self.tr("Success"), self.tr("Template '{0}' was saved successfully.").format(name), parent=self, duration=3000)

    def _update_detail_panel_state(self):
        is_enabled = self.current_selected_row >= 0
        self.detail_card.setEnabled(is_enabled)
        if not is_enabled: self._clear_detail_panel()

    def _populate_detail_panel(self, section_data: Dict):
        for w in [self.detail_section_name_edit, self.detail_start_page_spin, self.detail_end_page_spin, self.detail_filename_edit]:
            w.blockSignals(True)
        self.detail_section_name_edit.setText(section_data.get('name', ''))
        self.detail_start_page_spin.setValue(section_data.get('start_page', 0))
        self.detail_end_page_spin.setValue(section_data.get('end_page', 0))
        self.detail_filename_edit.setText(section_data.get('filename', ''))
        for w in [self.detail_section_name_edit, self.detail_start_page_spin, self.detail_end_page_spin, self.detail_filename_edit]:
            w.blockSignals(False)
            
    def _clear_detail_panel(self):
        for w in [self.detail_section_name_edit, self.detail_filename_edit]: w.clear()
        for w in [self.detail_start_page_spin, self.detail_end_page_spin]: w.setValue(0)
            
    def _browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, self.tr("Select PDF File"), "", "PDF Files (*.pdf)")
        if path: self.pdf_path_edit.setText(path)
        
    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, self.tr("Select Output Folder"), "")
        if path: self.output_dir_edit.setText(path)
        
    def _open_pdf(self):
        path = self.pdf_path_edit.text()
        if not path or not os.path.exists(path):
            MessageBox(self.tr("File Not Found"), self.tr("Please select a valid PDF file first."), self.window()).exec()
            return
        webbrowser.open(f'file:///{os.path.realpath(path)}')
        
    def _open_output_dir(self):
        path = self.output_dir_edit.text()
        if not path or not os.path.isdir(path):
            InfoBar.error(self.tr("Folder Not Found"), self.tr("Please select a valid output folder first."), parent=self)
            return
        try:
            if sys.platform == 'win32': os.startfile(os.path.realpath(path))
            elif sys.platform == 'darwin': subprocess.run(['open', os.path.realpath(path)], check=True)
            else: subprocess.run(['xdg-open', os.path.realpath(path)], check=True)
        except Exception as e:
            MessageBox(self.tr("Failed to Open Folder"), self.tr("Could not open the folder:\n{0}").format(e), self.window()).exec()
            
    def _split_pdf(self):
        pdf_path, out_dir = self.pdf_path_edit.text(), self.output_dir_edit.text()
        if not pdf_path or not out_dir or not self.sections:
            MessageBox(self.tr("Incomplete Input"), self.tr("Please provide an input PDF, an output folder, and at least one section."), self.window()).exec()
            return
        try:
            doc = fitz.open(pdf_path)
            os.makedirs(out_dir, exist_ok=True)
            for section in self.sections:
                start, end, filename = section['start_page'], section['end_page'], section['filename']
                if not (0 < start <= end <= len(doc)):
                    raise ValueError(f"Invalid page range for '{section['name']}' ({start}-{end}). Total pages: {len(doc)}")
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=start - 1, to_page=end - 1)
                if len(new_doc) > 0: new_doc.save(os.path.join(out_dir, filename))
                new_doc.close()
            doc.close()
            InfoBar.success(self.tr("Success"), self.tr("{0} PDF sections were successfully split.").format(len(self.sections)), duration=5000, parent=self)
        except Exception as e:
            MessageBox(self.tr("An Error Occurred"), self.tr("Failed to process the PDF: {0}").format(e), self.window()).exec()


class SettingInterface(QWidget):
    themeChanged = pyqtSignal()
    languageChanged = pyqtSignal()
    def __init__(self, settings: QSettings, parent=None):
        super().__init__(parent=parent)
        self.setObjectName("SettingInterface")
        self.settings = settings
        self.main_layout = QVBoxLayout(self)
        self.language_map = {"System Default": "System", "English": "en", "Bahasa Indonesia": "id"}
        self.theme_combo_items_en = ['System', 'Light', 'Dark']
        self.init_layout()
        self.retranslateUi()
        self.connect_signals()
        saved_theme = self.settings.value("theme", "System")
        for i, text in enumerate(self.theme_combo_items_en):
            if text == saved_theme: self.theme_combo.setCurrentIndex(i); break
        saved_lang = self.settings.value("language", "System")
        for text, code in self.language_map.items():
            if code == saved_lang: self.language_combo.setCurrentText(text); break
            
    def init_layout(self):
        self.main_layout.setContentsMargins(30, 10, 30, 20)
        self.main_layout.setSpacing(15)
        self.theme_group_label = SubtitleLabel(self)
        self.theme_label = BodyLabel(self)
        self.theme_combo = ComboBox(self)
        self.language_group_label = SubtitleLabel(self)
        self.language_label = BodyLabel(self)
        self.language_combo = ComboBox(self)
        self.language_combo.addItems(self.language_map.keys())
        self.main_layout.addWidget(self.theme_group_label)
        self.main_layout.addWidget(self.theme_label)
        self.main_layout.addWidget(self.theme_combo)
        self.main_layout.addWidget(self.language_group_label)
        self.main_layout.addWidget(self.language_label)
        self.main_layout.addWidget(self.language_combo)
        self.main_layout.addStretch(1)
        
    def retranslateUi(self):
        self.theme_group_label.setText(self.tr("Appearance"))
        self.theme_label.setText(self.tr("Application Theme"))
        self.theme_combo_items_tr = [self.tr('Follow System'), self.tr('Light'), self.tr('Dark')]
        current_selection = self.theme_combo.currentText()
        self.theme_combo.clear()
        self.theme_combo.addItems(self.theme_combo_items_tr)
        if current_selection in self.theme_combo_items_tr: self.theme_combo.setCurrentText(current_selection)
        self.language_group_label.setText(self.tr("Language"))
        self.language_label.setText(self.tr("Application Language"))
        
    def connect_signals(self):
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self.language_combo.currentTextChanged.connect(self._on_language_changed)
        
    def _on_theme_changed(self, index):
        if index < 0: return
        theme_setting = self.theme_combo_items_en[index]
        self.settings.setValue("theme", theme_setting)
        setTheme(Theme.LIGHT if theme_setting == 'Light' else Theme.DARK if theme_setting == 'Dark' else Theme.AUTO)
        self.themeChanged.emit()
        
    def _on_language_changed(self, text: str):
        lang_code = self.language_map.get(text, "System")
        self.settings.setValue("language", lang_code)
        self.languageChanged.emit()


class MainWindow(FluentWindow):
    def __init__(self, settings: QSettings, translator: QTranslator):
        super().__init__()
        self.settings = settings
        self.translator = translator
        self.main_interface = MainInterface(self)
        self.setting_interface = SettingInterface(self.settings, self)
        self.init_navigation()
        self.init_window()
        self.setting_interface.themeChanged.connect(self.main_interface.retranslateUi)
        self.setting_interface.languageChanged.connect(self.prompt_restart)
        
    def init_window(self):
        self.resize(1100, 850)
        self.setMicaEffectEnabled(sys.platform == 'win32')
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.retranslate_all()
        
    def prompt_restart(self):
        MessageBox(self.tr("Restart Required"), self.tr("Please restart the application for language changes to take full effect."), self.window()).exec()
        
    def retranslate_all(self):
        self.setWindowTitle(self.tr("Super Splitter PDF v15.4"))
        if hasattr(self, 'main_nav_item'): self.main_nav_item.setText(self.tr("PDF Splitter"))
        if hasattr(self, 'settings_nav_item'): self.settings_nav_item.setText(self.tr("Settings"))
        self.main_interface.retranslateUi()
        self.setting_interface.retranslateUi()
        
    def init_navigation(self):
        self.main_nav_item = self.addSubInterface(self.main_interface, FluentIcon.HOME, self.tr("PDF Splitter"))
        self.settings_nav_item = self.addSubInterface(self.setting_interface, FluentIcon.SETTING, self.tr("Settings"), position=NavigationItemPosition.BOTTOM)
        
    def event(self, e: QEvent) -> bool:
        if e.type() == QEvent.ApplicationPaletteChange:
            if self.settings.value("theme", "System") == 'System': 
                setTheme(Theme.AUTO)
                self.main_interface.retranslateUi()
        return super().event(e)


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('mycompany.supersplitter.v1')
    app = QApplication(sys.argv)
    QApplication.setOrganizationName("MyCompany")
    QApplication.setApplicationName("SuperSplitter")
    settings = QSettings()
    translator = QTranslator()
    lang_code = settings.value("language", "System")
    if lang_code == "System": lang_code = QLocale.system().name()
    if not lang_code.startswith('en'):
        if translator.load(f"i18n/{lang_code.split('_')[0]}.qm"):
            app.installTranslator(translator)
    
    saved_theme = settings.value("theme", "System")
    setTheme(Theme.LIGHT if saved_theme == 'Light' else Theme.DARK if saved_theme == 'Dark' else Theme.AUTO)
    
    w = MainWindow(settings, translator)
    w.show()
    sys.exit(app.exec_())