from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QFileDialog,
    QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QMessageBox
)
from PyQt5.QtCore import Qt
import pandas as pd

from services.mapping_service import MappingService
from services.theo_doi_reader import TheoDoiReader
from services.nova_nhap_processor import NovaNhapProcessor
from services.bang_ke_writer import BangKeWriter

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto Fill Bảng Kê – Nova Nhập")
        self.setMinimumWidth(700)

        self.mapping_path = None
        self.theo_doi_path = None
        self.bang_ke_path = None

        self._init_ui()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addLayout(self._file_picker("Mapping file:", self._pick_mapping))
        layout.addLayout(self._file_picker("File THEO DÕI:", self._pick_theo_doi))

        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("Sheets:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        sheet_layout.addWidget(self.sheet_combo)
        layout.addLayout(sheet_layout)

        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Chọn tháng:"))
        self.month_combo = QComboBox()
        for m in range(1, 13):
            self.month_combo.addItem(f"Tháng {m}", m)
        month_layout.addWidget(self.month_combo)
        layout.addLayout(month_layout)

        layout.addLayout(self._file_picker("File BẢNG KÊ:", self._pick_bang_ke))

        self.process_btn = QPushButton("PROCESSING")
        self.process_btn.setFixedHeight(40)
        self.process_btn.clicked.connect(self._on_process)
        layout.addWidget(self.process_btn, alignment=Qt.AlignCenter)

    def _file_picker(self, label, callback):
        layout = QHBoxLayout()
        layout.addWidget(QLabel(label))
        edit = QLineEdit()
        edit.setReadOnly(True)
        layout.addWidget(edit)
        btn = QPushButton("Chọn")
        btn.clicked.connect(lambda: callback(edit))
        layout.addWidget(btn)
        return layout

    def _pick_mapping(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.mapping_path = path
            edit.setText(path)

    def _pick_theo_doi(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.theo_doi_path = path
            edit.setText(path)
            xls = pd.ExcelFile(path)
            self.sheet_combo.clear()
            for s in xls.sheet_names:
                self.sheet_combo.addItem(s)
            idx = self.sheet_combo.findText("NOVA NHẬP")
            if idx >= 0:
                self.sheet_combo.setCurrentIndex(idx)

    def _pick_bang_ke(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.bang_ke_path = path
            edit.setText(path)

    def _on_process(self):
        if not all([self.mapping_path, self.theo_doi_path, self.bang_ke_path]):
            QMessageBox.warning(self, "Thiếu file", "Vui lòng chọn đủ file")
            return

        try:
            month = self.month_combo.currentData()

            mapping = MappingService(self.mapping_path).load_mapping("Nova Nhập")
            orders = TheoDoiReader(self.theo_doi_path).read_nova_nhap_by_month(month)

            processor = NovaNhapProcessor(mapping)
            processed_orders = processor.process(orders)

            writer = BangKeWriter(self.bang_ke_path)
            writer.write_orders(processed_orders, start_row=13)

            QMessageBox.information(
                self, "Hoàn tất",
                f"Đã xử lý xong NOVA NHẬP – Tháng {month}\nSố đơn: {len(processed_orders)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Lỗi", str(e))
