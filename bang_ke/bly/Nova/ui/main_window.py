from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QFileDialog,
    QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QMessageBox
)
from PyQt5.QtCore import Qt
import pandas as pd
from openpyxl import load_workbook
import traceback

from ..services import (MappingService, TheoDoiReader, NovaNhapProcessor, BangKeWriter, PhuPhiService,
                        ParallelSheetWriter)


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

        DEV_MAPPING_PATH = r"D:\GitHub\Tools\bang_ke\bly\Nova\mapping.xlsx"
        DEV_THEO_DOI_PATH = (r"D:\GitHub\Tools\bang_ke\bly\THEO DOI TAM UNG SONGLONG-MPF-CHINH NGUYEN 21.06.25.xlsx")
        DEV_BANG_KE_PATH = (r"D:\GitHub\Tools\bang_ke\bly\BẢNG KÊ CÔNG NỢ-STONE - Sample - Copy.xlsx")

        self.mapping_edit = QLineEdit()
        self.mapping_edit.setReadOnly(True)
        layout.addLayout(self._file_picker(
            "Mapping file:", self._pick_mapping, self.mapping_edit
        ))
        # ===== DEV ONLY – auto mapping =====
        if DEV_MAPPING_PATH:
            self.mapping_path = DEV_MAPPING_PATH
            self.mapping_edit.setText(DEV_MAPPING_PATH)

        self.theo_doi_edit = QLineEdit()
        self.theo_doi_edit.setReadOnly(True)
        layout.addLayout(self._file_picker(
            "File THEO DÕI:", self._pick_theo_doi, self.theo_doi_edit
        ))
        # ===== DEV ONLY =====
        if DEV_THEO_DOI_PATH:
            self.theo_doi_path = DEV_THEO_DOI_PATH
            self.theo_doi_edit.setText(DEV_THEO_DOI_PATH)

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

        self.bang_ke_edit = QLineEdit()
        self.bang_ke_edit.setReadOnly(True)
        layout.addLayout(self._file_picker(
            "File BẢNG KÊ:", self._pick_bang_ke, self.bang_ke_edit
        ))
        # ===== DEV ONLY =====
        if DEV_BANG_KE_PATH:
            self.bang_ke_path = DEV_BANG_KE_PATH
            self.bang_ke_edit.setText(DEV_BANG_KE_PATH)

        self.process_btn = QPushButton("PROCESSING")
        self.process_btn.setFixedHeight(40)
        self.process_btn.clicked.connect(self._on_process)
        layout.addWidget(self.process_btn, alignment=Qt.AlignCenter)

    def _file_picker(self, label, callback, edit=None):
        layout = QHBoxLayout()
        layout.addWidget(QLabel(label))

        if edit is None:
            edit = QLineEdit()
            edit.setReadOnly(True)

        layout.addWidget(edit)

        btn = QPushButton("Chọn")
        btn.clicked.connect(lambda: callback(edit))
        layout.addWidget(btn)
        return layout

    def _pick_mapping(self, edit):
        path, _ = QFileDialog.getOpenFileName(
            self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.mapping_path = path
            edit.setText(path)

    def _pick_theo_doi(self, edit):
        path, _ = QFileDialog.getOpenFileName(
            self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.theo_doi_path = path
            edit.setText(path)
            xls = pd.ExcelFile(path)
            self.sheet_combo.clear()
            for s in xls.sheet_names:
                self.sheet_combo.addItem(s)
            xls.close()
            idx = self.sheet_combo.findText("NOVA NHẬP")
            if idx >= 0:
                self.sheet_combo.setCurrentIndex(idx)

    def _pick_bang_ke(self, edit):
        path, _ = QFileDialog.getOpenFileName(
            self, "", "", "Excel Files (*.xlsx)")
        if path:
            self.bang_ke_path = path
            edit.setText(path)

    def _on_process(self):
        if not all([self.mapping_path, self.theo_doi_path, self.bang_ke_path]):
            QMessageBox.warning(self, "Thiếu file", "Vui lòng chọn đủ file")
            return

        try:
            month = self.month_combo.currentData()

            # ===============================
            # 1. LOAD MAPPING
            # ===============================
            mapping_service = MappingService(self.mapping_path)

            nova_nhap_mapping = mapping_service.load_mapping("Nova Nhập")
            nova_xuat_mapping = mapping_service.load_mapping("Nova Xuất")

            phu_phi_nhap_mapping = mapping_service.load_phu_phi("Phụ Phí Nhập")
            phu_phi_xuat_mapping = mapping_service.load_phu_phi("Phụ Phí Xuất")

            parallel_nhap_mapping = mapping_service.load_parallel_mapping("nhập")
            parallel_xuat_mapping = mapping_service.load_parallel_mapping("xuất")

            # ===============================
            # 2. READ THEO DÕI
            # ===============================
            reader = TheoDoiReader(self.theo_doi_path)

            orders_nhap = reader.read_nova_nhap_by_month(month)
            orders_xuat = reader.read_nova_xuat_by_month(month)

            if not orders_nhap or not orders_xuat:
                QMessageBox.information(self, "Không có dữ liệu", "Không có đơn nào")
                return

            # ===============================
            # 3. INIT SERVICES
            # ===============================
            # MỞ BẢNG KÊ 1 LẦN DUY NHẤT
            writer = BangKeWriter(
                self.bang_ke_path,
                sheet_name="NOVA STONE-T.2026"
            )

            # MỞ THEO DÕI 1 LẦN (READ-ONLY)
            theo_doi_wb = load_workbook(self.theo_doi_path, data_only=True)

            # ===============================
            # === LUỒNG NHẬP ===
            # ===============================
            processor = NovaNhapProcessor(nova_nhap_mapping)
            phu_phi_service = PhuPhiService(phu_phi_nhap_mapping)

            processed_nhap = processor.process(orders_nhap)

            current_row = 13

            for processed_order, order in zip(processed_nhap, orders_nhap):
                order_start_row = current_row

                row_after_nova = writer.write_orders(
                    [processed_order],
                    start_row=current_row
                )

                row_after_phu_phi = phu_phi_service.write_phu_phi(
                    order_row_idx=order.row_idx,
                    order_data=order.data,
                    theo_doi_ws=theo_doi_wb["NOVA NHẬP"],
                    bang_ke_writer=writer,
                    start_row=order_start_row,
                    order_start_row=order_start_row
                )

                writer.write_order_total(
                    order_start_row,
                    row_after_phu_phi - 1
                )

                current_row = row_after_phu_phi

            # ===============================
            # === LUỒNG XUẤT ===
            # ===============================
            processor = NovaNhapProcessor(nova_xuat_mapping)
            phu_phi_service = PhuPhiService(phu_phi_xuat_mapping)

            processed_xuat = processor.process(orders_xuat)

            for processed_order, order in zip(processed_xuat, orders_xuat):
                order_start_row = current_row

                row_after_nova = writer.write_orders(
                    [processed_order],
                    start_row=current_row
                )

                row_after_phu_phi = phu_phi_service.write_phu_phi(
                    order_row_idx=order.row_idx,
                    order_data=order.data,
                    theo_doi_ws=theo_doi_wb["NOVA XUẤT"],
                    bang_ke_writer=writer,
                    start_row=order_start_row,
                    order_start_row=order_start_row
                )

                writer.write_order_total(
                    order_start_row,
                    row_after_phu_phi - 1
                )

                current_row = row_after_phu_phi

            # ===============================
            # PARALLEL SHEETS
            # ===============================
            ParallelSheetWriter(
                writer.wb["nhập"]
            ).write_parallel(
                theo_doi_ws=theo_doi_wb["NOVA NHẬP"],
                orders=orders_nhap,
                mapping=parallel_nhap_mapping,
                start_row=6
            )

            ParallelSheetWriter(
                writer.wb["xuất"]
            ).write_parallel(
                theo_doi_ws=theo_doi_wb["NOVA XUẤT"],
                orders=orders_xuat,
                mapping=parallel_xuat_mapping,
                start_row=6
            )

            # ===============================
            # SAVE 1 LẦN DUY NHẤT
            # ===============================
            writer.wb.save(writer.path)

            QMessageBox.information(
                self,
                "Hoàn tất",
                f"Đã xử lý xong – Tháng {month}\n"
                f"\nNOVA NHẬP – Số đơn: {len(orders_nhap)}\n"
                f"\nNOVA XUẤT – Số đơn: {len(orders_xuat)}"
            )

        except Exception as e:
            print("===== FULL TRACEBACK =====")
            traceback.print_exc()
            print("===== END TRACEBACK =====")

            QMessageBox.critical(self, "Lỗi", str(e))
