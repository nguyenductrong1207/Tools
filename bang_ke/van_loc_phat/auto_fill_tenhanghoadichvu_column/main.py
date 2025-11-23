import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QTextEdit,
    QHBoxLayout,
)
from processor import process_bang_ke


class MainUI(QWidget):
    def __init__(self):
        super().__init__()

        self.mapping_path = None
        self.input_path = None

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Tool Auto Điền Cột G Bảng Kê Vạn Lộc Phát")

        layout = QVBoxLayout()

        # chọn mapping
        self.btn_mapping = QPushButton("Chọn file mapping.xlsx")
        self.btn_mapping.clicked.connect(self.pick_mapping)
        layout.addWidget(self.btn_mapping)

        self.lbl_mapping = QLabel("Chưa chọn mapping")
        layout.addWidget(self.lbl_mapping)

        # chọn file bảng kê
        self.btn_file = QPushButton("Chọn file Bảng Kê")
        self.btn_file.clicked.connect(self.pick_file)
        layout.addWidget(self.btn_file)

        self.lbl_file = QLabel("Chưa chọn file")
        layout.addWidget(self.lbl_file)

        # nút xử lý
        self.btn_process = QPushButton("Xử lý")
        self.btn_process.clicked.connect(self.handle_process)
        layout.addWidget(self.btn_process)

        # log
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

        self.setLayout(layout)

    def pick_mapping(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn mapping.xlsx", "", "Excel (*.xlsx)"
        )
        if path:
            self.mapping_path = path
            self.lbl_mapping.setText(path)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Bảng Kê", "", "Excel (*.xlsx)"
        )
        if path:
            self.input_path = path
            self.lbl_file.setText(path)

    def handle_process(self):
        if not self.mapping_path or not self.input_path:
            self.log_box.append(
                "⚠ Vui lòng chọn đầy đủ 2 file mapping.xlsx và file Bảng Kê.\n"
            )
            return

        output_path = self.input_path.replace(".xlsx", "_processed.xlsx")

        log_text = process_bang_ke(self.input_path, self.mapping_path, output_path)

        self.log_box.append(log_text)
        self.log_box.append(f"\nFile đã được lưu tại:\n{output_path}\n")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = MainUI()
    ui.resize(600, 600)
    ui.show()
    sys.exit(app.exec_())
