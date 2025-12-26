import pandas as pd
import re


class MappingService:
    def __init__(self, mapping_path):
        self.mapping_path = mapping_path

    def load_mapping(self, sheet_name):
        df = pd.read_excel(self.mapping_path, sheet_name=sheet_name, header=None)
        mapping = {}
        for _, row in df.iterrows():
            if pd.notna(row[0]) and pd.notna(row[1]):
                mapping[str(row[0]).strip().upper()] = str(row[1]).strip().upper()
        return mapping

    def load_phu_phi_nhap(self):
        df = pd.read_excel(self.mapping_path, sheet_name="Phụ Phí Nhập", header=None)

        fixed_rows = []
        for r in [1, 2]:  # excel row 2–3
            fixed_rows.append({
                "P": df.iloc[r, 1],
                "Q": df.iloc[r, 2],
                "T": df.iloc[r, 3],
            })

        footer_rows = {}
        for r in [19, 20, 21]:  # excel row 20–22
            footer_rows[r + 1] = {
                "P": df.iloc[r, 1],
                "Q": df.iloc[r, 2],
                "T": df.iloc[r, 3],
            }

        by_name = {}
        for i in range(len(df)):
            name = df.iloc[i, 0]
            if pd.notna(name):
                by_name[str(name).strip()] = {
                    "P": df.iloc[i, 1],
                    "Q": df.iloc[i, 2],
                    "T": df.iloc[i, 3],
                }

        return {
            "fixed": fixed_rows,
            "by_name": by_name,
            "footer": footer_rows
        }

    def load_parallel_mapping(self, sheet_name):
        """
        sheet_name: 'nhập' hoặc 'xuất'
        Mapping format:
            Col A: source column (THEO DÕI)
            Col B: target column (BẢNG KÊ)
        """

        df = pd.read_excel(
            self.mapping_path,
            sheet_name=sheet_name,
            header=None
        )

        col_pattern = re.compile(r"^[A-Z]{1,3}$")

        mappings = []

        # 🔥 BỎ QUA DÒNG 1 (HEADER)
        for i in range(1, len(df)):
            src = df.iloc[i, 0]
            dst = df.iloc[i, 1]

            if pd.isna(src) or pd.isna(dst):
                continue

            src = str(src).strip().upper()
            dst = str(dst).strip().upper()

            # chỉ nhận chữ cột hợp lệ
            if not col_pattern.match(src):
                continue
            if not col_pattern.match(dst):
                continue

            mappings.append({
                "src": src,
                "dst": dst
            })

        return mappings
