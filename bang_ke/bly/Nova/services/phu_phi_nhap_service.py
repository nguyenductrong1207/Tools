from openpyxl.utils import get_column_letter


class PhuPhiNhapService:
    def __init__(self, phu_phi_mapping: dict):
        self.mapping = phu_phi_mapping

    def _to_int(self, val):
        try:
            return int(float(val))
        except:
            return 0

    def write_phu_phi(
        self,
        order_row_idx: int,
        order_data: dict,
        theo_doi_ws,
        bang_ke_writer,
        start_row: int
    ):
        """
        Returns: last written row index
        """
        row = start_row

        # =========================
        # A. FIXED ROWS (2–3)
        # =========================
        for item in self.mapping["fixed"]:
            bang_ke_writer.write_phu_phi_row(
                row,
                p=item["P"],
                q=item["Q"],
                t_formula=item["T"]
            )
            row += 1

        # =========================
        # B. AJ–AT FLAGS
        # =========================
        header_row = 5

        for col_idx in range(36, 46):  # AJ=36, AT=46
            col_letter = get_column_letter(col_idx)

            flag_val = self._to_int(
                theo_doi_ws[f"{col_letter}{order_row_idx}"].value
            )

            if flag_val != 1:
                continue

            name = theo_doi_ws[f"{col_letter}{header_row}"].value
            if not name:
                continue

            name = str(name).strip()
            mapping_row = self.mapping["by_name"].get(name)
            if not mapping_row:
                continue

            bang_ke_writer.write_phu_phi_row(
                row,
                p=mapping_row["P"],
                q=mapping_row["Q"],
                t_formula=mapping_row["T"]
            )
            row += 1

        # =========================
        # C. FOOTER ROWS (20–22)
        # =========================
        for excel_row, item in self.mapping["footer"].items():
            repeat = 1

            # row 21 → repeat by column G
            if excel_row == 21:
                repeat = self._to_int(order_data.get("G"))

            for _ in range(repeat):
                bang_ke_writer.write_phu_phi_row(
                    row,
                    p=item["P"],
                    q=item["Q"],
                    t_formula=item["T"]
                )
                row += 1

        return row
