from openpyxl import load_workbook


class BangKeWriter:
    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb.active

    def write_orders(self, orders, start_row):
        row = start_row

        for order in orders:
            merge = order["merge"]
            start = row
            end = row + merge - 1

            # --- Insert rows if needed ---
            if merge > 1:
                self.ws.insert_rows(start + 1, merge - 1)

            # --- Write MERGED columns: ONLY top-left ---
            for col, val in order["base"].items():
                self.ws[f"{col}{start}"].value = val

            # --- Write conts & cars (NOT merged) ---
            for i in range(merge):
                if i < len(order["conts"]):
                    self.ws[f"I{start + i}"].value = order["conts"][i]
                if i < len(order["cars"]):
                    self.ws[f"J{start + i}"].value = order["cars"][i]

            # --- Write gross (merged column O) ---
            if order["gross"] is not None:
                self.ws[f"O{start}"].value = order["gross"]

            # --- Merge cells AFTER writing ---
            if merge > 1:
                for col in list("ABCDEFGH") + ["K", "L", "M", "N", "O"]:
                    self.ws.merge_cells(f"{col}{start}:{col}{end}")

            row = end + 1

        self.wb.save(self.path)
