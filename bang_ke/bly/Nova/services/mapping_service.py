import pandas as pd


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
