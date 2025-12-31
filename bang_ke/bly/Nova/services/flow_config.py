from dataclasses import dataclass


@dataclass
class FlowConfig:
    name: str  # "NHẬP" | "XUẤT"
    theo_doi_sheet: str  # "NOVA NHẬP" | "NOVA XUẤT"
    mapping_sheet: str  # "Nova Nhập" | "Nova Xuất"
    phu_phi_sheet: str  # "Phụ Phí Nhập" | "Phụ Phí Xuất"
    parallel_sheet: str  # "nhập" | "xuất"
