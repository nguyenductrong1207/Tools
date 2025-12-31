from .order_processor import OrderProcessor
from .phu_phi_nhap_service import PhuPhiService


class FlowRunner:
    def __init__(
            self,
            flow_config,
            mapping_service,
            theo_doi_reader,
            bang_ke_writer,
            theo_doi_wb
    ):
        self.cfg = flow_config
        self.mapping_service = mapping_service
        self.reader = theo_doi_reader
        self.writer = bang_ke_writer
        self.theo_doi_wb = theo_doi_wb

    def run(self, month: int, start_row: int) -> int:
        """
        Chạy 1 flow (NHẬP hoặc XUẤT)
        Return: dòng cuối cùng sau khi ghi
        """

        # 1. LOAD MAPPING
        mapping = self.mapping_service.load_mapping(self.cfg.mapping_sheet)
        phu_phi_mapping = self.mapping_service.load_phu_phi(self.cfg.phu_phi_sheet)

        # 2. READ THEO DÕI
        orders = self.reader.read_nova_by_month(
            sheet_name=self.cfg.theo_doi_sheet,
            month=month
        )

        if not orders:
            return start_row

        # 3. PROCESS
        processor = OrderProcessor(mapping)
        processed_orders = processor.process(orders)

        phu_phi_service = PhuPhiService(phu_phi_mapping)
        theo_doi_ws = self.theo_doi_wb[self.cfg.theo_doi_sheet]

        current_row = start_row

        # 4. WRITE
        for processed_order, order in zip(processed_orders, orders):
            order_start_row = current_row

            row_after_main = self.writer.write_orders(
                [processed_order],
                start_row=current_row
            )

            row_after_phu_phi = phu_phi_service.write_phu_phi(
                order_row_idx=order.row_idx,
                order_data=order.data,
                theo_doi_ws=theo_doi_ws,
                bang_ke_writer=self.writer,
                start_row=order_start_row,
                order_start_row=order_start_row
            )

            self.writer.write_order_total(
                order_start_row,
                row_after_phu_phi - 1
            )

            current_row = row_after_phu_phi

        return current_row
