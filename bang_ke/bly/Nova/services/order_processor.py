class OrderProcessor:
    def __init__(self, mapping: dict):
        self.mapping = mapping

    def process(self, orders: list) -> list:
        """
        Input : list OrderModel
        Output: list processed_order dict
        """
        processed = []

        for order in orders:
            base = {}
            conts = []
            cars = []

            for src_col, dst_col in self.mapping.items():
                val = order.data.get(src_col)
                if val not in (None, ""):
                    base[dst_col] = val

            # cont / xe (đã xử lý trước đó)
            cont_raw = order.data.get("N")
            if cont_raw:
                conts = [x.strip() for x in str(cont_raw).splitlines() if x.strip()]

            car_raw = order.data.get("P")
            if car_raw:
                cars = [x.strip() for x in str(car_raw).splitlines() if x.strip()]

            merge = max(len(conts), len(cars), 1)

            processed.append({
                "base": base,
                "conts": conts,
                "cars": cars,
                "merge": merge
            })

        return processed
