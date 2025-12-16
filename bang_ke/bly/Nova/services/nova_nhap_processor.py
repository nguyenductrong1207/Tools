class NovaNhapProcessor:
    def __init__(self, mapping):
        self.mapping = mapping

    def _split_lines(self, value):
        if not value:
            return []
        return [v.strip() for v in str(value).splitlines() if v.strip()]

    def process(self, orders):
        results = []

        for order in orders:
            data = order.data

            conts = self._split_lines(data.get("N"))
            cars = self._split_lines(data.get("P"))
            merge = max(len(conts), 1)

            base = {}
            for src, dst in self.mapping.items():
                val = data.get(src)
                if val not in (None, ""):
                    base[dst] = val

            base["H"] = "KHO NOVA"

            k = data.get("K") or 0
            l = data.get("L") or 0
            n = data.get("N") if isinstance(data.get("N"), (int, float)) else None
            gross = round(n / (k + l), 2) if n and (k + l) else None

            results.append({
                "base": base,
                "conts": conts,
                "cars": cars,
                "gross": gross,
                "merge": merge
            })

        return results
