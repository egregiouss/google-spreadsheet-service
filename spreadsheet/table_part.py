class TablePart:
    def __init__(
        self,
        head_config: dict[str, any],
        head_data: dict[str, any],
        body_config,
        body_data: list[dict[str, any]],
        kids: list["TablePart"],
    ):
        self.body_config = body_config
        self.head_config = head_config
        self.parts_kids = kids
        self.body_data = body_data
        self.head: dict[str, any] = self._make_head(head_data)
        self.body: list[dict[str, any]] = self._make_body()

    def _make_body(self) -> list[dict]:
        body = []
        for record in self.body_data:
            body_line = {}
            for header, handler in self.body_config.items():
                try:
                    body_line[header] = str(handler(record))
                except KeyError:
                    continue
            body.append(body_line)
        for kid in self.parts_kids:
            body.extend(kid.get_data())
        return body

    def _make_head(self, head_data):
        head = {}
        for header, handler in self.head_config.items():
            head[header] = handler(head_data)
        return head

    def get_data(self):
        data = [self.head]
        data.extend(self.body)
        return data
