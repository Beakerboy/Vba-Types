class VBAString:
    def __init__(self, value = "") -> None:
        self.value = value

    def __str__(self: T) -> str:
        return self.value
