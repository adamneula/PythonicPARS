class bondObjects:
    def __init__(self, cusip: str, face_value: float):
        self.cusip = cusip
        self.face_value = face_value

    def __hash__(self):
        return hash(self.cusip)

    def __eq__(self, other):
        if isinstance(other, bondObjects):
            return self.cusip == other.cusip
        return False

    