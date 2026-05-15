class VBADate:
    def __init__(self, dt: datetime):
        self.dt = dt
    
    def to_datetime(self) -> datetime:
        return self.dt
    
    @classmethod
    def from_datetime(cls, dt: datetime) -> 'VBADate':
        return cls(dt)
