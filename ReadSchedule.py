from urllib import request
from dataclasses import dataclass
import time







@dataclass
class Shedule:
    group: str
    date_time: time
    lesson_name: str
    lesson_type: str
    speaker: str
    auditorium: str
    even_week: bool

    def __repr__(self) -> str:
        return f"{self.group}: {self.lesson_name}"


