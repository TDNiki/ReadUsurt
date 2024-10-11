import time
import xlrd
from locale import setlocale, LC_ALL
from requests import get
from dataclasses import dataclass
from os import path

@dataclass
class Shedule:
    group: str
    date_time: time.struct_time
    lesson_name: str
    lesson_type: str
    speaker: str
    auditorium: str
    even_week: bool

    def __repr__(self) -> str:
        return f"{self.group}: {self.lesson_name}"
    
class ReadSchedule_Error(Exception):

    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class Connect_Error(ReadSchedule_Error):

    def __init__(self, msg: str = 'Connection failed', *args: object) -> None:
        super().__init__(msg, *args)

class DateParsing_Error(ReadSchedule_Error):
    def __init__(self, msg: str = 'Parse Error', *args: object) -> None:
        super().__init__(msg, *args)


class ReadSchedule:
    """"""
    __table: xlrd.sheet.Sheet
    __DEFAULT_EXT = '.xls'
    __DEFAULT_NAME = 'Temp_Shedule'
    __EWEEK_RUS = {'нечетная': 1, 'четная': 0}
    __HEAD_COORDS = (1, 0) #row col
    __TIME_COL = 1
    __GROUP_NAME_ROW = 2
    __even: bool
    __year: str
    

    def __init__(self, buffer_path: str, bb_link: str) -> None:
        if type(buffer_path) is not str or type(bb_link) is not str: raise TypeError
        self.__file = buffer_path
        #self.__get_file(bb_link) DEBUG
        a = xlrd.open_workbook(self.__DEFAULT_NAME + self.__DEFAULT_EXT)
        self.__table = a.sheet_by_index(0) # self.__file
        self.__head_parser() #additional info
    
    def get_all(self) -> list[Shedule]:
        """:RETURNS: list of dataclass - Shedule"""
        data = list()
        print(self.__table.ncols, self.__table.nrows)
        for i_row in range(3, self.__table.nrows):
            cur_date = 
            for i_col in range(1, self.__table.ncols):
                data.append(Shedule(
                    group = self.__table.cell_value(self.__GROUP_NAME_ROW, i_col),
                    date_time = 0
                ))
            break
    
    @staticmethod
    def __str_to_date(date: str, ltime: str, year: str) -> time.struct_time:
        """"""
        try:
            day, month = date.split(maxsplit = 2)[:2]
            date =  year + ' ' + day + ' ' + month[:3]  + ' ' + ltime.split('-', maxsplit = 1)[0] # 2024 07 окт 08:30
        except Exception as err:
            raise DateParsing_Error(f"Can't parse date part: {err}")

        return time.strptime(date, '%Y %d %b %H:%M')
    
    def __head_parser(self):
        try:
            info: list[str] = self.__table.cell_value(*self.__HEAD_COORDS).split()
            self.__even = not self.__EWEEK_RUS[info[-1].lower()]
            self.__year = info[-2].split('/')[0]
        except Exception as err:
            raise DateParsing_Error(f"Can't parse date part: {err}")

    def __get_file(self, link):
        res = get(link)
        if not res.ok: raise Connect_Error('Failed to connect bb')
        self.__file = path.join(self.__file, self.__DEFAULT_NAME + self.__DEFAULT_EXT)
        with open(self.__file, 'wb') as f:
            f.write(res.content)


setlocale(category = LC_ALL, locale = 'Russian')

a = ReadSchedule('','https://bb.usurt.ru/bbcswebdav/institution/%D0%A0%D0%B0%D1%81%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5/%D0%9E%D1%87%D0%BD%D0%B0%D1%8F%20%D1%84%D0%BE%D1%80%D0%BC%D0%B0%20%D0%BE%D0%B1%D1%83%D1%87%D0%B5%D0%BD%D0%B8%D1%8F/%D0%9D%D0%B5%D1%87%D0%B5%D1%82%D0%BD%D0%B0%D1%8F%20%D0%BD%D0%B5%D0%B4%D0%B5%D0%BB%D1%8F/%D0%9C%D0%B5%D1%85%D0%B0%D0%BD%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D0%B9%20%D1%84%D0%B0%D0%BA%D1%83%D0%BB%D1%8C%D1%82%D0%B5%D1%82/%D0%9C%D0%A4%201%20%D0%BA%D1%83%D1%80%D1%81%20%20%D0%BD%D0%B5%D1%87%D0%B5%D1%82%D0%BD%D0%B0%D1%8F.xls')
a.get_all()



