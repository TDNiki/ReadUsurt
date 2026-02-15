# Licensed under Creative Commons Attribution-NonCommercial 3.0 Unported
# https://creativecommons.org/licenses/by-nc/3.0/
# (c) Tsygankov Nikita, 2025

import tempfile
import xlrd
from requests import get
from dataclasses import dataclass
from os import path
from zoneinfo import ZoneInfo
from datetime import datetime



MONTH_SEP = 9
MONTH_DEC = 12
GROUP_FOR_GROUP_FLAG = 'п/г'

@dataclass
class Schedule:
    faculty: str
    group: str
    even_week: bool
    date_time: datetime = None #may be str type if FLAG is false
    lesson_name: str = None #may contain all cell value if FLAG is false
    lesson_type: str = None
    speaker: str = None
    auditorium: str = None
    parse_suc: bool = True #FLAG Successful parsing;
    def __repr__(self) -> str:
        return f"{self.group}: {self.lesson_name}"
    
class ReadSchedule_Error(Exception):

    def __init__(self, msg: str = 'Error, while reading schedule',  *args: object) -> None:
        super().__init__(msg, *args)

class Connect_Error(ReadSchedule_Error):

    def __init__(self, msg: str = 'Connection failed', *args: object) -> None:
        super().__init__(msg, *args)

class Parsing_Error(ReadSchedule_Error):
    def __init__(self, msg: str = 'Parse Error', *args: object) -> None:
        super().__init__(msg, *args)

class DateParsing_Error(ReadSchedule_Error):
    def __init__(self, msg: str = 'Parse Error', *args: object) -> None:
        super().__init__(msg, *args)

class ReadSchedule:
    """Reads excel shedule file\n\n
    :RETURNS: list of dataclass (Schedule)"""
    __table: xlrd.sheet.Sheet
    __DEFAULT_EXT = '.xls'
    __DEFAULT_NAME = 'Temp_Shedule'
    __EWEEK_RUS = {'нечетная': 1, 'четная': 0}
    __HEAD_COORDS = (1, 0) #row col
    __FACULTY_COORDS = (0, 0)
    __TIME_COL = 1
    __DATE_COL = 0
    __GROUP_NAME_ROW = 2
    __cur_pc_datetime: datetime
    __file_path: str #path to temp exel file
    __months_EN_RU= {
            'янв': 'Jan', 'фев': 'Feb', 'мар': 'Mar', 'апр': 'Apr',
            'май': 'May', 'июн': 'Jun', 'июл': 'Jul', 'авг': 'Aug',
            'сен': 'Sep', 'окт': 'Oct', 'ноя': 'Nov', 'дек': 'Dec'
        }
    corrupted_data = 0

    def __init__(self, bb_link: str) -> None:
        if type(bb_link) is not str or bb_link == '': raise TypeError
        self.bb_link = bb_link
        buffer_path = tempfile.TemporaryDirectory()
        self.__file = buffer_path.name
        self.buffer_path = buffer_path
        self.__get_file(bb_link)
        
        self.__table = xlrd.open_workbook(self.__file_path)
        self.__cur_pc_datetime = datetime.now(ZoneInfo('Asia/Yekaterinburg'))
    
    def get_all(self, date_start_scan: datetime = None) -> list[Schedule]:
        """:RETURNS: list of dataclass - Schedule"""
        if not date_start_scan: date_start_scan = self.__cur_pc_datetime

        data = list()
        sheet_count = self.__table.nsheets
         #additional info
        for sh in range(sheet_count):
            tb = self.__table.sheet_by_index(sh)
            header = self.__head_parser(tb)
            for i_row in range(3, tb.nrows):
                if tb.cell_value(i_row, self.__DATE_COL):
                    cur_date = tb.cell_value(i_row, self.__DATE_COL)
                
                if tb.cell_value(i_row, self.__TIME_COL):
                    cur_time = tb.cell_value(i_row, self.__TIME_COL)
                    
                for i_col in range(2, tb.ncols):
                    if tb.cell_value(i_row, i_col).isspace() or not tb.cell_value(i_row, i_col): continue # skips empty cell
                    
                    
                    try:

                        year = header[2] if MONTH_SEP <= self.__determine_year(cur_date) <= MONTH_DEC else header[3]
                        date_time = self.__str_to_date(cur_date, cur_time, year)

                        if date_time.date() < date_start_scan.date():
                            continue

                        temp_sh = Schedule(group = tb.cell_value(self.__GROUP_NAME_ROW, i_col), even_week = header[0], faculty = header[1])
                        temp_sh.date_time = date_time

                        l_info = self.__parse_lesson_info(tb.cell_value(i_row, i_col))
                        
                        if not l_info:
                            
                            temp_sh.lesson_name = tb.cell_value(i_row, i_col)
                            temp_sh.parse_suc = False
                        else:
                            temp_sh.lesson_name = l_info[0]
                            temp_sh.lesson_type = l_info[-1]

                            if GROUP_FOR_GROUP_FLAG in temp_sh.lesson_type:
                                temp_sh.lesson_type = "Л\б занятия " + temp_sh.lesson_type

                            if l_info[1].isspace():
                                temp_sh.parse_suc = False
                                
                            temp_sh.speaker = l_info[1]
                            temp_sh.auditorium = l_info[2]
                            

                            temp_sh.speaker = temp_sh.speaker.split(', ')[0]
                            


                        data.append(temp_sh)


                    except DateParsing_Error as er:
                        self.corrupted_data += 1
                        continue
                    except Exception as err:
                        raise ReadSchedule_Error(err)

                
        
        return data
    
    
    @staticmethod
    def __parse_lesson_info(cell_value: str) -> list:
        """:RETURNS: list of parsed info; (lesson_name: str, speaker_info: str, location: str, lesson_type: str)"""
        #EXAMPLE
        #-  Физическая культура и спорт (элективные дисциплины (модули))
        #-  Розенфельд Александр Семёнович, Профессор
        #-  Спорт компл.-10, Практические занятия
        try:

            info = []
            for i in cell_value.split('\n'):
                limit = len(i)
                while i[0].isalpha() == False and limit > 1:
                    i = i[1:]
                    limit -= 1
                
                info.append(i)

            if len(info) != 3: return None
            info.extend(info.pop().split(','))
        except Exception as err:
            raise Parsing_Error(f"Can't parse lesson_info part: {err}")
        
        return info
    


    def __determine_year(self, parsed_date: str) -> int:
        try:
            month_rus = parsed_date.split(maxsplit=2)[1][:3].lower()
            month_num = list(self.__months_EN_RU.keys()).index(month_rus) + 1

            return month_num
        except:
            raise DateParsing_Error("Can't determine year")


    def __str_to_date(self, date: str, ltime: str, year: str) -> datetime:
        """Convert str date information to datetime"""



        try:
            day, month = date.split(maxsplit = 2)[:2]
            date =  year + ' ' + day + ' ' + self.__months_EN_RU[month[:3].lower()]  + ' ' + ltime.split('-', maxsplit = 1)[0] # 2024 07 окт 08:30
        except Exception as err:
            raise DateParsing_Error(f"Can't parse date part: {err}")

        return datetime.strptime(date, '%Y %d %b %H:%M').replace(tzinfo = ZoneInfo('Asia/Yekaterinburg'))
    
    def __head_parser(self, table) -> tuple[int, str, int, int]:
        """Header parser"""
        try:
            header = table.cell_value(*self.__HEAD_COORDS).split()
            even = header[-1]
            even = not self.__EWEEK_RUS[even.lower()]
            year_start, year_end = header[-2].split('/', 1)

            return even, table.cell_value(*self.__FACULTY_COORDS), year_start, year_end
        except Exception as err:
            raise Parsing_Error(f"Can't parse header part: {err}")

    def __get_file(self, link):
        """Gets excel file from url request"""
        res = get(link)
        if not res.ok: raise Connect_Error('Failed to connect bb')
        self.__file_path = path.join(self.__file, self.__DEFAULT_NAME + self.__DEFAULT_EXT)
        with open(self.__file_path, 'wb') as f:
            f.write(res.content)

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.buffer_path.cleanup()








