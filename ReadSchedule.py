import time
import xlrd
from locale import setlocale, LC_ALL
from requests import get
from dataclasses import dataclass
from os import path, remove

@dataclass
class Shedule:
    group: str
    even_week: bool
    date_time: time.struct_time = None #may be str type if FLAG is false
    lesson_name: str = None #may contain all cell value if FLAG is false
    lesson_type: str = None
    speaker: str = None
    auditorium: str = None
    dparse_suc: bool = True #FLAG Successful parsing;
    parse_suc: bool = True #FLAG Successful parsing;
    #default = None, bcs I need init class without params and add them before 
    def __repr__(self) -> str:
        return f"{self.group}: {self.lesson_name}"
    
class ReadSchedule_Error(Exception):

    def __init__(self, msg: str = 'Error, while reading shedule',  *args: object) -> None:
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
    date_interval: tuple[time.struct_time, time.struct_time]
    __table: xlrd.sheet.Sheet
    __DEFAULT_EXT = '.xls'
    __DEFAULT_NAME = 'Temp_Shedule'
    __EWEEK_RUS = {'нечетная': 1, 'четная': 0}
    __HEAD_COORDS = (1, 0) #row col
    __TIME_COL = 1
    __DATE_COL = 0
    __GROUP_NAME_ROW = 2
    __START_INDEX_ROW = 3
    __even: bool
    __year: str
    __LOCALE = 'Russian'
    __file_path: str #path to temp exel file
    

    def __init__(self, buffer_path: str, bb_link: str) -> None:
        if type(buffer_path) is not str or type(bb_link) is not str: raise TypeError
        setlocale(category = LC_ALL, locale = self.__LOCALE) # For russian date visual
        self.__file = buffer_path
        self.__get_file(bb_link)
        self.__table = xlrd.open_workbook(self.__file_path).sheet_by_index(0)
        self.__head_parser() #additional info
        self.__get_date_interval() # for date sort
    
    def get_all(self) -> list[Shedule]:
        """:RETURNS: list of dataclass - Shedule"""
        self.__get_schedule(self.__START_INDEX_ROW, self.__table.nrows)
    
    def __get_schedule(self, start_index_date: int, end_index_date: int):
        data = list()

        for i_row in range(start_index_date, end_index_date):
            if self.__table.cell_value(i_row, self.__DATE_COL):
                cur_date = self.__table.cell_value(i_row, self.__DATE_COL)
            
            if self.__table.cell_value(i_row, self.__TIME_COL):
                cur_time = self.__table.cell_value(i_row, self.__TIME_COL)
                
            for i_col in range(2, self.__table.ncols):
                if self.__table.cell_value(i_row, i_col).isspace(): continue # skips empty cell
                temp_sh = Shedule(group = self.__table.cell_value(self.__GROUP_NAME_ROW, i_col), even_week = self.__even)
                try:
                    l_info = self.__parse_lesson_info(self.__table.cell_value(i_row, i_col))
                    temp_sh.date_time = self.__str_to_date(cur_date, self.__year, cur_time)
                    temp_sh.lesson_name = l_info[0]
                    temp_sh.lesson_type = l_info[-1]
                    temp_sh.speaker = l_info[1]
                    temp_sh.auditorium = l_info[2]
                    
                except DateParsing_Error:
                    temp_sh.date_time = cur_time
                    temp_sh.dparse_suc = False
                except Parsing_Error:
                    temp_sh.lesson_name = self.__table.cell_value(i_row, i_col)
                    temp_sh.parse_suc = False
                except Exception as err:
                    raise ReadSchedule_Error(err)

                data.append(temp_sh)
        
        return data
    
    def __get_date_interval(self):
        
        self.date_interval = self.__str_to_date(self.__table.cell_value(self.__START_INDEX_ROW, self.__DATE_COL), self.__year),\
        self.__str_to_date(self.__table.cell_value(self.__table.nrows, self.__DATE_COL), self.__year)
    
    @staticmethod
    def __parse_lesson_info(cell_value: str) -> list:
        """:RETURNS: list of parsed info; (lesson_name: str, speaker_info: str, location: str, lesson_type: str)"""
        #EXAMPLE
        #-  Физическая культура и спорт (элективные дисциплины (модули))
        #-  Розенфельд Александр Семёнович, Профессор
        #-  Спорт компл.-10, Практические занятия
        try:
            info = [i[2:] for i in cell_value.split('\n')]
            if len(info) != 3: raise Parsing_Error('Enter data is not valid to parse correct')
            info.extend(info.pop().split(','))
        except Exception as err:
            raise Parsing_Error(f"Can't parse lesson_info part: {err}")
        
        return info
        
    @staticmethod
    def __str_to_date(date: str, year: str, ltime: str = None) -> time.struct_time:
        """Convert str date information to time.struct_time"""
        try:
            day, month = date.split(maxsplit = 2)[:2]
            date =  year + ' ' + day + ' ' + month[:3] # 2024 07 окт
            if ltime:
               date += ' ' + ltime.split('-', maxsplit = 1)[0] # 2024 07 окт 08:30
               return time.strptime(date, '%Y %d %b %H:%M')
            return time.strptime(date, '%Y %d %b')
        
        except Exception as err:
            raise DateParsing_Error(f"Can't parse date part: {err}")
    
    def __head_parser(self):
        """Header parser"""
        try:
            info: list[str] = self.__table.cell_value(*self.__HEAD_COORDS).split()
            self.__even = not self.__EWEEK_RUS[info[-1].lower()]
            self.__year = info[-2].split('/')[0]
        except Exception as err:
            raise Parsing_Error(f"Can't parse header part: {err}")

    def __get_file(self, link):
        """Gets excel file from url request"""
        res = get(link)
        if not res.ok: raise Connect_Error('Failed to connect bb')
        self.__file_path = path.join(self.__file, self.__DEFAULT_NAME + self.__DEFAULT_EXT)
        with open(self.__file_path, 'wb') as f:
            f.write(res.content)

    def __del__(self):
        remove(self.__file_path)





