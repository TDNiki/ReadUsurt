import tempfile
import xlrd
from requests import get
from dataclasses import dataclass
from os import path
from zoneinfo import ZoneInfo
from datetime import datetime


CHANGE_YEAR = 11
GROUP_FOR_GROUP_FLAG = 'п/г'

@dataclass
class Schedule:
    group: str
    even_week: bool
    date_time: datetime = None #may be str type if FLAG is false
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
    __TIME_COL = 1
    __DATE_COL = 0
    __GROUP_NAME_ROW = 2
    __year: str
    __file_path: str #path to temp exel file
    __cur_pc_datetime: datetime
    __months_EN_RU= {
            'янв': 'Jan', 'фев': 'Feb', 'мар': 'Mar', 'апр': 'Apr',
            'май': 'May', 'июн': 'Jun', 'июл': 'Jul', 'авг': 'Aug',
            'сен': 'Sep', 'окт': 'Oct', 'ноя': 'Nov', 'дек': 'Dec'
        }
    corrupted_data = 0

    def __init__(self, bb_link: str) -> None:
        if type(bb_link) is not str or bb_link == '': raise TypeError
        buffer_path = tempfile.TemporaryDirectory()
        self.__file = buffer_path.name
        self.buffer_path = buffer_path
        self.__get_file(bb_link)
        self.__table = xlrd.open_workbook(self.__file_path)
        self.__cur_pc_datetime = datetime.now(ZoneInfo('Asia/Yekaterinburg'))
        self.__year = str(self.__cur_pc_datetime.year)
    
    def get_all(self, date_start_scan: datetime = None) -> list[Schedule]:
        """:RETURNS: list of dataclass - Schedule"""
        if date_start_scan is not datetime: date_start_scan = self.__cur_pc_datetime

        data = list()
        sheet_count = self.__table.nsheets
         #additional info
        for sh in range(sheet_count):
            tb = self.__table.sheet_by_index(sh)
            even = self.__head_parser(tb)
            for i_row in range(3, tb.nrows):
                if tb.cell_value(i_row, self.__DATE_COL):
                    cur_date = tb.cell_value(i_row, self.__DATE_COL)
                
                if tb.cell_value(i_row, self.__TIME_COL):
                    cur_time = tb.cell_value(i_row, self.__TIME_COL)
                    
                for i_col in range(2, tb.ncols):
                    if tb.cell_value(i_row, i_col).isspace(): continue # skips empty cell
                    
                    try:
                        date_time = self.__str_to_date(cur_date, cur_time, self.__year)

                        

                        if date_time.month - self.__cur_pc_datetime.month == CHANGE_YEAR:
                            date_time.replace(year= date_time.year - 1)

                        if date_time.date() < date_start_scan.date():
                            continue

                        temp_sh = Schedule(group = tb.cell_value(self.__GROUP_NAME_ROW, i_col), even_week = even)
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

                            temp_sh.speaker = l_info[1]
                            temp_sh.auditorium = l_info[2]
                            
                            if temp_sh.speaker[0] == ' ': temp_sh.speaker = temp_sh.speaker[1:]

                            temp_sh.speaker = temp_sh.speaker.split(', ')[0]


                        data.append(temp_sh)


                    except DateParsing_Error:
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

            info = [i.replace("- ", "", 1) for i in cell_value.split('\n')]
            if len(info) != 3: return None
            info.extend(info.pop().split(','))
        except Exception as err:
            raise Parsing_Error(f"Can't parse lesson_info part: {err}")
        
        return info
        


    def __str_to_date(self, date: str, ltime: str, year: str) -> datetime:
        """Convert str date information to time.struct_time"""



        try:
            day, month = date.split(maxsplit = 2)[:2]
            date =  year + ' ' + day + ' ' + self.__months_EN_RU[month[:3].lower()]  + ' ' + ltime.split('-', maxsplit = 1)[0] # 2024 07 окт 08:30
        except Exception as err:
            raise DateParsing_Error(f"Can't parse date part: {err}")

        return datetime.strptime(date, '%Y %d %b %H:%M').replace(tzinfo = ZoneInfo('Asia/Yekaterinburg'))
    
    def __head_parser(self, table):
        """Header parser"""
        try:
            info: list[str] = table.cell_value(*self.__HEAD_COORDS).split()
            return not self.__EWEEK_RUS[info[-1].lower()]
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





