from requests_html import HTMLSession as Session
from requests_html import HTMLResponse

class bb_connect_Error(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class ParsingLinks_Error(bb_connect_Error):
    def __init__(self, msg: str = "Error while parsing links", *args: object) -> None:
        super().__init__(msg, *args)

class bb_connect:
    """Gets all url .xls from url directory"""

    __EXT_EXCEL = '.xls'
    __FIND_PARAM = '[id$="xythosFileAnchor"]' 
    
    def __init__(self, bb_link: str) -> None:
        if type(bb_link) is not str: raise TypeError('Param bb_link must be str type')

        self.bb_link = bb_link
    
    def get_urls(self) -> list:
        """Gets all url .xls from url directory"""
        return self.__to_xld_files()
    
    def __to_xld_files(self) -> list:

        urls = list()
        ses = Session()
        get: HTMLResponse = ses.get(self.bb_link)
        #cur_get.html: HTML
        elements = get.html.find(self.__FIND_PARAM)

        while len(elements):
            try:
                link = elements[0].absolute_links.pop() #absolute_links: set; must be only one link in set
                if link.endswith(self.__EXT_EXCEL):
                    urls.append(link)
                else:
                    elements.extend(ses.get(link).html.find(self.__FIND_PARAM)) #merge lists
                elements.pop(0)

            except KeyError:
                pass #dir doesn't has any folders, files
            except Exception as err:
                raise ParsingLinks_Error(err)

        return urls
        
