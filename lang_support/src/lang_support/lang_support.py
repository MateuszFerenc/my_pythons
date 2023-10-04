from os import listdir
from os.path import abspath, join, basename, dirname
from re import match
from importlib import import_module
from sys import argv

dl_instance = None

class LangSupport:
    
    lang_file_template = \
        ".. <Enter lang file description>\n"\
        ".. file format - parameter#text\n"\
        ".. '#' ends parameter placeholder\n"\
        ".. '..' starts single line comment, not interpreted\n"\
        ".. '{}' use for string format"
    
    __slot__ = ["__init__", "create_lang_file", "get_languages", "set_language", "get_text", "ext_text"]
        
    def __init__(self, directory: str = None, ignore_file_error: bool = False, ignore_key_error: bool = False, ignore_dict_error: str = False) -> None:
        assert directory is not None
        assert type(ignore_file_error) is bool
        assert type(ignore_key_error) is bool
        assert type(ignore_dict_error) is bool
        self.lang_list = []
        self.language = "EN_us"  # default language
        self.dictionary = {}  # language dictionary
        self.path = dirname(abspath(argv[0]))
        self.directory = join(self.path, directory)

        self.ignore_file_error = ignore_file_error
        self.ignore_key_error = ignore_key_error
        self.ignore_dict_error = ignore_dict_error
        
        self.get_languages()  # initialise available languages
        self.set_language(self.language)
        
    def create_lang_file(self, lang: str) -> (True | False):
        if match('^[A-Z]{2}_[a-z]{2}$', str(lang)) is None:
            return False
        
        try:
            with open(join(self.directory, lang), "x") as new_file:
                for line in LangSupport.lang_file_template.split('\n'):
                    new_file.write(line + '\n')
        except FileExistsError:
            return False
        else:
            return True

    def get_languages(self) -> list:
        files = None
        try:
            files = listdir(self.directory)
        except FileNotFoundError:
            do_print(f"Directory {self.directory} does not exist", log_type=3)
            exit(3)
        self.lang_list = []
        for Lang in files:
            if match('^[A-Z]{2}_[a-z]{2}$', str(Lang)) is not None:  # add only files in name format 'XX_yy'
                self.lang_list.append(Lang)
        return self.lang_list

    def set_language(self, lang: str, dump: bool = False) -> ( None | dict ):
        assert type(lang) is str
        assert type(dump) is bool
        if len(self.lang_list):
            if lang not in self.lang_list:
                do_print(f"{lang} language not found.", log_type=1)
            self.language = lang
        else:
            do_print(f"Languages not indexed or not present in directory!", log_type=2)
            return None
        try:
            file = join(self.directory, self.language)
            with open(file, "r", encoding="utf-8") as lang_data:
                self.dictionary = {}
                for line in lang_data:
                    if not line.startswith(".."):  # double-dotted lines are not interpreted comment lines
                        split_line = line.strip().split("#", 1)  # separate parameters from values
                        # and remove escaping
                        self.dictionary[split_line[0]] = split_line[1]  # enter values by keys into dictionary
                lang_data.close()
                if dump:
                    return self.dictionary
        except FileNotFoundError:
            do_print(f"File {file} not found!", log_type=2)
            if not self.ignore_file_error:
                exit(3)
            do_print(f"Exit disabled by ignore_file_error flag.")

    def get_text(self, dict_key: str, *args) -> ( None | str ):
        text = None
        try:
            text = str(self.dictionary[dict_key])  # get text value based on key
        except KeyError:
            if len(self.dictionary):
                do_print(f"{dict_key} key not found in {self.language} language file.", log_type=2)
                if self.ignore_key_error:
                    do_print(f"Exit disabled by ignore_key_error flag.")
                    return dict_key
            else:
                do_print(f"Language: {self.language} not loaded!")
                if self.ignore_dict_error:
                    do_print(f"Exit disabled by ignore_dict_error flag.")
                    return dict_key
            exit(3)
        try:
            text = text.format(*args)   # try to format text with arguments, if any specified
        except IndexError:
            if text.find("{}"):
                do_print(f"Key: {dict_key} can be formatted, but no args were given.", log_type=1)
        return text.replace(r'\n', '\n').replace(r'\t', '\t')
    
    @staticmethod
    def ext_text(dict: dict, dict_key: str, *args) -> ( None | str ):
        text = None
        try:
            text = str(dict[dict_key])  # get text value based on key
        except KeyError:
            return dict_key
        try:
            text = text.format(*args)   # try to format text with arguments, if any specified
        except IndexError:
            pass
        return text.replace(r'\n', '\n').replace(r'\t', '\t')
    

class LangSupportDL(LangSupport):
    
    __slot__ = ["__init__"]

    def __init__(self, directory: str = None, ignore_file_error: bool = False, ignore_key_error: bool = False, ignore_dict_error: str = False):
        assert directory is not None

        global dl_instance 
        dl_module = import_module("simpldlogger.datalogger")
        dl_instance = dl_module.SimpleDataLogger(f"{basename(__file__).split('.')[0]}({directory})", 'logs')

        super().__init__(directory=directory, ignore_file_error=ignore_file_error, ignore_key_error=ignore_key_error, ignore_dict_error=ignore_dict_error)


def do_print(text: str, log_type: int = 0) -> None:
    assert type(text) is str
    assert type(log_type) is int

    if dl_instance is None:
        print(text)
    else:
        dl_instance.log(text, log_type)


if __name__ == "__main__":
    print("Fatal error! This file should not be run as a standalone.")
    exit(3)
