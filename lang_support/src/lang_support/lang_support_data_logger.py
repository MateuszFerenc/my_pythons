from os import listdir, path
from re import match
from datalogger import DataLogger


class LangSupportDL:
    def __init__(self, directory: str = None, ignore_file_error: bool = False, ignore_key_error: bool = False, ignore_dict_error: str = False) -> None:
        assert directory is not None
        assert type(ignore_file_error) is bool
        assert type(ignore_key_error) is bool
        assert type(ignore_dict_error) is bool
        self.dl = DataLogger(f"{path.basename(__file__).split('.')[0]}({directory})", 'logs')
        self.lang_list = []
        self.language = "EN_us"  # default language
        self.dictionary = {}  # language dictionary
        self.path = path.dirname(__file__)
        self.directory = path.join(self.path, directory)

        self.ignore_file_error = ignore_file_error
        self.ignore_key_error = ignore_key_error
        self.ignore_dict_error = ignore_dict_error

        self.get_languages()  # initialise available languages
        self.set_language(self.language)

    def get_languages(self) -> list:
        files = None
        try:
            files = listdir(self.directory)
        except FileNotFoundError:
            self.dl.log(f"Directory {self.directory} does not exist", log_type=3)
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
                self.dl.log(f"{lang} language not found.", log_type=1)
            self.language = lang
        else:
            self.dl.log(f"Languages not indexed or not present in directory!", log_type=2)
            return None
        try:
            file = path.join(self.directory, self.language)
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
            self.dl.log(f"File {file} not found!", log_type=2)
            if not self.ignore_file_error:
                exit(3)
            self.dl.log(f"Exit disabled by ignore_file_error flag.")

    def get_text(self, dict_key: str, *args) -> ( None | str ):
        text = None
        try:
            text = str(self.dictionary[dict_key])  # get text value based on key
        except KeyError:
            if len(self.dictionary):
                self.dl.log(f"{dict_key} key not found in {self.language} language file.", log_type=2)
                if self.ignore_key_error:
                    self.dl.log(f"Exit disabled by ignore_key_error flag.")
                    return dict_key
            else:
                self.dl.log(f"Language: {self.language} not loaded!", log_type=2)
                if self.ignore_dict_error:
                    self.dl.log(f"Exit disabled by ignore_dict_error flag.")
                    return dict_key
            exit(3)
        try:
            text = text.format(*args)   # try to format text with arguments, if any specified
        except IndexError:
            if text.find("{}"):
                self.dl.log(f"Key: {dict_key} can be formatted, but no args were given.", log_type=1)
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


if __name__ == "__main__":
    print("Fatal error! This file should not be run as a standalone.")
    exit(3)
