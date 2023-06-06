from os import listdir, path
from re import match


class LangSupport:
    
    lang_file_template = \
        ".. <Enter lang file description>\n"\
        ".. file format - parameter#text\n"\
        ".. '#' ends parameter placeholder\n"\
        ".. '..' starts single line comment, not interpreted\n"\
        ".. '{}' use for string format"
        
    def __init__(self, directory: str = None, ignore_file_error: bool = False, ignore_key_error: bool = False, ignore_dict_error: str = False) -> None:
        assert directory is not None
        assert type(ignore_file_error) is bool
        assert type(ignore_key_error) is bool
        assert type(ignore_dict_error) is bool
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
        
    def create_lang_file(self, lang: str) -> (True | False):
        if match('^[A-Z]{2}_[a-z]{2}$', str(lang)) is None:
            return False
        
        try:
            with open(path.join(self.directory, lang), "x") as new_file:
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
            print(f"Directory {self.directory} does not exist")
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
                print(f"{lang} language not found.")
            self.language = lang
        else:
            print(f"Languages not indexed or not present in directory!")
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
            print(f"File {file} not found!")
            if not self.ignore_file_error:
                exit(3)
            print(f"Exit disabled by ignore_file_error flag.")

    def get_text(self, dict_key: str, *args) -> ( None | str ):
        text = None
        try:
            text = str(self.dictionary[dict_key])  # get text value based on key
        except KeyError:
            if len(self.dictionary):
                print(f"{dict_key} key not found in {self.language} language file.")
                if self.ignore_key_error:
                    print(f"Exit disabled by ignore_key_error flag.")
                    return dict_key
            else:
                print(f"Language: {self.language} not loaded!")
                if self.ignore_dict_error:
                    print(f"Exit disabled by ignore_dict_error flag.")
                    return dict_key
            exit(3)
        try:
            text = text.format(*args)   # try to format text with arguments, if any specified
        except IndexError:
            if text.find("{}"):
                print(f"Key: {dict_key} can be formatted, but no args were given.")
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
