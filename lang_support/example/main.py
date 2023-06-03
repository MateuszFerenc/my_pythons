from lang_support import LangSupport
from os.path import join as path_join

# to eaisly run that example change directory to lang_support
# and then type "py -m example.main"

if __name__ == "__main__":
    lang = LangSupport(path_join("example", "Language"))    # path_join and hook to example dir used only to make this example able to run without douplicating lang_support.py file
    
    print(lang.lang_list)   # get list of indexed language files
    print(lang.language)    # get currently selected language
    
    print(lang.get_text("example_key0"))    # get text with currently selected language
    print(lang.get_text("example_key1", "example_formatter"))    # get text with currently selected language and if possible format it (place "example_formater0" into {} brackets)
    
    lang.set_language("PL_pl")      # switch language
    print(lang.language)    # get currently selected language
    
    print(lang.get_text("example_key0"))    # get text with currently selected language
    print(lang.get_text("example_key1", "example_formatter"))    # get text with currently selected language and if possible format it (place "example_formater0" into {} brackets)
    