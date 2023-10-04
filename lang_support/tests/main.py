from lang_support import LangSupportDL
from os.path import join as path_join

# to easily run that example change directory to lang_support
# and then type "py -m tests.main"

if __name__ == "__main__":
    lang = LangSupportDL("Language")    # path_join and hook to example dir used only to make this example able to run without duplicating lang_support.py file
    
    print(lang.lang_list)   # get list of indexed language files
    print(lang.language)    # get currently selected language
    
    print(lang.get_text("example_key0"))    # get text with currently selected language
    print(lang.get_text("example_key1", "example_formatter"))    # get text with currently selected language and if possible format it
    
    lang.set_language("PL_pl")      # switch language
    print(lang.language)    # get currently selected language
    
    print(lang.get_text("example_key0"))    # get text with currently selected language
    print(lang.get_text("example_key1", "example_formatter"))    # get text with currently selected language and if possible format it
    
    print(lang.create_lang_file("EN_gb"))       # create lang file template EN_gb