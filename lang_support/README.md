## My python program that allows you to use multiple languages in your programs

Class named **LangSupport** is in lang_support.py file.  

Install via pip:  
`pip install lang_support`

**main.py** file, in tests directory, is an example of use.  

To run example:
```bash
cd lang_support
py -m tests.main
```  
<br>

Import **LangSupport** class:  
`import lang_support.lang_support`
or
`from lang_support.lang_support import LangSupport`
<br>  

Import **LangSupportDL** class:  
`import lang_support.lang_support_data_logger`
or
`from lang_support.lang_support_data_logger import LangSupportDL`


### ***generate_translation.py*** file is a script which read specified language file and generate translation for selected language (deepl package and token needed)

Place DeepL token in ***data_token.json*** file (in directory with ***generate_translation.py*** file) like below:  
```json
{
    "deepl-token": "your-token"
}
```  
To run translation generator, enter following commands:
```bash
cd lang_support
pip install deepl   # if not installed
py -m tests.generate_translation.generate_translation.generate_translation --help
```