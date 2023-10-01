## My python program that allows you to use multiple languages in your programs

Class named **LangSupport** is in lang_support.py file.  

Install via pip:  
`pip install lang_support`

**main.py** file, in tests directory, is an example of use.  

To run example:
```bash
cd lang_support
py -m example.main
```  

There is also ***lang_support_data_logger.py*** file, which is the same LangSupport class but print functions were changed to logging functions (to enable logging output to file instead of standard output)  


### ***generate_translation.py*** file is a script which read one language file and generate translation for selected language (deepl package and token needed)

Place Deepl doken in ***data_token.json*** file like below:  
```json
{
    "deepl-token": "your-token"
}
```