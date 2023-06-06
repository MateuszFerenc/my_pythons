#-*- coding: utf-8 -*-

import deepl
import json
import argparse

if __name__ == "__main__":
    no_token = True
    try:
        with open("data_token.json", "r") as data_token:
            try:
                token_c = json.load(data_token)
            except json.JSONDecodeError:
                print("data_token.json is corrupted")
            else:
                no_token = False
                print("Token from data_token.json, successfully read")
    except FileNotFoundError:
        print("data_token.json, not found")

    if no_token or 'deepl-token' not in token_c:
        no_token = True
        token_c['deepl-token'] = input('No deepl-token in data_token.json, enter it please\n>')
        
    translator = deepl.Translator(token_c['deepl-token']) 
    langs = []
    for language in translator.get_target_languages():
        langs.append(language.code)
        
    parser = argparse.ArgumentParser(description="Translator Copyright (2023) Mateusz Ferenc")
    parser.add_argument("-i", "--input", type=str, help="text file to translate", required=True)
    parser.add_argument("-o", "--output", type=str, help="name of translated file", required=True)
    parser.add_argument("-l", "--language", type=str, help="Target language", required=True, choices=langs)
    parser.add_argument("-c", "--preserve_formatting", help="Perserve file structure", action="store_true")
    args = parser.parse_args()
    
    print()
    
    try:
        with open(args.input, 'r', encoding="utf-8") as to_translate:
            try:
                with open(args.output, 'x', encoding="utf-8") as translated:
                    for i, line in enumerate(to_translate):
                        text = f"Translating {i} line of {args.input} file."
                        print(end='\x1b[2K')
                        print(text, end="\r")
                        if not line.startswith(".."):
                            key, text = line.strip().split("#", 1)
                        else:
                            key = None
                            text = line
                        result = translator.translate_text(text, target_lang=args.language, preserve_formatting=args.preserve_formatting).text
                        if key is not None:
                            translated.write(f"{key}#{result}\n")
                        else:
                            translated.write(result)
            except FileExistsError:
                raise FileExistsError
    except Exception as e:
        print(e)
    
    if no_token:    
        with open("data_token.json", "w") as data_token_write:
            json.dump(token_c, data_token_write, indent=4)