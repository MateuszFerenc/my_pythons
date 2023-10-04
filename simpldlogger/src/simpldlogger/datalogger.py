from os import mkdir
from datetime import datetime
from os.path import join, dirname, abspath
from sys import argv

class SimpleDataLogger:
    
    instances = []

    __slot__ = ["__init__", "end", "config_logger", "log"]

    def __init__(self, log_name: str = None, directory: str = None, debug: bool = False) -> None:
        assert type(log_name) is str
        assert type(directory) is str
        assert type(debug) is bool

        self.__class__.instances.append(self)
        self.do_log = not debug
        self.path = dirname(abspath(argv[0]))
        self.directory = ""
        self.logextension = "log"
        self.log_name = log_name
        self.log_types = ["Info", "Warning", "Error", "Fatal Error"]

        if not debug:
            try:
                self.directory = join(self.path, directory)
                mkdir(self.directory)
            except FileExistsError:
                self.directory = join(self.path, directory)
            except TypeError:
                pass

        self.log("# Log begin #")

    def end(self) -> None:
        self.log("# Log end #")

    def config_logger(self, do_log: bool = True, new_log: str = None) -> None:
        assert type(do_log) is bool
        assert type(new_log) is str

        self.do_log = do_log
        if new_log is not None:
            self.log_name = new_log

    def log(self, text: str, log_type: int = 0, do_print: bool = False, do_ignore_newline: bool = True, do_ignore_tab: bool = True) -> None:
        assert type(log_type) is int
        assert 0 <= log_type < len(self.log_types)
        assert type(do_print) is bool
        assert type(do_ignore_newline) is bool
        assert type(do_ignore_tab) is bool

        now = datetime.now()
        text = text.replace('\n', r'\n') if do_ignore_newline else text
        text = text.replace('\t', r'\t') if do_ignore_tab else text
        log = f"{now.day}/{now.month}/{now.year} {now.hour}:{now.minute:02n}:{now.second:02n} {self.log_types[log_type]}: {text}\n"

        if self.do_log and not do_print:
            try:
                with open(join(self.directory, f"{self.log_name}.{self.logextension}"), "a") as log_file:
                    log_file.write(log)
            except Exception as e:
                print(e)
                exit()
        else:
            print(f"<{self.log_name}> {log}", end="")

if __name__ == "__main__":
    print("Fatal error! This file should not be run as a standalone.")
    exit(3)
