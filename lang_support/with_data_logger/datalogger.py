from os import mkdir, path
from datetime import datetime

class DataLogger:
    instances = []

    def __init__(self, log_name: str = None, directory: str = None, debug: bool = False):
        assert type(log_name) is str
        assert type(directory) is str
        assert type(debug) is bool
        self.__class__.instances.append(self)
        self.do_log = not debug
        self.path = path.dirname(__file__)
        self.directory = ""
        self.logextension = "log"
        self.log_name = log_name
        self.log_types = ["Info", "Warning", "Error", "Fatal Error"]
        if not debug:
            try:
                self.directory = path.join(self.path, directory)
                mkdir(self.directory)
            except FileExistsError:
                self.directory = path.join(self.path, directory)
            except TypeError:
                pass

        self.log("# Log begin #")

    def end(self):
        self.log("# Log end #")

    def config_logger(self, do_log: bool, new_log: str = None):
        assert type(do_log) is bool
        assert type(new_log) is str
        self.do_log = do_log
        if new_log is not None:
            self.log_name = new_log

    def log(self, text: str, log_type: int = 0, do_print: bool = False):
        assert log_type in range(0, len(self.log_types) + 1)
        assert type(do_print) is bool
        now = datetime.now()
        log = f"{now.day}/{now.month}/{now.year} {now.hour}:{now.minute:02n}:{now.second:02n} {self.log_types[log_type]}: {text}\n"
        if self.do_log and not do_print:
            try:
                with open(path.join(self.directory, f"{self.log_name}.{self.logextension}"), "a") as log_file:
                    log_file.write(log)
            except Exception as e:
                print(e)
                exit()
        else:
            print(f"<{self.log_name.name if self.do_log else self.log_name}> {log}", end="")


# DataLogger usage
# dl = DataLogger(<file_name>, <log_directory>, debug=<True/False>)
# dl.log(<text>, do_print=<True/False>)
#
# for inst in DataLogger.instances:
#       inst.end()

if __name__ == "__main__":
    print("Fatal error! This file could not be ran standalone")
    exit(3)
