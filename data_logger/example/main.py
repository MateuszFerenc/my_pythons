from datalogger import DataLogger

if __name__ == "__main__":
    dl = DataLogger("example_log", "any_choosen_dir")   # initialise DataLogger instance log file name "example_log" in directory "any_chosen_dir"
    
    dl.log("example without any arguments")                         # log info
    dl.log("example with arument log_type = 0", log_type=0)         # log info
    dl.log("example with arument log_type = 1", log_type=1)         # log warning
    dl.log("example with arument log_type = 2", log_type=2)         # log error
    dl.log("example with arument log_type = 3", log_type=3)         # log fatal error
    
    
    for inst in DataLogger.instances:
           inst.end()