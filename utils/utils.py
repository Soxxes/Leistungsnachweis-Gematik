import logging
import sys


# logging decorator
def add_logging(func):
    def create_logs(*args, **kwargs):
        try:
            res = func(*args, **kwargs)
            logging.info(f"Successfully ran function {func.__name__} with arguments: {args, kwargs}")
            return res
        except Exception as e:
            logging.info(f"Error in '{func.__name__}' function. Terminated with error:")
            logging.info(f"Function called with arguments: {args, kwargs}")
            logging.error(f"{e}")
            print("[ERROR] Unexpected behavior: Please send the log file to the developers.")
            input("Press any key to exit ...")
            sys.exit()
    return create_logs
