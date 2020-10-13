
from datetime import datetime

class LogWriter:
    def __init__(self, log_file, print_timestamp=True, mute=False):
        self.log_file = log_file
        self.mute = mute
        self.print_timestamp = print_timestamp

    def get_timestamp_str(self):
        return datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")

    def print(self, message):
        if self.print_timestamp:
            if message[0] == "\n" :
                messageString = "\n" + self.get_timestamp_str() + " " + message[1:]
            else:
                messageString = self.get_timestamp_str() + " " + message
        else:
            messageString = message
        if not self.mute:
            print(messageString)
            if self.log_file != None:
                log_file = open(file=self.log_file, mode="a", encoding="utf8")
                log_file.write("\n" + messageString)
                log_file.close()