import os

class Logger:
    def __init__(self, filename):
        if not os.path.exists(os.path.dirname(filename)):
            try:
                os.makedirs(os.path.dirname(filename))
            except OSError as e:
                print(e.strerror)

                self.is_open = False

                return

        self.log_file = open(filename, 'w', newline="\n")
        self.is_open = True

    def log(self, *args):
        for arg in args:
            if self.is_open == True:
                self.log_file.write(str(arg))
                self.log_file.write('\r\n')

            try:
                print(arg)
            except UnicodeEncodeError:
                print(arg.encode('ascii', 'ignore').decode('ascii'))