class Logger:
    def __init__(self, name):
        self._log_file = open(name, 'w', newline="\n")

    def log(self, *args):
        for arg in args:
            self._log_file.write(str(arg))
            self._log_file.write('')

            try:
                print(arg)
            except UnicodeEncodeError:
                print(arg.encode('ascii', 'ignore').decode('ascii'))