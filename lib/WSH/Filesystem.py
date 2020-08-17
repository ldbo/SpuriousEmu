class Filesystem:
    variables = []

    @staticmethod
    def FolderExists(interpreter, arguments):
        interpreter.add_file_event('FolderExists', arguments[0].value)
        return False
