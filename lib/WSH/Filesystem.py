class Filesystem:
    variables = []

    @staticmethod
    def FolderExists(interpreter, arguments):
        interpreter.add_file_event('FolderExists', arguments[0].value)
        return False

    @staticmethod
    def CreateTextFile(interpreter, arguments):
        interpreter.add_file_event('CreateTextFile', arguments[0].value)
        return interpreter.create_object('VBAEnv.WSH.FileHandler')
