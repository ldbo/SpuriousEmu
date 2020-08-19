class Filesystem:
    variables = []

    @staticmethod
    def FolderExists(interpreter, arguments):
        interpreter.add_file_event('FolderExists', arguments[1].value)
        return False

    @staticmethod
    def CreateTextFile(interpreter, arguments):
        path = arguments[1].value
        interpreter.add_file_event('CreateTextFile', path)
        file_handler = interpreter.create_object('VBAEnv.WSH.FileHandler')
        file_handler.value['path'] = path
        return file_handler

    @staticmethod
    def CreateFolder(interpreter, arguments):
        path = arguments[1].value
        interpreter.add_file_event('CreateFolder', path)
