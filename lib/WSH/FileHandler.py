class FileHandler:
    variables = ["path", "open"]

    @staticmethod
    def Write(interpreter, arguments):
        # TODO check open
        path = arguments[0].value['path']
        interpreter.add_file_event("Write", path, arguments[1].value)

    @staticmethod
    def Close(interpreter, arguments):
        # TODO set open to False
        pass
