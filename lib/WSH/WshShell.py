class WshShell:
    variables = []

    @staticmethod
    def Run(interpreter, arguments):
        command = arguments[0].value
        interpreter.add_command_execution(command)
