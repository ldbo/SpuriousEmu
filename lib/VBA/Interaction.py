def Environ(interpreter, arguments):
    return f"%{arguments[0].value}"


def msgBox(interpreter, arguments):
    interpreter.add_stdout(f"msgBox {arguments[0].value}")
