def CreateObject(interpreter, arguments):
    object_name = arguments[0].value

    if object_name == "WScript.Shell":
        return interpreter.create_object("VBAEnv.WSH.WshShell")
    else:
        return None
