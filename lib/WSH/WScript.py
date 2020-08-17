def CreateObject(interpreter, arguments):
    object_name = arguments[0].value

    if object_name == "WScript.Shell":
        return interpreter.create_object("VBAEnv.WSH.WshShell")
    elif object_name == "Scripting.FileSystemObject":
        return interpreter.create_object("VBAEnv.WSH.Filesystem")
    else:
        return None
