Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "c:\Users\rodrigo.soares\sistema-gestao-estabelecimentos-main"
objShell.Run """c:\Users\rodrigo.soares\sistema-gestao-estabelecimentos-main\.venv\Scripts\python.exe"" app.py", 0, False
