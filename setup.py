from cx_Freeze import setup, Executable

build_exe_options = {"excludes": ["tkinter", "sqlite3", "ctypes", "logging", "test", "pydoc_data",
                                  "flask", "asyncio", "concurrent", "unittest", "lib2to3",
                                  "distutils", "xlrd-1.2.0.dist-info"],
                     "optimize": 2,
                     "include_files": ["data.xlsx", "message.html"]}

setup(name="AutoMail",
      version="0.1",
      description="Send bulk emails in Outlook to a list of recipients",
      options={"build_exe": build_exe_options},
      executables=[Executable("script.py")])
