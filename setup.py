from cx_Freeze import setup, Executable

setup(name = "Tool Change Excel" ,
      version = "1.0.0" ,
      description = "DESCRIPTION" ,
      executables = [Executable("index.py", base = "Win32GUI")]
)