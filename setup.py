from cx_Freeze import setup, Executable

includefiles = ['config.json','exception_logger.py','helper.py','index.py','InputData.py','Interface.py','xlsxHelper.py']

setup(name = "Tool Change Excel" ,
      version = "1.0.0" ,
      description = "DESCRIPTION" ,
      executables = [Executable("index.py", base = "Win32GUI")],
      options = {'build_exe': {
      # 'includes': includes,
      # 'excludes': excludes,
      # 'packages': packages,
      'include_files': includefiles}
      }, 
)