import json
f = open('config.json')
data = json.load(f)
f.close()

APP_NAME = data["appName"]
WINDOW_SIZE = data["windowSize"]
PREFIX_MERGE = data["prefixMerge"]
OPEN_EXCEL_WHILE_RUNNING = data["isShowExcelWhileRunning"]
DEFAULT_MERGE_FILE_NAME = data["defaultMergeFileName"]
LANGUAGE_OPTION = data["languageOption"]
