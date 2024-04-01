
import json
f = open('config.json')
data = json.load(f)
f.close()


APP_NAME = data["app_name"]
WINDOW_SIZE = data["window_size"]
