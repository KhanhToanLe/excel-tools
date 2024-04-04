import datetime

def write(traceback):
    f = open("excception.txt", "a")
    log_message = str(datetime.datetime.now()) + " : " + traceback +'\n'
    f.write(log_message)
    f.close()