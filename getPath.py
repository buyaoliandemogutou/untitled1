import os
import time

nowTime=time.strftime("%Y-%m-%d", time.localtime())

def get_Path():
    path = os.path.split(os.path.realpath(__file__))[0]
    return path