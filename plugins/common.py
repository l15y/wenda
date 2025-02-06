from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from starlette.middleware.cors import CORSMiddleware  #引入 CORS中间件模块
from bottle import route, response, request, static_file, hook
import threading
import webbrowser
import re
import json
from yaml import load, dump
try:
    from yaml import CLoader as Loader, CDumper as Dumper
except ImportError:
    from yaml import Loader, Dumper

import argparse
parser = argparse.ArgumentParser(description='Wenda config')
parser.add_argument('-c', type=str, dest="Config",
                    default='config.yml', help="配置文件")
parser.add_argument('-p', type=int, dest="Port", help="使用端口号")
parser.add_argument('-l', type=bool, dest="Logging", help="是否开启日志")
parser.add_argument('-t', type=str, dest="LLM_Type", help="选择使用的大模型")
args = parser.parse_args()


class dotdict(dict):
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__


def object_hook(dict1):
    for key, value in dict1.items():
        if isinstance(value, dict):
            dict1[key] = dotdict(value)
        else:
            dict1[key] = value
    return dotdict(dict1)


green = "\033[1;32m"
red = "\033[1;31m"
white = "\033[1;37m"


def error_helper(e, doc_url):
    error_print(e)
    error_print("查看：", doc_url)
    # webbrowser.open_new(doc_url)


def error_print(*s):
    print(red, end="")
    print(*s)
    print(white, end="")


def success_print(*s):
    print(green, end="")
    print(*s)
    print(white, end="")


wenda_Config = args.Config
wenda_Port = str(args.Port)
wenda_Logging = str(args.Logging)
wenda_LLM_Type = str(args.LLM_Type)
print(args)
try:
    stream = open(wenda_Config, encoding='utf8')
except:
    error_print('加载配置失败，改为加载默认配置')
    stream = open('example.config.yml', encoding='utf8')
settings = load(stream, Loader=Loader)
settings = dotdict(settings)
stream.close()
if wenda_Port != 'None':
    settings.port = wenda_Port
if wenda_Logging != 'None':
    settings.logging = wenda_Logging
if wenda_LLM_Type != 'None':
    settings.llm_type = wenda_LLM_Type
try:
    settings.llm = settings.llm_models[settings.llm_type]
except:
    error_print("没有读取到LLM参数，可能是因为当前模型为API调用。")
del settings.llm_models


settings_str_toprint = dump(dict(settings))
settings_str_toprint = re.sub(r':', ":"+"\033[1;32m", settings_str_toprint)
settings_str_toprint = re.sub(r'\n', "\n\033[1;31m", settings_str_toprint)
print("\033[1;31m", end="")
print(settings_str_toprint, end="")
print("\033[1;37m")

settings_str = json.dumps(settings)
settings = json.loads(settings_str, object_hook=object_hook)


class CounterLock:
    def __init__(self):
        self.lock = threading.Lock()
        self.waiting_threads = 0
        self.waiting_threads_lock = threading.Lock()

    def acquire(self):
        with self.waiting_threads_lock:
            self.waiting_threads += 1
        acquired = self.lock.acquire()

    def release(self):
        self.lock.release()
        with self.waiting_threads_lock:
            self.waiting_threads -= 1

    def get_waiting_threads(self):
        with self.waiting_threads_lock:
            return self.waiting_threads

    def __enter__(self):  # 实现 __enter__() 方法，用于在 with 语句的开始获取锁
        self.acquire()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):  # 实现 __exit__() 方法，用于在 with 语句的结束释放锁
        self.release()


def allowCROS():
    response.set_header('Access-Control-Allow-Origin', '*')
    response.add_header('Access-Control-Allow-Methods', 'POST,OPTIONS')
    response.add_header('Access-Control-Allow-Headers',
                        'Origin, Accept, Content-Type, X-Requested-With, X-CSRF-Token')


app = FastAPI(title="Wenda",
              description="Wenda API",
              version="1.0.0",
              # docs_url=None,
              # redoc_url=None,
              openapi_url="/api/v1/openapi.json",
              docs_url="/api/v1/docs",
              redoc_url="/api/v1/redoc")

#设置允许访问的域名
origins = ["*"]  #也可以设置为"*"，即为所有。

#设置跨域传参
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,  #设置允许的origins来源
    allow_credentials=True,
    allow_methods=["*"],  # 设置允许跨域的http方法，比如 get、post、put等。
    allow_headers=["*"])  #允许跨域的headers，可以用来鉴别来源等作用。
