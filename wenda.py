from fastapi import FastAPI, WebSocket, WebSocketDisconnect
import time
from starlette.requests import Request
from fastapi.middleware.wsgi import WSGIMiddleware
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
import asyncio
import functools
import bottle
from bottle import route, response, request, static_file, hook
import datetime
import json
import os
import threading
import torch
from plugins.common import error_helper, error_print, success_print
from plugins.common import allowCROS
from plugins.common import settings
from plugins.common import app
# from md2pptx.md2pptxllm import writeMdFile, genPPTX
from md2pptx.md2ppt import writeMdFile, genPPTX
from img2analysis.imganalysis import gencontent
from starlette.responses import FileResponse
import logging
logging.captureWarnings(True)
logger = None
try:
    from loguru import logger
except:
    pass


def load_LLM():
    try:
        from importlib import import_module
        LLM = import_module('llms.llm_'+settings.llm_type)
        return LLM
    except Exception as e:
        logger and logger.exception(e)
        print("LLM模型加载失败，请阅读说明：https://github.com/l15y/wenda", e)


LLM = load_LLM()

logging = settings.logging
if logging:
    from plugins.defineSQL import session_maker, 记录


model = None
tokenizer = None


def load_model():
    LLM.load_model()
    torch.cuda.empty_cache()
    success_print("模型加载完成")


if __name__ == '__main__':
    thread_load_model = threading.Thread(target=load_model)
    thread_load_model.start()
zhishiku = None


def load_zsk():
    try:
        global zhishiku
        import plugins.zhishiku as zsk
        zhishiku = zsk
        success_print("知识库加载完成")
    except Exception as e:
        logger and logger.exception(e)
        error_helper(
            "知识库加载失败，请阅读说明", r"https://github.com/l15y/wenda#%E7%9F%A5%E8%AF%86%E5%BA%93")
        raise e


if __name__ == '__main__':
    thread_load_zsk = threading.Thread(target=load_zsk)
    thread_load_zsk.start()


@route('/llm')
def llm_js():
    noCache()
    return static_file('llm_'+settings.llm_type+".js", root="llms")


@route('/plugins')
def read_auto_plugins():
    noCache()
    plugins = []
    for root, dirs, files in os.walk("autos"):
        for file in files:
            if(file.endswith(".js")):
                file_path = os.path.join(root, file)
                with open(file_path, "r", encoding='utf-8') as f:
                    plugins.append({"name": file, "content": f.read()})
    return json.dumps(plugins)
# @route('/writexml', method=("POST","OPTIONS"))
# def writexml():
    # data = request.json
    # s=json2xml(data).decode("utf-8")
    # with open(os.environ['wenda_'+'Config']+"_",'w',encoding = "utf-8") as f:
    #     f.write(s)
    #     # print(j)
    #     return s


def noCache():
    response.set_header("Pragma", "no-cache")
    response.add_header("Cache-Control", "must-revalidate")
    response.add_header("Cache-Control", "no-cache")
    response.add_header("Cache-Control", "no-store")


def pathinfo_adjust_wrapper(func):
    # A wrapper for _handle() method
    @functools.wraps(func)
    def _(s, environ):
        environ["PATH_INFO"] = environ["PATH_INFO"].encode(
            "utf8").decode("latin1")
        return func(s, environ)
    return _


bottle.Bottle._handle = pathinfo_adjust_wrapper(
    bottle.Bottle._handle)  # 修复bottle在处理utf8 url时的bug


@hook('before_request')
def validate():
    REQUEST_METHOD = request.environ.get('REQUEST_METHOD')
    HTTP_ACCESS_CONTROL_REQUEST_METHOD = request.environ.get(
        'HTTP_ACCESS_CONTROL_REQUEST_METHOD')
    if REQUEST_METHOD == 'OPTIONS' and HTTP_ACCESS_CONTROL_REQUEST_METHOD:
        request.environ['REQUEST_METHOD'] = HTTP_ACCESS_CONTROL_REQUEST_METHOD


waiting_threads = 0


@route('/chat_now', method=('GET', "OPTIONS"))
def api_chat_now():
    allowCROS()
    noCache()
    return {'queue_length': waiting_threads}


@route('/find', method=("POST", "OPTIONS"))
def api_find():
    allowCROS()
    data = request.json
    if not data:
        return '0'
    prompt = data.get('prompt')
    step = data.get('step')
    if step is None:
        step = int(settings.library.step)
    return json.dumps(zhishiku.find(prompt, int(step)))

import sys
import runpy

@route('/genppt', method=("POST", "OPTIONS"))
def api_genppt():
    allowCROS()
    data = request.json
    if not data:
        return '0'
    content = data.get('content')
    # print(content)
    # inputfile = writeMdFile('template: CHATOVENS.pptx\n\n' + 'cardlayout: horizontal\n' +
    #                         'baseTextSize: 20\n' +
    #                         'CardColour: BACKGROUND 2\n' +
    #                         'CardTitlePosition: inside\n' +
    #                         'cardshadow: yes\n' +
    #                         'cardshape: rounde\n\n' + content)
    inputfile = writeMdFile(content)
    now = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
    outputfile = os.getcwd() + os.sep + 'md2pptx' + os.sep + 'data' + os.sep + now + r'output.pptx'
    # sys.argv.append(inputfile)
    # sys.argv.append(outputfile)
    #
    # runpy.run_path('./md2pptx/md2pptx', init_globals={'sys': sys})
    outputfile = genPPTX(outputfile, inputfile)
    return outputfile
    # print(outputfile)
    # return FileResponse(outputfile, 'output.pptx')

@app.get('/api/downloadppt')
def downloadppt(outputfile):
    allowCROS()
    noCache()
    # print(outputfile)
    return FileResponse(outputfile)

@route('/imganalysis', method=("POST", "OPTIONS"))
def imganalysis():
    allowCROS()
    # data = request.json
    imgbase64 = request.forms['imgbase64']
    # print(outputfile)
    return gencontent(imgbase64)

@route('/completions', method=("POST", "OPTIONS"))
def api_chat_box():
    response.content_type = "text/event-stream"
    response.add_header("Connection", "keep-alive")
    response.add_header("Cache-Control", "no-cache")
    response.add_header("X-Accel-Buffering", "no")
    data = request.json
    messages = data.get('messages')
    stream = data.get('stream')
    prompt = messages[-1]['content']
    data['prompt'] = prompt
    history = []
    for i, old_chat in enumerate(messages[0:len(messages)-1]):
        if old_chat['role'] == "user":
            history.append(old_chat)
        elif old_chat['role'] == "assistant":
            old_chat['role'] = "AI"
            history.append(old_chat)
        else:
            continue
    data['history'] = history
    data['level'] = 0
    from websocket import create_connection
    ws = create_connection("ws://127.0.0.1:"+str(settings.port)+"/ws")
    ws.send(json.dumps(data))
    if not stream:
        response.content_type = "application/json"
        temp_result = ''
        try:
            while True:
                result = ws.recv()
                if len(result) > 0:
                    temp_result = result
        except:
            pass
        yield json.dumps({"response": temp_result})

    else:
        try:
            while True:
                result = ws.recv()
                if len(result) > 0:
                    yield "data: %s\n\n" % json.dumps({"response": result})
        except:
            pass
        yield "data: %s\n\n" % "[DONE]"

    ws.close()


@route('/chat_stream', method=("POST", "OPTIONS"))
def api_chat_stream():
    allowCROS()
    response.add_header("Connection", "keep-alive")
    response.add_header("Cache-Control", "no-cache")
    response.add_header("X-Accel-Buffering", "no")
    data = request.json
    data = json.dumps(data)
    from websocket import create_connection
    ws = create_connection("ws://127.0.0.1:"+str(settings.port)+"/ws")
    ws.send(data)
    try:
        while True:
            result = ws.recv()
            if len(result) > 0:
                yield result
    except:
        pass
    ws.close()


@route('/chat', method=("POST", "OPTIONS"))
def api_chat():
    allowCROS()
    data = request.json
    data = json.dumps(data)
    from websocket import create_connection
    ws = create_connection("ws://127.0.0.1:"+str(settings.port)+"/ws")
    ws.send(data)
    try:
        while True:
            new_result = ws.recv()
            if len(new_result) > 0:
                result = new_result
    except:
        pass
    ws.close()
    print([result])
    return result


bottle.debug(True)


@app.middleware("http")
async def add_process_time_header(request: Request, call_next):

    start_time = time.time()
    response = await call_next(request)
    path=request.scope['path']
    if path.startswith('/static/') and not path.endswith(".html"):
        return response
        
    process_time = time.time() - start_time
    response.headers["X-Process-Times"] = str(process_time)
    response.headers["Pragma"] = "no-cache"
    response.headers["Cache-Control"] = "no-cache,no-store,must-revalidate"

    return response
users_count = [0]*4


def get_user_count_before(level):
    count = 0
    for i in range(level):
        count += users_count[i]
    return count


class AsyncContextManager:
    def __init__(self, level):
        self.level = level

    async def __aenter__(self):
        users_count[self.level] += 1

    async def __aexit__(self, exc_type, exc, tb):
        users_count[self.level] -= 1


Lock = AsyncContextManager


@app.websocket('/ws')
async def websocket_endpoint(websocket: WebSocket):
    global waiting_threads
    await websocket.accept()
    waiting_threads += 1
    # await asyncio.sleep(5)
    try:
        data = await websocket.receive_json()
        prompt = data.get('prompt')
        max_length = data.get('max_length')
        if max_length is None:
            max_length = 2048
        top_p = data.get('top_p')
        if top_p is None:
            top_p = 0.7
        temperature = data.get('temperature')
        if temperature is None:
            temperature = 0.9
        keyword = data.get('keyword')
        if keyword is None:
            keyword = prompt
        level = data.get('level')
        if level is None:
            level = 3
        history = data.get('history')
        history_formatted = LLM.chat_init(history)
        response = ''
        IP = websocket.client.host
        count_before = get_user_count_before(4)

        if count_before >= 4-level:
            time2sleep = (count_before+1)*level
            while time2sleep > 0:
                await websocket.send_text('正在排队，当前计算中用户数：'+str(count_before)+'\n剩余时间：'+str(time2sleep)+"秒")
                await asyncio.sleep(1)
                count_before = get_user_count_before(4)
                if count_before < 4-level:
                    break
                time2sleep -= 1
        lock = Lock(level)
        async with lock:
            print("\033[1;32m"+IP+":\033[1;31m"+prompt+"\033[1;37m")
            try:
                torch.cuda.empty_cache()
                for response in LLM.chat_one(prompt, history_formatted, max_length, top_p, temperature, data):
                    if (response):
                        # start = time.time()
                        await websocket.send_text(response)
                        await asyncio.sleep(0)
                        # end = time.time()
                        # cost+=end-start
            except Exception as e:
                error = str(e)
                torch.cuda.empty_cache()
                await websocket.send_text("错误"+ error)
                await websocket.close()
                raise e
            torch.cuda.empty_cache()
        if logging:
            with session_maker() as session:
                jl = 记录(时间=datetime.datetime.now(),
                        IP=IP, 问=prompt, 答=response)
                session.add(jl)
                session.commit()
        print(response)
        await websocket.close()
    except WebSocketDisconnect:
        pass
    waiting_threads -= 1


@app.get("/")
async def index(request: Request):
    return RedirectResponse(url="/index.html")

app.mount(path="/chat/", app=WSGIMiddleware(bottle.app[0]))
app.mount(path="/api/", app=WSGIMiddleware(bottle.app[0]))
app.mount("/txt/", StaticFiles(directory="txt"), name="txt")
app.mount("/", StaticFiles(directory="views"), name="static")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=settings.port,
                log_level='error', loop="asyncio")
