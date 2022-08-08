#! ipy2.7
# -*- encoding: utf-8 -*-


import clr
clr.AddReference('muia')  # muia.dll for uiwpfdriver.exe
from muia import WinWPFDriver
# from muiapy_dll import WinWPFDriver  # for debug uiwpfdriver.py

import sys
import base64
import json
import socket
import actions
from json.decoder import errmsg
actions.WINWPF = WinWPFDriver
WPFElement = actions.WPFElement
ADDR = ('127.0.0.1', 5820)

# ironpython's bug when pack to exe-->IronPython.Runtime.Exceptions.LookupException: unknown encoding: hex
# import it to avoid: (hex_codec, ctypes)
from encodings import hex_codec
import ctypes


def get_callable_class_method_names():
    isInstanceMethod = lambda attrname: not attrname.startswith("_") and hasattr(getattr(WPFElement, attrname), '__call__')
    return filter(isInstanceMethod, dir(WPFElement))


def resolve(b64str):
    return json.loads(base64.b64decode(b64str))


def crypt(rawstr):
    return base64.b64encode(json.dumps(rawstr))


def recv_end(the_socket):
    End = base64.b64encode("RockEND")
    total_data = []
    data = ''
    while True:
            data = the_socket.recv(8192)
            if End in data:
                total_data.append(data[:data.find(End)])
                break
            total_data.append(data)
            if len(total_data) > 1:
                #check if end_of_data was split
                last_pair=total_data[-2]+total_data[-1]
                if End in last_pair:
                    total_data[-2] = last_pair[:last_pair.find(End)]
                    total_data.pop()
                    break
            if not data:
                break
    return ''.join(total_data)


def send_end(sock, b64str):
    End = base64.b64encode("RockEND")
    sock.sendall(b64str+End)


def mapcore(b64req):
    request = {
        "req": {
            "action": "",
            "args": [],
            "kwargs": {},
            }
        }

    response = {
        "resp": {
            "result": True,
            "errmsg": "",
            "globals": {}
            }
        }

    request.update(resolve(b64req))

    action = request["req"]["action"]
    r_args = request["req"]["args"]
    r_kwargs = request["req"]["kwargs"]

    element_actions = get_callable_class_method_names()
    if action in element_actions:
        WPFElement.identifications, WPFElement.index, WPFElement.timeout = r_kwargs.get("identifications",{}), r_kwargs.get("index",0), r_kwargs.get("timeout",10)
        # after building to exe for command line, we can use handle to remamber the window which the previous action want to switch to.
        WPFElement.SetHandle(r_kwargs.get("CURRENT_HANDLE", 0))

        params = []
        if r_args:
            for arg in r_args:
                if isinstance(arg, str):
                    try:
                        params.append(arg.decode('utf-8'))
                    except:
                        params.append(arg)
                else:
                    params.append(arg)

        errmsg = ""
        try:
            if params:
                result = getattr(WPFElement, action)(*params)
            else:
                result = getattr(WPFElement, action)()
        except Exception as e:
            result, errmsg = False, str(e)
        finally:
            if isinstance(result, tuple):
                result, errmsg = result[0], result[1]
            result = False if result == False else True

    else:
        result,errmsg = False,"Do not hava action:%s" %action

    # GetVar also wiil response the handle named CURRENT_HANDLE
    response.update({
        "resp": {
                 "result": result,
                 "errmsg": errmsg,
                 "globals": WPFElement.GetVar()
             }
        }
    )

    resp = crypt(response)

    # this line is available for getting response from exe
    print "TestResultForWpfAction->%s" % resp
    # this line is available for back calling
    return resp


def main_server():
    # for server
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(ADDR)
    sock.listen(1)
    while True:
        print "Waiting for connection..."
        # 阻塞，并等待链接
        newsock, address = sock.accept()
        print "Connected to %s:%s" % (address[0], address[1])

        while True:
            request = recv_end(newsock)
            if not request:
                break
            response = mapcore(request)
            send_end(newsock, response)


def main_command():
    # for command line
    return resolve(mapcore(sys.argv[1]))


def usage():
    requests = [{"req": {
                 "action": "TypeInWin",
                 "args": ["pcteacher2"],
                 "kwargs":{"identifications": {"AutomationId": "PART_EditableTextBox"}},
                 }
            },
            {"req": {
                 "action": "TypeInWin",
                 "args": ["123456"],
                 "kwargs":{"identifications": {"AutomationId": "PbPassWord"}},
                 }
            },
            {"req": {
                 "action": "ClickWin",
                 "args": [],
                 "kwargs":{"identifications": {"AutomationId": "BtnLogin"}},
                 }
            },
            {"req": {
                 "action": "SelectItem",
                 "args": [],
                 "kwargs":{"identifications": {"ClassName": "TabItem"}, "index": 4},
                 }
            },
            {"req": {
                 "action": "ClickWin",
                 "args": [],
                 "kwargs":{"identifications": {"AutomationId": "MainItemBtn"}},
                 }
            },
            ]
    for i in requests:
        sys.argv = ["", base64.b64encode(json.dumps(i))]
        main_command()


def usage2(mtype = "command"):
    requests = [
        {'req': {
                 'action': 'StartApplication',
                 'args': ('D:\\auto\\pc_install\\npp.5.7.Installer.exe',),
                 'kwargs': {}
                 }
        },
        {'req': {
                 'action': 'SwitchToWindow',
                 'args': (u"Installer Language",),
                 'kwargs': {}
                 }
        },

        {'req': {
                 'action': 'MouseDragTo',
                 'args': (400, 400),
                 'kwargs': {'identifications': {"AutomationId": u"TitleBar"}}
                 }
        },
        {'req': {'action': 'TimeSleep', 'args': (1,), 'kwargs': {}}},
        {'req': {
                 'action': 'ClickWin',
                 'args': (),
                 'kwargs': {'index': 0, 'identifications': {'Name': 'OK'}, 'timeout': 10}
                 }
        },

        {'req': {
                 'action': 'SwitchToWindow',
                 'args': (u"Notepad++ v5.7 安装" ,),
                 'kwargs': {}
                 }
        },
        {'req': {'action': 'ClickWin', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'下一步(N) >'}, 'timeout': 10}}},
        {'req': {'action': 'ClickWin', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'我接受(I)'}, 'timeout': 10}}},
        {'req': {'action': 'TypeInWin', 'args': (u'd:\\hello input',), 'kwargs': {'index': 0, 'identifications': {'AutomationId': '1019'}, 'timeout': 10}}},
        {'req': {'action': 'ClickWin', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'下一步(N) >'}, 'timeout': 10}}},
        {'req': {'action': 'ClickWin', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'取消(C)'}, 'timeout': 10}}},
        {'req': {
                 'action': 'SwitchToDefaultWindow',
                 'args': (),
                 'kwargs': {}
                 }
        },
        #{'req': {'action': 'ClickWin', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'是(Y)'}, 'timeout': 10}}}
        {'req': {'action': 'MouseClick', 'args': (), 'kwargs': {'index': 0, 'identifications': {'Name': u'是(Y)'}, 'timeout': 10}}}
    ]

    if mtype == "command":
        resp = None
        for req in requests:
            if resp:
                #设置句柄， 生成exe后，使用句柄值，记录操作树
                req["req"]["kwargs"]["CURRENT_HANDLE"]=resp["resp"]["globals"]["CURRENT_HANDLE"]
            sys.argv=["",crypt(req)]
            resp = main_command()
            print resp
    elif mtype == "server":
        ## run main_server() first
        ## then call usage2("server") to send b64request to server 127.0.0.1:5820;
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect(ADDR)
        resp = None
        for req in requests:
            if resp:
                #设置句柄， 生成exe后，使用句柄值，记录操作树
                req["req"]["kwargs"]["CURRENT_HANDLE"]=resp["resp"]["globals"]["CURRENT_HANDLE"]
            request = crypt(req)
            send_end(sock, request)
            response=resolve(recv_end(sock))
            print response
        sock.close()


if __name__ == "__main__":
#     main_command()
    main_server()
#     usage2("command")
