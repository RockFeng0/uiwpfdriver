# encoding:utf-8
import base64,json,socket

def resolve(b64str):
    return json.loads(base64.b64decode(b64str))

def crypt(rawstr):
    return base64.b64encode(json.dumps(rawstr))

def recv_end(the_socket):
    End = base64.b64encode("RockEND")    
    total_data=[];data=''
    while True:
            data=the_socket.recv(8192)
            if End in data:
                total_data.append(data[:data.find(End)])
                break
            total_data.append(data)
            if len(total_data)>1:
                #check if end_of_data was split
                last_pair=total_data[-2]+total_data[-1]
                if End in last_pair:
                    total_data[-2]=last_pair[:last_pair.find(End)]
                    total_data.pop()
                    break
            if not data:
                break
    return ''.join(total_data)

def send_end(sock, b64str):
    End = base64.b64encode("RockEND")
    sock.sendall(b64str+End)
                
request1 = [{"req":{
                 "action":"TypeInWin",
                 "args":["pcteacher2"],
                 "kwargs":{"identifications":{"AutomationId":"PART_EditableTextBox"}},
                 }
            },
            {"req":{
                 "action":"TypeInWin",
                 "args":["123456"],
                 "kwargs":{"identifications":{"AutomationId":"PbPassWord"}},
                 }
            },
            {"req":{
                 "action":"ClickWin",
                 "args":[],
                 "kwargs":{"identifications":{"AutomationId":"BtnLogin"}},
                 }
            },   
            {"req":{
                 "action":"SelectItem",
                 "args":[],
                 "kwargs":{"identifications":{"ClassName":"TabItem"},"index":4},
                 }
            },
            {"req":{
                 "action":"ClickWin",
                 "args":[],
                 "kwargs":{"identifications":{"AutomationId":"MainItemBtn"}},
                 }
            },
            ]
request2 =  [
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
                 'args': (400,400),
                 'kwargs': {'identifications': {"AutomationId" : u"TitleBar"}}
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

## send b64request to server 127.0.0.1:5820
ADDR = ("127.0.0.1", 5820)
sock=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
sock.connect(ADDR)
resp = None
requests = request2   
for req in requests:
    if resp:
        #设置句柄，使用句柄值，记录操作树
        req["req"]["kwargs"]["CURRENT_HANDLE"]=resp["resp"]["globals"]["CURRENT_HANDLE"]
    request = crypt(req)
    send_end(sock, request)
    response=resolve(recv_end(sock))
    print response
sock.close()