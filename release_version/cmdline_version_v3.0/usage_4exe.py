# encoding:utf-8
import base64,subprocess,json,time

def resolve(b64str):
    return json.loads(base64.b64decode(b64str))

def crypt(rawstr):
    return base64.b64encode(json.dumps(rawstr))

def get_response(cmd):
    subp = subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)        
    result_code = 1
    result = None
    timeout_set = False
    while True:
        if not subp.poll()==None and result_code == 1:
            #进程已结束，仍然没有得到结果，超时30秒                          
            if not timeout_set:
                print 'listen response. just waiting 15 seconds...'
                time_now = time.time() + 15
                timeout_set = True
            
            if time.time() >= time_now:
                print 'listen response timeout at 15 seconds.'
                break
        
        next_line = subp.stdout.readline().decode('cp936')           

        if next_line:
            print next_line 
            if "TestResultForWpfAction->" in str(next_line):
                result = next_line.split("TestResultForWpfAction->")[-1]
#                     result = next_line.replace("TestResultForCurrentCase->","")
                result_code = 0                    
                break
            else:                    
                result_code = 1
    
    if subp.returncode:
        result_code = 1
        
    if result_code == 0:            
        return resolve(result)
            
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
request2 = [
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
        
requests = request2    
resp = None
for req in requests:
    if resp:
        #设置句柄， 生成exe后，使用句柄值，记录操作树
        req["req"]["kwargs"]["CURRENT_HANDLE"]=resp["resp"]["globals"]["CURRENT_HANDLE"]
    cmd = ["uiwpfdriver.exe", crypt(req)]
    print req["req"]["action"]
    resp = get_response(cmd)
    print resp
