# encoding:utf-8
import base64,json,socket,sys

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

## send b64request to server 127.0.0.1:5820
def wait_input(sock):
    WPF_ACTIONS = ['Check', 'CheckOff', 'CheckOn', 'ClickWin', 'CloseWin', 'ExpandOff', 'ExpandOn', 'GetTextDocumentAttribute', 'SwitchToWindow', 'SwitchToDefaultWindow', 'ScrollTo', 'SelectItem', 'SetWinStat', 'StartApplication', 'TimeSleep', 'TypeInWin', 'SetVar']
    WPF_PROPERTIES = ['AcceleratorKey', 'AccessKey', 'AutomationId', 'BoundingRectangle','ClassName', 'ClickablePoint', 'ControlType', 'Culture', 'HasKeyboardFocus',
                   'HelpText', 'IsContentElement', 'IsControlElement', 'IsEnabled', 'IsExists', 'IsKeyboardFocusable', 'IsOffscreen', 'IsPassword', 'LabeledBy', 'LocalizedControlType', 'Name', 'ProcessId', 'RuntimeId']
    MFC_MOUSE_ACTIONS = ["MouseMove", "MousePressButton", "MouseClick", "MouseDoubleClick", "MouseDrag", "MouseDragTo"]
    
    print u"""帮助命令: 
    action  查看所有的操作
    list    查看示例
    """
    # print u"'#'号后面输入操作; '>'号后输入操作的传入值; '-'号后面输入元素的识别属性"
    
    resp = None
    while True:
        action = raw_input(u"# ").strip()
        if action == "action":
            print WPF_ACTIONS+WPF_PROPERTIES+MFC_MOUSE_ACTIONS,"\n"
            continue
            
        if action == "list":
            print u"""
说明:
    符号'#'   操作符号,输入您的操作,示例: # ClickWin
    符号'>'   传参符号,输入您传入的参数,逗号隔开,示例: > username,passwd,sex
    符号'-'   元素符号,输入您元素识别属性值,逗号隔开,示例: - AutomationId=PbPassWord,Name=下一步
示例:
    # TypeIn
    > administrator
    - AutomationId=LoginEdit,ClassName=TextBox
    # MouseDrag
    > 100,100,200,200
    - ClassName=Button
            """
            continue
        
        args = raw_input("> ").strip().decode('cp936')
        print "--->%s" %repr(args)
        if args:
            args = args.split(",")
        else:
            args=()        
        
        idens = raw_input(u"- ").strip().decode('cp936').split(",")                
        try:            
            z = [i.split("=",1) for i in idens if len(i.split("=",1))==2]
            if z:
                identifications = dict(z)
            else:
                identifications = {}
            index = int(identifications.pop("index",0))
            timeout = int(identifications.pop("index",10))
            req = {'req': {
                         'action': action, 
                         'args': args, 
                         'kwargs': {'index': index, 'identifications': identifications, 'timeout': timeout}
                        }
                    }
            
            if resp:
                req["req"]["kwargs"]["CURRENT_HANDLE"]=resp["resp"]["globals"]["CURRENT_HANDLE"]
            
            request = crypt(req)
            send_end(sock, request)
            resp=resolve(recv_end(sock))
            
            if resp:
                print "\nResult:",resp["resp"]["result"],resp["resp"]["errmsg"]
        except Exception,e:
            print "\nError:",e
        print "-----------------\n\n"

if __name__ == "__main__":
    print u"UiDriver Tools Author: Bruce Luo(罗科峰)"
    
    try:
        ADDR = ("127.0.0.1", 5820)
        sock=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
        sock.connect(ADDR)        
    except:
        sys.exit(1)
        
    try:
        wait_input(sock)
        sock.close()
    except:
        pass
    
    
    