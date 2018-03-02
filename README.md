
# 痛点
 - muiapy相对底层一些，需要大量编码，为此，再次封装了一下
 - 只能使用ironpython编写 Window UI的自动化测试脚本，为此采用了三种方式封装: 命令行、交互、socket(简单RPC)
 - 建议采用RPC模式， 打开 uiwpfdriver.exe服务器，监听localhost:5820，然后按照约定的格式写出请求，加密成base64，发送即可

# uiwpfdriver工作原理

![](https://github.com/RockFeng0/uiwpfdriver/raw/master/pic/principle.png)

* * *
# uiwpfdriver 的几种使用方式

## 命令行模式

- 特点是 简单易用
- 在release_version中，下载 command_line_version 的最新版本

```
# 用法
import base64,json,os,subprocess

request = {'req': {
                 'action': 'StartApplication', 
                 'args': (r'D:\auto\python\uiwpfdriver\test_app\npp.5.7.Installer.exe',), 
                 'kwargs': {}
                 }
        }
bstr = base64.b64encode(json.dumps(request))
subp = subprocess.Popen(["uiwpfdriver.exe",bstr],stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
result = subp.stdout.readline().split("TestResultForWpfAction->")[-1]
print json.loads(base64.b64decode(result))

#控制台输出
#    {u'resp': {u'globals': {u'CURRENT_HANDLE': 0}, u'result': True, u'errmsg': u''}}

```

##  交互模式

- 特点是，UI的自动化，跟用户可以交互，如下git
- 在release_version中，下载  interaction_version 的最新版本

![](https://github.com/RockFeng0/uiwpfdriver/raw/master/pic/interactive.gif)


## RPC远程调用模式, 进阶编程

- 特点 是， ironpython环境下的，muiapy工程项目和封装的action，可以使用其他语言的编程，进行远程调用，达到MFC 或者 WPF 的windows窗口的UI自动化
- 在release_version中，下载   interaction_version 的最新版本  
- ** 注意，执行脚本前，先双击打开uiwpfdriver.exe **

```
import base64,json,os,subprocess

request = {'req': {
                 'action': 'StartApplication', 
                 'args': (r'D:\auto\python\uiwpfdriver\test_app\npp.5.7.Installer.exe',), 
                 'kwargs': {}
                 }
        }
bstr = base64.b64encode(json.dumps(request))

ADDR = ("127.0.0.1", 5820)
sock=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
sock.connect(ADDR)
# 注意 RockEND，是结束标志
End = base64.b64encode("RockEND")
# 发送调用请求
sock.sendall(bstr+End)
# 接受结果
result=sock.recv(8192)
print json.loads(base64.b64decode(result[:result.find(End)]))

#控制台输出
#    {u'resp': {u'globals': {u'CURRENT_HANDLE': 0}, u'result': True, u'errmsg': u''}}

```

* * *

# 约定的全局参数

> 只有三个:  index、 timeout、identifications

|参数名|描述|值或值类型 |
|:-:|-|-|
|identifications|属性键值对  ClassName、ConTrolType、AutomationId、Name等,用于标识和寻找元素，默认为空值,使用UiSpy工具查看|
| timeout        | 超时时间     |   默认10秒                                    |
| index		     | 标识元素     |  数字                                         |


# 约定的请求数据格式

- action -> UI执行的动作,参见后面的 API
- args -> action的参数，依据API的顺序传递
- kwargs -> 参见 全局参数

```
# 数据请求格式,
'{
	"req": {
		"action": "",
		"args": [],
		"kwargs": {
			
		}
	}
}'
```
示例，如下:

```
'{
	"req": {
		"action": "TypeInWin",
		"args": ["hello uiwpfdriver"],
		"kwargs": {
			"index": 0,
			"identifications": {
				"AutomationId": "1019"
			},
			"timeout": 10
		}
	}
}'
```


# 约定的响应数据的格式

- globals -> 使用action中 SetVar设置的全局变量
- result -> 执行action的操作结果
- errmsg -> 错误信息

```
'{
	"resp": {
		"globals": {
			
		},
		"result": true,
		"errmsg": ""
	}
}'
```

* * *

# 操作(Actions API)
> 带有【窗口】标记的，要配合全局参数使用
        
|   Actions API          | 参数                  |   返回值  |        描述              |
|-|-|-|-|
|Check                   |                       |           |点击复选框【窗口】        |
|CheckOff                |                       |           |不勾选复选框【窗口】      |
|CheckOn                 |                       |           |勾选复选框【窗口】        |
|ClickWin                |                       |           |点击【窗口】的按钮        |
|CloseWin                |                       |           |关闭指定【窗口】          |
|ExpandOff               |                       |           |关闭组合框【窗口】        |
|ExpandOn                |                       |           |展开组合框【窗口】        |
|GetTextDocumentAttribute|attr                   |           |获取文本组件的指定        |
|                        |                       |           |的attr属性值【窗口】      |
|SwitchToWindow          |win_title_name         |           |所有焦点切换到指定        |
|                        |                       |           |window标题的【窗口】      |
|SwitchToDefaultWindow   |                       |           |切换回默认【窗口】        |
|ScrollTo                |Horizon=-1,            |           |滚动滚动条至指定位        |
|                        |vertical=-1            |           |置【窗口】取值0至100      |
|SelectItem              |                       |           |展开组合框后或列表框      |
|                        |                       |           |选择指定选项【窗口】      |
|SetWinStat              |value->min/max/normal  |           |设置窗口状态【窗口】      |
|StartAppliaction        |app_path               |           |打开指定应用              |
|TimeSleep               |seconds                |           |睡眠                      |
|TypeInWin               |value                  |           |输入【窗口】              |
|SetVar                  |name ,value            |           |设置变量                  |
|IsContentElement        |                       |True/False |是否容器类【窗口】        |
|IsControlElement        |                       |True/False |是否控制类【窗口】        |
|IsEnabled               |                       |True/False |是否可控【窗口】          |
|IsExists                |                       |True/False |是否存在【窗口】          |
|IsKeyboardFocusable     |                       |True/False |是否键盘焦点【窗口】      |
|IsOffscreen             |                       |True/False |是否超出屏幕【窗口】      |
|IsPassword              |                       |True/False |是否密码【窗口】          |
|MouseDrag               |X1,Y1,X2,Y2            |           |鼠标拖拽坐标1             |
|                        |                       |           |至坐标2【窗口】           |
|MouseDoubleClick        |X,Y                    |           |鼠标双击x,y或者           |
|                        |                       |           |元素【窗口】              |
|MouseClick              |X,Y                    |           |左键点击x,y               |
|                        |                       |           |或者元素【窗口】          |
|MousePressButton        |X,Y                    |           |鼠标按住元素【窗口】      |
|                        |                       |           |如果指定了x,y坐标         |
|                        |                       |           |那么按住x,y               |
|MouseMove               |X，Y                   |           |光标移动至【窗口】        |
|                        |                       |           |如果指定了x,y坐标         |
|                        |                       |           |那么移动至x,y             |
|MouseDragTo             |X,Y                    |           |拖拽至坐标【窗口】        |
|AcceleratorKey          |                       |string     |(模型不适用)获取该值      |
|AccessKey               |                       |string     |(模型不适用)获取该值      |
|AutomationId            |                       |string     |(模型不适用)获取该值      |
|BoundingRectangle       |                       |string     |(模型不适用)获取该值      |
|ClassName               |                       |string     |(模型不适用)获取该值      |
|ClickablePoint          |                       |string     |(模型不适用)获取该值      |
|ControlType             |                       |string     |(模型不适用)获取该值      |
|Culture                 |                       |string     |(模型不适用)获取该值      |
|HasKeyboardFocus        |                       |string     |(模型不适用)获取该值      |
|HelpText                |                       |string     |(模型不适用)获取该值      |
|LabeledBy               |                       |string     |(模型不适用)获取该值      |
|LocalizedControlType    |                       |string     |(模型不适用)获取该值      |
|Name                    |                       |string     |(模型不适用)获取该值      |
|ProcessId               |                       |string     |(模型不适用)获取该值      |
|RuntimeId               |                       |string     |(模型不适用)获取该值      |


# muia底层封装实现

> Microsoft WPF UiAutomation的封装，参见: [muiapy项目](https://github.com/RockFeng0/muiapy)

