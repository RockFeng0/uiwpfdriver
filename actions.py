# -*- encoding: utf-8 -*-
'''
Current module: actions

Rough version history:
v1.0    Original version to use

********************************************************************
    @AUTHOR:  Administrator-Bruce Luo(罗科峰)
    MAIL:    lkf20031988@163.com
    RCS:      actions,v 1.0 2017年4月6日
    FROM:   2017年3月16日
********************************************************************

======================================================================

UI and Web Http automation frame for python.

'''


import os,time,subprocess
WINWPF = None;#for exe
# from muiapy_dll import WinWPFDriver as WINWPF;#for debug


class ElementProperty:
    ########################## AutomationElement
    elm = None
        
    #### General Accessibility
    @classmethod
    def AccessKey(cls):        
        if cls.elm: 
            return cls.elm.getProp("AccessKey")

    @classmethod
    def AcceleratorKey(cls):
        if cls.elm: 
            return cls.elm.getProp("AcceleratorKey")
        
    @classmethod
    def IsKeyboardFocusable(cls):
        if cls.elm: 
            return cls.elm.getProp("IsKeyboardFocusable")
        
    @classmethod
    def LabeledBy(cls):
        if cls.elm: 
            return cls.elm.getProp("LabeledBy")
        
    @classmethod
    def HelpText(cls):
        if cls.elm: 
            return cls.elm.getProp("HelpText")
        
    ##### State
    @classmethod
    def IsEnabled(cls):
        if cls.elm: 
            return cls.elm.getProp("IsEnabled")
    
    @classmethod
    def HasKeyboardFocus(cls):
        if cls.elm: 
            return cls.elm.getProp("HasKeyboardFocus")
    
    #### Identification
    @classmethod
    def ClassName(cls):
        if cls.elm: 
            return cls.elm.getProp("ClassName")
        
    @classmethod
    def ControlType(cls):
        if cls.elm: 
            return cls.elm.getProp("ControlType")
    
    @classmethod
    def Culture(cls):
        if cls.elm: 
            return cls.elm.getProp("Culture")
    
    @classmethod
    def AutomationId(cls):
        if cls.elm: 
            return cls.elm.getProp("AutomationId")
    
    @classmethod
    def LocalizedControlType(cls):
        if cls.elm: 
            return cls.elm.getProp("LocalizedControlType")
    
    @classmethod
    def Name(cls):
        if cls.elm: 
            return cls.elm.getProp("Name")
    
    @classmethod
    def ProcessId(cls):
        if cls.elm: 
            return cls.elm.getProp("ProcessId")
        
    @classmethod
    def RuntimeId(cls):
        if cls.elm: 
            return cls.elm.getProp("RuntimeId")
    
    @classmethod
    def IsPassword(cls):        
        if cls.elm: 
            return cls.elm.getProp("IsPassword")
    
    @classmethod
    def IsControlElement(cls):
        if cls.elm: 
            return cls.elm.getProp("IsControlElement")
    
    @classmethod
    def IsContentElement(cls):
        if cls.elm: 
            return cls.elm.getProp("IsContentElement")
    
    #### Visibility
    @classmethod
    def BoundingRectangle(cls):
        if cls.elm: 
            return cls.elm.getProp("BoundingRectangle")
    
    @classmethod
    def ClickablePoint(cls):
        if cls.elm: 
            return cls.elm.getProp("ClickablePoint")
    
    @classmethod
    def IsOffscreen(cls):
        if cls.elm: 
            return cls.elm.getProp("IsOffscreen")
    
    @classmethod
    def IsExists(cls):
        if cls.elm:
            return True
        return False
    
class WPFElement(ElementProperty):
    '''Window WPF Elements''' 
    (identifications,index,timeout)=({},0,10)
    __glob = {}
    __handle = 0
    
    @classmethod
    def SetVar(cls, name, value):
        ''' set static value
        :param name: glob parameter name
        :param value: parameter value
        '''
        cls.__glob.update({name:value})
                
    @classmethod
    def GetVar(cls, name=None):
        if name:
            return cls.__glob.get(name)
        cls.__glob.update({"CURRENT_HANDLE":cls.__handle})
        return cls.__glob
    
    @classmethod
    def SetHandle(cls, hex_or_int_handle):
        cls.__handle = hex_or_int_handle
    
    @classmethod
    def GetHandle(cls):
        return cls.__handle
    
    @classmethod
    def TimeSleep(cls,seconds):
        time.sleep(seconds)
    
    @classmethod
    def StartApplication(cls, app_path):
        if not os.path.exists(app_path):
            raise Exception('Not found "%s"' %app_path)
#         os.system('cmd /c start %s' %app_path)
        subprocess.Popen([app_path])
        time.sleep(0.2)
    
    @classmethod
    def SetWinStat(cls,value):
        # System.Windows.Automation.WindowVisualState  ->UIAutomationTypes.dll
        # Normal = 0; Maximized = 1; Minimized = 2
        result, errmsg = True, ""
        stat = {"normal":0, "max":1, "min":2}
        if not value in stat:
            errmsg = "window stat should be in [normal,max,min]"
            return False, errmsg
        
        try:  
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Window")
            elem_pattern.SetWindowVisualState(stat.get(value))
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def CloseWin(cls):
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Window")
#             print elem_pattern.Current.CanMaximize    
#             print elem_pattern.Current.CanMinimize
#             print elem_pattern.Current.IsModal    
#             print elem_pattern.Current.WindowVisualState    
#             print elem_pattern.Current.WindowInteractionState    
#             print elem_pattern.Current.IsTopmost
            elem_pattern.Close()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def ScrollTo(cls,horizontalPercent=-1,verticalPercent=-1):
        ''' 
        :param horizontalPercent=-1 表示纵向滚动条; verticalPercent=100，表示向下移动100%,即移动到底; verticalPercent=0，表示顶端
        :param verticalPercent=-1 表示横向滚动条; horizontalPercent=100，表示向右移动100%,即移动到最右; horizontalPercent=0，表示左侧       
        '''
        result, errmsg = True, ""
        try:
            horizontalPercent,verticalPercent = float(horizontalPercent),float(verticalPercent)
            for value in (horizontalPercent,verticalPercent):
                if  value != -1.0:
                    if value >100 or value < 0:
                        errmsg = "need float values(%s,%s) between [0-100 or -1]" %(horizontalPercent,verticalPercent)
                        return False,errmsg                
        except:
            errmsg = "need float values(%s,%s)" %(horizontalPercent,verticalPercent)
            return False, errmsg
            
                
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Scroll")            
#             print elem_pattern.Current.HorizontallyScrollable    
#             print elem_pattern.Current.HorizontalScrollPercent    
#             print elem_pattern.Current.HorizontalViewSize    
#             print elem_pattern.Current.VerticallyScrollable
#             print elem_pattern.Current.VerticalScrollPercent    
#             print elem_pattern.Current.VerticalViewSize
            elem_pattern.SetScrollPercent(horizontalPercent,verticalPercent)
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
            
    @classmethod
    def GetTextDocumentAttribute(cls,attr):
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Text") 
            return elem_pattern.DocumentRange.GetAttributeValue(getattr(elem_pattern,attr))
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    #### Control Patterns
    @classmethod
    def TypeInWin(cls,value):
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Value")
#             print elem_pattern.Current.Value
#             print elem_pattern.Current.IsReadOnly
            elem_pattern.SetValue(value)
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def ClickWin(cls):
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Invoke")
            elem_pattern.Invoke()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def Check(cls):
        '''
        :CheckBox 复选框 --可以多选, 如在方框中打勾，或填充 圆点
        :state 
            #    Indeterminate = 2
            #    On = 1
            #    Off = 0
        '''
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Toggle") 
            print "Current Toggle state: %s. Now Switch it." %elem_pattern.Current.ToggleState.value__
            elem_pattern.Toggle()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def CheckOn(cls):
        '''
        :CheckBox 复选框 --可以多选, 如在方框中打勾，或填充 圆点
        :state 
            #    Indeterminate = 2
            #    On = 1
            #    Off = 0
        '''
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Toggle")
            if elem_pattern.Current.ToggleState.value__ == 0:
                elem_pattern.Toggle()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def CheckOff(cls):
        '''
        :CheckBox 复选框 --可以多选, 如在方框中打勾，或填充 圆点
        :state 
            #    Indeterminate = 2
            #    On = 1
            #    Off = 0
        '''
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("Toggle") 
            if elem_pattern.Current.ToggleState.value__ == 1:
                elem_pattern.Toggle()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
                
    @classmethod
    def ExpandOn(cls):
        '''
        :ComboBox 组合框--->需要 下拉框选择的控件
        :state 有四种状态 
            #    LeafNode    =3
            #    PartiallyExpanded = 2
            #    Expanded = 1
            #    Collapsed = 0 
        '''
        result, errmsg = True, ""        
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("ExpandCollapse") 
            if elem_pattern.Current.ExpandCollapseState.value__ == 0:
                elem_pattern.Expand()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    @classmethod
    def ExpandOff(cls,value):
        '''
        :ComboBox 组合框--->需要 下拉框选择的控件
        :state 有四种状态 
            #    LeafNode    =3
            #    PartiallyExpanded = 2
            #    Expanded = 1
            #    Collapsed = 0 
        '''
        result, errmsg = True, ""        
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("ExpandCollapse") 
            if elem_pattern.Current.ExpandCollapseState.value__ == 1:
                elem_pattern.Collapse()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    @classmethod
    def SelectItem(cls):
        '''
        :ComboBox_ListBox 组合框 或者 列表框，展开后，选择条目
        :TabItem 选项卡项  ,选择条目
        '''
        result, errmsg = True, ""
        try:
            cls.elm = elm = cls.__wait()
            elem_pattern = elm.getControl("SelectionItem")
#                 print elem_pattern.Current.IsSelected
            elem_pattern.Select()
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result,errmsg
   
    @classmethod
    def SwitchToWindow(cls, win_title_name):
        ''' In order to find elements more faster, it's a good solution that switch to a window branch in windows tree            
        :return Will Set Handle and Return its value
        '''
        result, errmsg = True, ""
        try:
            win = WINWPF(timeout = cls.timeout).find_element(Name = win_title_name, ControlType = "ControlType.Window")
            cls.__handle = win.getProp("NativeWindowHandle")            
        except:
            result = False
            errmsg = "Can't switch to window '%s'" %win_title_name
        finally:
            return result,errmsg        
    
    @classmethod
    def SwitchToDefaultWindow(cls):
        ''' Set Handle to None. Then to default RootElement. '''
        result, errmsg = True, ""
        cls.__handle = 0
        return result,errmsg
    
    #### Mouse Event    
    @classmethod
    def MouseMove(cls, x=None, y=None):
        ''' move the mouse to the element or position
        defaultly move the element
        if x and y are given, will move to the positoin
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                pos = (int(x), int(y))
            else:
                pos = ()
            cls.__wait().mouse_move(pos)            
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    @classmethod
    def MousePressButton(cls,x = None, y = None):
        ''' press to the element or positon
        defaultly press the element
        if x and y are given, will press the positoin
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                pos = (int(x), int(y))
            else:
                pos = ()      
            cls.__wait().mouse_press_button(pos, button_up=False, button_name="left")
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
        
    @classmethod
    def MouseClick(cls, x = None, y = None, button_name="left"):
        ''' click the element or positon
        defaultly click the element
        if x and y are given, will click the positoin
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                pos = (int(x), int(y))
            else:
                pos = ()                
            cls.__wait().mouse_click(pos, button_name)           
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    @classmethod
    def MouseDoubleClick(cls, x = None, y = None, button_name="left"):
        ''' double click the element or positon
        defaultly double click the element
        if x and y are given, will double click the positoin
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                pos = (int(x), int(y))
            else:
                pos = () 
            cls.__wait().mouse_double_click(pos, button_name)           
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    @classmethod
    def MouseDrag(cls, x = None, y = None, x1 = None, y1 = None):
        ''' press the source position drag to the desitination position
        only for position
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                spos = (int(x), int(y))
            else:
                spos = () 
            
            if x1 and y1:
                dpos = (int(x1), int(y1))
            else:
                dpos = () 
                
            cls.__wait().mouse_drag(spos, dpos)           
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
                
    @classmethod
    def MouseDragTo(cls, x = None, y = None):
        ''' Drag element to the desitination position
        only for element
        '''
        result, errmsg = True, ""
        try:
            if x and y:
                dpos = (int(x), int(y))
            else:
                dpos = () 
                
            cls.__wait().mouse_drag_to(dpos)           
        except Exception,e:
            result, errmsg = False, e.message
        finally:
            return result, errmsg
    
    
    @classmethod
    def __wait(cls):
        if not isinstance(cls.identifications, dict):
            try:
                identify = eval(cls.identifications)
            except:
                identify = None
            finally:
                if not isinstance(identify, dict):
                    raise Exception("Invalid format of Element: %s" %(cls.identifications))
                cls.identifications = identify
                        
        scrope = WINWPF(timeout = cls.timeout)
        if cls.__handle != 0:
            scrope = getattr(WINWPF(timeout = cls.timeout), "find_element_by_handle")(cls.__handle)
            find = getattr(scrope, "find_element_by_handle")(cls.__handle)
            
        if cls.index != 0:
            find = getattr(scrope, "find_elements")
        else:
            find = getattr(scrope, "find_element")
        
        elements = find(**cls.identifications)
        if isinstance(elements, list) and len(elements):
            print "Got %d elements, choosed index=%d" %(len(elements),cls.index)
            return elements[cls.index]
        else:
            return elements
        


# class SelectionPattern:
#     pass
        
# class DockPattern:
#     pass
# 
# class GridItemPattern:
#     pass
# 
# class GridPattern:
#     pass
# class MultipleViewPattern:
#     pass
# 
# class RangeValuePattern:
#     pass    
# 
# class ScrollItemPattern:
#     pass
# 
# class TransformPattern:
#     pass

#     
def usage_for_mfc_app():    
    window_title1 = u"Installer Language"
    window_title2 = u"Notepad++ v5.7 安装" 
    WPFElement.StartApplication(r"F:\BaiduYunDownload\pcinstall\npp.5.7.Installer.exe")
    
    # Use UISpy or others to spy the WPF UI.    
    handle1 = WPFElement.SwitchToWindow(window_title1)
    print "handle1 ->", handle1  
    
    dpos = (400,400)
    WPFElement.identifications = {"AutomationId" : u"TitleBar"}
    WPFElement.MouseDragTo(*dpos)
    time.sleep(1)
    
    WPFElement.identifications = {"Name" : u"English"}
    WPFElement.SelectItem()
    
    WPFElement.identifications = {"Name" : u"Chinese (Simplified)"}
    WPFElement.SelectItem()
    
    WPFElement.identifications = {"Name" : u"OK"}  
    WPFElement.ClickWin()  
    print "---"
    
    handle2 = WPFElement.SwitchToWindow(window_title2)
    print "handle2 ->", handle2
    WPFElement.identifications = {"Name" : u"下一步(N) >"}
    WPFElement.ClickWin()
    
    WPFElement.identifications = {"Name" : u"我接受(I)"}
    WPFElement.ClickWin()
    
    WPFElement.identifications = {"AutomationId" : "1019"}
    WPFElement.TypeInWin(ur'd:\hello input你好')
    
    WPFElement.identifications = {"Name" : u"下一步(N) >"}
    WPFElement.ClickWin()
    
    WPFElement.identifications = {"Name" : u"取消(C)"}
    WPFElement.ClickWin()
    print "---"
    
    WPFElement.SwitchToDefaultWindow()
    WPFElement.identifications = {"Name" :  u"是(Y)"}
#     WPFElement.ClickWin()
    WPFElement.MouseClick()
    
def usage_for_wpf_app():
    dut = r"D:\auto\buffer\AiSchool\AiTeacherCenter\AiTeacherCenter\AiTeacher.exe"
    WPFElement.StartApplication(dut)
    
    #(identifications,prop,index,timeout)=({},None,0,10)
    WPFElement.identifications = {"AutomationId" : "txtUserName"}
    print "IsPassword: %s" %WPFElement.IsPassword()
    print "setting username value: Hello MUIA."
    WPFElement.TypeInWin("Hello MUIA")                
    print "---"
    
    WPFElement.identifications = {"AutomationId" : "PwdUser"}
    print "IsPassword: %s" %WPFElement.IsPassword()
    print "setting username value: 123456."
    WPFElement.TypeInWin("123456")                
    print "---"
        
    WPFElement.identifications = {"AutomationId" : "ckbIsSavePwd"}
    print "Name: %s" %WPFElement.Name()
    WPFElement.SwitchToggle()                    
    print "---"
    
    WPFElement.identifications = {"AutomationId" : "BtnLogin"}
    print "Name: %s" %WPFElement.Name()
    print "IsKeyboardFocusable: %s" %WPFElement.IsKeyboardFocusable()        
    WPFElement.ClickWin()
        
if __name__ == "__main__":    
    #ipy.exe driver.py
    ##### notepad++ installation example
    usage_for_mfc_app()
  
    #### AiSchool login example
#     usage_for_wpf_app()
    
