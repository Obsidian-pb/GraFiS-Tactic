Attribute VB_Name = "WorkWithKeys"
Option Explicit
'============= Функции для определения нажатой клавиши =================================

#If VBA7 Then
    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As VirtualKeys) As Integer
#Else
    Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As VirtualKeys) As Integer
#End If
 
 
Public Enum VirtualKeys    ' Virtual Keys, Standard Set
    VK_LBUTTON = &H1: VK_RBUTTON = &H2: VK_CANCEL = &H3: VK_MBUTTON = &H4
    'VK_MBUTTON = &H4 -  NOT contiguous with L RBUTTON
    VK_BACK = &H8: VK_TAB = &H9: VK_CLEAR = &HC: VK_RETURN = &HD
    VK_SHIFT = &H10: VK_CONTROL = &H11: VK_MENU = &H12: VK_PAUSE = &H13: VK_CAPITAL = &H14: VK_ESCAPE = &H1B
    VK_SPACE = &H20: VK_PRIOR = &H21: VK_NEXT = &H22: VK_END = &H23: VK_HOME = &H24
    VK_LEFT = &H25: VK_UP = &H26: VK_RIGHT = &H27: VK_DOWN = &H28: VK_SELECT = &H29: VK_PRINT = &H2A
    VK_EXECUTE = &H2B: VK_SNAPSHOT = &H2C: VK_INSERT = &H2D: VK_DELETE = &H2E: VK_HELP = &H2F
 
    ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
    ' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

    VK_NUMPAD0 = &H60: VK_NUMPAD1 = &H61: VK_NUMPAD2 = &H62: VK_NUMPAD3 = &H63: VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65: VK_NUMPAD6 = &H66: VK_NUMPAD7 = &H67: VK_NUMPAD8 = &H68: VK_NUMPAD9 = &H69
    VK_MULTIPLY = &H6A: VK_ADD = &H6B: VK_SEPARATOR = &H6C: VK_SUBTRACT = &H6D: VK_DECIMAL = &H6E: VK_DIVIDE = &H6F
    VK_F1 = &H70: VK_F2 = &H71: VK_F3 = &H72: VK_F4 = &H73: VK_F5 = &H74: VK_F6 = &H75: VK_F7 = &H76
    VK_F8 = &H77: VK_F9 = &H78: VK_F10 = &H79: VK_F11 = &H7A: VK_F12 = &H7B
    VK_F13 = &H7C: VK_F14 = &H7D: VK_F15 = &H7E: VK_F16 = &H7F: VK_F17 = &H80: VK_F18 = &H81
    VK_F19 = &H82: VK_F20 = &H83: VK_F21 = &H84: VK_F22 = &H85: VK_F23 = &H86: VK_F24 = &H87
    VK_NUMLOCK = &H90: VK_SCROLL = &H91
 
    '   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
    '   Used only as parameters to GetAsyncKeyState() and GetKeyState().
    '   No other API or message will distinguish left and right keys in this way.
    VK_LSHIFT = &HA0: VK_RSHIFT = &HA1: VK_LCONTROL = &HA2: VK_RCONTROL = &HA3: VK_LMENU = &HA4: VK_RMENU = &HA5
 
    VK_ATTN = &HF6: VK_CRSEL = &HF7: VK_EXSEL = &HF8: VK_EREOF = &HF9: VK_PLAY = &HFA
    VK_ZOOM = &HFB: VK_NONAME = &HFC: VK_PA1 = &HFD: VK_OEM_CLEAR = &HFE
End Enum
'==========================================================================================

Public Function KeyPressed(ByVal VKey As VirtualKeys) As Boolean
    KeyPressed = IIf(GetKeyState(VKey) < 0, True, False)
End Function
