Attribute VB_Name = "modTabControl"
' ======================================================================================
' Name:     modTabControl.bas
' Author:   Joshy Francis (joshylogicss@yahoo.co.in)
' Date:     14 May 2007
'
' Requires: None
'
' Copyright © 2000-2007 Joshy Francis
' --------------------------------------------------------------------------------------
'The implementation of TabControl in VB.All by API.
'you can freely use this code anywhere.But I wants you must include the copyright info
'All functions in this module written by me.
' --------------------------------------------------------------------------------------
'No updates.This is the first version.
'I Just included comments on every important lines.Sorry for my bad english.
'I developed this program by converting the C Documentation to VB and experiments with VB.
'You can improve this program by your experiments.I didn't done all parts of the
'TabControl.

Option Explicit

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_CLASSDC = &H40
Public Const CS_DBLCLKS = &H8
Public Const CS_HREDRAW = &H2
Public Const CS_INSERTCHAR = &H2000
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_NOCLOSE = &H200
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOMOVECARET = &H4000
Public Const CS_OWNDC = &H20
Public Const CS_PARENTDC = &H80
Public Const CS_PUBLICCLASS = &H4000
Public Const CS_SAVEBITS = &H800
Public Const CS_VREDRAW = &H1
'Public Declare Function INITCOMMONCONTROLSEX Lib "COMCTL32.DLL" Alias "InitCommonControlsEx" (ByVal hInstance As Long) As Long 'Boolean

Public Const WM_USER = &H400
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const CCM_FIRST = &H2000                    ' Common control shared messages
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)

'====== TAB CONTROL ==========================================================
Public Const TCM_FIRST = &H1300                    ' Tab control messages
Public Const WC_TABCONTROLA = "SysTabControl32"

' begin_r_commctrl

'#if (_WIN32_IE >= = &h0300)
Public Const TCS_SCROLLOPPOSITE = &H1           ' assumes multiline tab
Public Const TCS_BOTTOM = &H2
Public Const TCS_RIGHT = &H2
Public Const TCS_MULTISELECT = &H4             ' allow multi-select in button mode
'#End If
'#if (_WIN32_IE >= = &h0400)
Public Const TCS_FLATBUTTONS = &H8
'#End If
Public Const TCS_FORCEICONLEFT = &H10
Public Const TCS_FORCELABELLEFT = &H20
'#if (_WIN32_IE >= = &h0300)
Public Const TCS_HOTTRACK = &H40
Public Const TCS_VERTICAL = &H80
'#End If
Public Const TCS_TABS = &H0
Public Const TCS_BUTTONS = &H100
Public Const TCS_SINGLELINE = &H0
Public Const TCS_MULTILINE = &H200
Public Const TCS_RIGHTJUSTIFY = &H0
Public Const TCS_FIXEDWIDTH = &H400
Public Const TCS_RAGGEDRIGHT = &H800
Public Const TCS_FOCUSONBUTTONDOWN = &H1000
Public Const TCS_OWNERDRAWFIXED = &H2000
Public Const TCS_TOOLTIPS = &H4000
Public Const TCS_FOCUSNEVER = &H8000

'#if (_WIN32_IE >= = &h0400)
' EX styles for use with TCM_SETEXTENDEDSTYLE
Public Const TCS_EX_FLATSEPARATORS = &H1
Public Const TCS_EX_REGISTERDROP = &H2
'#End If

' end_r_commctrl

Public Const TCM_GETIMAGELIST = (TCM_FIRST + 2)
'Public Const TabCtrl_GetImageList(hwnd) \
'    (HIMAGELIST)SNDMSG((hwnd), TCM_GETIMAGELIST, 0, 0L)

Public Const TCM_SETIMAGELIST = (TCM_FIRST + 3)
'Public Const TabCtrl_SetImageList(hwnd, himl) \
'    (HIMAGELIST)SNDMSG((hwnd), TCM_SETIMAGELIST, 0, (LPARAM)(UINT)(HIMAGELIST)(himl))

Public Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)
'Public Const TabCtrl_GetItemCount(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETITEMCOUNT, 0, 0L)

Public Const TCIF_TEXT = &H1
Public Const TCIF_IMAGE = &H2
Public Const TCIF_RTLREADING = &H4
Public Const TCIF_PARAM = &H8
'#if (_WIN32_IE >= = &h0300)
Public Const TCIF_STATE = &H10

Public Const TCIS_BUTTONPRESSED = &H1
'#End If
'#if (_WIN32_IE >= = &h0400)
Public Const TCIS_HIGHLIGHTED = &H2
'#End If

'#if (_WIN32_IE >= = &h0300)
'Public Const TC_ITEMHEADERA = TCITEMHEADERA
'Public Const TC_ITEMHEADERW = TCITEMHEADERW
''#Else
'Public Const tagTCITEMHEADERA = TC_ITEMHEADERA
'Public Const TCITEMHEADERA = TC_ITEMHEADERA
'Public Const tagTCITEMHEADERW = TC_ITEMHEADERW
'Public Const TCITEMHEADERW = TC_ITEMHEADERW
'#End If
'Public Const TC_ITEMHEADER = TCITEMHEADER

Public Type TCITEMHEADERA 'tagTCITEMHEADERA
 mask As Long '   UINT mask;
 lpReserved1  As Long '   UINT lpReserved1;
lpReserved2    As Long '   UINT lpReserved2;
pszText As String '    LPSTR ;
cchTextMax   As Long '    int cchTextMax;
iImage   As Long '   int iImage;
'} TCITEMHEADERA, FAR *LPTCITEMHEADERA;
End Type
'Public Type tagTCITEMHEADERW
'    UINT mask;
'    UINT lpReserved1;
'    UINT lpReserved2;
'    LPWSTR pszText;
'    int cchTextMax;
'    int iImage;
'} TCITEMHEADERW, FAR *LPTCITEMHEADERW;
'
'#ifdef UNICODE
'Public Const  TCITEMHEADER          TCITEMHEADERW
'Public Const  LPTCITEMHEADER        LPTCITEMHEADERW
'#Else
'Public Const TCITEMHEADER = TCITEMHEADERA
'Public Const LPTCITEMHEADER = LPTCITEMHEADERA
'#End If

'#if (_WIN32_IE >= = &h0300)
'Public Const TC_ITEMA = TCITEMA
'Public Const TC_ITEMW = TCITEMW
''#Else
'Public Const tagTCITEMA = TC_ITEMA
'Public Const TCITEMA = TC_ITEMA
'Public Const tagTCITEMW = TC_ITEMW
'Public Const TCITEMW = TC_ITEMW
'#End If
'Public Const TC_ITEM = TCITEM

Public Type TCITEMA 'tagTCITEMA
mask As Long '    UINT mask;
'#if (_WIN32_IE >= = &h0300)
dwState   As Long '   DWORD dwState;
dwStateMask    As Long '   DWORD dwStateMask;
'#Else
lpReserved1   As Long '   UINT lpReserved1;
lpReserved2   As Long '   UINT lpReserved2;
'#End If
pszText As String '    LPSTR pszText;
cchTextMax  As Long '    int cchTextMax;
iImage  As Long '    int iImage;

 lParam    As Long '    LPARAM lParam;
'} TCITEMA, FAR *LPTCITEMA;
End Type
'Public Type tagTCITEMW
'{
'    UINT mask;
'#if (_WIN32_IE >= = &h0300)
'    DWORD dwState;
'    DWORD dwStateMask;
'#Else
'    UINT lpReserved1;
'    UINT lpReserved2;
'#End If
'    LPWSTR pszText;
'    int cchTextMax;
'    int iImage;
'
'    LPARAM lParam;
'} TCITEMW, FAR *LPTCITEMW;
'
'#ifdef UNICODE
'Public Const  TCITEM                 TCITEMW
'Public Const  LPTCITEM               LPTCITEMW
'#Else
'Public Const TCITEM = TCITEMA
'Public Const LPTCITEM = LPTCITEMA
'#End If

Public Const TCM_GETITEMA = (TCM_FIRST + 5)
Public Const TCM_GETITEMW = (TCM_FIRST + 60)

'#ifdef UNICODE
'Public Const TCM_GETITEM = TCM_GETITEMW
'#Else
Public Const TCM_GETITEM = TCM_GETITEMA
'#End If
'
'Public Const TabCtrl_GetItem(hwnd, iItem, pitem) \
'    (BOOL)SNDMSG((hwnd), TCM_GETITEM, (WPARAM)(int)iItem, (LPARAM)(TC_ITEM FAR*)(pitem))

Public Const TCM_SETITEMA = (TCM_FIRST + 6)
Public Const TCM_SETITEMW = (TCM_FIRST + 61)

'#ifdef UNICODE
'Public Const TCM_SETITEM = TCM_SETITEMW
'#Else
Public Const TCM_SETITEM = TCM_SETITEMA
'#End If

'Public Const TabCtrl_SetItem(hwnd, iItem, pitem) \
'    (BOOL)SNDMSG((hwnd), TCM_SETITEM, (WPARAM)(int)iItem, (LPARAM)(TC_ITEM FAR*)(pitem))

Public Const TCM_INSERTITEMA = (TCM_FIRST + 7)
Public Const TCM_INSERTITEMW = (TCM_FIRST + 62)

'#ifdef UNICODE
'Public Const TCM_INSERTITEM = TCM_INSERTITEMW
'#Else
Public Const TCM_INSERTITEM = TCM_INSERTITEMA
'#End If

'Public Const TabCtrl_InsertItem(hwnd, iItem, pitem)   \
'    (int)SNDMSG((hwnd), TCM_INSERTITEM, (WPARAM)(int)iItem, (LPARAM)(const TC_ITEM FAR*)(pitem))

Public Const TCM_DELETEITEM = (TCM_FIRST + 8)
'Public Const TabCtrl_DeleteItem(hwnd, i) \
'    (BOOL)SNDMSG((hwnd), TCM_DELETEITEM, (WPARAM)(int)(i), 0L)

Public Const TCM_DELETEALLITEMS = (TCM_FIRST + 9)
'Public Const TabCtrl_DeleteAllItems(hwnd) \
'    (BOOL)SNDMSG((hwnd), TCM_DELETEALLITEMS, 0, 0L)

Public Const TCM_GETITEMRECT = (TCM_FIRST + 10)
'Public Const TabCtrl_GetItemRect(hwnd, i, prc) \
'    (BOOL)SNDMSG((hwnd), TCM_GETITEMRECT, (WPARAM)(int)(i), (LPARAM)(RECT FAR*)(prc))

Public Const TCM_GETCURSEL = (TCM_FIRST + 11)
'Public Const TabCtrl_GetCurSel(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETCURSEL, 0, 0)

Public Const TCM_SETCURSEL = (TCM_FIRST + 12)
'Public Const TabCtrl_SetCurSel(hwnd, i) \
'    (int)SNDMSG((hwnd), TCM_SETCURSEL, (WPARAM)i, 0)

Public Const TCHT_NOWHERE = &H1
Public Const TCHT_ONITEMICON = &H2
Public Const TCHT_ONITEMLABEL = &H4
Public Const TCHT_ONITEM = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)

'#if (_WIN32_IE >= = &h0300)
'Public Const LPTC_HITTESTINFO = LPTCHITTESTINFO
'Public Const TC_HITTESTINFO = TCHITTESTINFO
''#Else
'Public Const tagTCHITTESTINFO = C_HITTESTINFO
'Public Const TCHITTESTINFO = TC_HITTESTINFO
'Public Const LPTCHITTESTINFO = LPTC_HITTESTINFO
'#End If

Public Type TCHITTESTINFO 'tagTCHITTESTINFO
pt As POINTAPI '    POINT pt;
 flags As Long '   UINT flags;
'} TCHITTESTINFO, FAR * LPTCHITTESTINFO;
End Type
Public Const TCM_HITTEST = (TCM_FIRST + 13)
'Public Const TabCtrl_HitTest(hwndTC, pinfo) \
'    (int)SNDMSG((hwndTC), TCM_HITTEST, 0, (LPARAM)(TC_HITTESTINFO FAR*)(pinfo))

Public Const TCM_SETITEMEXTRA = (TCM_FIRST + 14)
'Public Const TabCtrl_SetItemExtra(hwndTC, cb) \
'    (BOOL)SNDMSG((hwndTC), TCM_SETITEMEXTRA, (WPARAM)(cb), 0L)

Public Const TCM_ADJUSTRECT = (TCM_FIRST + 40)
'Public Const TabCtrl_AdjustRect(hwnd, bLarger, prc) \
'    (int)SNDMSG(hwnd, TCM_ADJUSTRECT, (WPARAM)(BOOL)bLarger, (LPARAM)(RECT FAR *)prc)

Public Const TCM_SETITEMSIZE = (TCM_FIRST + 41)
'Public Const TabCtrl_SetItemSize(hwnd, x, y) \
'    (DWORD)SNDMSG((hwnd), TCM_SETITEMSIZE, 0, MAKELPARAM(x,y))

Public Const TCM_REMOVEIMAGE = (TCM_FIRST + 42)
'Public Const TabCtrl_RemoveImage(hwnd, i) \
'        (void)SNDMSG((hwnd), TCM_REMOVEIMAGE, i, 0L)

Public Const TCM_SETPADDING = (TCM_FIRST + 43)
'Public Const TabCtrl_SetPadding(hwnd,  cx, cy) \
'        (void)SNDMSG((hwnd), TCM_SETPADDING, 0, MAKELPARAM(cx, cy))

Public Const TCM_GETROWCOUNT = (TCM_FIRST + 44)
'Public Const TabCtrl_GetRowCount(hwnd) \
'        (int)SNDMSG((hwnd), TCM_GETROWCOUNT, 0, 0L)
Public Const TCM_GETTOOLTIPS = (TCM_FIRST + 45)
'Public Const TabCtrl_GetToolTips(hwnd) \
'        (HWND)SNDMSG((hwnd), TCM_GETTOOLTIPS, 0, 0L)

Public Const TCM_SETTOOLTIPS = (TCM_FIRST + 46)
'Public Const TabCtrl_SetToolTips(hwnd, hwndTT) \
'        (void)SNDMSG((hwnd), TCM_SETTOOLTIPS, (WPARAM)hwndTT, 0L)

Public Const TCM_GETCURFOCUS = (TCM_FIRST + 47)
'Public Const TabCtrl_GetCurFocus(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETCURFOCUS, 0, 0)

Public Const TCM_SETCURFOCUS = (TCM_FIRST + 48)
'Public Const TabCtrl_SetCurFocus(hwnd, i) \
'    SNDMSG((hwnd),TCM_SETCURFOCUS, i, 0)

'#if (_WIN32_IE >= = &h0300)
Public Const TCM_SETMINTABWIDTH = (TCM_FIRST + 49)
'Public Const TabCtrl_SetMinTabWidth(hwnd, x) \
'        (int)SNDMSG((hwnd), TCM_SETMINTABWIDTH, 0, x)

Public Const TCM_DESELECTALL = (TCM_FIRST + 50)
'Public Const TabCtrl_DeselectAll(hwnd, fExcludeFocus)\
'        (void)SNDMSG((hwnd), TCM_DESELECTALL, fExcludeFocus, 0)
'#End If

'#if (_WIN32_IE >= = &h0400)

Public Const TCM_HIGHLIGHTITEM = (TCM_FIRST + 51)
'Public Const TabCtrl_HighlightItem(hwnd, i, fHighlight) \
'    (BOOL)SNDMSG((hwnd), TCM_HIGHLIGHTITEM, (WPARAM)i, (LPARAM)MAKELONG (fHighlight, 0))

Public Const TCM_SETEXTENDEDSTYLE = (TCM_FIRST + 52)    ' optional wParam == mask
'Public Const TabCtrl_SetExtendedStyle(hwnd, dw)\
'        (DWORD)SNDMSG((hwnd), TCM_SETEXTENDEDSTYLE, 0, dw)

Public Const TCM_GETEXTENDEDSTYLE = (TCM_FIRST + 53)
'Public Const TabCtrl_GetExtendedStyle(hwnd)\
'        (DWORD)SNDMSG((hwnd), TCM_GETEXTENDEDSTYLE, 0, 0)

Public Const TCM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
'Public Const TabCtrl_SetUnicodeFormat(hwnd, fUnicode)  \
'    (BOOL)SNDMSG((hwnd), TCM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)

Public Const TCM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
'Public Const TabCtrl_GetUnicodeFormat(hwnd)  \
'    (BOOL)SNDMSG((hwnd), TCM_GETUNICODEFORMAT, 0, 0)

'#End If     ' _WIN32_IE >= = &h0400

Public Const TCN_FIRST = 0
Public Const TCN_KEYDOWN = (TCN_FIRST - 0)

'#if (_WIN32_IE >= = &h0300)
'Public Const TC_KEYDOWN = NMTCKEYDOWN
''#Else
'Public Const tagTCKEYDOWN = TC_KEYDOWN
'Public Const NMTCKEYDOWN = TC_KEYDOWN
''#End If

Public Type NMTCKEYDOWN 'tagTCKEYDOWN
hdr As Long '    NMHDR hdr;
wVKey    As Long '   WORD wVKey;
flags    As Long '  UINT flags;
'} NMTCKEYDOWN;
End Type
Public Const TCN_SELCHANGE = (TCN_FIRST - 1)
Public Const TCN_SELCHANGING = (TCN_FIRST - 2)
'#if (_WIN32_IE >= = &h0400)
Public Const TCN_GETOBJECT = (TCN_FIRST - 3)
'#End If     ' _WIN32_IE >= = &h0400
'
'#End If     ' NOTABCONTROL
Public Const ICC_TAB_CLASSES = &H8
Public Enum ComCtlClasses
     ICC_LISTVIEW_CLASSES = &H1      ' listview, header
     ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
     ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
'     ICC_TAB_CLASSES = &H8            ' tab, tooltips
     ICC_UPDOWN_CLASS = &H10          ' updown
     ICC_PROGRESS_CLASS = &H20        ' progress
     ICC_HOTKEY_CLASS = &H40          ' hotkey
     ICC_ANIMATE_CLASS = &H80         ' animate
     ICC_WIN95_CLASSES = &HFF        '
     ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
     ICC_USEREX_CLASSES = &H200       ' comboex
     ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
    #If (WIN32_IE >= &H400) Then    '
         ICC_INTERNET_CLASSES &H800
         ICC_PAGESCROLLER_CLASS &H1000       ' page scroller
         ICC_NATIVEFNTCTL_CLASS &H2000       ' native font control
    #End If
End Enum
Public Type INITCOMMONCONTROLSEX
    dwSize As Long 'DWORD ;             // size of this structure
    dwICC As ComCtlClasses 'Long 'DWORD ;              // flags indicating which classes to be initialized
End Type '} INITCOMMONCONTROLSEX, *LPINITCOMMONCONTROLSEX;

Public Declare Function INITCOMMONCONTROLSEX Lib "COMCTL32.DLL" Alias "InitCommonControlsEx" (ICCClass As INITCOMMONCONTROLSEX) As Long 'Boolean

Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Wnd As Long 'Ony one Global variable for all operations

Function CreateTabControl(ByVal hwnd As Long) As Long
'This is the main function.It Creates a tabcontrol window and returns the window handle
'You can modify this function by including Coordinate Parameters,style etc.

Dim stl As Long
    Dim IX As INITCOMMONCONTROLSEX, Inited  As Long
        IX.dwICC = ICC_TAB_CLASSES
        IX.dwSize = Len(IX)
Inited = INITCOMMONCONTROLSEX(IX)
'If CommonControl library is not initialized the program does'nt work.
        If Inited <> 1 Then
            MsgBox "INITCOMMONCONTROLSEX Failed.", vbCritical
        End If
    stl = WS_CHILD Or WS_VISIBLE 'Or WS_THICKFRAME
'Here i gives the different styles for the tabcontrol

'stl = stl Or TCS_BOTTOM 'Tabs appear in bottom
'stl = stl Or TCS_BUTTONS 'Button Style tabs
'stl = stl Or TCS_FIXEDWIDTH 'Tabs have fixed width
stl = stl Or TCS_FLATBUTTONS 'FlatButton style
'stl = stl Or TCS_FORCELABELLEFT '
stl = stl Or TCS_HOTTRACK 'hottracking
stl = stl Or TCS_MULTILINE 'Multi-rows
stl = stl Or TCS_MULTISELECT 'Multiselction
stl = stl Or TCS_TABS 'Style Tabs
'stl = stl Or TCS_VERTICAL 'Vertical
'stl = stl Or TCS_RIGHTJUSTIFY
stl = stl Or TCS_TOOLTIPS ' You must create the tooltipcontrol
'stl = stl Or TCS_FOCUSONBUTTONDOWN 'I didn't deeply tested it
'stl = stl Or TCS_SCROLLOPPOSITE 'explorerbar style scrolling of tabs
'stl = stl Or TCS_SINGLELINE ' Adds  scrolling buttons

'Creates the tabcontrolwindow.
'It is very fast and safe function.In VB the dynamic control creation is not possible when the controls are kept in DLLS that not compatible with VB.
'If the control not created no error will occur the return value will be zero else return value will be handle.
'This is the style of C & CPP programs. The concept of pointers is done in VB by this way.

    Wnd = CreateWindowEx(TCS_EX_FLATSEPARATORS, WC_TABCONTROLA, ByVal vbNullString, stl, 0, 0, 400, 400, hwnd, 0, App.hInstance, ByVal 0)
'    Wnd = CreateWindowEx(0, WC_TABCONTROLA, ByVal vbNullString, stl, 0, 0, 400, 400, hwnd, 0, App.hInstance, ByVal 0)

If Wnd = 0 Then
' Window is not created . Load Window class from the DLL. You can test this line by commenting the above line 'Inited = INITCOMMONCONTROLSEX(IX)...'
'The below technique is very useful to crteate other controls used by other Programs.

        Dim Ret As Long, Class As WNDCLASS
    With Class
        .cbClsextra = Len(Class)
        .hInstance = LoadLibrary("COMCTL32.DLL")
        .lpszClassName = WC_TABCONTROLA
        .style = CS_PUBLICCLASS
    End With
    Ret = RegisterClass(Class)
        If Ret = 0 Then
            'If Register class failed exit
            Exit Function
        End If
    Wnd = CreateWindowEx(0, WC_TABCONTROLA, ByVal vbNullString, stl, 100, 100, 100, 100, hwnd, 0, App.hInstance, ByVal 0)
End If
If Wnd <> 0 Then
    CreateTabControl = Wnd
'    SetParent Wnd, hwnd
'    ShowWindow Wnd, 4
        'An Imagelist is created for put the tab images quickly.
            hIml = ImageList_Create(32, 32, ILC_COLOR Or ILD_TRANSPARENT, 1, 0)
        
        'Adds an image to the imagelist
            ImageList_AddIcon hIml, , frmTabControl.Icon.Handle
        'Set the Tabcontrol's imagelist
                SendMessage Wnd, TCM_SETIMAGELIST, 0, ByVal hIml
        'Adds some sample tabs
        Call AddTab(0, "Sample")
        Call AddTab(1, "Sample2")
End If
End Function

Function DestroyTabControl() As Long
If Wnd <> 0 Then
'Destroy the tabcontrol and free the memory
    DestroyWindow Wnd
        Wnd = 0
End If
End Function
Function AddTab(ByVal i As Long, ByVal str As String, Optional ByVal img As Long = 0) As Long
If Wnd = 0 Then Exit Function
Dim tb As TCITEMA, Ret As Long
'First Send TCITEM structure to the tabcontrol.
Ret = SendMessage(Wnd, TCM_INSERTITEMA, i, tb)
If Ret <> -1 Then
'If New Tab is added
'    With tb
'        .cchTextMax = Len(str)
'        .mask = TCIF_TEXT   '.dwStateMask
'        .pszText = str   'VarPtr(str)
'    End With
'   AddTab = SendMessage(Wnd, TCM_SETITEMA, Ret, tb)
'Change the new tab's properties such as caption & image
    Dim tcih As TCITEMHEADERA
        With tcih
            .cchTextMax = Len(str)
            .pszText = str
            .mask = TCIF_TEXT Or TCIF_IMAGE
            .iImage = img
        End With
'Set the new properties to the new tab
   AddTab = SendMessage(Wnd, TCM_SETITEMA, Ret, tcih)
    
End If
End Function
Function GetText(ByVal i As Long) As String
If Wnd = 0 Then Exit Function
Dim Ret As Long
'returns the caption of the tab by index
    Dim tcih As TCITEMHEADERA
        With tcih
            .cchTextMax = 260
            .pszText = Space$(260)
            .mask = TCIF_TEXT Or TCIF_IMAGE
        End With
    Ret = SendMessage(Wnd, TCM_GETITEMA, i, tcih)
GetText = Trim0(tcih.pszText)
End Function

Function GetCount() As Long
'returns the count of tabs+1
    GetCount = SendMessage(Wnd, TCM_GETITEMCOUNT, 0, ByVal 0)
End Function
Function GetRowCount() As Long
'returns count of rows
    GetRowCount = SendMessage(Wnd, TCM_GETROWCOUNT, 0, ByVal 0)
End Function

Function GetSelected() As Long
'returns the selected tab index
    GetSelected = SendMessage(Wnd, TCM_GETCURSEL, 0, ByVal 0)
End Function
Function GetFocused() As Long
'returns the focused tab index
    GetFocused = SendMessage(Wnd, TCM_GETCURFOCUS, 0, ByVal 0)
End Function
Function DelTab(ByVal i As Long) As Long
'deletes a tab by index
    DelTab = SendMessage(Wnd, TCM_DELETEITEM, i, ByVal 0)
End Function
Function ClearTabs() As Long
'clear all tabs
    ClearTabs = SendMessage(Wnd, TCM_DELETEALLITEMS, 0, ByVal 0)
End Function
Function SelTab(ByVal i As Long, Optional ByVal SetCurFocus As Boolean = True) As Long
'select a tab by index
    SelTab = SendMessage(Wnd, TCM_SETCURSEL, i, ByVal 0)
If SetCurFocus = True Then
    'set focus on the selected tab
'    Dim c As Long
'                c = GetSelected
'            If i = c Then c = 0
'            If i = c Then c = GetCount - 1
'                SendMessage Wnd, TCM_SETCURSEL, c, ByVal 0
'                 or
        DeselectAll
    SelTab = SendMessage(Wnd, TCM_SETCURSEL, i, ByVal 0)
    SendMessage Wnd, TCM_SETCURFOCUS, i, ByVal 0
End If
End Function
Function DeselectAll(Optional ByVal bExcludeFocus As Boolean = False) As Long
Dim i As Long
'Deselects all tabs option to exclude current focusing tab
If bExcludeFocus = True Then
    i = 1
Else
    i = 0
End If
    DeselectAll = SendMessage(Wnd, TCM_DESELECTALL, i, ByVal 0)
End Function

Function GetTabRect(ByVal i As Long) As RECT
'returns the tab coordinates in pixels by index
Dim R As RECT, c As Long
    c = SendMessage(Wnd, TCM_GETITEMRECT, i, R)
GetTabRect = R
End Function

Function AdjustRect(ByVal fLarger As Long, R As RECT) As Long
'fLarger
'Operation to perform. If this parameter is TRUE, prc specifies a display rectangle and receives the corresponding window rectangle. If this parameter is FALSE, prc specifies a window rectangle and receives the corresponding display area.
AdjustRect = SendMessage(Wnd, TCM_ADJUSTRECT, fLarger, R)
End Function
Function GetBottomTab() As Long
'this function is by me.
'I wrote this for identify the bottom tab when tabcontrol32 has the new XP style(TCS_SCROLLOPPOSITE).
'I Commented the style because this function is not accurate in number of tabs is very high.
'But the style gives the explorerbar style tot the tabcontrol

Dim i As Long, R As RECT, c As Long, br As Long, bt As Long, br2 As Long, tr As Long
Dim fr As Long, ft As Long
c = GetCount
'    For i = 0 To c - 1
            R = GetTabRect(i)
'            If br < R.Bottom Then
                br = R.Bottom
                bt = i
'            End If
'    Next
            R = GetTabRect(c - 1)
                tr = R.Bottom
        br2 = br
            br = 0
    For i = c - 1 To 0 Step -1
            R = GetTabRect(i)
'            If R.Bottom > tr And br = 0 And R.Bottom < br2 Then
            If R.Bottom > tr And R.Bottom < br2 Then
                    If fr = 0 Then
                        fr = R.Bottom
                        ft = i
                    End If
                br = R.Bottom
                tr = tr + br
                bt = i
'                    Exit For
            End If
    Next
            R = GetTabRect(c - 1)
                tr = R.Bottom
If ((br2 - br) * 2) = (br2 - tr) Then
    GetBottomTab = 0
Else
        Dim Ret As Long
            Ret = ((br2 - br) - (br2 - tr))
    If (Ret < tr And Ret > 1) Or (tr Mod Abs(Ret) < tr) Then
        GetBottomTab = 0
    Else
        If ((br - fr)) < fr Then
            GetBottomTab = ft
        Else
            GetBottomTab = bt
        End If
    End If
End If
End Function
Function SetItemSIze(cx As Integer, cy As Integer) As Long
'sets the tabs' Common ItemSize
Dim PrevSize As Long, nSize As Long
    nSize = MAKELPARAM(cx, cy)
'        If PrevSize <> 0 Then
'            nSize = PrevSize
'        End If
'Sets the amount of space (padding) around each tab's icon and label in a tab control.
PrevSize = SendMessage(Wnd, TCM_SETITEMSIZE, 0, ByVal nSize)
    cx = LoWord(PrevSize)
    cy = HiWord(PrevSize)
SetItemSIze = PrevSize
End Function
Function MAKELPARAM(ByVal cx As Integer, ByVal cy As Integer) As Long
'I Converted the this function by my experiments
'Print &HFFFF&
'65535
'Print &H8000&
'32768
'Print &H10000
'65536
'Print &HFFFF0000
'-65536
    MAKELPARAM = (cy) * &H10000 + cx
End Function
Function SetPadding(ByVal cx As Integer, ByVal cy As Integer) As Long
'Sets the amount of space (padding) around each tab's icon and label in a tab control.
SetPadding = SendMessage(Wnd, TCM_SETPADDING, 0, ByVal MAKELPARAM(cx, cy))
End Function
Function SetTabState(ByVal i As Long, Optional ByVal nState As Long = TCIS_HIGHLIGHTED) As Long
'Highlights,Presses a tab by index
Dim tb As TCITEMA
    With tb
        .mask = TCIF_STATE  '.dwStateMask
        .dwState = nState 'TCIS_HIGHLIGHTED ' TCIS_BUTTONPRESSED 'Or TCIS_HIGHLIGHTED 'bState
    End With
SetTabState = SendMessage(Wnd, TCM_SETITEMW, i, tb)
End Function

