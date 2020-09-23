VERSION 5.00
Begin VB.Form frmTabControl 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command21 
      Caption         =   "set Item Size"
      Height          =   450
      Left            =   8535
      TabIndex        =   25
      Top             =   2040
      Width           =   1680
   End
   Begin VB.CommandButton Command20 
      Caption         =   "single line scroll"
      Height          =   315
      Left            =   7455
      TabIndex        =   24
      Top             =   6735
      Width           =   2595
   End
   Begin VB.CommandButton Command19 
      Caption         =   "multiline mode"
      Height          =   330
      Left            =   7440
      TabIndex        =   23
      Top             =   6270
      Width           =   2625
   End
   Begin VB.CommandButton Command18 
      Caption         =   "set tabs to top"
      Height          =   300
      Left            =   7500
      TabIndex        =   22
      Top             =   5100
      Width           =   2550
   End
   Begin VB.CommandButton Command17 
      Caption         =   "set tabs to horizontal"
      Height          =   315
      Left            =   7425
      TabIndex        =   21
      Top             =   5880
      Width           =   2625
   End
   Begin VB.CommandButton Command16 
      Caption         =   "set tabs to vertical"
      Height          =   300
      Left            =   7455
      TabIndex        =   20
      Top             =   5460
      Width           =   2625
   End
   Begin VB.CommandButton Command15 
      Caption         =   "set tabs to bottom"
      Height          =   390
      Left            =   7440
      TabIndex        =   19
      Top             =   4620
      Width           =   2550
   End
   Begin VB.CommandButton Command14 
      Caption         =   "set style tabs"
      Height          =   285
      Left            =   7335
      TabIndex        =   18
      Top             =   4275
      Width           =   2760
   End
   Begin VB.CommandButton Command13 
      Caption         =   "set Style flat buttons"
      Height          =   375
      Left            =   7290
      TabIndex        =   17
      Top             =   3810
      Width           =   2820
   End
   Begin VB.CommandButton Command9 
      Caption         =   "set style buttons"
      Height          =   390
      Left            =   7215
      TabIndex        =   16
      Top             =   3345
      Width           =   2910
   End
   Begin VB.CommandButton Command12 
      Caption         =   "del last tab"
      Height          =   405
      Left            =   9345
      TabIndex        =   15
      Top             =   795
      Width           =   915
   End
   Begin VB.CommandButton Command11 
      Caption         =   "get bottom tab"
      Height          =   390
      Left            =   7245
      TabIndex        =   14
      Top             =   2940
      Width           =   1470
   End
   Begin VB.CommandButton Command10 
      Caption         =   "getRowCount"
      Height          =   420
      Left            =   7245
      TabIndex        =   13
      Top             =   2475
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9690
      Top             =   1455
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   3150
      ScaleHeight     =   1560
      ScaleWidth      =   1860
      TabIndex        =   12
      Top             =   4515
      Width           =   1920
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Set Padding"
      Height          =   330
      Left            =   8475
      TabIndex        =   11
      Top             =   2520
      Width           =   1845
   End
   Begin VB.TextBox txtCY 
      Height          =   285
      Left            =   8940
      TabIndex        =   10
      Text            =   "16"
      Top             =   1725
      Width           =   645
   End
   Begin VB.TextBox txtCX 
      Height          =   345
      Left            =   8940
      TabIndex        =   8
      Text            =   "64"
      Top             =   1275
      Width           =   690
   End
   Begin VB.CommandButton Command8 
      Caption         =   "deselect all"
      Height          =   435
      Left            =   7080
      TabIndex        =   6
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton Command6 
      Caption         =   "select tab 2"
      Height          =   510
      Left            =   6945
      TabIndex        =   5
      Top             =   1305
      Width           =   1380
   End
   Begin VB.CommandButton Command5 
      Caption         =   "del selected tab"
      Height          =   465
      Left            =   7905
      TabIndex        =   4
      Top             =   765
      Width           =   1350
   End
   Begin VB.CommandButton Command4 
      Caption         =   "clear"
      Height          =   540
      Left            =   7065
      TabIndex        =   3
      Top             =   750
      Width           =   780
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Selected tab"
      Height          =   450
      Left            =   8820
      TabIndex        =   2
      Top             =   225
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add tab"
      Height          =   510
      Left            =   7860
      TabIndex        =   1
      Top             =   150
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tab count"
      Height          =   510
      Left            =   6975
      TabIndex        =   0
      Top             =   150
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "cy"
      Height          =   210
      Left            =   8460
      TabIndex        =   9
      Top             =   1635
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "cx"
      Height          =   300
      Left            =   8445
      TabIndex        =   7
      Top             =   1320
      Width           =   420
   End
End
Attribute VB_Name = "frmTabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' Name:     frmTabControl.frm
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
Dim PrevClickedItem As Long

Private Sub Command1_Click()
MsgBox GetCount
End Sub

Private Sub Command10_Click()
MsgBox GetRowCount
End Sub

Private Sub Command11_Click()
Dim bt As Long
bt = GetBottomTab
MsgBox GetText(bt), , bt
End Sub

Private Sub Command12_Click()
DelTab GetCount - 1
SelTab GetCount - 1
End Sub

Private Sub Command13_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_BUTTONS) = TCS_BUTTONS Then
Else
    stl = stl Or TCS_BUTTONS
End If
If (stl And TCS_FLATBUTTONS) = TCS_FLATBUTTONS Then
Else
    stl = stl Or TCS_FLATBUTTONS
End If
    SetWindowLong Wnd, GWL_STYLE, stl

End Sub

Private Sub Command14_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_BUTTONS) = TCS_BUTTONS Then
    stl = stl And Not TCS_BUTTONS
End If
If (stl And TCS_FLATBUTTONS) = TCS_FLATBUTTONS Then
    stl = stl And Not TCS_FLATBUTTONS
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command15_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_BOTTOM) = TCS_BOTTOM Then
Else
    stl = stl Or TCS_BOTTOM
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command16_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_VERTICAL) = TCS_VERTICAL Then
Else
    stl = stl Or TCS_VERTICAL
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command17_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_VERTICAL) = TCS_VERTICAL Then
    stl = stl And Not TCS_VERTICAL
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command18_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_BOTTOM) = TCS_BOTTOM Then
    stl = stl And Not TCS_BOTTOM
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command19_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_MULTILINE) = TCS_MULTILINE Then
Else
    stl = stl Or TCS_MULTILINE
End If
    SetWindowLong Wnd, GWL_STYLE, stl

End Sub

Private Sub Command2_Click()
'Add new tab
Dim c As Long
    c = GetCount
Dim str As String
    str = "Tab " & c
AddTab c, str
    SelTab c
End Sub


Private Sub Command20_Click()
'Changes the Style of Tabcontrol
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_MULTILINE) = TCS_MULTILINE Then
    stl = stl And Not TCS_MULTILINE
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command21_Click()
'Sets the ItemSize
Dim cx As Integer, cy As Integer
    cx = Val(txtCX)
    cy = Val(txtCY)
SetItemSIze cx, cy
    txtCX = cx
    txtCY = cy
End Sub

Private Sub Command3_Click()
Dim c As Long
    c = GetSelected
MsgBox GetText(c), , c
End Sub

Private Sub Command4_Click()
ClearTabs
End Sub

Private Sub Command5_Click()
DelTab GetSelected
End Sub

Private Sub Command6_Click()
SelTab 1
End Sub

Private Sub Command7_Click()
SetPadding Val(txtCX), Val(txtCY)
End Sub

Private Sub Command8_Click()
DeselectAll
End Sub

Private Sub Command9_Click()
Dim stl As Long
    stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And TCS_BUTTONS) = TCS_BUTTONS Then
Else
    stl = stl Or TCS_BUTTONS
End If
If (stl And TCS_FLATBUTTONS) = TCS_FLATBUTTONS Then
    stl = stl And Not TCS_FLATBUTTONS
End If
    SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Form_Load()
'Very simple way to create the Tabcontrol
CreateTabControl hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload the tabcontrol
DestroyTabControl
End Sub
Sub TabClicked(ByVal PrevTab As Long)
'The Main Event
'I did not included the subclassing.Because I found this way is very safe and useful.

    Dim tr As RECT, cR As RECT, stl As Long, RC As Long, i As Long, c As RECT, bt As Long
On Error Resume Next
    Picture1.BorderStyle = 1
        GetClientRect Wnd, cR
            i = GetSelected
    tr = GetTabRect(i)
'        cR.Top = cR.Top + tR.Bottom
        cR.Top = tr.Bottom
    stl = GetWindowLong(Wnd, GWL_STYLE)
    RC = GetRowCount
        If (stl And TCS_SCROLLOPPOSITE) = TCS_SCROLLOPPOSITE And RC > 1 Then
            cR.Bottom = cR.Bottom - tr.Bottom
'            If (tR.Top - tR.Bottom <= 0) Then 'And ((tR.Bottom / tR.Top) > 2) Then
'                    If tR.Top < tR.Bottom Then RC = 2
'''                    If tR.Top < tR.Bottom Then RC = 1
'''                For stl = 1 To RC
'''                    cR.Bottom = cR.Bottom - tR.Bottom
'''                Next
'                    cR.Bottom = cR.Bottom - (tR.Bottom * (RC - 1))
'            End If
                bt = GetBottomTab
            c = GetTabRect(bt)
                If bt = 0 Then
                    cR.Bottom = cR.Bottom - tr.Bottom
                Else
                   cR.Bottom = tr.Bottom - c.Bottom
                End If
        ElseIf (stl And TCS_BUTTONS) = TCS_BUTTONS Then
            cR.Bottom = cR.Bottom - tr.Bottom
                If RC > 1 Then
                    c = GetTabRect(GetCount - 1)
                    cR.Top = c.Bottom
                    cR.Bottom = cR.Bottom + tr.Bottom
                    cR.Bottom = cR.Bottom - c.Bottom
                End If
        ElseIf (stl And TCS_BOTTOM) = TCS_BOTTOM Then
            MoveWindow Picture1.hwnd, 0, 0, 0, 0, 1
            Exit Sub
        ElseIf (stl And TCS_VERTICAL) = TCS_VERTICAL Then
            MoveWindow Picture1.hwnd, 0, 0, 0, 0, 1
            Exit Sub
        Else
            cR.Bottom = cR.Bottom - tr.Bottom
                If RC > 1 Then
                    c = GetTabRect(GetCount - 1)
                    cR.Top = tr.Bottom
                End If
        End If
MoveWindow Picture1.hwnd, cR.Left + 5, cR.Top + 5, cR.Right - 5, cR.Bottom - 5, 1
    Picture1.Cls
        Picture1.Print GetText(i)

End Sub
Private Sub Timer1_Timer()
'Timer Used to TabEvent
Dim i As Long
    i = GetSelected
If PrevClickedItem <> i Then
    TabClicked PrevClickedItem
        PrevClickedItem = i
End If
End Sub
