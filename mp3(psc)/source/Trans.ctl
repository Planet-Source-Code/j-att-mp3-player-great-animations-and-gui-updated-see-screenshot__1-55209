VERSION 5.00
Begin VB.UserControl Trans 
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ScaleHeight     =   90
   ScaleWidth      =   90
End
Attribute VB_Name = "Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LW_KEY     As Long = &H1
Private Const G_E        As Long = (-20)
Private Const W_E        As Long = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
''Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
''Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, _
                                                                      ByVal crKey As Long, _
                                                                      ByVal bAlpha As Byte, _
                                                                      ByVal dwFlags As Long) As Long

Public Sub remap()

  
  Dim Ret As Long

    Ret = GetWindowLong(UserControl.Parent.hwnd, G_E)
    Ret = Ret Or W_E
    With UserControl
        SetWindowLong .Parent.hwnd, G_E, Ret
        SetLayeredWindowAttributes .Parent.hwnd, vbBlue, 0, LW_KEY
        .BackColor = .Parent.BackColor
    End With 'UserControl

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next
    UserControl.BackColor = UserControl.Parent.BackColor
    On Error GoTo 0

End Sub
