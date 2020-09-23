VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmOpen2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbExtension 
      Height          =   315
      ItemData        =   "frmOpen2.frx":000C
      Left            =   1560
      List            =   "frmOpen2.frx":0016
      TabIndex        =   8
      Text            =   "Mp3 Files (*.mp3)"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   3120
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5040
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   480
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4524
      _Version        =   393217
      Indentation     =   397
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      HotTracking     =   -1  'True
      ImageList       =   "TreeImageListeDir"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList TreeImageListeDir 
      Left            =   6840
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen2.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen2.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen2.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen2.frx":0B06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image command1 
      Height          =   240
      Left            =   5940
      Picture         =   "frmOpen2.frx":0EA0
      Top             =   150
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Files of type:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "File name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Look in:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Extension          As String
Private Path               As String
Private CFm_Filename       As String
Private CFm_SetName        As Boolean
Private CFm_ShortName      As String

Public Property Get Filename() As String


    Filename = CFm_Filename

End Property

Public Property Let Filename(ByVal PropVal As String)


    CFm_Filename = PropVal

End Property

Public Property Get SetNames() As Boolean


    SetNames = CFm_Filename

End Property

Public Property Let SetNames(ByVal PropVal As Boolean)


    CFm_SetName = PropVal

End Property

Public Property Get ShortName() As String


    ShortName = CFm_ShortName

End Property

Public Property Let ShortName(ByVal PropVal As String)


    CFm_ShortName = PropVal

End Property

Private Sub AddFile(Optional ByVal FolderName As String)

  
  Dim MyPath As String
  Dim MyName As String
  Dim MyFile As String

    On Error Resume Next
    MyPath = FolderName
    If Right$(MyPath, 1) <> "\" Then
        MyPath = MyPath & "\"
    End If
    MyName = Dir(MyPath, vbDirectory)
    Do While LenB(MyName)
        If MyName <> "." Then
            If MyName <> ".." Then
                If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                    TreeView1.Nodes.Add MyPath, tvwChild, MyPath & MyName & "\", MyName, 1, 2
                End If
            End If
        End If
        MyName = Dir
    Loop
    MyFile = Dir(MyPath & "*.*")
    Do While LenB(MyFile)
        If LenB(Extension) Then
            If Right$(MyFile, 3) = Extension Then
                TreeView1.Nodes.Add MyPath, tvwChild, "|" & MyPath & MyFile, MyFile, 3
            End If
         Else 'LENB(EXTENSION) = FALSE/0
            TreeView1.Nodes.Add MyPath, tvwChild, "|" & MyPath & MyFile, MyFile, 3
        End If
        MyFile = Dir
    Loop
    On Error GoTo 0

End Sub

Private Sub cmbExtension_click()

    If cmbExtension.ListIndex = 1 Then
        Extension = ""
     Else 'NOT CMBEXTENSION.LISTINDEX...
        Extension = "mp3"
    End If
    Call Drive1_Change
End Sub

Private Sub cmdCancel_Click()

    CFm_Filename = ""
    Unload Me

End Sub

Private Sub cmdOpen_Click()
    
    If Me.Caption = "Save" Then
        CFm_Filename = Path & txtFilename.Text
    End If
    Unload Me

End Sub

Private Sub Command1_Click()

    With TreeView1
        If Len(.SelectedItem.Key) > 3 Then
        .SelectedItem = .SelectedItem.Parent
        .SelectedItem.Expanded = False
        .SetFocus
        End If
    End With 'TreeView1

End Sub

Private Sub Drive1_Change()

  Dim DriveFullPath As String

    TreeView1.Nodes.Clear
    If Right$(Drive1.Drive, 1) <> "\" Then
        DriveFullPath = Drive1.Drive & "\"
     Else 'NOT RIGHT$(DRIVE1.DRIVE,...
        DriveFullPath = Drive1.Drive
    End If
    TreeView1.Nodes.Add , , DriveFullPath, "My Computer (" & DriveFullPath & ")", 4
    AddFile DriveFullPath

End Sub

Private Sub Form_Load()

    Extension = "mp3"
    'TreeView1.Nodes.Add , , "C:\", "My Computer (C:\)", 4
    'AddFile "C:\"
    Call Drive1_Change

End Sub

Private Sub TreeView1_NodeClick(ByVal node As MSComctlLib.node)

  Dim CurrentNodeKey As String

    CurrentNodeKey = node.Key
    If Left$(CurrentNodeKey, 1) <> "|" Then
        AddFile CurrentNodeKey
        Path = CurrentNodeKey
     Else 'NOT LEFT$(CURRENTNODEKEY,...
        txtFilename.Text = node.Text
        If CFm_SetName Then CFm_ShortName = node.Text
        CFm_Filename = Right$(CurrentNodeKey, Len(CurrentNodeKey) - 1)
    End If

End Sub
