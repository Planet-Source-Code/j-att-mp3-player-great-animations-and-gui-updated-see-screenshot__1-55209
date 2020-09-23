Attribute VB_Name = "mod_GE_PAN"
Option Explicit
'==========================================
'=       yar interactive software         =
'=      2d game engine Version 1.4        =
'=      --------------------------        =
'= Polygon Animation file playback engine =
'==========================================
'== Call:                                        ==
'==                                              ==
'== LoadPAN - to load a Polygon Animation file   ==
'==      for playback using this game engine.    ==
'==                                              ==
'== UnloadPAN - To unload a Polygon Animation    ==
'==      from memory.                            ==
'== DrawPANFrame - To draw a specified frame of  ==
'== a loaded Polygon Animation to a HDC          ==
'==
'==================================================
'polygon drawing stuff
Public Type POINTAPI
    X                As Long
    Y                As Long
End Type
'use this to set the fillmode to 2 for optimal speed
'drawing options stuff
'dest rectangle
Private Type RECT
    x1               As Long
    y1               As Long
    x2               As Long
    y2               As Long
End Type
'PAN files
'Polygon data Structure
Private Type PolyShape
    PolyType         As Byte    '0 = polygon, 1 = rect, 2=line, 3=Ellips
    PolyPnt()        As POINTAPI
    PntCount         As Long    'if its a polygon, this is the count of points
    PolyColor        As Long
End Type
'Frame Data Structure
Private Type PolyFrame
    PolyShp()        As PolyShape 'multiple shapes/polygons
    PolyCount        As Byte    'number of shapes/polygons
End Type
'File Data Structure
Public Type polyPAN
    Polys()          As PolyFrame
    OutLineColor     As Long    'what color is everything outlined in?
    FrameCount       As Long
End Type
Private sprite     As Long
''Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
                                              lpPoint As POINTAPI, _
                                              ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal x1 As Long, _
                                                ByVal y1 As Long, _
                                                ByVal x2 As Long, _
                                                ByVal y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, _
                                              ByVal x1 As Long, _
                                              ByVal y1 As Long, _
                                              ByVal x2 As Long, _
                                              ByVal y2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
''Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
''Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
''Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Sub DrawPANFrame(FrameZ As Long, _
                        Animation As polyPAN, _
                        lngHdc As Long, _
                        ByVal Pic As Boolean)

  'On Error Resume Next
  
  Dim I As Long
  Dim P As POINTAPI

    If FrameZ > Animation.FrameCount Then
        Exit Sub
    End If
    On Error GoTo Hell
    'draw color...
    'we use invisible pens to make things faster now so we got rid of this...
    'DeleteObject SelectObject(Hdc, CreatePen(0, 1, Animation.OutLineColor))
    If Pic Then
        BitBlt lngHdc, 0, 0, 282, 150, sprite, 0, 0, vbSrcCopy
     Else 'PIC = FALSE/0
        DeleteObject SelectObject(lngHdc, CreateSolidBrush(vbBlue))
        Rectangle lngHdc, 0, 0, 95, 190
    End If
    'for each shape in this frame
    With Animation
        For I = 1 To .Polys(FrameZ).PolyCount
            'select color for this polygon...
            DeleteObject SelectObject(lngHdc, CreateSolidBrush(.Polys(FrameZ).PolyShp(I).PolyColor))
            'what type are we drawing...
            Select Case .Polys(FrameZ).PolyShp(I).PolyType
             Case Is = 0 'polygon
                Polygon lngHdc, .Polys(FrameZ).PolyShp(I).PolyPnt(1), .Polys(FrameZ).PolyShp(I).PntCount
             Case Is = 1 'rect
                Rectangle lngHdc, .Polys(FrameZ).PolyShp(I).PolyPnt(1).X, .Polys(FrameZ).PolyShp(I).PolyPnt(1).Y, .Polys(FrameZ).PolyShp(I).PolyPnt(2).X, .Polys(FrameZ).PolyShp(I).PolyPnt(2).Y
             Case Is = 2 'line
                MoveToEx lngHdc, .Polys(FrameZ).PolyShp(I).PolyPnt(1).X, .Polys(FrameZ).PolyShp(I).PolyPnt(1).Y, P
                LineTo lngHdc, .Polys(FrameZ).PolyShp(I).PolyPnt(2).X, .Polys(FrameZ).PolyShp(I).PolyPnt(2).Y
             Case Is = 3 'ellipse
                Ellipse lngHdc, .Polys(FrameZ).PolyShp(I).PolyPnt(1).X, .Polys(FrameZ).PolyShp(I).PolyPnt(1).Y, .Polys(FrameZ).PolyShp(I).PolyPnt(2).X, .Polys(FrameZ).PolyShp(I).PolyPnt(2).Y
            End Select
        Next '  I
    End With 'Animation

Exit Sub

Hell:
    MsgBox "An error occured, e-mail albert@yarsoft.com with the following:" & vbNewLine & _
           vbNewLine & _
           "Error drawing frame '" & FrameZ & "', FrameCount: " & Animation.FrameCount & vbNewLine & _
           vbNewLine & _
           "The program will attempt to keep running.", vbExclamation, "Error..."
    On Error GoTo 0

End Sub

Private Function LoadGraphicDC(sFilename As String) As Long

  'On Error Resume Next
  
  Dim LoadGraphicDCTEMP As Long

    LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
    SelectObject LoadGraphicDCTEMP, LoadPicture(sFilename)
    LoadGraphicDC = LoadGraphicDCTEMP
    On Error GoTo 0

End Function

Public Sub LoadPAN(ByVal strFilename As String, _
                   zPolyInfo As polyPAN, _
                   Optional bgpic As String)

  
  Dim IntBinaryFile As Long

    'Loads Polygon animation file into memory
    On Error GoTo errOut
    IntBinaryFile = FreeFile
    Open strFilename For Binary Access Read Lock Write As IntBinaryFile
    'Extract the data
    Get IntBinaryFile, 1, zPolyInfo
    Close IntBinaryFile
    If Len(bgpic) > 0 Then
        sprite = LoadGraphicDC(bgpic)
    End If

Exit Sub

errOut:
    MsgBox "Loading Polygon-Movie File Faild:" & vbNewLine & _
           "  The following error occured when trying to load this Polygon-Movie file:" & vbNewLine & _
           Err.Description, vbExclamation, "Polygon Movie Load ERROR"

End Sub

Public Sub UnloadPAN(panStruct As polyPAN)

  Dim A As Long

    'Dim C As Long
  Dim I As Long
    'call this subroutine to erase all memory occupied by
    'a polyPAN data/type-structure.
    With panStruct
        For I = 1 To .FrameCount
            For A = 1 To .Polys(I).PolyCount
                'clear point data
                Erase .Polys(I).PolyShp(A).PolyPnt
            Next '  A
            'clear shape data
            Erase .Polys(I).PolyShp
        Next '  I
    End With 'panStruct
    'clear frame data array from memory
    Erase panStruct.Polys

End Sub
