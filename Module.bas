Attribute VB_Name = "Module"
Public RainDrop(1 To 1000) As RAIN

Public Type RAIN
    XVal As Integer
    YVal As Integer
End Type

Public Fishy(1 To 25) As FISH

Public Type FISH
    XVal As Integer
    YVal As Integer
    Type As Integer
    DescentSpeed As Integer
    XSpeed As Integer
End Type

Public Const SRCAND As Long = &H8800C6
Public Const SRCPAINT  As Long = &HEE0086
Public Const SRCCOPY  As Long = &HCC0020
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
   ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
   ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
   ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Sub Main()
If App.PrevInstance Then End
If InStr(Command, "/c") <> 0 Then
    frmAbout.Show
ElseIf InStr(Command, "/p") <> 0 Then
    'Preview - Don't do anything
Else
    frmSaver.Show
End If
End Sub
