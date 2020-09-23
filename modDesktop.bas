Attribute VB_Name = "modDesktop"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''NOT MY CODE - I FORGET WHO WROTE IT.'''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Public DispRec As RECT
Public DeskBmp As BITMAP ' Bitmap copy of the desktop
Public DeskDC As Long    ' Desktop device context handle
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
Public Sub InitDeskDC(OutHdc As Long, OutBmp As BITMAP, DispRec As RECT)
Dim DskHwnd As Long     ' hWnd of desktop
Dim DskRect As RECT     ' rect size of desktop
Dim DskHdc As Long      ' hdc of desktop
Dim hOutBmp As Long     ' handle to output bitmap
DskHwnd = GetDesktopWindow() ' Get src - HWND of Desktop
DskHdc = GetWindowDC(DskHwnd) ' Get src HDC - Handle to device context
Call GetWindowRect(DskHwnd, DskRect) ' Get src Rectangle dimentions
With DispRec
   ' Create handle to compatible output bitmap
   hOutBmp = CreateCompatibleBitmap(DskHdc, (.Right - .Left + 1), (.Bottom - .Top + 1))
   Call GetObject(hOutBmp, Len(OutBmp), OutBmp) ' Get handle to bitmap
   OutHdc = CreateCompatibleDC(DskHdc) ' Create compatible hdc
   Call SelectObject(OutHdc, hOutBmp) ' Copy bitmap structure into output dc
   Call StretchBlt(OutHdc, 0, 0, _
        (.Right - .Left + 1), _
        (.Bottom - .Top + 1), _
        DskHdc, 0, 0, _
        (DskRect.Right - DskRect.Left + 1), _
        (DskRect.Bottom - DskRect.Top + 1), _
        vbSrcCopy) ' Paint bitmap desk dc to output dc
End With
Call DeleteObject(hOutBmp)      ' Delete handle to output bitmap
Call ReleaseDC(DskHwnd, DskHdc) ' Clean up - Release src HDC
End Sub
