Attribute VB_Name = "Module1"
Public hr As Long

Public Const WINDING = 2

Public Declare Function CreatePolygonRgn Lib "gdi32" ( _
    lpPoint As POINTAPI, _
    ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Sub ReleaseCapture Lib "user32" ()

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public bDown As Boolean

Public Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

Public Const pi As Double = 3.14159265358979

Public Const bEventID As Byte = 10

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public myPolygon() As POINTAPI

Public bShift As Double

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
    Call Initialize_Region
    bShift = bShift + 0.2
    If bShift > 12.6 Then bShift = 0.2
End Sub

Public Sub Initialize_Region()
    Dim i As Long
    Dim q As Long
    For i = 0 To frmMain.ScaleWidth * 2
        ReDim Preserve myPolygon(i)
        If i <= frmMain.ScaleWidth Then
            myPolygon(i).X = i * 16
            myPolygon(i).Y = frmMain.ScaleHeight / 16 + frmMain.ScaleHeight / 16 * Sin(bShift + i / 10 * pi)
        Else
            myPolygon(i).X = frmMain.ScaleWidth - (i - frmMain.ScaleWidth) * 16
            myPolygon(i).Y = frmMain.ScaleHeight - (frmMain.ScaleHeight / 16) - frmMain.ScaleHeight / 16 * Sin(bShift + i / 10 * pi)
        End If
    Next
    hr = CreatePolygonRgn(myPolygon(0), UBound(myPolygon) + 1, WINDING)
    Call SetWindowRgn(frmMain.hwnd, hr, True)
End Sub

Public Sub Main()
    bShift = 0.2
    Load frmMain
    Call Initialize_Region
    frmMain.Show
End Sub
