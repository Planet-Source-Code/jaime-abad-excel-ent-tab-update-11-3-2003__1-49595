Attribute VB_Name = "mdlTab"
'JAIME ABAD 1/11/2003

Option Explicit

Public Const Bottom = 1
Public Const Top = 0
Public Const WINDING = 2
Public Const BSIDE = 10
Const GWL_WNDPROC = (-4)
Const WM_MOUSEWHEEL = &H20A
'Const WM_KILLFOCUS = 8
'Const WM_SETFOCUS = &H7

Type PointApi
   X As Long
   Y As Long
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private SetedProced As Long
Private lngHwnd As Long
Private clsMouseEvent As MouseEvent

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointApi, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Declare Function CreatePolyPolygonRgn& Lib "gdi32" (lpPoint As PointApi, ByVal nCount As Long, ByVal nPolyfillMode As Long, lpPolyCount As Long)
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As PointApi, ByVal nCount As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As PointApi) As Long

Public Function ReadMsg(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_MOUSEWHEEL
            If wParam > 0 Then
               clsMouseEvent.Up
            Else
               clsMouseEvent.Down
            End If
    End Select
    ReadMsg = CallWindowProc(SetedProced, hWnd, Msg, wParam, lParam)
End Function

Public Sub SetTheProc(lhwnd As Variant, clsMouse As MouseEvent)
    Set clsMouseEvent = clsMouse
    lngHwnd = lhwnd
    SetedProced = GetWindowLong(lngHwnd, GWL_WNDPROC)
    SetWindowLong lngHwnd, GWL_WNDPROC, AddressOf ReadMsg
End Sub

Public Sub LostTheProc()
    SetWindowLong lngHwnd, GWL_WNDPROC, SetedProced
End Sub
