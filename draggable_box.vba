Rem Types
Public Type POINT
    X As Long
    Y As Long
End Type

Rem Imports
Public Declare PtrSafe Sub GetCursorPos Lib "User32.dll" (ByRef lpPoint As POINT)
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Rem Constants
Public Const VK_LBUTTON = 1

Rem Functions
Rem Deprecated
Public Function GetTagNameAsIndex(name As String, sh As PowerPoint.shape) As Long
    Dim i As Long
    For i = 1 To sh.Tags.Count
        If StrComp(sh.Tags.name(i), name, vbBinaryCompare) = 0 Then
            GetTagNameAsIndex = i
            Exit Function
        End If
    Next i
    
    GetTagNameAsIndex = -1
End Function

Rem Entry Point
Sub DraggableObjectEntry(sh As PowerPoint.shape)
    Dim p As POINT
                
            Dim SlideHeight As Long
            Dim SlideWidth As Long
                
            Dim MonitorWidth As Long
            Dim MonitorHeight As Long
                
            Dim HorizontalRatio As Double
            Dim VerticalRatio As Double
                
            Dim offsetX As Double
            Dim offsetY As Double
                
            SlideHeight = ActivePresentation.PageSetup.SlideHeight
            SlideWidth = ActivePresentation.PageSetup.SlideWidth
                
            MonitorWidth = GetSystemMetrics(0)
            MonitorHeight = GetSystemMetrics(1)
                
            HorizontalRatio = SlideWidth / MonitorWidth
            VerticalRatio = SlideHeight / MonitorHeight
                
            offsetX = sh.Width / 2
            offsetY = sh.Height / 2
    
            Do While GetAsyncKeyState(VK_LBUTTON) >= 0
                GetCursorPos p
                
                sh.Left = p.X * HorizontalRatio - offsetX
                sh.Top = p.Y * VerticalRatio - offsetY
                
                DoEvents
            Loop
End Sub
