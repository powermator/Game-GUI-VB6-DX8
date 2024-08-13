Attribute VB_Name = "calls"
'                        ***general declaration***
' SOME FUNCTIONS USED whith mouse

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'--------call to librarry that empty a text file
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type
