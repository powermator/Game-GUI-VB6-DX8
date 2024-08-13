Attribute VB_Name = "errmodule"
Public errmsg As String
Public errnumber As Byte

Public Sub err()
If errnumber = 1 Then
errmsg = "This name allready exist !! Please try another name..."
Back_Buffer.DrawText 230, 510, errmsg, False
End If
End Sub

