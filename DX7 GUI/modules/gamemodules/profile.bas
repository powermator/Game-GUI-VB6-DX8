Attribute VB_Name = "profilemodule"

'PROFILE.DAT STRUCTURE


'''''''''''''''''''''''''''''''''''''''''''''''''
'                                0 OR TEST      '1
'''''''''''''''''''''''''''''''''''''''''''''''''
'           DEFAULT PLAYER              ||    0 '2
'''''''''''''''''''''''''''''''''''''''''''''''''
'                                       ||      '3
'''''''''''''''''''''''''''''''''''''''''''''''''
'                                       ||      '4
'''''''''''''''''''''''''''''''''''''''''''''''''
'                                               '
'                                               '
'                                               '
'                                               '
'                                               '
'                                               '
'                                               '


Type playerpro
number      As String
playername  As String
End Type


Public playernow As playerpro

Public test        As Byte



Public Sub checkprofilename()

Open App.Path & "\profiles\profile.dat" For Random As 1
Get #1, 1, test


'-------------
If test = 0 Then
Get #1, 2, playernow
start = 0: namee = 1
End If

If Not test = 0 Then
Get #1, test + 1, playernow
start = 1: namee = 0
End If

End Sub

Public Sub writeprofilename(name As String) 'write a profile message-------************

phrase = name + " profile loaded ... "

If mesh1posx = -6 Then
num2 = num2 + 1

If num2 <= Len(phrase) Then msg = Mid$(phrase, 1, num2)

If num2 < Len(name) Then Back_Buffer.DrawText 100, 650, msg + "_", False
If num2 >= Len(name) Then Back_Buffer.DrawText 100, 650, msg, False
If num2 > 200 Then num2 = 200
End If

                                            '*****************


End Sub
