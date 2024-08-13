Attribute VB_Name = "variables"
Public start       As Byte
Public namee       As Byte
Public show1       As Byte
Public mesh1posx   As Single
Public flag        As Byte

Public botonposx   As Single
Public botonflag   As Byte
Public finalquit   As Byte
Public transvalue  As Single
Public transvalue2 As Single
Public transvalue3 As Single
Public transvalue4 As Single
Public transvalue5 As Single


Public num2        As Byte
Public num         As Byte
Public msg         As String
Public pn          As String
Public phrase      As String

Public curspeed    As Byte
Public sit         As Byte
Public b           As POINTAPI
Public oknamebord  As Byte

Public pname       As String
Public playername  As String
Public quitbord    As Byte
Public quit        As Byte
Public nameflag    As Byte

Public noquit      As Byte

Public creditbord  As Byte
Public credit      As Byte

Public coulor      As Long

Public profile     As Byte
Public probord     As Byte
Public proADD      As Byte


Public t(20)       As String
Public ypos(21)    As Long
Public time        As Long


Public upbord     As Byte
Public downbord   As Byte

Public tempplayer As playerpro  'hold any playername inthe profile menue

Public textposinscroll  As Integer
Public scrollstart As Integer
Public scrollflag  As Byte


Public prouse_bord  As Byte

Public options As Byte
Public optionbord As Byte

Public saveandquitbord As Byte
Public videobord As Byte
Public video As Byte

Public title As String
'#######################################
Type currentsetting
mode As Byte
resolution As Integer
coulor As Byte
texture As Byte
effect  As Byte
Light  As Byte
gamma As Integer
fps   As Byte
soundvolume  As Integer
musicvolume  As Integer
sondeffects  As Byte

mousespeed   As Integer
mousescroll  As Integer
mousesensitivity  As Integer
mouseinverseleftright  As Byte
mouseinverseupdown  As Byte
End Type
'######################################
Public aa As currentsetting
Public currentstring As String

Public audiobord    As Byte
Public audio        As Byte

Public controlsbord  As Byte
Public CONTROLLS        As Byte

