VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5670
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
'Hi , this program was made by jaberani@yahoo.com                  =
'***PLEASE VOTE FOR ME ********************************************=
'feel free to change or edit code or whatever U want :-)           =
'                                                                  =
'==========DIRECTX 7 D3DRM        ======================29/8/2004===

Private Sub Form_Click()
If finalquit = 1 Then Set D3d = Nothing: Set Dd = Nothing: Set dx = Nothing: End
errnumber = 0
If quitbord = 1 And start = 1 Then quit = 1: start = 0

If oknamebord = 1 And namee = 1 Or proADD = 1 And profile = 1 Then '<----------------------'-
nameflag = 1                                                                                '!

For i = 0 To Len("                              ")                                          '!
If pn = Mid$("                        ", 1, i) Then                                         '!
nameflag = 0                                                                                '!
End If                                                                                      '!
Next i
If pn = "" Then nameflag = 0
                                                              
                                                                                            '!
If nameflag = 1 Then      '[[[[you typ a name
'|||||||||||| this procedure to check whether the name is allready exist                    '!
For i = 2 To test + 1

Get #1, i, tempplayer
If pn = tempplayer.playername Then errnumber = 1                                            '!


Next i

'||||||||||||

If errnumber = 0 Then                                                                      '!
test = test + 1                                                                            '!
playernow.number = test
playernow.playername = pn

Put #1, 1, test: Put #1, test + 1, playernow                                               '!
End If




End If                                                                                     '!


If namee = 1 And start = 0 Then start = 1: namee = 0
                                                                                           '!
                                                                                           '@
End If                                                '<----------------------------------@@@



If noquit = 1 Then quit = 0: start = 1
If creditbord = 1 And start = 1 Then start = 0: credit = 1: time = 0: setcredittextpos

If probord = 1 Then probord = 0: profile = 1: start = 0


textposinscroll = scrollstart + 2 * 11
If textposinscroll <= 320 Then
If profile = 1 And upbord = 1 Then scrollstart = scrollstart + 15
End If

If scrollflag = 1 And profile = 1 And downbord = 1 Then scrollstart = scrollstart - 15


If profile = 1 And prouse_bord = 1 And start = 0 Then start = 1: prouse_bord = 0: num2 = 0: profile = 0

If optionbord = 1 And start = 1 Then options = 1: start = 0: title = "OPTION MENUE ..."

''==== when quit from option menue*********

If options = 1 And saveandquitbord = 1 And start = 0 Then
start = 1: options = 0: saveandquitbord = 0: video = 0: audio = 0: CONTROLLS = 0
Put #2, 15, aa
End If

''========                        *****************


'=================for video menue
If video = 1 And b.x > 236 And b.x < 258 And b.y > 212 And b.y < 233 Or b.x > 449 And b.x < 468 And b.y > 212 And b.y < 233 Then aa.mode = aa.mode + 1
If video = 1 And b.x > 236 And b.x < 258 And b.y > 243 And b.y < 263 Or b.x > 449 And b.x < 468 And b.y > 243 And b.y < 263 Then aa.resolution = aa.resolution + 1
If video = 1 And b.x > 236 And b.x < 258 And b.y > 272 And b.y < 293 Or b.x > 449 And b.x < 468 And b.y > 272 And b.y < 293 Then aa.coulor = aa.coulor + 1
If video = 1 And b.x > 236 And b.x < 258 And b.y > 302 And b.y < 322 Or b.x > 449 And b.x < 468 And b.y > 302 And b.y < 322 Then aa.texture = aa.texture + 1
If video = 1 And b.x > 236 And b.x < 258 And b.y > 331 And b.y < 353 Or b.x > 449 And b.x < 468 And b.y > 331 And b.y < 353 Then aa.effect = aa.effect + 1

If video = 1 And b.x > 236 And b.x < 258 And b.y > 362 And b.y < 383 Then aa.gamma = aa.gamma - 1
If video = 1 And b.x > 449 And b.x < 468 And b.y > 362 And b.y < 383 Then aa.gamma = aa.gamma + 1

If video = 1 And b.x > 236 And b.x < 258 And b.y > 391 And b.y < 414 Then aa.fps = aa.fps - 1
If video = 1 And b.x > 449 And b.x < 468 And b.y > 391 And b.y < 414 Then aa.fps = aa.fps + 1

If video = 1 And b.x > 236 And b.x < 258 And b.y > 421 And b.y < 444 Or b.x > 449 And b.x < 468 And b.y > 421 And b.y < 444 Then aa.Light = aa.Light + 1

If audio = 1 And b.x > 165 And b.x < 187 And b.y > 163 And b.y < 184 Or b.x > 329 And b.x < 349 And b.y > 163 And b.y < 184 Then aa.sondeffects = aa.sondeffects + 1

If audio = 1 And b.x > 165 And b.x < 187 And b.y > 101 And b.y < 124 Then aa.soundvolume = aa.soundvolume - 1

If audio = 1 And b.x > 329 And b.x < 349 And b.y > 101 And b.y < 124 Then aa.soundvolume = aa.soundvolume + 1

If audio = 1 And b.x > 165 And b.x < 187 And b.y > 133 And b.y < 153 Then aa.musicvolume = aa.musicvolume - 1
If audio = 1 And b.x > 329 And b.x < 349 And b.y > 133 And b.y < 153 Then aa.musicvolume = aa.musicvolume + 1

If CONTROLLS = 1 And b.x > 211 And b.x < 225 And b.y > 54 And b.y < 76 Then aa.mousespeed = aa.mousespeed - 1
If CONTROLLS = 1 And b.x > 362 And b.x < 378 And b.y > 54 And b.y < 76 Then aa.mousespeed = aa.mousespeed + 1

If CONTROLLS = 1 And b.x > 211 And b.x < 225 And b.y > 94 And b.y < 115 Then aa.mousescroll = aa.mousescroll - 1
If CONTROLLS = 1 And b.x > 362 And b.x < 378 And b.y > 94 And b.y < 115 Then aa.mousescroll = aa.mousescroll + 1

If CONTROLLS = 1 And b.x > 211 And b.x < 225 And b.y > 133 And b.y < 153 Then aa.mousesensitivity = aa.mousesensitivity - 1
If CONTROLLS = 1 And b.x > 362 And b.x < 378 And b.y > 133 And b.y < 153 Then aa.mousesensitivity = aa.mousesensitivity + 1

If CONTROLLS = 1 And b.x > 281 And b.x < 297 And b.y > 172 And b.y < 192 Or b.x > 357 And b.x < 374 And b.y > 172 And b.y < 192 Then aa.mouseinverseupdown = aa.mouseinverseupdown + 1

If CONTROLLS = 1 And b.x > 281 And b.x < 297 And b.y > 210 And b.y < 230 Or b.x > 357 And b.x < 374 And b.y > 210 And b.y < 230 Then aa.mouseinverseleftright = aa.mouseinverseleftright + 1



If aa.mode > 2 Then aa.mode = 1
If aa.resolution > 4 Then aa.resolution = 3
If aa.coulor > 6 Then aa.coulor = 5
If aa.texture > 8 Then aa.texture = 7
If aa.effect > 10 Then aa.effect = 9
If aa.gamma < 265 Then aa.gamma = 265
If aa.gamma > 441 Then aa.gamma = 441
If aa.fps < 12 Then aa.fps = 12
If aa.fps > 32 Then aa.fps = 32
If aa.Light > 14 Then aa.Light = 13
If aa.sondeffects > 12 Then aa.sondeffects = 11
If aa.soundvolume > 322 Then aa.soundvolume = 322
If aa.soundvolume < 193 Then aa.soundvolume = 193

If aa.musicvolume > 322 Then aa.musicvolume = 322
If aa.musicvolume < 193 Then aa.musicvolume = 193

If aa.mousespeed < 235 Then aa.mousespeed = 235
If aa.mousespeed > 355 Then aa.mousespeed = 355

If aa.mousescroll < 235 Then aa.mousescroll = 235
If aa.mousescroll > 355 Then aa.mousescroll = 355

If aa.mousesensitivity < 235 Then aa.mousesensitivity = 235
If aa.mousesensitivity > 355 Then aa.mousesensitivity = 355

If aa.mouseinverseupdown > 17 Then aa.mouseinverseupdown = 16
If aa.mouseinverseleftright > 19 Then aa.mouseinverseleftright = 18
''===============

If options = 1 And audiobord = 1 Then audio = 1: audiobord = 0: video = 0: CONTROLLS = 0
If options = 1 And videobord = 1 Then audio = 0: video = 1: CONTROLLS = 0

If options = 1 And controlsbord = 1 Then CONTROLLS = 1: audio = 0: video = 0


End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If namee = 1 Or profile = 1 And start = 0 Then

If Not pn = "" Then If KeyCode = 8 Then pn = Mid$(pn, 1, Len(pn) - 1)
If Len(pn) <= 23 Then
If KeyCode >= 65 And KeyCode <= 105 Then pn = pn + Chr$(KeyCode)
If KeyCode = 32 Then pn = pn + " "
End If

End If
End Sub
Private Sub Form_Load()
errnumber = 0    ' I MEAN THER IS NO ERROR
checkprofilename

'''''

Open App.Path & "\setting\setting.dat" For Random As #2
Get #2, 15, aa


''''




'-------- switching some variables
noquit = 0
finalquit = 0
nameflag = 1
show1 = 0
mesh1posx = -10
botonposx = 1

transvalue = 0.5  ' for namee menue only
transvalue2 = 0   'for start menue  only
transvalue3 = 0   'for quit menue   only
transvalue4 = 0   'for profilemenue only
transvalue5 = 0   'for optionmenue  only

sit = 0
pn = ""
curspeed = 0
quit = 0
creditbord = 0

time = 0
profile = 0
options = 0
optionbord = 0
audio = 0
CONTROLLS = 0

scrollstart = 300 ' related to profilemenue that specify text location



'----------text position

t(1) = "               -DX7 GUI-                "
t(2) = "  © 2005 - 2006  All rights reserved    "
t(3) = "   Programing By : Jaber Moyyad         "
t(4) = "                                        "
t(5) = "           jaberani@yahoo.com           "
t(6) = "                                        "
t(7) = "   Dont forget to vote for This Code :-)"
t(8) = "                                        "
t(9) = "   feel free to change or edit code     "
t(10) = "  or whatever U want                   "
t(11) = " I know this GUI hase some programing & designing        "
t(12) = " errors  but it is for learning purpose only             "
t(13) = "  A Special Thanks to:                 "
t(14) = "  PlanetSourceCode.com                 "
t(15) = "            Maxforums.net              "
t(16) = "            PClord                     "
t(17) = "            Fabio calvi                "
t(18) = "            Simon Price                "
t(19) = "            Egy_Tiger                  "
t(20) = "            Enjoy This Game !!!        "

time = 0
setcredittextpos


'----------------------------------------------dd
Set Dd = dx.DirectDraw4Create("")
Dd.SetCooperativeLevel Me.hWnd, DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
Dd.SetDisplayMode 1024, 768, 16, 0, DDSDM_DEFAULT
Show
Ddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
Ddsd.ddsCaps.lCaps = DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
Ddsd.lBackBufferCount = 2

Set Primary = Dd.CreateSurface(Ddsd)
Ddsd.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
Set Back_Buffer = Primary.GetAttachedSurface(Ddsd.ddsCaps)

'--------------------------------------------d3d
Set D3d = dx.Direct3DRMCreate

Set Device = D3d.CreateDeviceFromSurface("IID_IDirect3DHALDevice", Dd, Back_Buffer, D3DRMDEVICE_DEFAULT)
Device.SetBufferCount 2
Device.SetQuality D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD
Device.SetTextureQuality D3DRMTEXTURE_NEAREST
Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY

Set WorldFrame = D3d.CreateFrame(Nothing)
Set CameraFrame = D3d.CreateFrame(WorldFrame)
Set LightFrame = D3d.CreateFrame(WorldFrame)

Set Light = D3d.CreateLightRGB(D3DRMLIGHT_AMBIENT, 2, 2, 2)
Set Shadow = D3d.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
LightFrame.AddLight Light
LightFrame.AddLight Shadow

Set ViewPort = D3d.CreateViewport(Device, CameraFrame, 0, 0, 1024, 768)



'-------  the line
Set MeshFramex = D3d.CreateFrame(WorldFrame)
Set Meshx = D3d.CreateMeshBuilder
Meshx.LoadFromFile App.Path & "\xfiles\line.x", 0, 0, Nothing, Nothing
Meshx.ScaleMesh 1, 1, 1
MeshFramex.AddVisual Meshx
MeshFramex.SetPosition Nothing, 0, 0, 20
'MeshFramex.SetRotation Nothing, 0, 0, 1, 0.1

'=====================================TEXTURES
Set tex = D3d.LoadTexture(App.Path & "\pic\start\usa.bmp")
Set tex2 = D3d.LoadTexture(App.Path & "\pic\start\b.bmp")
Set tex3 = D3d.LoadTexture(App.Path & "\pic\buttom\1.bmp")
Set tex4 = D3d.LoadTexture(App.Path & "\pic\buttom\3.bmp")

'****************
'****************
For i = 1 To 21

Set MeshFrame(i) = D3d.CreateFrame(WorldFrame)
Set Mesh(i) = D3d.CreateMeshBuilder

Set dxf(i) = D3d.CreateFace

dxf(i).AddVertex -50, -50, 0
dxf(i).AddVertex -50, 50, 0
dxf(i).AddVertex 50, 50, 0
dxf(i).AddVertex 50, -50, 0

dxf(i).SetTextureCoordinates 0, 0, 0
dxf(i).SetTextureCoordinates 1, 1, 0
dxf(i).SetTextureCoordinates 2, 1, 1
dxf(i).SetTextureCoordinates 3, 0, 1

dxf(i).SetTexture tex3

Mesh(i).AddFace dxf(i)
MeshFrame(i).AddVisual Mesh(i)
Mesh(i).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0.5)
Next i

'************************
dxf(1).SetTexture tex2
dxf(20).SetTexture tex2
dxf(3).SetTexture tex4
 


'=============================================the small rect in the bottom
Mesh(1).ScaleMesh 0.2, 0.025, 0.05
'==============================================namee menue==========================================
'ok and name box of
For j = 2 To 3
Mesh(j).ScaleMesh 0.05, 0.009, 0.05
MeshFrame(j).SetPosition Nothing, 0.4, -1.8, 20
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0.5)
Next j
'==========================================for start menue bottons====================================
For j = 4 To 8
Mesh(j).ScaleMesh 0.05, 0.009, 0.05
MeshFrame(j).SetPosition Nothing, 0.4, 6.5 - j, 20
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0)
Mesh(12).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0) ' profile botton
Next j

Mesh(12).ScaleMesh 0.04, 0.01, 1
MeshFrame(12).SetPosition Nothing, -8, -4.3, 20 'the back
Mesh(12).SetTexture tex3

'=========================================yes and no of quit menue
Mesh(9).ScaleMesh 0.025, 0.009, 0.05
Mesh(10).ScaleMesh 0.025, 0.009, 0.05
Mesh(20).ScaleMesh 0.1, 0.06, 1 'the back of quit menue and nameemenue


'=============================================profile menue =============================================
Mesh(11).ScaleMesh 0.2, 0.06, 1
MeshFrame(11).SetPosition Nothing, 0, 0, 20     'the back
Mesh(11).SetTexture tex2
Mesh(11).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0)

Mesh(13).ScaleMesh 0.025, 0.009, 0.05
MeshFrame(13).SetPosition Nothing, -7, 0, 20     'use
Mesh(13).SetTexture tex2
Mesh(13).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0)

Mesh(14).ScaleMesh 0.025, 0.009, 0.05
MeshFrame(14).SetPosition Nothing, -7, -1, 20     'del
Mesh(14).SetTexture tex2
Mesh(14).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0)

Mesh(15).ScaleMesh 0.025, 0.009, 0.05
MeshFrame(15).SetPosition Nothing, -7, -2, 20     'add
Mesh(15).SetTexture tex2
Mesh(15).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0)


'============================================option menuue

For j = 16 To 19
Mesh(j).ScaleMesh 0.05, 0.009, 0.05
Next j

Mesh(21).ScaleMesh 0.2, 0.12, 1
Mesh(21).SetTexture tex2


MeshFrame(16).SetPosition Nothing, -6.5, -5.4, 20
MeshFrame(17).SetPosition Nothing, 0, -5.4, 20
MeshFrame(18).SetPosition Nothing, 6.5, -5.4, 20
MeshFrame(19).SetPosition Nothing, 0, -6.5, 20

'==the background
MeshFrame(21).SetPosition Nothing, 0, 1.4, 20


'+++++++++++++++++++++++++++++++++++++++++++++

WorldFrame.SetSceneBackgroundImage tex
'-----font
Back_Buffer.SetForeColor vbWhite
Back_Buffer.SetFont Me.Font
'-------
'CreateGamma
Do
WorldFrame.Move 0.05
ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER


'----
If start = 1 Then startmenue
If namee = 1 Then namemenue
If quit = 1 Then quitmenue
If credit = 1 Then creditmenue
If profile = 1 Then profilemenue
If errnumber = 1 Then err
If options = 1 Then optionmenue
If video = 1 Then videomenue
If audio = 1 Then audiomenue
If CONTROLLS = 1 Then controlsmenue
''''''''''''

GetCursorPos b

Back_Buffer.SetForeColor vbWhite

Back_Buffer.DrawText 900, 200, gn, False

''''''''''''
Device.Update
ViewPort.Render WorldFrame
Primary.Flip Nothing, DDFLIP_WAIT
DoEvents
Loop
End Sub


Private Sub startmenue()
noquit = 0
quit = 0
quitbord = 0
creditbord = 0
probord = 0
optionbord = 0


phrase = playernow.playername + " profile loaded ... "

Back_Buffer.SetForeColor vbWhite

transvalue = transvalue - 0.09
If transvalue < 0 Then transvalue = 0

goup
sleepoptionbottons 0.02

Mesh(2).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue)
Mesh(3).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue)
Mesh(20).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue)

If transvalue = 0 Then 'when namee menue is gone start the start!!!
wakeupfive
End If

If transvalue2 = 0.5 Then
If b.x > 400 And b.x < 660 And b.y > 236 And b.y < 279 Then Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawText 465, 250, " singleplay", False
Back_Buffer.SetForeColor vbWhite

If b.x > 400 And b.x < 660 And b.y > 287 And b.y < 330 Then Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawText 465, 300, " multiplay ", False
Back_Buffer.SetForeColor vbWhite

If b.x > 400 And b.x < 660 And b.y > 337 And b.y < 387 Then Back_Buffer.SetForeColor vbBlack: optionbord = 1
Back_Buffer.DrawText 465, 350, " options   ", False
Back_Buffer.SetForeColor vbWhite

If b.x > 400 And b.x < 660 And b.y > 390 And b.y < 433 Then Back_Buffer.SetForeColor vbBlack: creditbord = 1
Back_Buffer.DrawText 465, 400, " credits   ", False
Back_Buffer.SetForeColor vbWhite


If b.x > 400 And b.x < 660 And b.y > 439 And b.y < 483 Then Back_Buffer.SetForeColor vbBlack: quitbord = 1
Back_Buffer.DrawText 465, 450, " quit      ", False
Back_Buffer.SetForeColor vbWhite


If b.x > 0 And b.x < 204 And b.y > 580 And b.y < 628 Then Back_Buffer.SetForeColor vbBlack: probord = 1
Back_Buffer.DrawText 50, 595, " profile   ", False
Back_Buffer.SetForeColor vbWhite


writeprofilename playernow.playername



End If


'  when back from quit menue bottons
sleepquitmenue

sleepprofilemenue

End Sub


Private Sub namemenue()
sleepoptionbottons 0.02
oknamebord = 0
sit = sit + 1
If sit > 2 Then sit = 1
If sit = 2 Then num = num + 1

curspeed = curspeed + 1
If curspeed = 30 Then
coulor = vbWhite
End If
If curspeed = 60 Then
coulor = vbRed
curspeed = 0
End If

Back_Buffer.SetForeColor coulor
Back_Buffer.DrawText 410, 350, pn + "_", False

Back_Buffer.SetForeColor vbWhite
Back_Buffer.DrawText 410, 350, pn, False

If num <= Len("Type your name please ... ") Then msg = Mid$("Type your name please ... ", 1, num)


If num < Len("Type your name please ... ") Then Back_Buffer.DrawText 300, 255, msg + "_", False
If num >= Len("Type your name please ... ") Then Back_Buffer.DrawText 300, 255, msg, False


If b.x > 400 And b.x < 660 And b.y > 453 And b.y < 498 Then Back_Buffer.SetForeColor vbBlack: oknamebord = 1

MeshFrame(20).SetPosition Nothing, 0.4, 0, 20 'the back
MeshFrame(2).SetPosition Nothing, 0.4, -1.8, 20 'ok
MeshFrame(3).SetPosition Nothing, 0.4, 0.5, 20 'namebox

Back_Buffer.DrawText 520, 470, "OK ", False

''''
If num > 200 Then num = 200
End Sub

Private Sub quitmenue()

Back_Buffer.SetForeColor vbWhite
finalquit = 0
godown
sleepfive 0.02
wakeupquitmenue
noquit = 0
If mesh1posx <= -9 Then
If b.x > 470 And b.x < 595 And b.y > 413 And b.y < 457 Then Back_Buffer.SetForeColor vbbrown: Back_Buffer.DrawText 520, 427, "NO", False: noquit = 1
If b.x > 470 And b.x < 595 And b.y > 465 And b.y < 508 Then Back_Buffer.SetForeColor vbbrown: Back_Buffer.DrawText 517, 477, "Yes", False: finalquit = 1
End If
End Sub

Private Sub creditmenue()


sleepfive 0.01
godown

For i = 1 To 20
ypos(i) = ypos(i) - 1

If ypos(i) > 400 Then
Back_Buffer.DrawText 300, ypos(i), t(i), False
End If
Next i

If ypos(20) < 300 Then
credit = 0: start = 1: setcredittextpos
End If


End Sub

Private Sub wakeupfive()

transvalue2 = transvalue2 + 0.01
If transvalue2 > 0.5 Then transvalue2 = 0.5
For j = 4 To 8
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue2)
Mesh(12).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue2) ' profile botton
Next j

End Sub

Private Sub sleepfive(x As Single)

transvalue2 = transvalue2 - x
If transvalue2 < 0 Then transvalue2 = 0
For j = 4 To 8
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue2)
Mesh(12).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue2) ' profile botton

Next j

End Sub

Private Sub goup()

mesh1posx = mesh1posx + 0.1
If mesh1posx > -6 Then mesh1posx = -6
MeshFrame(1).SetPosition Nothing, 0, mesh1posx, 20

End Sub
Private Sub godown()

mesh1posx = mesh1posx - 0.05
If mesh1posx < -10 Then mesh1posx = -10
MeshFrame(1).SetPosition Nothing, 0, mesh1posx, 20
End Sub

Private Sub wakeupquitmenue()

MeshFrame(20).SetPosition Nothing, 0.4, 0, 20 'the back


transvalue3 = transvalue3 + 0.01
If transvalue3 > 0.5 Then transvalue3 = 0.5

Mesh(20).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)
MeshFrame(9).SetPosition Nothing, 0.4, -2, 20
Mesh(9).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)
MeshFrame(10).SetPosition Nothing, 0.4, -1, 20
Mesh(10).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)

If transvalue3 = 0.5 Then
Back_Buffer.DrawText 310, 240, "Quit Game ...", False
Back_Buffer.DrawText 430, 340, "Are You Sure ?", False
Back_Buffer.DrawText 520, 427, "NO", False
Back_Buffer.DrawText 517, 477, "Yes", False
End If

End Sub

Private Sub sleepquitmenue()

transvalue3 = transvalue3 - 0.01
If transvalue3 < 0 Then transvalue3 = 0


Mesh(9).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)
Mesh(10).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)
Mesh(20).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue3)

End Sub

Private Sub profilemenue()

curspeed = curspeed + 1
If curspeed = 30 Then
coulor = vbWhite
End If
If curspeed = 60 Then
coulor = vbRed
curspeed = 0
End If

proADD = 0
prouse_bord = 0


Back_Buffer.SetForeColor vbWhite
sleepfive 0.01

wakeupprofilemenue

If transvalue4 = 0.5 Then

Back_Buffer.SetForeColor vbWhite

Back_Buffer.DrawText 20, 250, "Profile Menue", False

Back_Buffer.SetForeColor vbWhite
If b.x > 92 And b.x < 216 And b.y > 362 And b.y < 405 Then Back_Buffer.SetForeColor vbBlack: proADD = 1
Back_Buffer.DrawText 135, 375, "Add", False

Back_Buffer.SetForeColor vbWhite
If b.x > 92 And b.x < 216 And b.y > 414 And b.y < 457 Then Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawText 130, 428, "Delet", False


Back_Buffer.SetForeColor vbWhite
If b.x > 92 And b.x < 216 And b.y > 465 And b.y < 508 Then Back_Buffer.SetForeColor vbBlack: prouse_bord = 1
Back_Buffer.DrawText 135, 477, "Use", False


Mesh(3).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, 0.5)
MeshFrame(3).SetPosition Nothing, -5.7, 1.3, 20 'namebox

Back_Buffer.SetForeColor coulor
Back_Buffer.DrawText 102, 315, pn + "_", False

Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawText 102, 315, pn, False

Back_Buffer.DrawRoundedBox 403, 303, 797, 507, 100, 100
Back_Buffer.DrawRoundedBox 400, 300, 800, 510, 100, 100

Back_Buffer.DrawLine 450, 303, 450, 507
Back_Buffer.DrawLine 451, 303, 451, 507

Back_Buffer.DrawBox 450, 320, 700, 490

drawtriangle 430, 350, 1
drawtriangle 430, 460, -1


']]]]]]]]]]
For i = 2 To test + 1

Get #1, i, tempplayer
scrollflag = 0

textposinscroll = scrollstart + i * 15

Back_Buffer.SetForeColor vbWhite
If Not textposinscroll >= 475 And Not textposinscroll <= 320 Then Back_Buffer.DrawText 460, textposinscroll, tempplayer.playername, False


Next i
']]]]]]]]]]]]

If textposinscroll >= 475 Then scrollflag = 1   '\\\it just check if we need to use scroll bar for down

End If



writeprofilename "ffff"
drawtextlines


End Sub

Private Sub drawtriangle(x As Integer, y As Integer, direction As Integer)

Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawBox x - 10, 340, x + 10, 470 'outbox

upbord = 0
downbord = 0

If b.x > x - 10 And b.x < x + 10 And b.y > 340 And b.y < 350 Then upbord = 1
If b.x > x - 10 And b.x < x + 10 And b.y > 460 And b.y < 470 Then downbord = 1

If direction = 1 Then '////////////////////it mean up
For i = 1 To 10
Back_Buffer.DrawLine x - 10, y - i, x + 10, y - i
Next i


If upbord = 1 Then
Back_Buffer.SetForeColor vbWhite
End If

Back_Buffer.DrawLine x - 10, y, x + 10, y
Back_Buffer.DrawLine x - 10, y, x, y - 10
Back_Buffer.DrawLine x, y - 10, x + 10, y
End If                '///////////////////////////////////




If direction = -1 Then 'it means down
Back_Buffer.SetForeColor vbBlack
For i = 1 To 10
Back_Buffer.DrawLine x - 10, y + i, x + 10, y + i
Next i

If downbord = 1 Then
Back_Buffer.SetForeColor vbWhite
End If

Back_Buffer.DrawLine x - 10, y, x + 10, y
Back_Buffer.DrawLine x - 10, y, x, y + 10
Back_Buffer.DrawLine x + 10, y, x, y + 10

End If

If upbord = 1 Or downbord = 1 Then Back_Buffer.SetForeColor vbWhite: Back_Buffer.DrawBox x - 10, 340, x + 10, 470 'outbox

End Sub


Private Sub drawtextlines()  'the sub that draw lines like '----------------
                                                           '----------------

For i = 2 To test + 1


textposinscroll = scrollstart + i * 15

If b.x >= 460 And b.x <= 700 And b.y >= textposinscroll + 3 And b.y <= textposinscroll + 11 Then
If Not b.y > 475 And Not b.y < 330 Then '00000
Back_Buffer.SetForeColor vbWhite
Back_Buffer.DrawLine 460, textposinscroll + 3, 700, textposinscroll + 3
Back_Buffer.DrawLine 460, textposinscroll + 15, 700, textposinscroll + 15

End If
End If '00000
Next i

End Sub

Private Sub wakeupprofilemenue()
transvalue4 = transvalue4 + 0.01
If transvalue4 > 0.5 Then transvalue4 = 0.5
Mesh(11).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)

Mesh(13).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
Mesh(14).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
Mesh(15).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
End Sub


Private Sub sleepprofilemenue()
transvalue4 = transvalue4 - 0.05
If transvalue4 < 0 Then transvalue4 = 0
Mesh(11).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)

Mesh(13).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
Mesh(14).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
Mesh(15).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue4)
End Sub


Private Sub setcredittextpos()
time = 0
For i = 1000 To 1800 Step 40
time = time + 1
ypos(time) = i
Next i
End Sub

Private Sub optionmenue()

optionbord = 0
 videobord = 0
 audiobord = 0
saveandquitbord = 0
controlsbord = 0
sleepfive 0.03
wakeupoptionbottons

If transvalue5 = 0.5 Then

Back_Buffer.SetForeColor vbWhite
Back_Buffer.DrawText 10, 30, title, False

If b.x > 385 And b.x < 639 And b.y > 694 And b.y < 738 Then Back_Buffer.SetForeColor vbBlack: saveandquitbord = 1
Back_Buffer.DrawText 445, 707, "SAVE AND QUIT", False


Back_Buffer.SetForeColor vbWhite
If b.x > 385 And b.x < 639 And b.y > 639 And b.y < 681 Then Back_Buffer.SetForeColor vbBlack: audiobord = 1
Back_Buffer.DrawText 445, 655, "   AUDIO", False


Back_Buffer.SetForeColor vbWhite
If b.x > 53 And b.x < 307 And b.y > 639 And b.y < 681 Then Back_Buffer.SetForeColor vbBlack: videobord = 1
Back_Buffer.DrawText 130, 655, "  VIDEO", False


Back_Buffer.SetForeColor vbWhite
If b.x > 718 And b.x < 972 And b.y > 639 And b.y < 681 Then Back_Buffer.SetForeColor vbBlack: controlsbord = 1
Back_Buffer.DrawText 806, 655, "CONTROLS", False

End If



End Sub

Private Sub wakeupoptionbottons()


Mesh(21).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue5)
'====

transvalue5 = transvalue5 + 0.01
If transvalue5 > 0.5 Then transvalue5 = 0.5

For j = 16 To 19
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue5)
Next j

End Sub

Private Sub sleepoptionbottons(f As Single)
transvalue5 = transvalue5 - f
If transvalue5 < 0 Then transvalue5 = 0
Mesh(21).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue5)
For j = 16 To 19
Mesh(j).SetColor dx.CreateColorRGBA(1, 0.5, 0.5, transvalue5)
Next j

End Sub

Private Sub videomenue()
title = "VIDEO PERFOMANCE"
Back_Buffer.SetForeColor vbWhite
Back_Buffer.DrawText 10, 30, title, False


Back_Buffer.SetForeColor vbWhite
Get #2, aa.mode, currentstring
Back_Buffer.DrawText 20, 210, "Video Mode------------" + "        " + currentstring, False

Get #2, Val(aa.resolution), currentstring
Back_Buffer.DrawText 20, 240, "Game Resolution-------" + "      " + currentstring, False

Get #2, Val(aa.coulor), currentstring
Back_Buffer.DrawText 20, 270, "Coulor Depth----------" + "          " + currentstring, False

Get #2, Val(aa.texture), currentstring
Back_Buffer.DrawText 20, 300, "Texture Quality-------" + "         " + currentstring, False

Get #2, Val(aa.effect), currentstring
Back_Buffer.DrawText 20, 330, "Video Effects---------", False
Back_Buffer.SetForeColor vbBlack

Back_Buffer.DrawText 340, 330, currentstring, False

Back_Buffer.SetForeColor vbWhite
Back_Buffer.DrawText 20, 360, "Gamma-----------------", False

Get #2, 12, currentstring
Back_Buffer.DrawText 20, 390, "F.P.S-----------------" + "         " + Str(aa.fps), False

Get #2, Val(aa.Light), currentstring
Back_Buffer.DrawText 20, 420, "Lighting & Shadow-----" + "          " + currentstring, False

Back_Buffer.SetForeColor vbBlack

For i = 210 To 420 Step 30

Back_Buffer.DrawBox 238, i, 470, i + 25
Back_Buffer.DrawBox 238, i, 260, i + 25
Back_Buffer.DrawBox 448, i, 470, i + 25


Back_Buffer.DrawLine 238, i + 12.5, 260, i
Back_Buffer.DrawLine 238, i + 12.5, 260, i + 25

Back_Buffer.DrawLine 448, i, 470, i + 12.5
Back_Buffer.DrawLine 448, i + 25, 470, i + 12.5


Next i

For i = 0 To 5

Back_Buffer.DrawBox aa.gamma - 7 + i, 360 + i, aa.gamma + 7 - i, 385 - i
Next i


End Sub

Private Sub audiomenue()
Back_Buffer.SetForeColor vbWhite
title = "AUDIO PERFOMANCE"
Get #2, Val(aa.sondeffects), currentstring

Back_Buffer.DrawText 20, 100, "SOUND VOLUME", False
Back_Buffer.DrawText 20, 130, "MUSIC VOLUME", False
Back_Buffer.DrawText 20, 160, "SOUND EFFACTS" + "          " + currentstring, False

Back_Buffer.SetForeColor vbBlack
For i = 100 To 160 Step 30
Back_Buffer.DrawBox 165, i, 350, i + 25

Back_Buffer.DrawBox 165, i, 187, i + 25
Back_Buffer.DrawBox 328, i, 350, i + 25

Back_Buffer.DrawLine 165, i + 12.5, 187, i
Back_Buffer.DrawLine 165, i + 12.5, 187, i + 25

Back_Buffer.DrawLine 328, i, 350, i + 12.5
Back_Buffer.DrawLine 328, i + 25, 350, i + 12.5


Next i


For i = 0 To 5
Back_Buffer.DrawBox aa.soundvolume - 7 + i, 100 + i, aa.soundvolume + 7 - i, 125 - i
Back_Buffer.DrawBox aa.musicvolume - 7 + i, 130 + i, aa.musicvolume + 7 - i, 155 - i
Next i

End Sub

Private Sub controlsmenue()
Back_Buffer.SetForeColor vbBlack
Back_Buffer.DrawBox 7, 20, 220, 50
title = "MOUSE CONTROLLS"
Back_Buffer.DrawText 410, 30, "KEYBOARD CONTROLLS", False
Back_Buffer.DrawBox 10, 50, 990, 550
Back_Buffer.DrawBox 11, 51, 989, 549
For i = 1 To 6
Back_Buffer.DrawLine 400 + i, 50, 400 + i, 550
Next i
Back_Buffer.DrawBox 405, 20, 600, 50

Back_Buffer.SetForeColor vbWhite

Back_Buffer.DrawText 20, 50, "MOUSE SPEED", False

Back_Buffer.DrawText 20, 90, "SCROLL SPEED", False

Back_Buffer.DrawText 20, 130, "MOUSE SENSITIVITY", False

Back_Buffer.DrawText 20, 170, "INVERSE MOUSE UP/DOWN", False


Back_Buffer.DrawText 20, 210, "INVERSE MOUSE RIGHT/LEFT", False


Back_Buffer.SetForeColor vbBlack

For i = 53 To 133 Step 39

Back_Buffer.DrawBox 210, i, 380, i + 25
Back_Buffer.DrawBox 210, i, 230, i + 25
Back_Buffer.DrawBox 380, i, 360, i + 25

Back_Buffer.DrawLine 230, i, 210, i + 12.5
Back_Buffer.DrawLine 210, i + 12.5, 230, i + 25

Back_Buffer.DrawLine 360, i, 380, i + 12.5
Back_Buffer.DrawLine 380, i + 12.5, 360, i + 25

Next i

For j = 1 To 5
Back_Buffer.DrawBox aa.mousespeed - 7 + j, 53 + j, aa.mousespeed + 7 - j, 53 + 25 - j
Back_Buffer.DrawBox aa.mousescroll - 7 + j, 92 + j, aa.mousescroll + 7 - j, 92 + 25 - j
Back_Buffer.DrawBox aa.mousesensitivity - 7 + j, 131 + j, aa.mousesensitivity + 7 - j, 131 + 25 - j
Next j

For i = 170 To 230 Step 38
Back_Buffer.DrawBox 278, i, 375, i + 25
  Back_Buffer.DrawBox 278, i, 298, i + 25
  Back_Buffer.DrawBox 355, i, 375, i + 25

 
 Back_Buffer.DrawLine 278, i + 12.5, 298, i
 Back_Buffer.DrawLine 278, i + 12.5, 298, i + 25
 
 Back_Buffer.DrawLine 355, i, 375, i + 12.5
 Back_Buffer.DrawLine 355, i + 25, 375, i + 12.5
Next i

Back_Buffer.SetForeColor vbWhite
Get #2, Val(aa.mouseinverseupdown), currentstring
Back_Buffer.DrawText 311, 172, currentstring, False

Get #2, Val(aa.mouseinverseleftright), currentstring
Back_Buffer.DrawText 311, 210, currentstring, False

End Sub
