Attribute VB_Name = "DX7INPUT"

'               DIRECT INPUT FOR DX7 WITH VB6
'               2005-2006
'               BY:LION FORCE
'               FARSTAR_MAN@YAHOO.COM




Public di As directinput
Public didev As DirectInputDevice
Public dikey As DIKEYBOARDSTATE


Sub directinput()
Set di = Dx.DirectInputCreate
Set didev = di.CreateDevice("GUID_SYSKEYBOARD")
didev.SetCooperativeLevel Form1.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
didev.SetCommonDataFormat DIFORMAT_KEYBOARD
didev.Acquire
End Sub
