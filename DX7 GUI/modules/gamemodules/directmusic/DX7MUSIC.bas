Attribute VB_Name = "DX7MUSIC"
Option Explicit
Public Dx As New DirectX7
Public DmLoad As DirectMusicLoader
Public DmPer As DirectMusicPerformance
Public Seg As DirectMusicSegment

Private Sub Command1_Click()
DmPer.Stop Seg, Nothing, 0, 0
End Sub

Private Sub directmusic()

Set DmLoad = Dx.DirectMusicLoaderCreate
Set DmPer = Dx.DirectMusicPerformanceCreate
DmPer.Init Nothing, Me.hWnd
DmPer.SetPort -1, 1
Set Seg = DmLoad.LoadSegment(App.Path & "\midi.mid")
Seg.SetStandardMidiFile
DmPer.SetMasterAutoDownload True
Seg.Download DmPer
DmPer.PlaySegment Seg, 0, 0

End Sub


