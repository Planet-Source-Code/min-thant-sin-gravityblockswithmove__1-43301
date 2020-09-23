Attribute VB_Name = "basMiscellaneous"
Option Explicit

Public Sub PlaySound()
      Const SOUND_FILE = "\drop.wav"
      If Dir(App.Path & SOUND_FILE) = "" Then Exit Sub
      'If FileLen(App.Path & soundfile) > 4000 Then Exit Sub
      
      sndPlaySound App.Path & SOUND_FILE, SND_ASYNC
End Sub

'As the name suggests...
Public Sub Delay(ByVal sec As Single)
      Dim Marker As Single
            
      Marker = Timer
      Do Until Timer > Marker + sec
            DoEvents
      Loop
End Sub

Public Sub CleanUpCollections()
      Set colSameBlocks = Nothing
      Set colBlocksLeft = Nothing
      Set colBlocksRemoved = Nothing
      Set colHintBlocks = Nothing
End Sub
