Attribute VB_Name = "basMovingBlocks"
Option Explicit

'////////////////////////////////////////////////////////////////////
'These subroutines are very similar.
'As the block moves or drops, it exchanges places with the block
'next to it or under it. And we update their X & Y coords blah blah
'blah according to their new locations.
'////////////////////////////////////////////////////////////////////

Public Sub MoveLeft(ByVal Index As Integer)
      Dim LeftBlock As Integer
      Dim temp As Integer
      
      LeftBlock = GetBlockFromCoord(Blocks(Index).XCoord - 1, Blocks(Index).YCoord)
      'Out of range.  We check this for safety.
      If LeftBlock = -1 Then Exit Sub
                  
      'First blank the block that is being moved
      Call BlitBlank(Index)
      
      'Swap the blocks' locations
      temp = Index
      With Blocks(Index)
            Board(.XCoord, .YCoord) = LeftBlock
            Board(.XCoord - 1, .YCoord) = temp
      End With
      
      'Update their data
      Blocks(Index).XCoord = Blocks(Index).XCoord - 1
      Blocks(LeftBlock).XCoord = Blocks(LeftBlock).XCoord + 1
      
      Blocks(Index).Left = Blocks(Index).XCoord * BlockWidth
      Blocks(LeftBlock).Left = Blocks(LeftBlock).XCoord * BlockWidth
      
      'Display the block's image that is being moved
      Call BlitImage(Index)
      
      
      '//////////////////////////////////////////////////////////
      Dim BottomBlock As Integer
                        
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      
      'If it is not out of range.
      If BottomBlock <> -1 Then
            If Blocks(BottomBlock).Exists = False Then
                  DropUntilNoMoreBlocks Index, BottomBlock
            End If
      End If
      
End Sub

Public Sub MoveRight(ByVal Index As Integer)
      Dim RightBlock As Integer
      Dim temp As Integer
      
      RightBlock = GetBlockFromCoord(Blocks(Index).XCoord + 1, Blocks(Index).YCoord)
      'Out of range.  We check this for safety.
      If RightBlock = -1 Then Exit Sub
            
      Call BlitBlank(Index)
      
      temp = Index
      With Blocks(Index)
            Board(.XCoord, .YCoord) = RightBlock
            Board(.XCoord + 1, .YCoord) = temp
      End With
      
      Blocks(Index).XCoord = Blocks(Index).XCoord + 1
      Blocks(RightBlock).XCoord = Blocks(RightBlock).XCoord - 1
      
      Blocks(Index).Left = Blocks(Index).XCoord * BlockWidth
      Blocks(RightBlock).Left = Blocks(RightBlock).XCoord * BlockWidth
      
      Call BlitImage(Index)
      
      '//////////////////////////////////////////////////////////
      Dim BottomBlock As Integer
                        
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      
      'If it is not out of range.
      If BottomBlock <> -1 Then
            If Blocks(BottomBlock).Exists = False Then
                  DropUntilNoMoreBlocks Index, BottomBlock
            End If
      End If
      
End Sub

Public Sub MoveRightUntilNoMoreBlocks(ByVal Index As Integer, ByVal rBlock As Integer)
      Dim RightBlock As Integer
      Dim temp As Integer
      
      RightBlock = rBlock
      
      boolProcessing = True
      
      Do While Blocks(RightBlock).Exists
            Call BlitBlank(RightBlock)
            
            temp = Blocks(RightBlock).ID
            Board(Blocks(RightBlock).XCoord, Blocks(RightBlock).YCoord) = Blocks(Index).ID
            Board(Blocks(RightBlock).XCoord - 1, Blocks(RightBlock).YCoord) = temp
            
            Blocks(RightBlock).XCoord = Blocks(RightBlock).XCoord - 1
            Blocks(Index).XCoord = Blocks(Index).XCoord + 1
            
            Blocks(RightBlock).Left = Blocks(RightBlock).XCoord * BlockWidth
            Blocks(Index).Left = Blocks(Index).XCoord * BlockWidth
            
            Call BlitImage(RightBlock)
            
            RightBlock = GetBlockFromCoord(Blocks(Index).XCoord + 1, Blocks(Index).YCoord)
            'Out of range.
            If RightBlock = -1 Then Exit Do
      Loop
      
      boolProcessing = False
End Sub

Public Sub Drop(ByVal Index As Integer)
      Dim BottomBlock As Integer
      Dim temp As Integer
      
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      'Out of range. We check this for safety.
      If BottomBlock = -1 Then Exit Sub
      
      Call BlitBlank(Index)
      
      temp = Blocks(Index).ID
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord) = Blocks(BottomBlock).ID
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord + 1) = temp
      
      Blocks(Index).YCoord = Blocks(Index).YCoord + 1
      Blocks(BottomBlock).YCoord = Blocks(BottomBlock).YCoord - 1
      
      Blocks(Index).Top = Blocks(Index).YCoord * BlockHeight
      Blocks(BottomBlock).Top = Blocks(BottomBlock).YCoord * BlockHeight
      
      Call BlitImage(Index)
      
      'Playing sound
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      'Touch the ground.
      If BottomBlock = -1 Then
            If boolPlaySound Then PlaySound
      Else
            'Touch another block
            If Blocks(BottomBlock).Exists Then
                  If boolPlaySound Then PlaySound
            End If
      End If
End Sub

Public Sub DropUntilNoMoreBlocks(ByVal Index As Integer, ByVal BottomBlock As Integer)
      Dim bottom As Integer
      Dim temp As Integer
            
      bottom = BottomBlock
      
      boolProcessing = True
      
      Do While Blocks(bottom).Exists = False
            Call BlitBlank(Index)
            
            temp = Blocks(Index).ID
            Board(Blocks(Index).XCoord, Blocks(Index).YCoord) = Blocks(bottom).ID
            Board(Blocks(Index).XCoord, Blocks(Index).YCoord + 1) = temp
            
            Blocks(Index).YCoord = Blocks(Index).YCoord + 1
            Blocks(bottom).YCoord = Blocks(bottom).YCoord - 1
            
            Blocks(Index).Top = Blocks(Index).YCoord * BlockHeight
            Blocks(bottom).Top = Blocks(bottom).YCoord * BlockHeight
            
            Call BlitImage(Index)
            
            DoEvents
            Delay (0.1)
            
            bottom = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
            If bottom = -1 Then Exit Do
            frmPuzzle.picBoard.Refresh
      Loop
      
      If boolPlaySound Then PlaySound
      boolProcessing = False
      boolDragging = False
End Sub
