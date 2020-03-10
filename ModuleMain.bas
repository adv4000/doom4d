Attribute VB_Name = "ModuleMain"
'--- API Function to Show or Hide Mouse Cursor
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
' ShowCursor True  <--- Show Cursor
' ShowCursor False <--- Hide Cursor
'--------------------------------------------------------

'---- Type of Textures to Loading.....
Public GameType As Byte ' 1 -> Earth
                        ' 2 -> Hell
'--------------------------------------------------------
Public Area As String  ' Death or Earth
'--------------------------------------------------------

'Video Mode Settings-------------------------------------
Public ScreenW As Integer  ' Weidth of Screen
Public ScreenH As Integer ' Height of Screen
Public BitPP   As Byte   ' Bit Per Pixel

'Set Selected Video Mode Settings
'Resolution of Video Mode
' Test PASSED:
' 640x480x16   800x600x16   1024x768x16   1280x1024x16
' 640x480x32   800x600x32   1024x768x32   1280x1024x32

Public Sub SetVideoSettings(W As Integer, H As Integer, BPP As Byte)
 ScreenW = W
 ScreenH = H
 BitPP = BPP
End Sub
'--------------------------------------------------------
