Attribute VB_Name = "modFile"
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit
 
 Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
 Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long

 Private Type OpenFileName
  lStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
 End Type

 Private OFName      As OpenFileName, i     As Long
 Private DataFile    As Integer, FileLength As Long
 Private Chunks      As Integer, Chunk()    As Byte
 Private SmallChunks As Integer
 
 Private Const ChunkSize As Integer = 1024

Public Function ShowOpen(ByVal hWnd As Long, Optional ByVal ShowImage As Boolean = False) As String
 Dim iPos As Integer, iText As String
 
 '* Muestra el cuadro de Diálogo Abrir.
 OFName.hWndOwner = hWnd
 Call IniFile(ShowImage)
 ShowOpen = ""
 If (GetOpenFileName(OFName) <> 0) Then
  iText = Trim$(OFName.lpstrFile)
  ShowOpen = iText
  iPos = InStrRev(ShowOpen, "\")
  Call ChDir(Mid$(ShowOpen, 1, Abs(iPos - 1)))
 ElseIf (ShowImage = True) Then
  iText = Trim$(OFName.lpstrFile)
  If (iText <> "") Then ShowOpen = Mid$(iText, 1, Len(iText) - 1)
 Else
  ShowOpen = ""
 End If
End Function

Private Sub IniFile(ByVal ShowImage As Integer)
 '* Establece propiedades iniciales del Commondialog.
 OFName.lStructSize = Len(OFName)
 OFName.hInstance = App.hInstance
 If (ShowImage = 0) Then
  OFName.lpstrFilter = "Zip file (*.zip)" + Chr$(0) + "*.zip" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  OFName.lpstrDefExt = ".zip"
  OFName.lpstrTitle = "Open file Zip"
 ElseIf (ShowImage = 1) Then
  OFName.lpstrFilter = "Files of Bitmap (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) & _
                       "JPEG Filter (*.jpg;*.jpeg)" + Chr$(0) + "*.jpg;*.jpeg" + Chr$(0) & _
                       "GIF (*.gif)" + Chr$(0) + "*.gif" + Chr$(0) & _
                       "All image files (*.bmp;*.jpg;*.jpeg;*.gif)" + Chr$(0) + "*.bmp;*.jpg;*.jpeg;*.gif" + Chr$(0) & _
                       "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  OFName.lpstrDefExt = ".jpg"
  OFName.lpstrTitle = "Open Image"
  OFName.nFilterIndex = 4
 Else
  OFName.lpstrFilter = "Readme file (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  OFName.lpstrDefExt = ".txt"
  OFName.lpstrTitle = "Open Readme file"
 End If
 OFName.lpstrFile = Space$(254)
 OFName.nMaxFile = 255
 OFName.lpstrFileTitle = Space$(254)
 OFName.nMaxFileTitle = 255
 OFName.Flags = &H80000 + &H400 + &H1000 + &H4 + &H8 + &H2 + &H800
End Sub

'------------------------------------------------------------------
' Disigned by Rodney Safe Computing Tiger software
' You are free to distribute this code.
' But do not forget to include my name somewhere in
' your comments.
' Have a nice.
' Rodney Godfried.
'------------------------------------------------------------------
Public Sub SavePhoto(ByVal PhotoFileName As String, ByVal FieldName As Field)
On Error GoTo Out
 '---------------------------------------------
 ' If there is no image file exits.
 '---------------------------------------------
 If (Len(PhotoFileName) = 0) Then Exit Sub
 DataFile = FreeFile
 '---------------------------------------------
 ' Open the image file.
 '---------------------------------------------
 Open PhotoFileName For Binary Access Read As DataFile
  FileLength = LOF(DataFile) '* Length of data in file.
  '---------------------------------------------
  ' If the imagefile is empty exits.
  '---------------------------------------------
  If (FileLength = 0) Then
   Close DataFile
   Exit Sub
  End If
  '---------------------------------------------
  ' Calculate the bytes(Chunks)pakages to write.
  '---------------------------------------------
  Chunks = FileLength \ ChunkSize
  SmallChunks = FileLength Mod ChunkSize
  '---------------------------------------------
  ' Resize the chunck array to adjust the firts
  ' bytes package to be copied.
  '---------------------------------------------
  ReDim Chunk(SmallChunks)
  Get DataFile, , Chunk()
  '---------------------------------------------
  ' Write the bytes to the given database
  ' fieldname.
  '---------------------------------------------
  Call FieldName.AppendChunk(Chunk())
  '---------------------------------------------
  ' Adjust the chunck array for the rest bytes
  ' packages to be copied.
  '---------------------------------------------
  ReDim Chunk(ChunkSize)
  For i = 1 To Chunks
   Get DataFile, , Chunk()
   Call FieldName.AppendChunk(Chunk())
  Next
 Close DataFile
 Exit Sub
Out:
End Sub

Public Function LoadPhoto(ByVal FieldName As Field) As StdPicture
 Dim lngOffset As Long, lngTotalSize As Long
 Dim strChunk  As String
 
On Error GoTo Out
 DataFile = FreeFile
 Open AppDir & "RscPic.tmp" For Binary Access Write As DataFile
  lngTotalSize = FieldName.ActualSize
  Chunks = lngTotalSize \ ChunkSize
  SmallChunks = lngTotalSize Mod ChunkSize
  ReDim Chunk(ChunkSize)
  Chunk() = FieldName.GetChunk(ChunkSize)
  Put DataFile, , Chunk()
  lngOffset = lngOffset + ChunkSize
  Do While (lngOffset < lngTotalSize)
   Chunk() = FieldName.GetChunk(ChunkSize)
   Put DataFile, , Chunk()
   lngOffset = lngOffset + ChunkSize
  Loop
 Close DataFile
 '============================================
 ' Load the picture into the image box.
 '============================================
 Set LoadPhoto = LoadPicture(AppDir & "RscPic.tmp")
On Error GoTo 0
 Exit Function
Out:
 Set LoadPhoto = Nothing
 Close DataFile
 Call Kill(AppDir & "RscPic.tmp")
On Error GoTo 0
End Function
