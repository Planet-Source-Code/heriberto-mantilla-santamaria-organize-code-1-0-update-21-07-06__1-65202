VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWeb 
   Caption         =   "Search and Copy from PSC"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "frmWeb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin OrganizeCode.SOfficeButton SOffBtnExtract 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      Caption         =   "  Set"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayIcon        =   0   'False
      MouseIcon       =   "frmWeb.frx":0802
      MousePointer    =   99
      Picture         =   "frmWeb.frx":0B1C
      PictureAlign    =   1
      SetBorder       =   -1  'True
      ShadowText      =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
   End
   Begin SHDocVwCtl.WebBrowser WbBrw 
      Height          =   1215
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   1260
      ExtentX         =   2222
      ExtentY         =   2143
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Only you can use this function with the post that take a .zip file."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   210
      Width           =   7140
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2006  *'
'* Programado por HACKPRO TM © Copyright 2006 *'
'**********************************************'
Option Explicit

 Private code      As String
 Private LendPos   As Long    '* The end pos of the string inbetween searchword1 and searchword2.
 Private LStartPos As Long    '* The start pos of the string inbetween searchword1 and searchword2.
 Private tCode(6)  As String
 
Private Sub Form_Load()
 '* Establezco el ancho del formulario.
 Call WbBrw.Navigate("http://www.pscode.com/")
End Sub

Private Sub Form_Resize()
On Error Resume Next
 With WbBrw
  .Top = 600
  .Left = 0
  .Height = ScaleHeight - 600
  .Width = ScaleWidth
 End With
On Error GoTo 0
End Sub

Private Sub SOffBtnExtract_Click()
 Dim Partir1 As String, i As Integer
 
On Error Resume Next
 '* Extraer los datos principales.
 code = WbBrw.Document.body.innerHTML
 '* Ahora saco el código fuente necesario para saber el autor y demás cosas importantes.
 Partir1 = FuncInstrSandwich(1, code, "href=""#zip""", "Terms of Agreement:&nbsp;&nbsp;&nbsp;<BR>", True)
 tCode(0) = Trim$(FuncInstrSandwich(1, Partir1, "<b>By:</b>", "&nbsp;"))  '* Author.
 tCode(1) = Trim$(FuncInstrSandwich(1, code, "<b>Compatibility:</b>", "<br>")) '* Lenguaje.
 tCode(2) = Trim$(FuncInstrSandwich(1, Partir1, "Users have accessed this&nbsp;code&nbsp;", "<td colspan=3>&nbsp;</td></tr>")) '* Descripción.
 tCode(3) = Trim$(FuncInstrSandwich(1, code, "<!--title start-->", "<!--title end-->")) '* Post Name.
 For i = 0 To 3
  tCode(i) = Replace(tCode(i), "&nbsp;", " ")
  tCode(i) = ReplaceTags(tCode(i))
 Next
 tCode(3) = Replace(tCode(3), vbCrLf, "")
 tCode(3) = Replace(tCode(3), vbCr, "")
 tCode(3) = Replace(tCode(3), vbLf, "")
 tCode(3) = Replace(tCode(3), vbNewLine, "")
 isEdit = False
 With frmNew
  .txtFields(1).Text = Trim$(tCode(0))
  .cmbLanguage.ListIndex = FindInCombo(.cmbLanguage, tCode(1))
  .txtFields(2).Text = Trim$(tCode(2))
  .txtFields(0).Text = Trim$(Mid$(tCode(3), 1, Len(tCode(3)) - 12))
  Call .Show(1)
 End With
On Error GoTo 0
End Sub

Private Function FindInCombo(ByVal tCombo As ComboBox, Optional ByVal SearchText As String) As Long
 Dim i As Long, tFind As Boolean, tText As String
 
 '* Busca si el elemento se encuentra en la lista.
 tFind = False
 FindInCombo = -1
 For i = 0 To tCombo.ListCount
  tText = LCase$(tCombo.list(i))
  If (i = 11) Then tText = LCase$("VB")
  If (InStr(1, LCase$(SearchText), tText, vbTextCompare) <> 0) Then
   tFind = True
   Exit For
  End If
 Next
 If (tFind = True) Then FindInCombo = i
End Function

'* This is a function that greatly expands the usefulness of the instr _
   function..this function looks for what is between 2 other strings _
   for example: _
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: _
 _
   lets say your looking to extract the host address of a full url. _
   In other words..what lies between "http//www." and the first "/" _
   followwing the http/www. _
   You could do an instr looking for the "http" part _
   then do an instr for the "/" part _
   and if they both return a nonzero then the word _
   were looking for starts at the first instr + the _
   len of the first word ("htp://www."-which is 10) _
   and ends at the second instr..so we would use the _
   mid() function. _
   This function does all this for you _
-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.- _
  PARAMETERS:
 '  lngStartPos:   [required]  :  The start point of the search for the first _
                                  word (strFindFirst$) same as instr _
    strToSearch$:  [required]  :  The string were doing this function on :-8 _
    strFindFirst$: [required]  :  The first string to search 4 in strToSearch$ _
    strFindEnd$:   [required]  :  The second string to search 4 instrToSearch$ _
    bCaseMatters   [optional]  :  Whether or not 2 take case into consideration _
    lngMaxSpreadLen[optional]  :  The maximum allowabled len between the 2 _
                                  strings (strFindFirst$ and strFindEnd$) _
-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-
Private Function FuncInstrSandwich(lngStartPos&, strToSearch$, strFindFirst$, strFindEnd$, Optional ByVal bCaseMatters As Boolean = False, Optional ByVal lngMaxSpreadLen As Long = -1) As String
 Dim l2Start As Long, l2 As Long, l1 As Long
 Dim nCase   As String
 
 '* Set starting pos to avoid errors.
 If (lngStartPos < 1) Then lngStartPos = 1
 '* I case is irrelevant the lower _
    case everything.
 nCase = strToSearch$
 If (bCaseMatters = False) Then
  strToSearch$ = LCase$(strToSearch$)
  strFindFirst$ = LCase$(strFindFirst$)
  strFindEnd$ = LCase$(strFindEnd$)
 End If
 '* Find the starting pos of the first string.
 l1 = InStr(lngStartPos, strToSearch$, strFindFirst$)
 '* If it's found, search for the second part, _
    the start pos being return of the first _
    inst (l1) + len of that string.
 If (l1 <> 0) Then
  '* Property.
  l2Start = (Len(strFindFirst$) + l1)
  l2 = InStr(l2Start, strToSearch$, strFindEnd$)
  '* If the second string is found...
  If (l2 <> 0) Then
   '* This means user HAS NOT specified a max spead len or _
      he HAS specified a max spread len and the len of the _
      string between searchword1 and searchword2 <= lngMaxSpreadLen.
   If (lngMaxSpreadLen <= 0) Or _
      (lngMaxSpreadLen > 0) And _
      ((l2 - l2Start) <= lngMaxSpreadLen) _
   Then
    '* The actual string that lies between l2Start and l2.
    FuncInstrSandwich = Mid$(nCase, l2Start, (l2 - l2Start))
    '* Return the start and end pos of the sandwich string.
    Let LStartPos = l2Start
    Let LendPos = l2
   Else
    FuncInstrSandwich = ""
   End If
  Else
   '* Second string not founds so return -1.
   FuncInstrSandwich = ""
  End If
  '* First string not founds so return -1.
 Else
  FuncInstrSandwich = ""
 End If
End Function

'* Removes HTML Tags
'* By mukthar m
'* Contact mukthar@onesourceindia.com
'* Visit www.onesourceindia.com to send free sms to anywhere in india.
'* Report any bugs to mukthar@onesourceindia.com
Private Function ReplaceTags(ByVal Char As String)
 '* Remover carácteres no necesarios.
 Dim SourceStr         As String, j As Long
 Dim TargetStr         As String
 Dim Opened            As Boolean
 Dim LenOfString       As String
 Dim CurrentChar       As String
 Dim CurrentRunningTag As String
 Dim TagPos            As Integer
 Dim ExcludeList       As String
 Dim LineBreak         As String
 Dim OpenedWith        As String
 Dim CharPos           As Long
 Dim CloseWith         As String
 Const OpeningChars    As String = "<&"
 Const ClosingChars    As String = ">;"
 
 ExcludeList = ""
 LineBreak = "</P><BR></TD></TBODY>"
 SourceStr = Char
 LenOfString = Len(SourceStr)
 For j = 1 To LenOfString
  CurrentChar = Mid(SourceStr, j, 1)
  If (Opened = True) Then
   CurrentRunningTag = CurrentRunningTag & CurrentChar
   If (UCase$(CurrentChar) = UCase$(CloseWith)) Then
    Opened = False
    If (InStr(UCase$(LineBreak), UCase$(CurrentRunningTag)) > 0) Then TargetStr = TargetStr & vbNewLine
   End If
  ElseIf (InStr(UCase$(OpeningChars), UCase$(CurrentChar)) > 0) Then
   CharPos = InStr(UCase$(OpeningChars), UCase$(CurrentChar))
   CloseWith = Mid$(ClosingChars, CharPos, 1)
   Opened = True
   OpenedWith = CurrentChar
   CurrentRunningTag = CurrentChar
  Else
   If (InStr(UCase$(ExcludeList), UCase$(CurrentRunningTag)) = 0) Then TargetStr = TargetStr & CurrentChar
  End If
 Next
 ReplaceTags = TargetStr
End Function
