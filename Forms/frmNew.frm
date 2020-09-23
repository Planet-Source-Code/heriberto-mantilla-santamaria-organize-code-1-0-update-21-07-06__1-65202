VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Post"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin OrganizeCode.SOfficeButton SOffZip 
      Height          =   280
      Left            =   3840
      TabIndex        =   16
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNew.frx":038A
      MousePointer    =   99
      SetBorder       =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
      TipText         =   "Zip File."
   End
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      ItemData        =   "frmNew.frx":06A4
      Left            =   1080
      List            =   "frmNew.frx":06CF
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   735
      Width           =   3015
   End
   Begin VB.ComboBox cmbCat 
      Height          =   315
      ItemData        =   "frmNew.frx":0734
      Left            =   1080
      List            =   "frmNew.frx":0756
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1110
      Width           =   3015
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   405
      Width           =   3015
   End
   Begin OrganizeCode.SOfficeButton SOffSave 
      Height          =   375
      Left            =   3210
      TabIndex        =   20
      Top             =   5325
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      Caption         =   "     &Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNew.frx":07C4
      MousePointer    =   99
      Picture         =   "frmNew.frx":0ADE
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
      TipText         =   "Save a New Post."
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   18
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   15
      Top             =   4200
      Width           =   2745
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   4
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2865
      Width           =   3015
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Index           =   2
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1485
      Width           =   3015
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   75
      Width           =   3015
   End
   Begin OrganizeCode.SOfficeButton SOffAuthor 
      Height          =   375
      Left            =   1665
      TabIndex        =   21
      Top             =   5325
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      Caption         =   "&Author Picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNew.frx":0E78
      MousePointer    =   99
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
      TipText         =   "Show the author image."
   End
   Begin OrganizeCode.SOfficeButton SOffScreen 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5325
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   661
      Caption         =   "S&creenshot"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNew.frx":1192
      MousePointer    =   99
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
      TipText         =   "Screenshot of the code."
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      Height          =   195
      Index           =   7
      Left            =   105
      TabIndex        =   4
      Top             =   795
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   2
      Top             =   435
      Width           =   510
   End
   Begin VB.Image imgNoStart 
      Height          =   240
      Left            =   510
      Picture         =   "frmNew.frx":14AC
      Stretch         =   -1  'True
      Top             =   6255
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStart 
      Height          =   240
      Left            =   180
      Picture         =   "frmNew.frx":1EAE
      Top             =   6255
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPoint 
      Height          =   240
      Index           =   4
      Left            =   2265
      MouseIcon       =   "frmNew.frx":28B0
      MousePointer    =   99  'Custom
      Picture         =   "frmNew.frx":2BBA
      Top             =   4980
      Width           =   240
   End
   Begin VB.Image imgPoint 
      Height          =   240
      Index           =   3
      Left            =   1950
      MouseIcon       =   "frmNew.frx":35BC
      MousePointer    =   99  'Custom
      Picture         =   "frmNew.frx":38C6
      Top             =   4980
      Width           =   240
   End
   Begin VB.Image imgPoint 
      Height          =   240
      Index           =   2
      Left            =   1665
      MouseIcon       =   "frmNew.frx":42C8
      MousePointer    =   99  'Custom
      Picture         =   "frmNew.frx":45D2
      Top             =   4980
      Width           =   240
   End
   Begin VB.Image imgPoint 
      Height          =   240
      Index           =   1
      Left            =   1365
      MouseIcon       =   "frmNew.frx":4FD4
      MousePointer    =   99  'Custom
      Picture         =   "frmNew.frx":52DE
      Top             =   4980
      Width           =   240
   End
   Begin VB.Image imgPoint 
      Height          =   240
      Index           =   0
      Left            =   1080
      MouseIcon       =   "frmNew.frx":5CE0
      MousePointer    =   99  'Custom
      Picture         =   "frmNew.frx":5FEA
      Top             =   4980
      Width           =   240
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Points:"
      Height          =   195
      Index           =   9
      Left            =   105
      TabIndex        =   19
      Top             =   5010
      Width           =   480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   195
      Index           =   8
      Left            =   105
      TabIndex        =   17
      Top             =   4590
      Width           =   705
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zip File:"
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   14
      Top             =   4230
      Width           =   555
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Height          =   195
      Index           =   5
      Left            =   105
      TabIndex        =   12
      Top             =   2835
      Width           =   420
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptors:"
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   10
      Top             =   2550
      Width           =   840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   1515
      Width           =   840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categorie:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   465
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private isFileS As String, xCount As Integer
 Private isFileA As String, i      As Integer

Private Sub Form_Load()
 Call NoStart
 Set imgPoint(0).Picture = imgStart.Picture
 imgPoint(0).Tag = "Ok"
 cmbCat.ListIndex = 0
 cmbLanguage.ListIndex = 0
On Error GoTo myErr
 If (isEdit = False) Then Exit Sub
 frmNew.Caption = "Edit Post"
 Call CargarTabla(xSQL, False)
 With Tabla
  txtFields(0).Text = .Fields("Name")
  txtFields(1).Text = .Fields("Author")
  cmbLanguage.Text = .Fields("Language")
  cmbCat.Text = .Fields("Categorie")
  txtFields(2).Text = .Fields("Description")
  txtFields(3).Text = .Fields("Descriptors")
  txtFields(4).Text = .Fields("Code")
  xCount = .Fields("Points")
  txtFields(5).Text = .Fields("ZipFile")
  txtFields(6).Text = .Fields("WebSite")
 End With
 Call CerrarTabla
 For i = 0 To xCount - 1
  Set imgPoint(i).Picture = imgStart.Picture
  imgPoint(i).Tag = "Ok"
 Next
 Exit Sub
myErr:
End Sub

Private Sub imgPoint_Click(Index As Integer)
 Dim i As Integer
 
 Call NoStart
 For i = 0 To Index
  Set imgPoint(i).Picture = imgStart.Picture
  imgPoint(i).Tag = "Ok"
 Next
End Sub

Private Sub NoStart()
 Dim i As Integer

 '* Coloca los datos en su forma inicial.
 For i = 0 To imgPoint.UBound
  Set imgPoint(i).Picture = imgNoStart.Picture
  imgPoint(i).Tag = ""
 Next
End Sub

Private Sub SOffAuthor_Click()
 isFileA = Trim$(ShowOpen(Me.hWnd, True))
End Sub

Private Sub SOffSave_Click()
 Dim xCount As Integer, i As Integer
 
 If (isEdit = False) Then
  SQL = modSQL.Get_Select("tabInfCodes", "Name", "Name = '" & txtFields(0).Text & "'")
  If (CargarTabla(SQL, True) > 0) Then
   Call MsgBox("This name already exists in the Database.", vbCritical + vbOKOnly, Ttl)
   Exit Sub
  End If
 End If
 xCount = 0
 For i = 0 To imgPoint.UBound
  If (imgPoint(i).Tag <> "") Then xCount = xCount + 1
 Next
On Error Resume Next
 Set Tabla = Nothing
 '* Crea y devuelve una referencia a unobjeto ActiveX.
 Set Tabla = CreateObject("ADODB.RecordSet")
 If (isEdit = True) Then
  SQL = xSQL
 Else
  SQL = modSQL.Get_Select("tabInfCodes", "*")
 End If
 With Tabla
  '* Averiguo si objeto esta abierto. Si entonces lo cierro primero.
  If (.State = adStateOpen) Then .Close
  .ActiveConnection = Datos     '* Indica a qué objeto Connection pertenece actualmente el objeto Command o el objeto Recordset especificado.
  .CursorLocation = adUseClient '* Establece o devuelve la posición de un servicio de cursores.
  .LockType = adLockOptimistic  '* Indica el tipo de bloqueo que se pone en los registros durante el proceso de edición.
  .CursorType = adOpenKeyset    '* Indica el tipo de cursor que se usa en un objeto Recordset.
  .Source = SQL                 '* Indica el origen de los datos contenidos en un objeto Recordset (un objeto Command, una instrucción SQL, un nombre de tabla o un procedimiento almacenado).
  Call .Open
  If (isEdit = False) Then Call .AddNew
  .Fields("Name") = txtFields(0).Text
  .Fields("Author") = txtFields(1).Text
  .Fields("Language") = cmbLanguage.Text
  .Fields("Categorie") = cmbCat.Text
  .Fields("Description") = txtFields(2).Text
  .Fields("Descriptors") = txtFields(3).Text
  .Fields("Code") = txtFields(4).Text
  .Fields("Points") = xCount
  .Fields("ZipFile") = txtFields(5).Text
  .Fields("WebSite") = txtFields(6).Text
  If (isFileS <> "") Then Call SavePhoto(isFileS, !ScreenshotFile)
  Call .UpDate
 End With
 Call CerrarTabla
 SQL = modSQL.Get_Select("tab_PicUser", "*", modSQL.Get_Simp_Set("Author", txtFields(1).Text))
 If (CargarTabla(SQL) = 0) Then
  With Tabla
   If (.RecordCount = 0) Then Call .AddNew
   !Author = txtFields(1).Text
   If (isFileA <> "") Then
    Call SavePhoto(isFileA, !AuthorPic)
    .UpDate
   End If
  End With
  Call CerrarTabla
 Else
  Call CerrarTabla
 End If
 Call frmPpal.LoadCodes
 Call Unload(frmNew)
 Set frmNew = Nothing
End Sub

Private Sub SOffScreen_Click()
 isFileS = Trim$(ShowOpen(Me.hWnd, 1))
End Sub

Private Sub SOffZip_Click()
 txtFields(5).Text = ShowOpen(Me.hWnd, 0)
End Sub
