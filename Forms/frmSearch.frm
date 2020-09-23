VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search..."
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvResult 
      Height          =   3120
      Left            =   120
      TabIndex        =   5
      Top             =   1335
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Categorie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Language"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Zip File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "WebSite"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Descriptor"
         Object.Width           =   2540
      EndProperty
   End
   Begin OrganizeCode.SOfficeButton SOffSearch 
      Height          =   390
      Left            =   2985
      TabIndex        =   4
      Top             =   855
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   688
      Caption         =   "    Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSearch.frx":038A
      MousePointer    =   99
      Picture         =   "frmSearch.frx":06A4
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
      TipText         =   "Search in the database."
   End
   Begin VB.TextBox txtField 
      DataField       =   "Name"
      Height          =   285
      Left            =   1095
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox cmbFindFor 
      Height          =   315
      ItemData        =   "frmSearch.frx":0A3E
      Left            =   1095
      List            =   "frmSearch.frx":0A57
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Search:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   915
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find For:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   615
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private ListResult As ListItem

Private Sub Form_Activate()
On Error Resume Next
 txtField.SetFocus
On Error GoTo 0
End Sub

Private Sub Form_Load()
 cmbFindFor.ListIndex = 0
 SOffSearch.TipTitle = Ttl
End Sub

Private Sub SOffSearch_Click()
 Dim i As Long
 
 '* Buscar en la base de datos.
On Error GoTo myErr
 Select Case cmbFindFor.ListIndex
  Case 0: xText = "Author"
  Case 1: xText = "Categorie"
  Case 2: xText = "Descriptors"
  Case 3: xText = "Language"
  Case 4: xText = "Name"
  Case 5: xText = "ZipFile"
  Case 6: xText = "WebSite"
 End Select
 yText = txtField.Text
 Call lsvResult.ListItems.Clear
 SQL = modSQL.Get_Select("tabInfCodes", "*", modSQL.Get_Simp_Set(xText, yText, "LIKE"))
 Call CargarTabla(SQL)
 With Tabla
  For i = 0 To .RecordCount - 1
   Set ListResult = lsvResult.ListItems.Add(, , IsConvertNullEmpty(.Fields("Author")))
   ListResult.SubItems(1) = IsConvertNullEmpty(.Fields("Name"))
   ListResult.SubItems(2) = IsConvertNullEmpty(.Fields("Categorie"))
   ListResult.SubItems(3) = IsConvertNullEmpty(.Fields("Language"))
   ListResult.SubItems(4) = IsConvertNullEmpty(.Fields("ZipFile"))
   ListResult.SubItems(5) = IsConvertNullEmpty(.Fields("WebSite"))
   ListResult.SubItems(6) = IsConvertNullEmpty(.Fields("Descriptors"))
   .MoveNext
  Next
 End With
 Exit Sub
myErr:
End Sub
