VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organize Code by HACKPRO TM"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9240
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin OrganizeCode.SOfficeButton SOpt 
      Height          =   285
      Index           =   3
      Left            =   225
      TabIndex        =   22
      Top             =   195
      Width           =   315
      _ExtentX        =   741
      _ExtentY        =   741
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPpal.frx":038A
      MousePointer    =   99
      Picture         =   "frmPpal.frx":06A4
      SetBorder       =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipCentered     =   0   'False
      TipForeColor    =   0
      TipText         =   "Search a specific code."
      TipTitle        =   "Organize Code"
   End
   Begin OrganizeCode.ucTreeView ucTreeView 
      Height          =   5595
      Left            =   270
      TabIndex        =   2
      Top             =   615
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   9869
   End
   Begin OrganizeCode.SOfficeButton SOpt 
      Height          =   345
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Top             =   6330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "    Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPpal.frx":0A3E
      MousePointer    =   99
      Picture         =   "frmPpal.frx":0D58
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipCentered     =   0   'False
      TipForeColor    =   0
      TipText         =   "Delete Post or Delete Author Post."
      TipTitle        =   "Organize Code"
   End
   Begin OrganizeCode.SOfficeButton SOpt 
      Height          =   345
      Index           =   1
      Left            =   195
      TabIndex        =   5
      Top             =   6330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "    Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPpal.frx":10F2
      MousePointer    =   99
      Picture         =   "frmPpal.frx":140C
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipCentered     =   0   'False
      TipForeColor    =   0
      TipText         =   "Edit an exist Post."
      TipTitle        =   "Organize Code"
      XPosPicture     =   3
   End
   Begin OrganizeCode.SOfficeButton SOpt 
      Height          =   345
      Index           =   0
      Left            =   2325
      TabIndex        =   3
      Top             =   6330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "    New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPpal.frx":17A6
      MousePointer    =   99
      Picture         =   "frmPpal.frx":1AC0
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      ShadowText      =   -1  'True
      TipActive       =   -1  'True
      TipBackColor    =   14811135
      TipCentered     =   0   'False
      TipForeColor    =   0
      TipText         =   "Add New Post."
      TipTitle        =   "Organize Code"
      XPosPicture     =   3
   End
   Begin OrganizeCode.SOfficeButton SCode 
      Height          =   345
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   165
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   609
      Caption         =   "Author's Code List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTitle        =   -1  'True
      PictureAlign    =   1
      SetBorder       =   -1  'True
      SetGradient     =   -1  'True
      SetHighLight    =   0   'False
      ShadowText      =   -1  'True
      TipBackColor    =   14811135
      TipForeColor    =   0
   End
   Begin OrganizeCode.SOfficeButton SCode 
      Height          =   6615
      Index           =   1
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   11668
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SetBorder       =   -1  'True
      SetHighLight    =   0   'False
      TipBackColor    =   14811135
      TipForeColor    =   0
      Begin MSComctlLib.ImageList imgLstIcons 
         Left            =   480
         Top             =   855
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPpal.frx":1E5A
               Key             =   ""
               Object.Tag             =   "Code"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPpal.frx":21F4
               Key             =   ""
               Object.Tag             =   "Info"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPpal.frx":278E
               Key             =   ""
               Object.Tag             =   "Zip"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPpal.frx":2D28
               Key             =   ""
               Object.Tag             =   "StartOut"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPpal.frx":373A
               Key             =   ""
               Object.Tag             =   "StartIn"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   0
      Left            =   3495
      ScaleHeight     =   6255
      ScaleWidth      =   5625
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "OK"
      Top             =   465
      Width           =   5625
      Begin VB.PictureBox picPhoto 
         BorderStyle     =   0  'None
         Height          =   2205
         Left            =   105
         ScaleHeight     =   2205
         ScaleWidth      =   1965
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   105
         Width           =   1965
      End
      Begin VB.TextBox txtDescr 
         Height          =   300
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5835
         Width           =   4275
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3285
         Left            =   1245
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2460
         Width           =   4275
      End
      Begin OrganizeCode.SOfficeButton SOffBtnURL 
         Height          =   255
         Left            =   2145
         TabIndex        =   12
         Top             =   1545
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         Caption         =   "Not Available"
         CaptionAlign    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         GrayIcon        =   0   'False
         MouseIcon       =   "frmPpal.frx":414C
         MousePointer    =   99
         MultiLine       =   -1  'True
         SetBorderH      =   0   'False
         ShadowText      =   -1  'True
         SystemColor     =   0   'False
         TipBackColor    =   14811135
         TipForeColor    =   0
         TipTitle        =   "Organize Code"
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   9
         Top             =   465
         Width           =   3330
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descriptors:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   5865
         Width           =   1035
      End
      Begin VB.Image imgPoint 
         Height          =   240
         Index           =   0
         Left            =   2850
         Picture         =   "frmPpal.frx":4466
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgPoint 
         Height          =   240
         Index           =   1
         Left            =   3135
         Picture         =   "frmPpal.frx":4E68
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgPoint 
         Height          =   240
         Index           =   2
         Left            =   3435
         Picture         =   "frmPpal.frx":586A
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgPoint 
         Height          =   240
         Index           =   3
         Left            =   3720
         Picture         =   "frmPpal.frx":626C
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgPoint 
         Height          =   240
         Index           =   4
         Left            =   4035
         Picture         =   "frmPpal.frx":6C6E
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2145
         TabIndex        =   11
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   2415
         Width           =   1035
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2145
         TabIndex        =   10
         Top             =   780
         Width           =   3330
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2145
         TabIndex        =   8
         Top             =   180
         Width           =   3330
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgPhoto 
         Height          =   2205
         Left            =   105
         Picture         =   "frmPpal.frx":7670
         Stretch         =   -1  'True
         Top             =   105
         Width           =   1965
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   3
      Left            =   3495
      ScaleHeight     =   6255
      ScaleWidth      =   5625
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "OK"
      Top             =   465
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ListView lsvZip 
         Height          =   5595
         Left            =   120
         TabIndex        =   19
         Top             =   165
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   9869
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Packed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ratio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
      End
      Begin OrganizeCode.SOfficeButton SOpt 
         Height          =   345
         Index           =   4
         Left            =   4245
         TabIndex        =   20
         Top             =   5835
         Width           =   1260
         _ExtentX        =   1773
         _ExtentY        =   609
         Caption         =   "    &View File"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPpal.frx":939A
         MousePointer    =   99
         Picture         =   "frmPpal.frx":96B4
         PictureAlign    =   1
         SetBorder       =   -1  'True
         SetGradient     =   -1  'True
         ShadowText      =   -1  'True
         TipActive       =   -1  'True
         TipBackColor    =   14811135
         TipForeColor    =   0
         TipText         =   "Contents of the file."
         XPosPicture     =   3
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   2
      Left            =   3495
      ScaleHeight     =   6255
      ScaleWidth      =   5625
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "OK"
      Top             =   465
      Visible         =   0   'False
      Width           =   5625
      Begin OrganizeCode.ASPictureBox imgScreenShot 
         Height          =   6000
         Left            =   105
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   10583
         BorderStyle     =   0
      End
      Begin VB.Image imgNoPic 
         Height          =   6255
         Left            =   105
         Picture         =   "frmPpal.frx":A0C6
         Stretch         =   -1  'True
         Top             =   465
         Width           =   5400
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   1
      Left            =   3495
      ScaleHeight     =   6255
      ScaleWidth      =   5625
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "OK"
      Top             =   465
      Visible         =   0   'False
      Width           =   5625
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6045
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   120
         Width           =   5430
      End
   End
   Begin OrganizeCode.GpTabStrip GpTabStrip 
      Height          =   6540
      Left            =   3555
      TabIndex        =   24
      Top             =   195
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   11536
      BackColor       =   14215660
      BorderStyle     =   2
      ForeColor       =   0
      Placement       =   1
      Style           =   1
      TabBorderColor  =   12937777
      TabColor        =   12937777
      TabColorActive  =   12937777
      TabFixedHeight  =   280
      TabFixedWidth   =   1200
      TabStyle        =   2
      TabWidthStyle   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuChanged 
         Caption         =   "&Changed Database File"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load PSC Readme File"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup File"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Database"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchData 
         Caption         =   "Search in &Database"
      End
      Begin VB.Menu mnuSearchPSC 
         Caption         =   "Search and Copy from &PSC"
      End
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private hBook   As Long, lKey  As Long, nAuthor As String
 Private ListZip As ListItem, i As Long, Obj     As Object
 
 Private Const SW_SHOWNORMAL As Long = 1
 
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
 For i = 0 To picContainer.UBound
  Call picContainer(i).Move(3495, 465, 5625, 6255)
 Next
 For Each Obj In frmPpal
  If (TypeOf Obj Is SOfficeButton) Then Obj.TipTitle = Ttl
 Next
 With ucTreeView
  Call .Move(270, 615, 2940, 5595)
  Call .Initialize
  Call .InitializeImageList
  Call .AddIcon(LoadResPicture(101, vbResIcon))
  Call .AddIcon(LoadResPicture(102, vbResIcon))
  Call .AddIcon(LoadResPicture(103, vbResIcon))
  Call .AddIcon(LoadResPicture(104, vbResIcon))
  Call .AddIcon(LoadResPicture(105, vbResIcon))
  Call .AddIcon(LoadResPicture(106, vbResIcon))
  Call .AddIcon(LoadResPicture(107, vbResIcon))
  Call .AddIcon(LoadResPicture(108, vbResIcon))
  Call .AddIcon(LoadResPicture(109, vbResIcon))
  .BorderStyle = bsNone
  .BackColor = SCode(0).BackColor
  .ItemHeight = 18
  .HasButtons = True
  .HasLines = True
  .HasRootLines = True
  Call .SetRedrawMode(Enable:=False)
  Call LoadCodes
  .SelectedNode = .NodeRoot
  Call .SetRedrawMode(Enable:=True)
 End With
 Call SCode(0).Move(195, 165, 3105, 345)
 Call SCode(1).Move(120, 120, 3255, 6610)
 'Call SCode(0).DisabledNormal
 With GpTabStrip
  Call .Move(3480, 120, 5670, 6615)
  .BackColor = vbButtonFace
  .BorderStyle = GpTabBorderStyle3DThin
  .Style = GpTabStyleWinXP
  .Placement = GpTabPlacementTopRight
  .TabStyle = GpTabTrapezoid
  .TabWidthStyle = GpTabFixed
  .TabFixedHeight = 280
  .TabFixedWidth = 1200 '80
  .TabColor = SCode(0).BorderColor
  .TabBorderColor = SCode(0).BorderColor
  '.XPColorScheme = GpTabUseWindows
  .HotTracking = True
  .AutoBackColor = False
  .Tabs.Clear
  Call .Tabs.Add(, "Zip", "Zip File ", imgLstIcons.ListImages(3).Picture)
  Call .Tabs.Add(, "Scr", "Screenshot ")
  Call .Tabs.Add(, "Code", "Code ", imgLstIcons.ListImages(1).Picture)
  Call .Tabs.Add(, "Wh", "What's? ", imgLstIcons.ListImages(2).Picture)
  Call .SelectTab(4)
 End With
 SOpt(3).Enabled = CargarTabla("SELECT * FROM tabInfCodes", True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 Call Unload(frmWeb)
 Set frmWeb = Nothing
 Call Kill(AppDir & "RscPic.tmp")
On Error GoTo 0
End Sub

Private Sub GpTabStrip_TabClick()
 Dim i As Integer
 
On Error Resume Next
 For i = 0 To picContainer.UBound
  picContainer(i).Visible = False
 Next
 Select Case GpTabStrip.SelectTabItem.Index
  Case 4: picContainer(0).Visible = True
  Case 3: picContainer(1).Visible = True
  Case 2: picContainer(2).Visible = True
  Case 1: picContainer(3).Visible = True
 End Select
On Error GoTo 0
End Sub

Private Sub mnuBackup_Click()
 Dim cData As String
 
 '* Crear un Backup de la base de datos.
 cData = Replace(Now, "/", "_")
 cData = Replace(cData, ":", "_")
 Call CreateDatabase(AppDir & "Database\" & lData, AppDir & "Backups\" & Replace(lData, ".mdb", cData & ".backup"))
End Sub

Private Sub mnuChanged_Click()
 Dim xText As String, Temp As String

 '* Cambiar el nombre de la base de datos.
 xText = ""
 Do While (xText = "")
  xText = Trim$(InputBox("Set the database file.", Ttl, lData))
 Loop
 ToLine = FreeFile
 Temp = "##################################" & vbCrLf
 Temp = Temp & "#  Copyright HACKPRO TM © 2006   #" & vbCrLf
 Temp = Temp & "# Configuration of Organize Code #" & vbCrLf
 Temp = Temp & "##################################" & vbCrLf & vbCrLf
 Temp = Temp & "# Database Name." & vbCrLf
 Temp = Temp & "Database = " & xText
 Open AppDir & "Config/Config.txt" For Output As #ToLine
  Print #ToLine, Temp;
 Close #ToLine
End Sub

Private Sub mnuLoad_Click()
 Dim isFileS As String, iPos As Long
 Dim xData() As String
 
 '* Cargar los datos venidos del archivo readme de PSC.
On Error Resume Next
 isFileS = Trim$(ShowOpen(Me.hWnd, 2))
 If (isFileS = "") Then Exit Sub
 ToLine = FreeFile
 iPos = 0
 ReDim xData(3)
 Open isFileS For Input As ToLine
  '* Read till End-Of-File.
  Do While Not (EOF(ToLine))
   Line Input #ToLine, LineCode
   '* Read a Text line.
   If (iPos = 0) Then
    xData(0) = Trim$(Mid$(LineCode, Len(Left$(LineCode, 7))))
    iPos = iPos + 1
   ElseIf (iPos = 1) Then
    xData(1) = xData(1) & Trim$(Mid$(LineCode, Len(Left$(LineCode, 13)))) & vbCrLf
    If (LineCode <> "This file came from Planet-Source-Code.com...the home millions of lines of source code") Then iPos = iPos + 1
   ElseIf (iPos >= 2) And (LineCode <> "This file came from Planet-Source-Code.com...the home millions of lines of source code") Then
    xData(2) = Trim$(Mid$(LineCode, Len(Left$(LineCode, 58))))
    Exit Do
   End If
  Loop
 Close ToLine
 xData(1) = Mid$(xData(1), 1, (Len(xData(1)) - Len(vbCrLf)))
 SQL = modSQL.Get_Select("tabInfCodes", "Name", "Name = '" & xData(0) & "'")
 If (CargarTabla(SQL, True) > 0) Then
  Call MsgBox("This name already exists in the Database.", vbCritical + vbOKOnly, Ttl)
  Exit Sub
 Else
  isEdit = False
  With frmNew
   .txtFields(0).Text = xData(0)
   .txtFields(2).Text = xData(1)
   .txtFields(6).Text = xData(2)
   Call .Show(1)
  End With
 End If
On Error GoTo 0
End Sub

Private Sub mnuRestore_Click()
 '* Restaurar la base de datos.
 Call CompactJetDatabase(AppDir & "Database\" & lData)
End Sub

Private Sub mnuSearchData_Click()
 '* Buscar un registro en particular.
 Call frmSearch.Show(1)
End Sub

Private Sub mnuSearchPSC_Click()
 '* Cargar el formulario para el Web.
 Call frmWeb.Show
End Sub

Private Sub SOffBtnURL_Click()
 '* Ejecutar la URL.
 Call ShellExecute(frmPpal.hWnd, vbNullString, SOffBtnURL.Caption, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub SCode_ChangedTheme(Index As Integer)
 GpTabStrip.BackColor = vbButtonFace
 SOffBtnURL.BackColor = SCode(Index).BackColor
End Sub

Private Sub SOpt_Click(Index As Integer)
On Error GoTo myErr
 Select Case Index
  Case 0 '* Agregar un nuevo registro.
   isEdit = False
   Set frmNew.Icon = SOpt(0).Picture
   frmNew.Caption = "New Post"
   Call frmNew.Show(1)
  Case 1 '* Modificar uno ya existente.
   isEdit = True
   Set frmNew.Icon = SOpt(1).Picture
   frmNew.Caption = "Edit Post"
   Call frmNew.Show(1)
  Case 2 '* Eliminar un registro ó varios registros.
   If (SOpt(1).Enabled = False) Then
    SQL = modSQL.Get_Delete("tabInfCodes", "*", modSQL.Get_Simp_Set("Author", nAuthor))
   Else
    SQL = xSQL
   End If
   Call CargarTabla(SQL, True)
   Call LoadCodes
  Case 3 '* Buscar un registro en particular.
   Call frmSearch.Show(1)
  Case 4 '* Abrir un archivo del archivo .zip.
   Call frmViewer.Show(1)
 End Select
 Exit Sub
myErr:
End Sub

Public Sub LoadCodes()
 Dim imagePos As Integer
 
On Error Resume Next
 '* Cargar los códigos principales.
 lKey = 1
 i = 0
 Call ucTreeView.Clear
 SQL = modSQL.Get_Select("tabInfCodes", "Author", , , "Author")
 Call CargarTabla(SQL)
 With Tabla
  Do Until (.EOF = True)
   lKey = lKey + 1
   hBook = ucTreeView.AddNode(, , lKey, IsConvertNullEmpty(.Fields(0)), 5, 5)
   SQL = modSQL.Get_Select("tabInfCodes", "Name, Categorie", "Author = '" & .Fields(0) & "'", "Name")
   Call CargarTabla(SQL, False, True)
   Do Until (Tebla.EOF = True)
    lKey = lKey + 1
    Select Case Tebla.Fields(1)
     Case "Events":               imagePos = 0
     Case "Class":                imagePos = 2
     Case "Subs", "Functions":    imagePos = 3
     Case "Modules":              imagePos = 4
     Case "Properties":           imagePos = 6
     Case "Usercontrols":         imagePos = 7
     Case "Complete Application": imagePos = 8
     Case Else:                   imagePos = 1
    End Select
    Call ucTreeView.AddNode(hBook, , lKey, Tebla.Fields(0), imagePos, imagePos)
    Tebla.MoveNext
   Loop
   i = i + 1
   .MoveNext
  Loop
 End With
 If (i > 0) Then
  SOpt(1).Enabled = True
  SOpt(2).Enabled = True
  SOffBtnURL.Enabled = True
 Else
  SOpt(1).Enabled = False
  SOpt(2).Enabled = False
  SOffBtnURL.Enabled = False
 End If
On Error GoTo 0
End Sub

Private Sub ucTreeView_NodeClick(ByVal hNode As Long)
 Dim lChild    As String, lPos As Long
 Dim lAuthor   As String
 Dim clsShaper As New clsRgnShaper
 
On Error Resume Next
 For i = 0 To imgPoint.UBound
  Set imgPoint(i).Picture = imgLstIcons.ListImages(4).Picture
 Next
 SOpt(4).Enabled = False
 lsvZip.ListItems.Clear
 txtDesc.Text = ""
 txtDescr.Text = ""
 lblShow(0).Caption = "Author"
 lblShow(1).Caption = "Language"
 lblShow(2).Caption = "Category"
 lblShow(3).Caption = "Not Available"
 imgScreenShot.Visible = False
 imgNoPic.Visible = True
 txtCode.Text = "Non available Source Code."
 With ucTreeView
  lChild = .NodeFullPath(hNode)
  lPos = InStrRev(lChild, "\", , vbTextCompare)
  If (lPos > 0) Then
   lAuthor = Mid$(lChild, 1, lPos - 1)
   nAuthor = lAuthor
   SOpt(1).Enabled = True
  Else
   nAuthor = lChild
   SOpt(1).Enabled = False
  End If
  lChild = Mid$(lChild, lPos + 1, Len(lChild))
  yText = lChild & "|" & lAuthor
  xText = modSQL.Get_Simp_Set("Author", nAuthor)
  SQL = modSQL.Get_Select("tab_PicUser", "*", xText)
  Call CargarTabla(SQL)
  If (IsNull(Tabla.Fields("AuthorPic")) = False) And (Len(Tabla.Fields("AuthorPic")) > 0) Then
   Set imgPhoto.Picture = LoadPhoto(Tabla!AuthorPic)
   Set picPhoto.Picture = imgPhoto.Picture
   If (imgPhoto.Picture = 0) Then
    Set imgPhoto.Picture = imgNoPic.Picture
    imgPhoto.Visible = True
    picPhoto.Visible = False
   Else
    Call clsShaper.RegionFromBitmap(picPhoto.Picture, picPhoto.hWnd, vbWhite)
    imgPhoto.Visible = False
    picPhoto.Visible = True
   End If
   Call Kill(AppDir & "RscPic.tmp")
  Else
   picPhoto.Visible = False
   imgPhoto.Visible = True
  End If
  Call CerrarTabla
  xText = modSQL.Get_Mult_Set("Name|Author", yText, "AND")
  SQL = modSQL.Get_Select("tabInfCodes", "*", xText)
  xSQL = SQL
  Call CargarTabla(SQL)
  With Tabla
   If (.BOF = True) And (.EOF = True) Then
    SOffBtnURL.Enabled = False
    Exit Sub
   Else
    SOffBtnURL.Enabled = True
   End If
   txtDesc.Text = .Fields("Description")
   txtDescr.Text = .Fields("Descriptors")
   lblShow(0).Caption = .Fields("Author")
   lblShow(1).Caption = .Fields("Language")
   lblShow(2).Caption = .Fields("Categorie")
   SOffBtnURL.Caption = IsConvertNullEmpty(.Fields("WebSite"), "Not Available")
   SOffBtnURL.TipText = SOffBtnURL.Caption
   SOffBtnURL.TipActive = True
   SOffBtnURL.TipCentered = False
   If (SOffBtnURL.Caption = "Not Available") Then
    SOffBtnURL.Enabled = False
   Else
    SOffBtnURL.Enabled = True
   End If
   txtCode.Text = IsConvertNullEmpty(.Fields("Code"), "Non available Source Code.")
   For i = 0 To .Fields("Points") - 1
    Set imgPoint(i).Picture = imgLstIcons.ListImages(5).Picture
   Next
   If (IsNull(.Fields("ScreenshotFile")) = False) And (Len(.Fields("ScreenshotFile")) > 0) Then
    Set imgScreenShot.Picture = LoadPhoto(!ScreenshotFile)
    Call Kill(AppDir & "RscPic.tmp")
    imgNoPic.Visible = False
    imgScreenShot.Visible = True
   Else
    Set imgScreenShot.Picture = imgNoPic.Picture
    imgNoPic.Visible = True
   End If
   Call Wait(0.4)
   If (.Fields("ZipFile") <> "") And (FileExits(.Fields("ZipFile")) = True) Then Call ViewZip(.Fields("ZipFile"))
  End With
  Call CerrarTabla
 End With
On Error GoTo 0
End Sub

Private Sub CompactJetDatabase(ByVal Location As String)
On Error GoTo myErr
 '* Restaurar la Base de datos.
 Dim JRO    As JRO.JetEngine
 Dim srcDB  As String
 Dim destDB As String
 
 Set JRO = New JRO.JetEngine
 srcDB = Location
 destDB = AppDir & "Database\" & "backup.mdb"
 JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & srcDB & ";Jet OLEDB:Database Password=HeryId1304", _
 "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & destDB & ";Jet OLEDB:Database Password=HeryId1304" & ";Jet OLEDB:Engine Type=4"
 Call Kill(srcDB)
 DoEvents
 Name destDB As srcDB
myErr:
End Sub

Private Sub CreateDatabase(ByVal FileName As String, ByVal CopyName As String)
 '* Crea una nueva B.D.
On Error GoTo myErr
 If Not (Datos Is Nothing) Then
  Datos.Close
  Set Datos = Nothing
 End If
 If (FileExits(AppDir & "Database\Database.ldb")) = True Then Set Datos = Nothing
 Call FileCopy(FileName, CopyName)
 Call Wait(0.4)
 Screen.MousePointer = vbDefault
 If (Datos Is Nothing) Then Call CargarBD
 Exit Sub
myErr:
 Screen.MousePointer = vbDefault
 If (Datos Is Nothing) Then Call CargarBD
End Sub

Private Sub ViewZip(ByVal ZipFile As String)
 Dim ParseFilename As String, A As Long, FileT     As String
 Dim FirstKey      As String, b As Long, ParsePath As String
 
 '* Muestra los archivos comprimidos.
 With lsvZip
  For A = 1 To 4000
   CompressedFileName(A) = ""
   CompressedDateTime(A) = ""
   UncompressedSize(A) = ""
   CompressedSize(A) = ""
   CompressedRatio(A) = ""
   CompressedPath(A) = ""
   CompressedFileType(A) = ""
  Next
  uZipFileName = ZipFile
  '-- Init Global Message Variables
  uZipInfo = ""
  uZipNumber = 0 '* Holds The Number Of Zip Files
  '-- Select UNZIP32.DLL Options - Change As Required!
  '-- Change The Next Line To Do The Actual Unzip!
  uExtractList = 1 '* 1 = List Contents Of Zip 0 = Extract
  '-- Select All Files
  uExcludeNames.uzFiles(0) = vbNullString
  uNumberXFiles = 0
  '-- Change The Next 2 Lines As Required!
  '-- These Should Point To Your Directory
  uExtractDir = ""
  If (uExtractDir <> "") Then uExtractList = 0 '* Unzip if dir specified
  '-- Let's Go And Unzip Them!
  Call VBUnZip32
  '-- Display The Returned Code Or Error!
  If (RetCode <> 0) Then
   Call ErrorHandler
   Exit Sub
  End If
  ' Display all the zip file data in the listview.
  For A = 1 To CompressedTotal
   '* Takes a full file specification and returns the filename
   ParseFilename = CompressedFileName(A)
   For b = Len(CompressedFileName(A)) To 1 Step -1
    If (Mid$(CompressedFileName(A), b, 1) = "\") Or (Mid$(CompressedFileName(A), b, 1) = "/") Then
     ParseFilename = Mid$(CompressedFileName(A), b + 1)
     If (Len(ParseFilename) <> Len(CompressedFileName(A))) Then ParsePath = Left(CompressedFileName(A), Len(CompressedFileName(A)) - Len(ParseFilename))
     Exit For
    End If
   Next
   If (Len(ParseFilename) > 0) Then
    Set ListZip = .ListItems.Add(, , ParseFilename)
    FileT = Right(ParseFilename, 4)
    FirstKey = GetKey(HKEY_CLASSES_ROOT, FileT, "")
    CompressedFileType(A) = GetKey(HKEY_CLASSES_ROOT, FirstKey, "")
    If (Len(CompressedFileType(A)) <= 0) Then CompressedFileType(A) = UCase(Right(ParseFilename, 3)) & " File"
    ListZip.SubItems(1) = CompressedFileType(A)
    ListZip.SubItems(2) = CompressedDateTime(A)
    ListZip.SubItems(3) = UncompressedSize(A)
    ListZip.SubItems(4) = CompressedSize(A)
    ListZip.SubItems(5) = CompressedRatio(A)
    'ListZip.SubItems(6) = CompressedPath(A)
    ListZip.SubItems(6) = ParsePath
   End If
  Next
  If (.ListItems.Count > 0) Then SOpt(4).Enabled = True
 End With
End Sub

Private Sub Wait(ByVal Segundos As Single)
 Dim ComienzoSeg As Double, FinSeg As Double

 '* Wait X Seconds.
 Segundos = Segundos * 0.225 '* Incrementa ó disminuye el tiempo de retardo del ciclo.
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While (FinSeg > Timer)
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
 Loop
End Sub
