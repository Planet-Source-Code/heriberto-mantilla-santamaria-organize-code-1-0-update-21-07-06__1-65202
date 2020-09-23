VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Viewer"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgLstIcon 
      Left            =   2760
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewer.frx":6964
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLstIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Select All"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtFile 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmViewer.frx":6A76
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private A As Long, OpenFileName As String

Private Sub Form_Load()
On Error Resume Next
 '-- Init Global Message Variables
 uZipInfo = ""
 uZipNumber = 0    '* Holds The Number Of Zip Files.
 '-- Select UNZIP32.DLL Options - Change As Required!
 uExtractList = 0  '* 1 = List Contents Of Zip 0 = Extract.
 uNumberFiles = 0
 For A = 1 To frmPpal.lsvZip.ListItems.Count
  If (frmPpal.lsvZip.ListItems(A).Selected = True) Then
   'uZipNames.uzFiles(0) = frmPpal.lsvZip.ListItems(A).Text
   uZipNames.uzFiles(0) = frmPpal.lsvZip.ListItems(A).SubItems(6) & frmPpal.lsvZip.ListItems(A).Text
   OpenFileName = frmPpal.lsvZip.ListItems(A).Text
   uNumberFiles = 1
  End If
 Next
 If (uNumberFiles = 0) Then
  Call MsgBox("Please select a file to view.", vbInformation + vbOKOnly, Ttl)
  Call Unload(Me)
 End If
 '-- Change The Next 2 Lines As Required!
 '-- These Should Point To Your Directory
 uExtractDir = App.Path '* Directory to extract zip file to.
 '-- Let's Go And Unzip Them!
 Call VBUnZip32
On Error GoTo HandleIt
 If (UCase(Right(OpenFileName, 3)) = "RTF") Then
  Call rtxtFile.LoadFile(AppDir & OpenFileName, rtfRTF)
 Else
  Call rtxtFile.LoadFile(AppDir & OpenFileName, rtfText)
 End If
 Exit Sub
HandleIt:
 If (Err.Number = 75) Then
  Call MsgBox("Unable to open and view the selected file.", vbCritical + vbOKOnly, Ttl)
  Call Unload(Me)
 End If
On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
 Call Kill(AppDir & OpenFileName)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
  Case "Copy"
   Call Clipboard.SetText(rtxtFile.SelText)
  Case "SelectAll"
   rtxtFile.SelStart = 0
   rtxtFile.SelLength = Len(rtxtFile.Text)
 End Select
End Sub
