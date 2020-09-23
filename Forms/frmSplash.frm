VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   4200
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSeconds 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -100
      Top             =   -900
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private Const HWND_TOPMOST = -1
 Private Const HWND_NOTOPMOST = -2
 Private Const SWP_NOSIZE = &H1
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOACTIVATE = &H10
 Private Const SWP_SHOWWINDOW = &H40
 
 Private l As Integer

 Private Declare Sub SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
 'Call CreateSkin(frmSplash)
 Call SetWindowPos(frmSplash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
 tmrSeconds.Enabled = True
End Sub

Private Sub tmrSeconds_Timer()
 If (l > 20) Then
  Call SetWindowPos(frmSplash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
  Call frmPpal.Show
  Call Unload(frmSplash)
  Set frmSplash = Nothing
 Else
  l = l + 1
 End If
End Sub
