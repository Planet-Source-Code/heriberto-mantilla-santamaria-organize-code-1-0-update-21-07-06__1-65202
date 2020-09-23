VERSION 5.00
Begin VB.UserControl ASPictureBox 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   PropertyPages   =   "ASPictureBox.ctx":0000
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ToolboxBitmap   =   "ASPictureBox.ctx":0010
   Begin VB.VScrollBar vsbScroll 
      Height          =   2295
      Left            =   3960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   270
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox picTwo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox picOne 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ASPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*******************************************
' Copyright © 2000 by Alexander Anikin
' E-Mail: pegas@poshuk.com
' http://www.poshuk.com/pegas/index.htm
'*******************************************
Option Explicit

 Event Click()
 Event DblClick()
 Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub hsbScroll_Change()
 Call UpdatePicTwo
End Sub

Private Sub hsbScroll_Scroll()
 Call hsbScroll_Change
End Sub

Private Sub picOne_Change()
 If (picOne.ScaleWidth <= picTwo.ScaleWidth) Or (picOne.Picture = LoadPicture()) Then hsbScroll.Visible = False
 If (picOne.ScaleHeight <= picTwo.ScaleHeight) Or (picOne.Picture = LoadPicture()) Then vsbScroll.Visible = False
 If (picOne.Picture = LoadPicture()) Then
  picTwo.Picture = LoadPicture()
  Exit Sub
 End If
 picTwo.Picture = picOne.Picture
 If (picOne.ScaleWidth > picTwo.ScaleWidth) Then hsbScroll.Visible = True
 If (picOne.ScaleHeight > picTwo.ScaleHeight) Then vsbScroll.Visible = True
 Call hsbScrollSett_Refresh
 Call vsbScrollSett_Refresh
End Sub

Private Sub picTwo_Click()
 RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
 UserControl.ScaleMode = vbPixels
 picOne.ScaleMode = vbPixels
 picTwo.ScaleMode = vbPixels
End Sub

Private Sub UserControl_Resize()
 If (UserControl.Height < 1500) Then
  UserControl.Height = 1500
 ElseIf (UserControl.Width < 1500) Then
  UserControl.Width = 1500
 End If
 '************************
 picTwo.Height = UserControl.ScaleHeight - hsbScroll.Height
 picTwo.Width = UserControl.ScaleWidth - vsbScroll.Width
 '************************
 vsbScroll.Left = picTwo.Width
 vsbScroll.Height = picTwo.Height
 hsbScroll.Top = picTwo.Height
 hsbScroll.Width = picTwo.Width
 Call picOne_Change
End Sub

Private Sub vsbScroll_Change()
 Call UpdatePicTwo
End Sub

Private Sub vsbScroll_Scroll()
 Call vsbScroll_Change
End Sub

Private Sub UpdatePicTwo()
 If (hsbScroll.Visible = False) And (vsbScroll.Visible = False) Then Exit Sub
 Call picTwo.PaintPicture(picOne.Picture, 0, 0, picTwo.ScaleWidth, picTwo.ScaleHeight, hsbScroll.Value, vsbScroll.Value, picTwo.ScaleWidth, picTwo.ScaleHeight, vbSrcCopy)
End Sub

Public Property Get Picture() As Picture
 Set Picture = picOne.Picture
End Property

Public Property Let Picture(ByVal New_Picture As IPictureDisp)
 Set Picture = New_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
 Set picOne.Picture = New_Picture
 PropertyChanged "Picture"
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = picTwo.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 picTwo.BackColor() = New_BackColor
 Call UpdatePicTwo
 PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As Integer
 BorderStyle = picTwo.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
 picTwo.BorderStyle() = New_BorderStyle
 PropertyChanged "BorderStyle"
End Property

Private Sub picTwo_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub picTwo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picTwo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picTwo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Set Picture = PropBag.ReadProperty("Picture", Nothing)
 picTwo.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
 picTwo.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Picture", Picture, Nothing)
 Call PropBag.WriteProperty("BackColor", picTwo.BackColor, &H8000000F)
 Call PropBag.WriteProperty("BorderStyle", picTwo.BorderStyle, 1)
End Sub

Private Sub hsbScrollSett_Refresh()
 hsbScroll.Value = 0
 If (picOne.ScaleWidth <= picTwo.ScaleWidth) Then Exit Sub
 hsbScroll.Max = picOne.ScaleWidth - picTwo.ScaleWidth
 If (hsbScroll.Max < 25) Then
  hsbScroll.LargeChange = 1
  hsbScroll.SmallChange = 1
 Else
  hsbScroll.LargeChange = hsbScroll.Max \ 10
  hsbScroll.SmallChange = hsbScroll.Max \ 25
 End If
End Sub

Private Sub vsbScrollSett_Refresh()
 vsbScroll.Value = 0
 If (picOne.ScaleHeight <= picTwo.ScaleHeight) Then Exit Sub
 vsbScroll.Max = picOne.ScaleHeight - picTwo.ScaleHeight
 If (vsbScroll.Max < 25) Then
  vsbScroll.LargeChange = 1
  vsbScroll.SmallChange = 1
 Else
  vsbScroll.LargeChange = vsbScroll.Max \ 10
  vsbScroll.SmallChange = vsbScroll.Max \ 25
 End If
End Sub
