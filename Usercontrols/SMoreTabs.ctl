VERSION 5.00
Begin VB.UserControl SMoreTabs 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   44
   ToolboxBitmap   =   "SMoreTabs.ctx":0000
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   195
      Top             =   60
   End
End
Attribute VB_Name = "SMoreTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Private Type POINTAPI
  X As Long
  Y As Long
 End Type
 
 Private Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
 End Type

 Private b_Enabled      As Boolean
 Private iVisibleCount  As Long
 Private iCount         As Long
 Private oBackColor     As Long
 Private oDisabledColor As Long
 Private oForeColor     As Long
 Private oFrameColor    As Long
 
 Public Event Click()
 
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Sub tmrFocus_Timer()
 If (IsMouseOver = False) Then
  tmrFocus.Enabled = False
  Call Refresh
 End If
End Sub

Private Sub UserControl_ExitFocus()
 Call Refresh
End Sub

Private Sub UserControl_InitProperties()
 iVisibleCount = 0
 iCount = 0
 oBackColor = Sys2RGB(vbButtonFace)
 oDisabledColor = Sys2RGB(vbGrayText)
 oFrameColor = Sys2RGB(vbHighlight)
 oForeColor = Sys2RGB(vbButtonText)
 UserControl.BackColor = Sys2RGB(oBackColor)
 Call Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim iPos As Integer
 
 If (X >= 2) And (X <= 22) Then
  If (iVisibleCount <> 0) Then iVisibleCount = iVisibleCount - 1
  iPos = 1
 ElseIf (X >= 23) And (X <= 31) Then
  If (iVisibleCount < iCount) Then iVisibleCount = iVisibleCount + 1
  iPos = 2
 End If
 Call Refresh(iPos)
 RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim iPos As Integer
 
 If (X >= 2) And (X <= 16) Then
  iPos = 1
 ElseIf (X >= 23) And (X <= 31) Then
  iPos = 2
 End If
 Call Refresh(iPos)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  Count = .ReadProperty("Count", 0)
  BackColor = .ReadProperty("BackColor", Sys2RGB(vbButtonFace))
  DisabledColor = .ReadProperty("DisabledColor", Sys2RGB(vbGrayText))
  ForeColor = .ReadProperty("ForeColor", Sys2RGB(vbButtonText))
  FrameColor = .ReadProperty("FrameColor", Sys2RGB(vbHighlight))
  Value = .ReadProperty("Value", 0)
 End With
End Sub

Private Sub UserControl_Resize()
 If Not (Ambient.UserMode = True) Then Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  '.WriteProperty "Alignment", eTab
  Call .WriteProperty("BackColor", oBackColor, Sys2RGB(vbButtonFace))
  Call .WriteProperty("Count", iCount, 0)
  Call .WriteProperty("DisabledColor", oDisabledColor, Sys2RGB(vbGrayText))
  Call .WriteProperty("Enabled", b_Enabled, True)
  Call .WriteProperty("ForeColor", oForeColor, Sys2RGB(vbButtonText))
  Call .WriteProperty("FrameColor", oFrameColor, Sys2RGB(vbButtonText))
  Call .WriteProperty("Value", iVisibleCount, 0)
 End With
End Sub

Public Property Get BackColor() As OLE_COLOR
 BackColor = oBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
 oBackColor = Sys2RGB(NewBackColor)
 UserControl.BackColor = oBackColor
 Call PropertyChanged("BackColor")
 Call Refresh
End Property

Public Property Get DisabledColor() As OLE_COLOR
 DisabledColor = oDisabledColor
End Property

Public Property Let DisabledColor(ByVal NewDisabledColor As OLE_COLOR)
 oDisabledColor = Sys2RGB(NewDisabledColor)
 UserControl.ForeColor = oDisabledColor
 Call PropertyChanged("DisabledColor")
 Call Refresh
End Property

Public Property Get Enabled() As Boolean
 Enabled = b_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 b_Enabled = New_Enabled
 UserControl.Enabled = b_Enabled
 Call PropertyChanged("Enabled")
End Property

Public Property Get Count() As Variant
 Count = iCount
End Property

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = oForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
 oForeColor = Sys2RGB(NewForeColor)
 UserControl.ForeColor = oForeColor
 Call PropertyChanged("ForeColor")
 Call Refresh
End Property

Public Property Get FrameColor() As OLE_COLOR
 FrameColor = oFrameColor
End Property

Public Property Let FrameColor(ByVal NewFrameColor As OLE_COLOR)
 oFrameColor = Sys2RGB(NewFrameColor)
 Call PropertyChanged("FrameColor")
 Call Refresh
End Property

Public Property Let Count(ByVal vNewCount As Variant)
 iCount = vNewCount
 Call PropertyChanged("Count")
End Property

Public Property Get Value() As Long
 Value = iVisibleCount
End Property

Public Property Let Value(ByVal iCount As Long)
 iVisibleCount = iCount
 Call PropertyChanged("Value")
End Property

Private Sub APILine(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 '* Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(hDC, hPen)
 Call MoveToEx(hDC, x1, y1, PT)
 Call LineTo(hDC, x2, y2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long, Optional ByVal SetBackground As Boolean = True)
 Dim hBrush As Long, TempRect As RECT

 '* Crea un área rectangular con un color específico.
 TempRect.Left = X
 TempRect.Top = Y
 TempRect.Right = X + Width
 TempRect.Bottom = Y + Height
 hBrush = CreateSolidBrush(ColorBorder)
 Call FrameRect(hDC, TempRect, hBrush)
 Call DeleteObject(hBrush)
 If (SetBackground = True) Then
  TempRect.Left = X + 1
  TempRect.Top = Y + 1
  TempRect.Right = X + Width - 1
  TempRect.Bottom = Y + Height - 1
  hBrush = CreateSolidBrush(ColorFill)
  Call FillRect(hDC, TempRect, hBrush)
  Call DeleteObject(hBrush)
 End If
End Sub

Private Function IsMouseOver() As Boolean
 Dim PT As POINTAPI
 
 '* Mouse inside the button.
 Call GetCursorPos(PT)
 IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hWnd)
End Function

Private Sub Refresh(Optional ByVal iPos As Integer = 0)
 Dim iOffSet As Long
 
 UserControl.Height = 320
 UserControl.Width = 660
 UserControl.Cls
 If (iPos = 1) And (iVisibleCount > 0) Then
  Call DrawRectangle(UserControl.hDC, 0, 2, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(oFrameColor, 180), oFrameColor, True)
 ElseIf (iPos = 2) And (iVisibleCount < iCount) Then
  Call DrawRectangle(UserControl.hDC, 14, 2, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(oFrameColor, 180), oFrameColor, True)
 End If
 If (Value <= 0) Or (iCount <= 0) Then
  iOffSet = oDisabledColor
 Else
  iOffSet = oForeColor
 End If
 '* Arrow Left.
 Call APILine(UserControl.hDC, 10, 9, 10, 16, iOffSet)
 Call APILine(UserControl.hDC, 9, 10, 9, 15, iOffSet)
 Call APILine(UserControl.hDC, 8, 11, 8, 14, iOffSet)
 Call APILine(UserControl.hDC, 7, 12, 7, 13, iOffSet)
 If (Value = iCount) Or (iCount <= 0) Then
  iOffSet = oDisabledColor
 Else
  iOffSet = oForeColor
 End If
 '* Arrow Right.
 Call APILine(UserControl.hDC, UserControl.ScaleWidth - 23, 9, UserControl.ScaleWidth - 23, 16, iOffSet)
 Call APILine(UserControl.hDC, UserControl.ScaleWidth - 22, 10, UserControl.ScaleWidth - 22, 15, iOffSet)
 Call APILine(UserControl.hDC, UserControl.ScaleWidth - 21, 11, UserControl.ScaleWidth - 21, 14, iOffSet)
 Call APILine(UserControl.hDC, UserControl.ScaleWidth - 20, 12, UserControl.ScaleWidth - 20, 13, iOffSet)
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, ByVal Base As Long) As Long
 Dim Red   As Long, Blue  As Long
 Dim Delta As Long, Green As Long
   
 '* Devuelve un Color con menos intensidad.
 Blue = ((theColor \ &H10000) Mod &H100)
 Green = ((theColor \ &H100) Mod &H100)
 Red = (theColor And &HFF)
 Delta = &HFF - Base
 Blue = Base + Blue * Delta \ &HFF
 Green = Base + Green * Delta \ &HFF
 Red = Base + Red * Delta \ &HFF
 If (Red > 255) Then Red = 255
 If (Green > 255) Then Green = 255
 If (Blue > 255) Then Blue = 255
 ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Private Function Sys2RGB(ByVal RGBCol As Long) As Long
 If (RGBCol < 0) Then
  Call OleTranslateColor(RGBCol, 0&, Sys2RGB)
 Else
  Sys2RGB = RGBCol
 End If
End Function
