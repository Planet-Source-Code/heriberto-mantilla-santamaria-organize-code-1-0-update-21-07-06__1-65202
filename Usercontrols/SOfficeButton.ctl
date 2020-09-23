VERSION 5.00
Begin VB.UserControl SOfficeButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   PropertyPages   =   "SOfficeButton.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "SOfficeButton.ctx":0035
End
Attribute VB_Name = "SOfficeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
'*                   Version 1.0.3                     *'
'*******************************************************'
'* Control:       SOfficeButton                        *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Description:   This usercontrol simulates a Office  *'
'*                Button.                              *'
'*                                                     *'
'*                This button is based on the origi-   *'
'*                nal code of fred.cpp, please see     *'
'*                the [CodeId = 56053].                *'
'*                                                     *'
'*                Also many thanks to Paul Caton for   *'
'*                it's spectacular self-subclassing    *'
'*                usercontrol template, please see     *'
'*                the [CodeId = 54117].                *'
'*******************************************************'
'* Started on:    Sunday, 09-jan-2005.                 *'
'*******************************************************'
'* Release date:  Monday, 18-jul-2005.                 *'
'*******************************************************'
'*                                                     *'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
Option Explicit

'* Private Types.
 Private Type RECT
  xLeft    As Long
  xTop     As Long
  xRight   As Long
  xBottom  As Long
 End Type
 
'*******************************************************'
'*                Subclasser Declarations              *'
'*                                                     *'
'* Author: Paul Caton.                                 *'
'* Mail:   Paul_Caton@hotmail.com                      *'
'* Web:    None                                        *'
'*******************************************************'
 
 '-uSelfSub declarations---------------------------------------------------------------------------
 Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
 End Enum

 Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
 End Enum

 Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
 End Type

 Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
 Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
 Private Const CODE_LEN      As Long = 240                                   'Thunk length in bytes
 Private Const WNDPROC_OFF   As Long = &H30                                  'WndProc execution offset
 Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
 Private Const PAGE_RWX      As Long = &H40                                  'Allocate executable memory
 Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
 Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
 Private Const IDX_SHUTDOWN  As Long = 1                                     'Shutdown flag data index
 Private Const IDX_HWND      As Long = 2                                     'hWnd data index
 Private Const IDX_EBMODE    As Long = 3                                     'EbMode data index
 Private Const IDX_CWP       As Long = 4                                     'CallWindowProc data index
 Private Const IDX_SWL       As Long = 5                                     'SetWindowsLong data index
 Private Const IDX_FREE      As Long = 6                                     'VirtualFree data index
 Private Const IDX_ME        As Long = 7                                     'Owner data index
 Private Const IDX_WNDPROC   As Long = 8                                     'Original WndProc data index
 Private Const IDX_CALLBACK  As Long = 9                                     'zWndProc data index
 Private Const IDX_BTABLE    As Long = 10                                    'Before table data index
 Private Const IDX_ATABLE    As Long = 11                                    'After table data index
 Private Const IDX_EBX       As Long = 14                                    'Data code index
 
 Private z_Code(29)          As Currency                                     'Thunk machine-code initialised here
 Private z_Data(552)         As Long                                         'Array whose data pointer is re-mapped to arbitary memory addresses
 Private z_DataDataPtr       As Long                                         'Address of z_Data()'s SafeArray data pointer
 Private z_DataOrigData      As Long                                         'Address of z_Data()'s original data
 Private z_hWnds             As Collection                                   'hWnd/thunk-address collection
 
 Private Declare Function CallWindowProcA Lib "USER32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
 Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
 Private Declare Function IsWindow Lib "USER32" (ByVal hWnd As Long) As Long
 Private Declare Function SetWindowLongA Lib "USER32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
 Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
 Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
 
 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 Private Declare Function TrackMouseEvent Lib "USER32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 Public Event MouseEnter()
 Public Event MouseLeave()
 
 Private bTrack                As Boolean
 Private bTrackUser32          As Boolean
 Private isInCtrl              As Boolean
 
 Private Const WM_MOUSEMOVE         As Long = &H200
 Private Const WM_MOUSELEAVE        As Long = &H2A3
 Private Const WM_THEMECHANGED      As Long = &H31A
 Private Const WM_SYSCOLORCHANGE    As Long = &H15 '21
'*******************************************************'

'*******************************************************'
'*                     Tool Tip Class                  *'
'*                                                     *'
'* Author: Mark Mokoski                                *'
'* Mail: markm@cmtelephone.com                         *'
'* Web:  www.rjillc.com                                *'
'*******************************************************'

 '******************************************************
 '* API Functions.                                     *
 '******************************************************
 Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
 Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Private Declare Function DestroyWindow Lib "USER32" (ByVal hWnd As Long) As Long
 Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
 Private Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
 
 '******************************************************
 '* Constants.                                         *
 '******************************************************
   
 '* Windows API Constants.
 Private Const CW_USEDEFAULT = &H80000000
 Private Const HWND_TOPMOST = -1
 Private Const SWP_NOACTIVATE = &H10
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const WM_USER = &H400

 '* Tooltip Window Constants.
 Private Const TTF_CENTERTIP = &H2
 Private Const TTF_SUBCLASS = &H10
 Private Const TTM_ACTIVATE = (WM_USER + 1)
 Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
 Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
 Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
 Private Const TTM_SETTITLE = (WM_USER + 32)
 Private Const TTM_ADDTOOLA = (WM_USER + 4)
 Private Const TTS_ALWAYSTIP = &H1
 Private Const TTS_BALLOON = &H40
 Private Const TTS_NOPREFIX = &H2

 '* Tool Tip Icons.
 Private Const TTI_ERROR                   As Long = 3
 Private Const TTI_INFO                    As Long = 1
 Private Const TTI_NONE                    As Long = 0
 Private Const TTI_WARNING                 As Long = 2
 
 '* Tool Tip API Class.
 Private Const TOOLTIPS_CLASSA = "tooltips_class32"

 '******************************************************
 '* Types.                                             *
 '******************************************************

 '* Tooltip Window Types.
 Private Type TOOLINFO
  lSize                             As Long
  lFlags                            As Long
  lhWnd                             As Long
  lId                               As Long
  lpRect                            As RECT
  hInstance                         As Long
  lpStr                             As String
  lParam                            As Long
 End Type

 '******************************************************
 '* Local Class variables and Data .                   *
 '******************************************************

 '* Local variables to hold property values.
 Private ToolActive                        As Boolean
 Private ToolBackColor                     As Long
 Private ToolCentered                      As Boolean
 Private ToolForeColor                     As Long
 Private ToolIcon                          As ToolIconType
 Private TOOLSTYLE                         As ToolStyleEnum
 Private ToolText                          As String
 Private ToolTitle                         As String

 '* Private Data for Class.
 Private m_ltthWnd                         As Long
 Private TI                                As TOOLINFO
 
 Public Enum ToolIconType
  TipNoIcon = TTI_NONE            '= 0
  TipIconInfo = TTI_INFO          '= 1
  TipIconWarning = TTI_WARNING    '= 2
  TipIconError = TTI_ERROR        '= 3
 End Enum

 Public Enum ToolStyleEnum
  StyleStandard = 0
  StyleBalloon = 1
 End Enum
'*******************************************************'

 '* Private Types.
 Private Type POINTAPI
  x      As Long
  y      As Long
 End Type
  
 '* Private Enum's.
 Public Enum OfficeAlign
  ACenter = &H0
  ALeft = &H1
  ARight = &H2
  ATop = &H3
  ABottom = &H4
 End Enum
 
 Public Enum OfficeState
  OfficeNormal = &H0
  OfficeHighLight = &H1
  OfficeHot = &H2
  OfficeDisabled = &H3
 End Enum
 
 Public Enum ShapeBorder
  Rectangle = &H0
  [Round Rectangle] = &H1
 End Enum
  
 '* Private variables.
 Private g_Font           As StdFont
 Private isAutoSizePic    As Boolean
 Private isBackColor      As OLE_COLOR
 Private isBorderColor    As OLE_COLOR
 Private isButtonShape    As ShapeBorder
 Private isCaption        As String
 Private isDisabledColor  As OLE_COLOR
 Private isEnabled        As Boolean
 Private isFocus          As Boolean
 Private isFontAlign      As OfficeAlign
 Private isForeColor      As OLE_COLOR
 Private isHeight         As Long
 Private isHighLightColor As OLE_COLOR
 Private isHotColor       As OLE_COLOR
 Private isHotTitle       As Boolean
 Private isMultiLine      As Boolean
 Private isPicture        As StdPicture
 Private isPictureAlign   As OfficeAlign
 Private isPictureSize    As Integer
 Private isSetBorder      As Boolean
 Private isSetBorderH     As Boolean
 Private isSetGradient    As Boolean
 Private isSetHighLight   As Boolean
 Private isShadowText     As Boolean
 Private isShowFocus      As Boolean
 Private isState          As OfficeState
 Private isSystemColor    As Boolean
 Private isTxtRect        As RECT
 Private isWidth          As Long
 Private isXPos           As Integer
 Private isYPos           As Integer
 Private m_bGrayIcon      As Boolean
 Private RectButton       As RECT
 
 '* Private Constants.
 Private Const defBackColor      As Long = vbButtonFace
 Private Const defBorderColor    As Long = vbHighlight
 Private Const defDisabledColor  As Long = vbGrayText
 Private Const defForeColor      As Long = vbButtonText
 Private Const defHighLightColor As Long = vbHighlight
 Private Const defHotColor       As Long = vbHighlight
 Private Const defShape          As Integer = &H0
 Private Const DSS_DISABLED      As Long = &H20
 Private Const DSS_MONO          As Long = &H80
 Private Const DSS_NORMAL        As Long = &H0
 Private Const DST_BITMAP        As Long = &H4
 Private Const DST_ICON          As Long = &H3
 Private Const DT_BOTTOM         As Long = &H8
 Private Const DT_CENTER         As Long = &H1
 Private Const DT_LEFT           As Long = &H0
 Private Const DT_RIGHT          As Long = &H2
 Private Const DT_SINGLELINE     As Long = &H20
 Private Const DT_TOP            As Long = &H0
 Private Const DT_VCENTER        As Long = &H4
 Private Const DT_WORDBREAK      As Long = &H10
 Private Const DT_WORD_ELLIPSIS  As Long = &H40000
 Private Const PS_SOLID          As Long = 0
 Private Const SW_SHOWNORMAL     As Long = 1
 Private Const Version           As String = "SOfficeButon 1.0.3 By HACKPRO TM"
 
 '* API's Windows Call.
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawFocusRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT) As Long
 Private Declare Function DrawState Lib "USER32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal Flags As Long) As Long
 Private Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function FrameRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
 Private Declare Function InflateRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
 Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Private Declare Function WindowFromPoint Lib "USER32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 
 '* Public Events.
 Public Event Click()
Attribute Click.VB_MemberFlags = "200"
 Public Event ChangedTheme()
 
 '* For Create GrayIcon --> MArio Florez.
 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
 Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
 Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
 Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
 Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
 Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
 Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
 Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
 Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
 
 ' Type - GetObjectAPI.lpObject
 Private Type BITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
 End Type

 ' Type - CreateIconIndirect / GetIconInfo
 Private Type ICONINFO
  fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies an icon; FALSE specifies a cursor.
  xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
  hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
 End Type

'*******************************************************'
'* Public Properties.                                  *'
'*******************************************************'
Public Property Get AutoSizePicture() As Boolean
 AutoSizePicture = isAutoSizePic
End Property

'* English: Adjusts the control to the picture size.
Public Property Let AutoSizePicture(ByVal TheAutoSize As Boolean)
 '* Ajusta el control al tamaño de la imagen.
 isAutoSizePic = TheAutoSize
 Call PropertyChanged("AutoSizePicture")
 Call Refresh(isState)
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = isBackColor
End Property

'* English: Returns/Sets the background color used to display text and graphics in an object.
Public Property Let BackColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del Usercontrol.
 isBackColor = ConvertSystemColor(theColor)
 Call PropertyChanged("BackColor")
 Call Refresh(isState)
End Property

Public Property Get BorderColor() As OLE_COLOR
 BorderColor = isBorderColor
End Property

'* English: Returns/Sets the color of border of the Object.
Public Property Let BorderColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del borde del objeto.
 isBorderColor = ConvertSystemColor(theColor)
 Call PropertyChanged("BorderColor")
 If (isSetBorder = True) Then Call Refresh(isState)
End Property

Public Property Get ButtonShape() As ShapeBorder
 ButtonShape = isButtonShape
End Property

'* English: Returns/Sets the type of border of the control.
Public Property Let ButtonShape(ByVal theButtonShape As ShapeBorder)
 '* Devuelve ó establece el tipo de borde del botón.
 isButtonShape = theButtonShape
 Call PropertyChanged("ButtonShape")
 If (isSetBorder = True) Then Call Refresh(isState)
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
 Caption = isCaption
End Property

'* English: Returns/Sets "Caption" property.
Public Property Let Caption(ByVal TheCaption As String)
 '* Devuelve ó establece el texto del Objeto.
 isCaption = TheCaption
 Call SetAccessKey(isCaption)
 Call PropertyChanged("Caption")
 Call Refresh(isState)
End Property

Public Property Get CaptionAlign() As OfficeAlign
 CaptionAlign = isFontAlign
End Property

'* English: Returns/Sets alignment of the text.
Public Property Let CaptionAlign(ByVal theAlign As OfficeAlign)
 '* Devuelve ó establece la alineación del texto.
 isFontAlign = theAlign
 Call PropertyChanged("CaptionAlign")
 Call Refresh(isState)
End Property

Public Property Get DisabledColor() As OLE_COLOR
 DisabledColor = isDisabledColor
End Property

'* English: Returns/Sets the color of the disabled text.
Public Property Let DisabledColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del texto deshabilitado.
 isDisabledColor = ConvertSystemColor(theColor)
 Call PropertyChanged("DisabledColor")
 Call Refresh(isState)
End Property

Public Property Get Enabled() As Boolean
 Enabled = isEnabled
End Property

'* English: Returns/Sets the Enabled property of the control.
Public Property Let Enabled(ByVal TheEnabled As Boolean)
 '* Devuelve ó establece si el Usercontrol esta habilitado ó deshabilitado.
 isEnabled = TheEnabled
 UserControl.Enabled = isEnabled
 Call PropertyChanged("Enabled")
 If (isEnabled = True) Then
  isState = OfficeNormal
 Else
  isState = OfficeDisabled
 End If
 Call Refresh(isState)
End Property

Public Property Get Font() As StdFont
 Set Font = g_Font
End Property

'* English: Returns/Sets the Font of the control.
Public Property Set Font(ByVal New_Font As StdFont)
 '* Devuelve ó establece el tipo de fuente del texto.
On Error Resume Next
 With g_Font
  .Name = New_Font.Name
  .Size = New_Font.Size
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 Call PropertyChanged("Font")
 Call Refresh(isState)
End Property

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = isForeColor
End Property

'* English: Use this color for drawing Normal Font.
Public Property Let ForeColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color de la fuente.
 isForeColor = ConvertSystemColor(theColor)
 Call PropertyChanged("ForeColor")
 Call Refresh(isState)
End Property

'* English: Control Version.
Public Property Get GetControlVersion() As String
 '* Español: Version del Control.
 GetControlVersion = Version & " © " & Year(Now)
End Property

Public Property Let GrayIcon(ByVal bGrayIcon As Boolean)
 m_bGrayIcon = bGrayIcon
 Call PropertyChanged("GrayIcon")
 Call Refresh
End Property

Public Property Get GrayIcon() As Boolean
 GrayIcon = m_bGrayIcon
End Property

Public Property Get HighLightColor() As OLE_COLOR
 HighLightColor = isHighLightColor
End Property

'* English: Use this color for drawing.
Public Property Let HighLightColor(ByVal theColor As OLE_COLOR)
 '* Color de fondo cuando el mouse pasa sobre el Objeto.
 isHighLightColor = ConvertSystemColor(theColor)
 Call PropertyChanged("HighLightColor")
 Call Refresh(isState)
End Property

Public Property Get HotColor() As OLE_COLOR
 HotColor = isHotColor
End Property

'* English: Use this color for drawing.
Public Property Let HotColor(ByVal theColor As OLE_COLOR)
 '* Color de fondo cuando se tiene presionado el Objeto.
 isHotColor = ConvertSystemColor(theColor)
 Call PropertyChanged("HotColor")
 Call Refresh(isState)
End Property

Public Property Get HotTitle() As Boolean
 HotTitle = isHotTitle
End Property

'* English: Use this color for drawing.
Public Property Let HotTitle(ByVal theTitle As Boolean)
 '* Color de fondo cuando se tiene presionado el Objeto.
 isHotTitle = theTitle
 Call PropertyChanged("HotTitle")
End Property

'* English: Returns a handle to the control.
Public Property Get hWnd() As Long
 '* Devuelve el controlador del control.
 hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = UserControl.MouseIcon
End Property

'* English: Sets a custom mouse icon.
Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)
 '* Devuelve ó establece un icono de mouse personalizado.
 Set UserControl.MouseIcon = MouseIcon
 Call PropertyChanged("MouseIcon")
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = UserControl.MousePointer
End Property

'* English: Returns/Sets the type of mouse pointer displayed when over part of an object.
Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)
 '* Devuelve ó establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.
 UserControl.MousePointer = MousePointer
 Call PropertyChanged("MousePointer")
End Property

Public Property Get MultiLine() As Boolean
 MultiLine = isMultiLine
End Property

'* English: Returns/Sets if the text is shown in multiple lines.
Public Property Let MultiLine(ByVal theMultiLine As Boolean)
 '* Devuelve ó establece si el texto se muestra en múltiples líneas.
 isMultiLine = theMultiLine
 Call PropertyChanged("MultiLine")
 Call Refresh(isState)
End Property

Public Property Get Picture() As StdPicture
 Set Picture = isPicture
End Property

'* English: Returns/Sets the image of the control.
Public Property Set Picture(ByVal thePicture As StdPicture)
 '* Devuelve ó establece la imagen del control.
 Set isPicture = thePicture
 Call PropertyChanged("Picture")
 Call Refresh(isState)
End Property

Public Property Get PictureAlign() As OfficeAlign
 PictureAlign = isPictureAlign
End Property

'* English: Returns/Sets the alignment of the image.
Public Property Let PictureAlign(ByVal theAlign As OfficeAlign)
 '* Devuelve ó establece la alineación de la imagen.
 isPictureAlign = theAlign
 Call PropertyChanged("PictureAlign")
 Call Refresh(isState)
End Property

Public Property Get PictureSize() As Integer
 PictureSize = isPictureSize
End Property

'* English: Returns/Sets the picture size.
Public Property Let PictureSize(ByVal theSize As Integer)
 '* Devuelve ó establece el tamaño de la imagen.
 isPictureSize = theSize
 Call PropertyChanged("PictureSize")
 Call Refresh(isState)
End Property

Public Property Get SetBorder() As Boolean
 SetBorder = isSetBorder
End Property

'* English: Returns/Sets if it's always shown the border.
Public Property Let SetBorder(ByVal theSetBorder As Boolean)
 '* Devuelve ó establece si se muestra siempre un borde.
 isSetBorder = theSetBorder
 Call PropertyChanged("SetBorder")
 Call Refresh(isState)
End Property

Public Property Get SetBorderH() As Boolean
 SetBorderH = isSetBorderH
End Property

'* English: Returns/Sets if it's always shown the Hot border.
Public Property Let SetBorderH(ByVal theSetBorderH As Boolean)
 '* Devuelve ó establece si se muestra siempre un borde.
 isSetBorderH = theSetBorderH
 Call PropertyChanged("SetBorderH")
End Property

Public Property Get SetGradient() As Boolean
 SetGradient = isSetGradient
End Property

'* English: Returns/Sets if the background is gradient.
Public Property Let SetGradient(ByVal theSetGradient As Boolean)
 '* Devuelve ó establece si el fondo es en degradado.
 isSetGradient = theSetGradient
 Call PropertyChanged("SetGradient")
 Call Refresh(isState)
End Property

Public Property Get SetHighLight() As Boolean
 SetHighLight = isSetHighLight
End Property

'* English: Returns/Sets if the background change is shown.
Public Property Let SetHighLight(ByVal theSetHighLight As Boolean)
 '* Devuelve ó establece si se muestra el cambio de fondo.
 isSetHighLight = theSetHighLight
 Call PropertyChanged("SetHighLight")
End Property

Public Property Get ShadowText() As Boolean
 ShadowText = isShadowText
End Property

'* English: Returns/Sets if a shadow is shown in the text of the button.
Public Property Let ShadowText(ByVal theShadowText As Boolean)
 '* Devuelve ó establece si se muestra una sombra en el texto del botón.
 isShadowText = theShadowText
 Call PropertyChanged("ShadowText")
End Property

Public Property Get ShowFocus() As Boolean
 ShowFocus = isShowFocus
End Property

'* English: Do you want to show the focus?
Public Property Let ShowFocus(ByVal theFocus As Boolean)
 '* Permite ver el enfoque del control.
 isShowFocus = theFocus
 Call PropertyChanged("ShowFocus")
End Property

Public Property Get SystemColor() As Boolean
 SystemColor = isSystemColor
End Property

'* English: Take the system color.
Public Property Let SystemColor(ByVal theSystemColor As Boolean)
 '* Toma los colores del Sistema.
 isSystemColor = theSystemColor
 Call PropertyChanged("SystemColor")
 Call Refresh(isState)
End Property

Public Property Get XPosPicture() As Integer
 XPosPicture = isXPos
End Property

'* English: Returns/Sets the Position X of the image.
Public Property Let XPosPicture(ByVal theXPos As Integer)
 '* Devuelve ó establece la Posición X de la imagen.
 isXPos = theXPos
 Call PropertyChanged("XPosPicture")
 Call Refresh(isState)
End Property

Public Property Get YPosPicture() As Integer
 YPosPicture = isYPos
End Property

'* English: Returns/Sets the Position Y of the image.
Public Property Let YPosPicture(ByVal theYPos As Integer)
 '* Devuelve ó establece la Posición Y de la imagen.
 isYPos = theYPos
 Call PropertyChanged("YPosPicture")
 Call Refresh(isState)
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
 If (isEnabled = True) Then RaiseEvent Click
End Sub

Private Sub UserControl_Click()
 If (isHotTitle = False) Then
  Call Refresh(OfficeHighLight)
  RaiseEvent Click
 End If
End Sub

Private Sub UserControl_GotFocus()
 If (isHotTitle = False) Then
  isFocus = True
  Call Refresh(isState)
 End If
End Sub

Private Sub UserControl_InitProperties()
 isAutoSizePic = False
 isBackColor = ConvertSystemColor(defBackColor)
 isBorderColor = ConvertSystemColor(defBorderColor)
 isButtonShape = defShape
 isCaption = Ambient.DisplayName
 isDisabledColor = ConvertSystemColor(defDisabledColor)
 isEnabled = True
 isFontAlign = ACenter
 isForeColor = ConvertSystemColor(defForeColor)
 isHighLightColor = ConvertSystemColor(defHighLightColor)
 isHotColor = ConvertSystemColor(defHotColor)
 isHotTitle = False
 isMultiLine = False
 isPictureAlign = ACenter
 isPictureSize = 16
 isSetBorder = False
 isSetGradient = False
 isSetHighLight = True
 isShadowText = False
 isShowFocus = False
 isSystemColor = True
 isXPos = 4
 isYPos = 4
 m_bGrayIcon = False
 Set g_Font = Ambient.Font
 Set isPicture = Nothing
 ToolActive = False
 ToolBackColor = vbInfoBackground
 ToolCentered = True
 ToolForeColor = vbInfoText
 ToolIcon = 1
 TOOLSTYLE = 1
 ToolTitle = "HACKPRO TM"
 ToolText = Extender.ToolTipText
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13, 32 '* Enter.
   RaiseEvent Click
  Case 37, 38 '* Left Arrow and Up.
   Call SendKeys("+{TAB}")
  Case 39, 40 '* Right Arrow and Down.
   Call SendKeys("{TAB}")
 End Select
End Sub

Private Sub UserControl_LostFocus()
 If (isHotTitle = False) Then
  isFocus = False
  Call Refresh(isState)
 End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If (isHotTitle = False) And (Button = vbLeftButton) And (isEnabled = True) Then
  Call Refresh(OfficeHot)
 End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim tmpState As Integer
 
 If (isEnabled = True) And (isHotTitle = False) Then
  If (IsMouseOver = True) Then
   Call Refresh(isState)
  Else
   tmpState = isState
   Call Refresh(OfficeNormal)
   isState = tmpState
  End If
 End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
 With PropBag
  AutoSizePicture = .ReadProperty("AutoSizePicture", False)
  BackColor = .ReadProperty("BackColor", ConvertSystemColor(defBackColor))
  BorderColor = .ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
  ButtonShape = .ReadProperty("ButtonShape", defShape)
  Caption = .ReadProperty("Caption", Ambient.DisplayName)
  CaptionAlign = .ReadProperty("CaptionAlign", &H0)
  DisabledColor = .ReadProperty("DisabledColor", ConvertSystemColor(defDisabledColor))
  Enabled = .ReadProperty("Enabled", True)
  ForeColor = .ReadProperty("ForeColor", ConvertSystemColor(defForeColor))
  GrayIcon = PropBag.ReadProperty("GrayIcon", True)
  HighLightColor = .ReadProperty("HighlightColor", ConvertSystemColor(defHighLightColor))
  HotColor = .ReadProperty("HotColor", ConvertSystemColor(defHotColor))
  HotTitle = .ReadProperty("HotTitle", False)
  MultiLine = .ReadProperty("MultiLine", False)
  PictureAlign = .ReadProperty("PictureAlign", &H0)
  PictureSize = .ReadProperty("PictureSize", 16)
  SetBorder = .ReadProperty("SetBorder", False)
  SetBorderH = .ReadProperty("SetBorderH", True)
  SetGradient = .ReadProperty("SetGradient", False)
  Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
  SetHighLight = .ReadProperty("SetHighLight", True)
  Set isPicture = .ReadProperty("Picture", Nothing)
  Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
  ShadowText = .ReadProperty("ShadowText", False)
  ShowFocus = .ReadProperty("ShowFocus", False)
  SystemColor = .ReadProperty("SystemColor", True)
  TipActive = .ReadProperty("TipActive", False)
  TipBackColor = .ReadProperty("TipBackColor", vbInfoBackground)
  TipCentered = .ReadProperty("TipCentered", True)
  TipForeColor = .ReadProperty("TipForeColor", vbInfoText)
  TipIcon = .ReadProperty("TipIcon", 1)
  TipStyle = .ReadProperty("TipStyle", 1)
  TipTitle = .ReadProperty("TipTitle", "HACKPRO TM")
  TipText = .ReadProperty("TipText", "")
  UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
  XPosPicture = .ReadProperty("XPosPicture", 4)
  YPosPicture = .ReadProperty("YPosPicture", 4)
 End With
 If (Ambient.UserMode = True) Then
  bTrack = True
  bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  If Not (bTrackUser32 = True) Then
   If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
    bTrack = False
   End If
  End If
  If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
   '* Start subclassing the UserControl.
   Call sc_Subclass(hWnd)
   Call sc_AddMsg(hWnd, WM_MOUSEMOVE)
   Call sc_AddMsg(hWnd, WM_MOUSELEAVE)
   Call sc_AddMsg(hWnd, WM_THEMECHANGED)
   Call sc_AddMsg(hWnd, WM_SYSCOLORCHANGE)
  End If
 End If
End Sub

Private Sub UserControl_Resize()
 If (isHotTitle = False) Then Call Refresh(isState) '* Call the Refresh Sub.
End Sub

'* The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
On Error GoTo Catch
 Call TipRemove
 If (Ambient.UserMode = True) Then Call sc_Terminate '* Stop all subclassing.
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
 With PropBag
  Call .WriteProperty("AutoSizePicture", isAutoSizePic, False)
  Call .WriteProperty("BackColor", isBackColor, ConvertSystemColor(defBackColor))
  Call .WriteProperty("BorderColor", isBorderColor, ConvertSystemColor(defBorderColor))
  Call .WriteProperty("ButtonShape", isButtonShape, defShape)
  Call .WriteProperty("Caption", isCaption, Ambient.DisplayName)
  Call .WriteProperty("CaptionAlign", isFontAlign, &H0)
  Call .WriteProperty("DisabledColor", isDisabledColor, ConvertSystemColor(defDisabledColor))
  Call .WriteProperty("Enabled", isEnabled, True)
  Call .WriteProperty("Font", g_Font, Ambient.Font)
  Call .WriteProperty("ForeColor", isForeColor, ConvertSystemColor(defForeColor))
  Call .WriteProperty("GrayIcon", m_bGrayIcon, True)
  Call .WriteProperty("HighlightColor", isHighLightColor, ConvertSystemColor(defHighLightColor))
  Call .WriteProperty("HotColor", isHotColor, ConvertSystemColor(defHotColor))
  Call .WriteProperty("HotTitle", isHotTitle, False)
  Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call .WriteProperty("MousePointer", MousePointer, vbDefault)
  Call .WriteProperty("MultiLine", isMultiLine, False)
  Call .WriteProperty("Picture", isPicture, Nothing)
  Call .WriteProperty("PictureAlign", isPictureAlign, &H0)
  Call .WriteProperty("PictureSize", isPictureSize, 16)
  Call .WriteProperty("SetBorder", isSetBorder, False)
  Call .WriteProperty("SetBorderH", isSetBorderH, True)
  Call .WriteProperty("SetGradient", isSetGradient, False)
  Call .WriteProperty("SetHighLight", isSetHighLight, True)
  Call .WriteProperty("ShadowText", isShadowText, False)
  Call .WriteProperty("ShowFocus", isShowFocus, False)
  Call .WriteProperty("SystemColor", isSystemColor, True)
  Call .WriteProperty("TipActive", ToolActive, False)
  Call .WriteProperty("TipBackColor", ToolBackColor, vbInfoBackground)
  Call .WriteProperty("TipCentered", ToolCentered, True)
  Call .WriteProperty("TipForeColor", ToolForeColor, vbInfoText)
  Call .WriteProperty("TipIcon", ToolIcon, 1)
  Call .WriteProperty("TipStyle", TOOLSTYLE, 1)
  Call .WriteProperty("TipText", ToolText, "")
  Call .WriteProperty("TipTitle", ToolTitle, "HACKPRO TM")
  Call .WriteProperty("XPosPicture", isXPos, 4)
  Call .WriteProperty("YPosPicture", isYPos, 4)
 End With
On Error GoTo 0
End Sub

'*******************************************************'
'* Private Subs and Functions.                         *'
'*******************************************************'

'* English: Paints lines in a simple and faster.
Private Sub APILine(ByVal whDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
 Dim pt As POINTAPI, hPen As Long, hPenOld As Long
 
 '* Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(whDC, hPen)
 Call MoveToEx(whDC, X1, Y1, pt)
 Call LineTo(whDC, X2, Y2)
 Call SelectObject(whDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

'* English: Convert Long to System Color.
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 '* Convierte un long en un color del sistema.
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

'* English: Paints a rectangle with oval border.
Private Sub DrawBox(ByVal hdc As Long, ByVal Offset As Long, ByVal Radius As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long, ByVal isWidth As Long, ByVal isHeight As Long)
 Dim pRect As RECT, hPen As Long, hBrush As Long
 
 '* Crea un rectángulo con border ovalados.
On Error Resume Next
 pRect.xLeft = -4
 pRect.xRight = isWidth - IIf(isCaption = "", 1, -2)
 pRect.xTop = -3
 pRect.xBottom = isHeight - 1
 hPen = SelectObject(hdc, CreatePen(PS_SOLID, 1, ColorBorder))
 hBrush = SelectObject(hdc, CreateSolidBrush(ColorFill))
 Call InflateRect(pRect, -Offset, -Offset)
 Call RoundRect(hdc, pRect.xLeft, pRect.xTop, pRect.xRight, pRect.xBottom, Radius, Radius)
 Call InflateRect(pRect, Offset, Offset)
 Call DeleteObject(SelectObject(hdc, hPen))
 Call DeleteObject(SelectObject(hdc, hBrush))
On Error GoTo 0
End Sub

'* English: Draw the text on the Object.
Private Sub DrawCaption(ByVal iColor1 As Long, ByVal iColor2 As Long)
 Dim lColor As Long, isFAlign As Long
   
 '* Dibuja el texto sobre el Objeto.
 If (isMultiLine = True) Then lColor = DT_WORDBREAK Else lColor = DT_SINGLELINE
 Select Case isFontAlign
  Case ACenter
   isFAlign = DT_CENTER Or DT_VCENTER Or lColor Or DT_WORD_ELLIPSIS
  Case ALeft
   isFAlign = DT_VCENTER Or DT_LEFT Or lColor Or DT_WORD_ELLIPSIS
  Case ARight
   isFAlign = DT_VCENTER Or DT_RIGHT Or lColor Or DT_WORD_ELLIPSIS
  Case ATop
   isFAlign = DT_CENTER Or DT_TOP Or lColor Or DT_WORD_ELLIPSIS
  Case ABottom
   isFAlign = DT_CENTER Or DT_BOTTOM Or lColor Or DT_WORD_ELLIPSIS
 End Select
 If (isState <> OfficeDisabled) Then
  lColor = iColor2
 Else
  lColor = iColor1
 End If
 If (isShadowText = True) And ((isState = &H1) Or (isState = &H2)) Then
  isTxtRect.xLeft = isTxtRect.xLeft + 1.5
  isTxtRect.xTop = isTxtRect.xTop + 1.5
  Call SetTextColor(UserControl.hdc, ShiftColorOXP(lColor))
  Call DrawText(UserControl.hdc, isCaption, -1, isTxtRect, isFAlign)
  isTxtRect.xLeft = isTxtRect.xLeft - 1.5
  isTxtRect.xTop = isTxtRect.xTop - 1.5
 End If
 Call SetTextColor(UserControl.hdc, lColor)
 Call DrawText(UserControl.hdc, isCaption, -1, isTxtRect, isFAlign)
End Sub

'* English: Show focus of control.
Private Sub DrawFocus()
 Dim iPos As Integer
 
 '* Muestra el enfoque del control.
 If (isFocus = True) And (isShowFocus = True) Then
  If (isButtonShape = &H0) Then '* Shape Rectangle.
   Call DrawFocusRect(UserControl.hdc, RectButton)
  Else
   For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4)
    Call SetPixel(UserControl.hdc, iPos, RectButton.xTop + 1, &H1DD6B7)
    Call SetPixel(UserControl.hdc, iPos, RectButton.xTop + isHeight - 3, &H1DD6B7)
   Next
   For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5
    Call SetPixel(UserControl.hdc, RectButton.xLeft, iPos, &H1DD6B7)
    Call SetPixel(UserControl.hdc, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, &H1DD6B7)
   Next
   For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4) Step 2
    Call SetPixel(UserControl.hdc, iPos, RectButton.xTop + 1, &H24427A)
    Call SetPixel(UserControl.hdc, iPos, RectButton.xTop + isHeight - 3, &H24427A)
   Next
   For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5 Step 2
    Call SetPixel(UserControl.hdc, RectButton.xLeft, iPos, &H24427A)
    Call SetPixel(UserControl.hdc, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, &H24427A)
   Next
   Call SetPixel(UserControl.hdc, RectButton.xLeft + 1, 2, vbBlack)
   Call SetPixel(UserControl.hdc, RectButton.xRight - IIf(isCaption = "", 5, 2), 2, vbBlack)
   Call SetPixel(UserControl.hdc, RectButton.xLeft + 1, RectButton.xTop + isHeight - 4, vbBlack)
   Call SetPixel(UserControl.hdc, RectButton.xRight - IIf(isCaption = "", 5, 2), RectButton.xTop + isHeight - 4, vbBlack)
  End If
 End If
End Sub

'* English: Draws a degraded one in vertical form.
Private Sub DrawVGradient(ByVal whDC As Long, ByVal lEndColor As Long, ByVal lStartColor As Long, ByVal x As Long, ByVal y As Long, ByVal X2 As Long, ByVal Y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, ni As Long
 Dim sR As Single, sG As Single, Sb As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* Dibuja un degradado en forma vertical.
 sR = (lStartColor And &HFF)
 sG = (lStartColor \ &H100) And &HFF
 Sb = (lStartColor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / Y2
 dG = (sG - eG) / Y2
 dB = (Sb - eB) / Y2
 For ni = 0 To Y2
  Call APILine(whDC, x, y + ni, X2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
 Next
End Sub

'* English: Draw a rectangle area with a specific color.
Private Sub DrawRectangle(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long, Optional ByVal SetBackground As Boolean = True)
 Dim hBrush As Long, TempRect As RECT

 '* Crea un área rectangular con un color específico.
 TempRect.xLeft = x
 TempRect.xTop = y
 TempRect.xRight = x + Width
 TempRect.xBottom = y + Height
 hBrush = CreateSolidBrush(ColorBorder)
 Call FrameRect(hdc, TempRect, hBrush)
 Call DeleteObject(hBrush)
 If (SetBackground = True) Then
  TempRect.xLeft = x + 1
  TempRect.xTop = y + 1
  TempRect.xRight = x + Width - 1
  TempRect.xBottom = y + Height - 1
  hBrush = CreateSolidBrush(ColorFill)
  Call FillRect(hdc, TempRect, hBrush)
  Call DeleteObject(hBrush)
 End If
End Sub

'* English: Draw a picture in the Object.
Private Sub DrawPicture()
 Dim isType As Long, isValue As Long
 
 '* Crea la imagen sobre el Objeto.
On Error Resume Next
 If Not (isPicture Is Nothing) Then
  If (Picture <> 0) Then
   Dim iX As Long, iY As Long
   
   If (isPictureSize <= 0) Then isPictureSize = 16
   Select Case isPicture.Type
    Case 1, 4: isType = DST_BITMAP
    Case 3:    isType = DST_ICON
   End Select
   If (isPictureAlign = &H0) Then
    iX = (isWidth - isPictureSize) / 2
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H1) Then
    iX = isXPos
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H2) Then
    iX = isWidth - isPictureSize - isXPos
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H3) Then
    iX = (isWidth - isPictureSize) / 2
    iY = isYPos
   ElseIf (isPictureAlign = &H4) Then
    iX = (isWidth - isPictureSize) / 2
    iY = isHeight - isPictureSize - isYPos
   End If
  End If
  If (isEnabled = False) Then
   isValue = DSS_DISABLED
   If (m_bGrayIcon = False) Then
    Call DrawState(UserControl.hdc, 0, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or isValue)
   Else
    Call RenderIconGrayscale(UserControl.hdc, isPicture.handle, iX, iY, isPictureSize, isPictureSize)
   End If
  Else
   isValue = DSS_NORMAL
   If (isState = OfficeHot) Then
    iX = iX - 1
    iY = iY - 1
   ElseIf (isState = OfficeHighLight) Then
    isValue = CreateSolidBrush(RGB(136, 141, 157))
    Call DrawState(UserControl.hdc, isValue, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or DSS_MONO)
    iX = iX - 2
    iY = iY - 2
    isValue = DSS_NORMAL
    Call DrawState(UserControl.hdc, 0, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or isValue)
    Call DeleteObject(isValue)
    Exit Sub
   End If
   Call RenderIconGrayscale(UserControl.hdc, isPicture.handle, iX, iY, isPictureSize, isPictureSize, False)
  End If
 End If
End Sub

'* English: Return, if the mouse is over the Object.
Private Function IsMouseOver() As Boolean
 Dim pt As POINTAPI
 
 '* Devuelve si el mouse esta sobre el objeto.
 Call GetCursorPos(pt)
 IsMouseOver = (WindowFromPoint(pt.x, pt.y) = hWnd)
End Function

'* English: Executable file or a document file.
Public Function OpenLink(ByVal sLink As String) As Long
 '* Ejecuta un archivo ó documento cualquiera.
On Error Resume Next
 OpenLink = ShellExecute(Parent.hWnd, vbNullString, sLink, vbNullString, "C:\", SW_SHOWNORMAL)
On Error GoTo 0
End Function

'* English: Draw appearance of the control.
Private Sub Refresh(Optional ByVal State As OfficeState = 0)
 Dim lColor  As Long, lBase   As Long, iColor1 As Long
 Dim iColor2 As Long, iColor3 As Long, iColor4 As Long
 Dim iColor5 As Long, iColor6 As Long, lBase1  As Integer
 
 '* Crea la apariencia del control.
 If (isEnabled = False) Then State = OfficeDisabled
 If (isSystemColor = False) Then
  iColor1 = isBackColor
  iColor2 = isBorderColor
  iColor3 = isDisabledColor
  iColor4 = isForeColor
  iColor5 = isHighLightColor
  iColor6 = isHotColor
 Else
  iColor1 = ConvertSystemColor(defBackColor)
  iColor2 = ConvertSystemColor(defBorderColor)
  iColor3 = ConvertSystemColor(defDisabledColor)
  iColor4 = ConvertSystemColor(defForeColor)
  iColor5 = ConvertSystemColor(defHighLightColor)
  iColor6 = ConvertSystemColor(defHotColor)
 End If
 If (isEnabled = False) Then iColor2 = iColor3
 With UserControl
  isHeight = .ScaleHeight
  isWidth = .ScaleWidth
  .AutoRedraw = True
  .ScaleMode = vbPixels
  .Cls
 On Error Resume Next
  Set .Font = g_Font
  Call GetClientRect(.hWnd, RectButton)
  Call GetClientRect(.hWnd, isTxtRect)
  .BackColor = iColor1
  lBase = &HB0
  lBase1 = 1
  If Not (isButtonShape = &H0) Then lBase1 = 4
  'If (State > &H0) And (State < &H3) And (isSetGradient = True) Then State = &H0
  Select Case State
   Case &H0 '* Normal State.
    If (isSetGradient = True) Then Call DrawVGradient(.hdc, iColor1, ShiftColorOXP(iColor1, &H72), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
    If (isSetBorder = True) Then
     If (isButtonShape = &H0) Then
      Call DrawRectangle(.hdc, 0, 0, isWidth, isHeight, iColor1, iColor2, IIf(isSetGradient = True, False, True))
     Else
      Call DrawBox(.hdc, 4, 5, iColor1, iColor2, RectButton.xRight + 2, RectButton.xBottom + 3)
     End If
    ElseIf (isSetGradient = False) Then
     Call DrawRectangle(.hdc, 0, 0, isWidth, isHeight, iColor1, iColor1)
    End If
   Case &H1, &H2 '* HighLight or Hot State.
    If (isSetHighLight = True) Then
     If (State = &H1) Then
      lColor = ShiftColorOXP(iColor5, &H40)
      If (isSetGradient = True) Then Call DrawVGradient(.hdc, iColor1, ShiftColorOXP(iColor5, &H122), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
     Else
      lColor = ShiftColorOXP(iColor6, &H10)
      lBase = &H9C
      If (isSetGradient = True) Then Call DrawVGradient(.hdc, iColor1, ShiftColorOXP(iColor6, &H40), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
     End If
    ElseIf (isSetBorderH = True) Then
     lColor = iColor1
     lBase = 0
    End If
    If (isSetBorderH = True) And (isButtonShape = &H0) Then
     Call DrawRectangle(.hdc, 0, 0, isWidth, isHeight, ShiftColorOXP(lColor, lBase), iColor2, IIf(isSetGradient = True, False, True))
    ElseIf (isSetBorderH = True) Then
     Call DrawBox(.hdc, 4, 5, ShiftColorOXP(lColor, lBase), iColor2, RectButton.xRight + 2, RectButton.xBottom + 3)
    End If
   Case &H3 '* Disabled State.
    lColor = iColor3
    If (isSetBorder = True) Then
     If (isButtonShape = &H0) Then
      Call DrawRectangle(.hdc, 0, 0, isWidth, isHeight, iColor1, lColor)
     Else
      Call DrawBox(.hdc, 4, 5, iColor1, lColor, RectButton.xRight + 2, RectButton.xBottom + 3)
     End If
    End If
  End Select
  isState = State
  If (isAutoSizePic = True) Then
   .Width = isPicture.Width
   .Height = isPicture.Height
   isHeight = .ScaleHeight
   isWidth = .ScaleWidth
  End If
  Call DrawCaption(iColor3, iColor4)
  Call DrawPicture
  If (isState <> &H3) Then Call DrawFocus
 End With
End Sub

'* English: Returns or sets a string that contains the keys that will act as the access keys (or hot keys for the control.)
Private Sub SetAccessKey(ByVal Caption As String)
 Dim AmperSandPos As Long, isText As String

 '* Devuelve ó establece una cadena que contiene las teclas que funcionarán como teclas de acceso (o teclas aceleradoras) del control.
 With UserControl
  .AccessKeys = ""
  If (Len(Caption) > 1) Then
   AmperSandPos = InStr(1, Caption, "&", vbTextCompare)
   If (AmperSandPos < Len(Caption)) And (AmperSandPos > 0) Then
    isText = Mid$(Caption, AmperSandPos + 1, 1)
    If (isText <> "&") Then
     .AccessKeys = LCase$(isText)
    Else
     AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
     isText = Mid$(Caption, AmperSandPos + 1, 1)
     If (isText <> "&") Then .AccessKeys = LCase$(isText)
    End If
   End If
  End If
 End With
End Sub

'* English: Shift a color.
Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
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

Public Property Get TipActive() As Boolean
 '* Retrieving value of a property, Boolean responce (true/false).
 '* Syntax: BooleanVar = object.TipActive.
 TipActive = ToolActive
End Property

Public Property Let TipActive(ByVal ToolData As Boolean)
 '* If True, activate (show) ToolTip, False deactivate (hide) tool tip.
 '* Syntax: object.TipActive = True/False.
 ToolActive = ToolData
 Call PropertyChanged("TipActive")
End Property

Public Property Get TipBackColor() As OLE_COLOR
 '* Retrieving value of a property, returns RGB as Long.
 '* Syntax: LongVar = object.BackColor.
 TipBackColor = ToolBackColor
End Property

Public Property Let TipBackColor(ByVal ToolData As OLE_COLOR)
 '* Assigning a value to the property, set RGB value as Long.
 '* Syntax: object.BackColor = RGB (as Long). Since 0 is _
    Black (no RGB), and the API thinks 0 is the default _
    color ("off" yellow), we need to "fudge" Black a bit _
    (yes set bit "1" to "1",). I couldn't resist the _
    pun!. So, in module or form code, if setting to Black, _
    make it "1", if restoring the default color, make it _
    "0".
 ToolBackColor = ConvertSystemColor(ToolData)
 Call PropertyChanged("TipBackColor")
End Property

Public Property Get TipCentered() As Boolean
 '* Retrieving value of a property, returns Boolean true/false.
 '* Syntax: BooleanVar = object.TipCentered.
 TipCentered = ToolCentered
End Property

Public Property Let TipCentered(ByVal ToolData As Boolean)
 '* Assigning a value to the property, Set Boolean true/false if ToolTip. _
    Is TipCentered on the parent control.
 '* Syntax: object.TipCentered = True/False.
 ToolCentered = ToolData
 Call PropertyChanged("TipCentered")
End Property

Public Property Get TipForeColor() As OLE_COLOR
 '* Retrieving value of a property, returns RGB value as Long.
 '* Syntax: LongVar = object.ForeColor.
 TipForeColor = ToolForeColor
End Property

Public Property Let TipForeColor(ByVal ToolData As OLE_COLOR)
 '* Assigning a value to the property, set RGB value as Long.
 '* Syntax: object.ForeColor = RGB(As Long).
 '* Since 0 is Black (no RGB), and the API thinks 0 is _
    the default color ("off" yellow), we need to "fudge" _
    Black a bit (yes set bit "1" to "1",). I couldn't _
    resist the pun!. So, in module or form code, if _
    setting to Black, make it "1" if restoring _
    the default color, make it "0".
 '* Syntax: object.ForeColor = RGB(as long).
 ToolForeColor = ConvertSystemColor(ToolData)
 Call PropertyChanged("TipForeColor")
End Property

Public Property Get TipIcon() As ToolIconType
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipIcon.
 TipIcon = ToolIcon
End Property

Public Property Let TipIcon(ByVal ToolData As ToolIconType)
 '* Assigning a value to the property, set TipIcon TipStyle with type var.
 '* Syntax: object.TipIcon = IconStyle.
 '* TipIcon Styles are: INFO, WARNING And ERROR (TipNoIcom, TipIconInfo, TipIconWarning, TipIconError).
 ToolIcon = ToolData
 Call PropertyChanged("TipIcon")
End Property

Public Property Get TipStyle() As ToolStyleEnum
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipStyle.
 TipStyle = TOOLSTYLE
End Property

Public Property Let TipStyle(ByVal ToolData As ToolStyleEnum)
 '* Assigning a value to the property, set TipStyle param Standard or Balloon
 '* Syntax: object.TipStyle = TipStyle.
 TOOLSTYLE = ToolData
 Call PropertyChanged("TipStyle")
End Property

Public Property Get TipText() As String
 '* Retrieving value of a property, returns string..
 '* Syntax: StringVar = object.TipText.
 TipText = ToolText
End Property

Public Property Let TipText(ByVal ToolData As String)
 '* Assigning a value to the property, Set as String.
 '* Syntax: object.TipText = StringVar.
 '* Multi line Tips are enabled in the Create sub.
 '* To change lines, just add a vbCrLF between text.
 '* ex. object.TipText = "Line 1 text" & vbCrLF & _
    "Line 2 text".
 ToolText = ToolData
 Call PropertyChanged("TipText")
End Property

Public Property Get TipTitle() As String
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipTitle.
 TipTitle = ToolTitle
End Property

Public Property Let TipTitle(ByVal ToolData As String)
 '* Assigning a value to the property, set as string.
 '* Syntax: object.TipTitle = StringVar.
 ToolTitle = ToolData
 Call PropertyChanged("TipTitle")
End Property

'* Private sub used with Create and Update subs/functions.
Private Sub CreateToolTip()
 Dim lpRect As RECT, lWinStyle As Long
 
 '* If Tool Tip already made, destroy it and reconstruct.
 Call TipRemove
 lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
 '* Create Baloon TipStyle if desired.
 If (TOOLSTYLE = StyleBalloon) Then lWinStyle = lWinStyle Or TTS_BALLOON
 '* The parent control has to be set first.
 If (UserControl.hWnd <> &H0) Then
  m_ltthWnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
  Call SendMessage(m_ltthWnd, TTM_ACTIVATE, CInt(ToolActive), TI)
  '* Make our ToolTip window a topmost window.
  Call SetWindowPos(m_ltthWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
  '* Get the rectangle of the parent control.
  Call GetClientRect(UserControl.hWnd, lpRect)
  '* Now set up our ToolTip info structure.
  With TI
   '* If we want it TipCentered, then set that flag.
   If (ToolCentered = True) Then
    .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
   Else
    .lFlags = TTF_SUBCLASS
   End If
   '* Set the hWnd prop to our Parent Control's hWnd.
   .lhWnd = UserControl.hWnd
   .lId = 0
   .hInstance = App.hInstance
   .lpRect = lpRect
   .lpStr = ToolText
  End With
  '* Add the ToolTip Structure.
  Call SendMessage(m_ltthWnd, TTM_ADDTOOLA, 0&, TI)
  '* Set Max Width to 32 characters, and enable Multi Line Tool Tips.
  Call SendMessage(m_ltthWnd, TTM_SETMAXTIPWIDTH, 0&, &H20)
  If (ToolIcon <> TipNoIcon) Or (ToolTitle <> vbNullString) Then
   '* If we want a TipTitle or we want an TipIcon.
   Call SendMessage(m_ltthWnd, TTM_SETTITLE, CLng(ToolIcon), ByVal ToolTitle)
  End If
  If (ToolForeColor <> Empty) Then
   '* 0 (zero) or Null is seen by the API as the default color. _
      See ForeColor property for more datails.
   Call SendMessage(m_ltthWnd, TTM_SETTIPTEXTCOLOR, ToolForeColor, 0&)
  End If
  If (ToolBackColor <> Empty) Then
   '* 0 (zero) or Null is seen by the API as the default color. _
      See BackColor property for more datails.
   Call SendMessage(m_ltthWnd, TTM_SETTIPBKCOLOR, ToolBackColor, 0&)
  End If
 End If
End Sub

Public Sub TipRemove()
 '* Kills Tool Tip Object.
 If (m_ltthWnd <> 0) Then Call DestroyWindow(m_ltthWnd)
End Sub

Private Sub UpDate()
 '* Used to update tooltip parameters that require reconfiguration of _
    subclass to envoke.
 If (ToolActive = True) Then Call CreateToolTip '* Refresh the object.
End Sub

' See post: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58622&lngWId=1
' Thanks MArio Florez.
Private Function RenderIconGrayscale(ByVal Dest_hDC As Long, ByVal hIcon As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim hBMP_Mask As Long, hBMP_Image As Long
 Dim hBMP_Prev As Long, hIcon_Temp As Long
 Dim hDC_Temp  As Long

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hIcon = 0) Then Exit Function
 ' Extract the bitmaps from the icon
 If (GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False) Then Exit Function
 ' Create a memory DC to work with
 hDC_Temp = CreateCompatibleDC(0)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' Make the image bitmap gradient
 If (RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , GrayC) = False) Then GoTo CleanUp
 ' Extract the gradient bitmap out of the DC
 Call SelectObject(hDC_Temp, hBMP_Prev)
 ' Take the newly gradient bitmap and make a gradient icon from it
 hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
 If (hIcon_Temp = 0) Then GoTo CleanUp
 ' Draw the newly created gradient icon onto the specified DC
 If (DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0) Then
  RenderIconGrayscale = True
 End If
CleanUp:
 Call DestroyIcon(hIcon_Temp): hIcon_Temp = 0
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 Call DeleteObject(hBMP_Mask): hBMP_Mask = 0
 Call DeleteObject(hBMP_Image): hBMP_Image = 0
End Function

Public Function GetIconBitmaps(ByVal hIcon As Long, ByRef Return_hBmpMask As Long, ByRef Return_hBmpImage As Long) As Boolean
 Dim TempICONINFO As ICONINFO

 If (GetIconInfo(hIcon, TempICONINFO) = 0) Then Exit Function
 Return_hBmpMask = TempICONINFO.hbmMask
 Return_hBmpImage = TempICONINFO.hbmColor
 GetIconBitmaps = True
End Function

'=============================================================================================================
Private Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim TempBITMAP As BITMAP, hScreen   As Long
 Dim hDC_Temp   As Long, hBMP_Prev   As Long
 Dim MyCounterX As Long, MyCounterY  As Long
 Dim NewColor   As Long, hNewPicture As Long
 Dim DeletePic  As Boolean

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hBitmap = 0) Then Exit Function
 ' Get the handle to the screen DC
 hScreen = GetDC(0)
 If (hScreen = 0) Then Exit Function
 ' Create a memory DC to work with the picture
 hDC_Temp = CreateCompatibleDC(hScreen)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' If the user specifies NOT to alter the original, then make a copy of it to use
 DeletePic = False
 hNewPicture = hBitmap
 ' Select the bitmap into the DC
 hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
 ' Get the height / width of the bitmap in pixels
 If (GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0) Then GoTo CleanUp
 If (TempBITMAP.bmHeight <= 0) Or (TempBITMAP.bmWidth <= 0) Then GoTo CleanUp
 ' Loop through each pixel and conver it to it's grayscale equivelant
 If (GrayC = True) Then
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
   For MyCounterY = 0 To TempBITMAP.bmHeight - 1
    NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
    If (NewColor <> -1) Then
     Select Case NewColor
      ' If the color is already a grey shade, no need to convert it
      Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
       NewColor = NewColor
      Case Else
       NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + 0.11 * ((NewColor \ 65536) Mod 256)
       NewColor = RGB(NewColor, NewColor, NewColor)
     End Select
     Call SetPixel(hDC_Temp, MyCounterX, MyCounterY, NewColor)
    End If
   Next
  Next
 End If
 ' Display the picture on the specified hDC
 Call BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy)
 RenderBitmapGrayscale = True
CleanUp:
 Call ReleaseDC(0, hScreen): hScreen = 0
 Call SelectObject(hDC_Temp, hBMP_Prev)
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 If (DeletePic = True) Then
  Call DeleteObject(hNewPicture)
  hNewPicture = 0
 End If
End Function

Private Function CreateIconFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long) As Long
 Dim TempICONINFO As ICONINFO

 If (hBMP_Mask = 0) Or (hBMP_Image = 0) Then Exit Function
 TempICONINFO.fIcon = 1
 TempICONINFO.hbmMask = hBMP_Mask
 TempICONINFO.hbmColor = hBMP_Image
 CreateIconFromBMP = CreateIconIndirect(TempICONINFO)
End Function

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hMod
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      TrackMouseEvent tme
    Else
      TrackMouseEventComCtl tme
    End If
  End If
End Sub

'-uSelfSub code-----------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long) As Boolean             'Subclass the specified window handle
  Dim nAddr As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError "sc_Subclass", "Invalid window handle"
  End If
  
  If z_hWnds Is Nothing Then
    RtlMoveMemory VarPtr(z_DataDataPtr), VarPtrArray(z_Data), 4             'Get the address of z_Data()'s SafeArray header
    z_DataDataPtr = z_DataDataPtr + 12                                      'Bump the address to point to the pvData data pointer
    RtlMoveMemory VarPtr(z_DataOrigData), z_DataDataPtr, 4                  'Get the value of z_Data()'s SafeArray pvData data pointer
  
    nAddr = zGetCallback                                                    'Get the address of this UserControl's zWndProc callback routine
    
    'Initialise the machine-code thunk
    z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
    
    zMap VarPtr(z_Code(0))                                                  'Map the address of z_Code()'s first element to the z_Data() array
    z_Data(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    z_Data(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                  'Store CallWindowProc function address in the thunk data
    z_Data(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                   'Store the SetWindowLong function address in the thunk data
    z_Data(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                   'Store the VirtualFree function address in the thunk data
    z_Data(IDX_ME) = ObjPtr(Me)                                             'Store my object address in the thunk data
    z_Data(IDX_CALLBACK) = nAddr                                            'Store the zWndProc address in the thunk data
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    
    Set z_hWnds = New Collection                                            'Create the window-handle/thunk-memory-address collection
  End If

  nAddr = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                    'Allocate executable memory
  RtlMoveMemory nAddr, VarPtr(z_Code(0)), CODE_LEN                          'Copy the machine-code to the allocated memory

  On Error GoTo Catch                                                       'Catch double subclassing
    z_hWnds.Add nAddr, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
  On Error GoTo 0

  zMap nAddr                                                                'Map z_Data() to the subclass thunk machine-code
  z_Data(IDX_EBX) = nAddr                                                   'Patch the data address
  z_Data(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
  z_Data(IDX_BTABLE) = nAddr + CODE_LEN                                     'Store the address of the before table in the thunk data
  z_Data(IDX_ATABLE) = z_Data(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
  nAddr = nAddr + WNDPROC_OFF                                               'Execution address of the thunk's WndProc
  z_Data(IDX_WNDPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, nAddr)        'Set the new WndProc and store the original WndProc in the thunk data
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
  sc_Subclass = True                                                        'Indicate success
  Exit Function                                                             'Exit

Catch:
  zError "sc_Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i     As Long
  Dim nAddr As Long

  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
  Else
    With z_hWnds
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        nAddr = .Item(i)                                                    'Map z_Data() to the hWnd thunk address
        If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already freed itself
          zMap nAddr                                                        'Map the thunk memory to the z_Data() array
          sc_UnSubclass z_Data(IDX_HWND)                                    'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    
    Set z_hWnds = Nothing                                                   'Destroy the window-handle/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Public Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "UnSubclass", "Subclassing hasn't been started", False
  Else
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                              'Delete all before messages
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                              'Delete all after messages
    zMap_hWnd lng_hWnd                                                      'Map the thunk memory to the z_Data() array
    z_Data(IDX_SHUTDOWN) = -1                                               'Set the shutdown indicator
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    z_hWnds.Remove "h" & lng_hWnd                                           'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
  End If

  If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
  End If

  If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  sc_CallOrigWndProc = CallWindowProcA(z_Data(IDX_WNDPROC), lng_hWnd, uMsg, _
                                                            wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = z_Data(0)                                                      'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
      Exit Sub
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = 0 Then                                                 'If the element is free...
        z_Data(i) = uMsg                                                    'Use this element
        Exit Sub                                                            'Bail
      ElseIf z_Data(i) = uMsg Then                                          'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    z_Data(nCount) = uMsg                                                   'Store the message in the appended table entry
  End If

  z_Data(0) = nCount                                                        'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    z_Data(0) = 0                                                           'Zero the table entry count
  Else
    nCount = z_Data(0)                                                      'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = uMsg Then                                              'If the message is found...
        z_Data(i) = 0                                                       'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  
  MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  If bEnd Then
    End
  End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map z_Data() to the specified address
Private Sub zMap(ByVal nAddr As Long)
  RtlMoveMemory z_DataDataPtr, VarPtr(nAddr), 4                             'Set z_Data()'s SafeArray data pointer to the specified address
End Sub

'Map z_Data() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started", True
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    zMap_hWnd = z_hWnds("h" & lng_hWnd)                                     'Get the thunk address
    zMap zMap_hWnd                                                          'Map z_Data() to the thunk address
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Determine the address of the final private method, zWndProc
Private Function zGetCallback() As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte                                                         'Value pointed at by the vTable entry
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Upper bound of z_Data()
  Dim k     As Long                                                         'vTable entry value
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                                'Get the address of my vTable
  zMap nAddr + &H7A4                                                        'Map z_Data() to the first possible vTable entry for a UserControl

  j = UBound(z_Data())                                                      'Get the upper bound of z_Data()
  
  For i = 0 To j                                                            'Loop through the vTable looking for the first method entry
    k = z_Data(i)                                                           'Get the vTable entry
    
    If k <> 0 Then                                                          'Skip implemented interface entries
      RtlMoveMemory VarPtr(bVal), k, 1                                      'Get the first byte pointed to by this vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'If a method (pcode or native)
        bSub = bVal                                                         'Store which of the method markers was found (pcode or native)
        Exit For                                                            'Method found, quit loop and scan methods
      End If
    End If
  Next i
  
  For i = i To j                                                            'Loop through the remaining vTable entries
    k = z_Data(i)                                                           'Get the vTable entry
    
    If IsBadCodePtr(k) Then                                                 'Is the vTable entry an invalid code address...
      Exit For                                                              'Bad code pointer, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), k, 1                                        'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      Exit For                                                              'Bad method signature, quit loop
    End If
  Next i
  
  If i > j Then                                                             'Loop completed without finding the last method
    zError "zGetCallback", "z_Data() overflow. Increase the number of elements in the z_Data() array"
  End If
  
  'Uncomment the following line to determine the minimum number of elements needed by the z_Data() array
  'Debug.Print "Optimal dimension: z_Data(" & IIf(i > IIf(MSG_ENTRIES > IDX_EBX, MSG_ENTRIES, IDX_EBX), i, IIf(MSG_ENTRIES > IDX_EBX, MSG_ENTRIES, IDX_EBX)) & ")"
 
  zGetCallback = z_Data(i - 1)                                              'Return the last good vTable entry address
End Function

'-Subclass callback: must be private and the last method in the source file-----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
 '*************************************************************************************************
 '* bBefore  - Indicates whether the callback is before or after the original WndProc. Usually you
 '*            will know unless the callback for the uMsg value is specified as MSG_BEFORE_AFTER
 '*            (both before and after the original WndProc).
 '* bHandled - In a before original WndProc callback, setting bHandled to True will prevent the
 '*            message being passed to the original WndProc and (if set to do so) the after
 '*            original WndProc callback.
 '* lReturn  - WndProc return value. Set as per the MSDN documentation for the message value,
 '*            and/or, in an after the original WndProc callback, act on the return value as set
 '*            by the original WndProc.
 '* hWnd     - Window handle.
 '* uMsg     - Message value.
 '* wParam   - Message related data.
 '* lParam   - Message related data.
 '*************************************************************************************************
 Select Case uMsg
  Case WM_MOUSEMOVE
   If (isSetHighLight = False) Or (isHotTitle = True) Then Exit Sub
   If Not (isInCtrl = True) Then
    isInCtrl = True
    Call TrackMouseLeave(lng_hWnd)
    Call Refresh(OfficeHighLight)
    Call UpDate
    RaiseEvent MouseEnter
   End If
  Case WM_MOUSELEAVE
   If (isSetHighLight = False) Then Exit Sub
   isInCtrl = False
   Call Refresh(OfficeNormal)
   RaiseEvent MouseLeave
  Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
   Call UserControl_Resize
   RaiseEvent ChangedTheme
 End Select
End Sub
