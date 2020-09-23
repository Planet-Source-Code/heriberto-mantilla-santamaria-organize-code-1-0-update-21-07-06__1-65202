VERSION 5.00
Begin VB.UserControl GpTabStrip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   PropertyPages   =   "GpTabs.ctx":0000
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   149
   ToolboxBitmap   =   "GpTabs.ctx":0014
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "GpTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'
' ***************************************************************************************
' * Project  | GpTabStrip                                                               *
' *----------|--------------------------------------------------------------------------*
' * Version  | V1.0                                                                     *
' *----------|--------------------------------------------------------------------------*
' * Author   | Genghis Khan(GuangJian Guo)                                              *
' *----------|--------------------------------------------------------------------------*
' * WebSite  | http://www.itkhan.com                                                    *
' *----------|--------------------------------------------------------------------------*
' * MailTo   | webmaster@itkhan.com                                                     *
' *----------|--------------------------------------------------------------------------*
' * Date     | 13 April 2003                                                            *
' ***************************************************************************************

' ======================================================================================
' Constants
' ======================================================================================

Private Const MODULE_NAME = "GpTabs"
Private Const XPBorderColor = &H808080       '&H733C00  ' XP·ç¸ñ±ß¿òµÄÑÕÉ«
Private Const XPFlatBorderColor = &HE0A193
Private Const XPFlatTabColor = &HFFC3B3
'Private Const XPFlatTabColorActive = vbWhite
Private Const XPFlatTabColorActive = &HEAEAEA
Private Const XPFlatTabColorHover = &HFFA49B
Private Const TabsInterval = 2          ' Ã¿¸öTabÖ®¼äµÄ¼ä¸ô¾àÀë
Private Const RoundRectSize = 1         ' Ô²½ÇµÄ´óÐ¡
Private Const DiscrepancyHeight = 2     ' Ñ¡ÖÐµÄTabÓëÃ»ÓÐÑ¡ÖÐµÄTabµÄ¸ß¶È²î
Private Const InflateFontHeight = 6     ' ÓëTabµÄCaptionÔÚµ±Ç°×ÖÌåµÄÊµ¼Ê¸ß¶ÈÏà¼ÓµÄµÃTabµÄÄ¬ÈÏ¸ß¶È
Private Const InflateFontWidth = 4      ' ÓëTabµÄCaptionÔÚµ±Ç°×ÖÌåµÄÊµ¼Ê¿í¶ÈÏà¼ÓµÄµÃTabµÄÄ¬ÈÏ¿í¶È
Private Const InflateIconHeight = 2     ' ÓëTabµÄIconµÄÊµ¼Ê¸ß¶ÈÏà¼ÓµÄµÃTabµÄÄ¬ÈÏ¸ß¶È
Private Const InflateIconWidth = 0      ' ÓëTabµÄIconµÄÊµ¼Ê¿í¶ÈÏà¼ÓµÄµÃTabµÄÄ¬ÈÏ¿í¶È

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

#Const DEBUGMODE = 0

'Set of bit flags that indicate which common control classes will be loaded
'from the DLL. The dwICC value of tagINITCOMMONCONTROLSEX can
'be a combination of the following:
 Const ICC_LISTVIEW_CLASSES = &H1          '/* listview, header
 Const ICC_TREEVIEW_CLASSES = &H2          '/* treeview, tooltips
 Const ICC_BAR_CLASSES = &H4               '/* toolbar, statusbar, trackbar, tooltips
 Const ICC_TAB_CLASSES = &H8               '/* tab, tooltips
 Const ICC_UPDOWN_CLASS = &H10             '/* updown
 Const ICC_PROGRESS_CLASS = &H20           '/* progress
 Const ICC_HOTKEY_CLASS = &H40             '/* hotkey
 Const ICC_ANIMATE_CLASS = &H80            '/* animate
 Const ICC_WIN95_CLASSES = &HFF            '/* loads everything above
 Const ICC_DATE_CLASSES = &H100            '/* month picker, date picker, time picker, updown
 Const ICC_USEREX_CLASSES = &H200          '/* ComboEx
 Const ICC_COOL_CLASSES = &H400            '/* Rebar (coolbar) control


' Ö¸¶¨´°¿ÚµÄ½á¹¹ÖÐÈ¡µÃÐÅÏ¢£¬ÓÃÓÚGetWindowLong¡¢SetWindowLongº¯Êý
 Const GWL_EXSTYLE = (-20)                 '/* À©Õ¹´°¿ÚÑùÊ½ */
 Const GWL_HINSTANCE = (-6)                '/* ÓµÓÐ´°¿ÚµÄÊµÀýµÄ¾ä±ú */
 Const GWL_HWNDPARENT = (-8)               '/* ¸Ã´°¿ÚÖ®¸¸µÄ¾ä±ú¡£²»ÒªÓÃSetWindowWordÀ´¸Ä±äÕâ¸öÖµ */
 Const GWL_ID = (-12)                      '/* ¶Ô»°¿òÖÐÒ»¸ö×Ó´°¿ÚµÄ±êÊ¶·û */
 Const GWL_STYLE = (-16)                   '/* ´°¿ÚÑùÊ½ */
 Const GWL_USERDATA = (-21)                '/* º¬ÒåÓÉÓ¦ÓÃ³ÌÐò¹æ¶¨ */
 Const GWL_WNDPROC = (-4)                  '/* ¸Ã´°¿ÚµÄ´°¿Úº¯ÊýµÄµØÖ· */
 Const DWL_DLGPROC = 4                     '/* Õâ¸ö´°¿ÚµÄ¶Ô»°¿òº¯ÊýµØÖ· */
 Const DWL_MSGRESULT = 0                   '/* ÔÚ¶Ô»°¿òº¯ÊýÖÐ´¦ÀíµÄÒ»ÌõÏûÏ¢·µ»ØµÄÖµ */
 Const DWL_USER = 8                        '/* º¬ÒåÓÉÓ¦ÓÃ³ÌÐò¹æ¶¨ */


' GetDeviceCapsË÷Òý±í£¬ÓÃÓÚGetDeviceCapsº¯Êý
 Const DRIVERVERSION = 0                   '/* ±¸Çý¶¯³ÌÐò°æ±¾
 Const BITSPIXEL = 12                      '/*
 Const LOGPIXELSX = 88                     '/*  Logical pixels/inch in X
 Const LOGPIXELSY = 90                     '/*  Logical pixels/inch in Y

' Windows¶ÔÏó³£Êý±í£¬º¯ÊýGetSysColor
 Const COLOR_ACTIVEBORDER = 10             '/* »î¶¯´°¿ÚµÄ±ß¿ò
 Const COLOR_ACTIVECAPTION = 2             '/* »î¶¯´°¿ÚµÄ±êÌâ
 Const COLOR_ADJ_MAX = 100                 '/*
 Const COLOR_ADJ_MIN = -100                '/*
 Const COLOR_APPWORKSPACE = 12             '/* MDI×ÀÃæµÄ±³¾°
 Const COLOR_BACKGROUND = 1                '/*
 Const COLOR_BTNDKSHADOW = 21              '/*
 Const COLOR_BTNLIGHT = 22                 '/*
 Const COLOR_BTNFACE = 15                  '/* °´Å¥
 Const COLOR_BTNHIGHLIGHT = 20             '/* °´Å¥µÄ3D¼ÓÁÁÇø
 Const COLOR_BTNSHADOW = 16                '/* °´Å¥µÄ3DÒõÓ°
 Const COLOR_BTNTEXT = 18                  '/* °´Å¥ÎÄ×Ö
 Const COLOR_CAPTIONTEXT = 9               '/* ´°¿Ú±êÌâÖÐµÄÎÄ×Ö
 Const COLOR_GRAYTEXT = 17                 '/* »ÒÉ«ÎÄ×Ö£»ÈçÊ¹ÓÃÁË¶¶¶¯¼¼ÊõÔòÎªÁã
 Const COLOR_HIGHLIGHT = 13                '/* Ñ¡¶¨µÄÏîÄ¿±³¾°
 Const COLOR_HIGHLIGHTTEXT = 14            '/* Ñ¡¶¨µÄÏîÄ¿ÎÄ×Ö
 Const COLOR_INACTIVEBORDER = 11           '/* ²»»î¶¯´°¿ÚµÄ±ß¿ò
 Const COLOR_INACTIVECAPTION = 3           '/* ²»»î¶¯´°¿ÚµÄ±êÌâ
 Const COLOR_INACTIVECAPTIONTEXT = 19      '/* ²»»î¶¯´°¿ÚµÄÎÄ×Ö
 Const COLOR_MENU = 4                      '/* ²Ëµ¥
 Const COLOR_MENUTEXT = 7                  '/* ²Ëµ¥ÕýÎÄ
 Const COLOR_SCROLLBAR = 0                 '/* ¹ö¶¯Ìõ
 Const COLOR_WINDOW = 5                    '/* ´°¿Ú±³¾°
 Const COLOR_WINDOWFRAME = 6               '/* ´°¿ò
 Const COLOR_WINDOWTEXT = 8                '/* ´°¿ÚÕýÎÄ
Const COLORONCOLOR = 3

' º¯ÊýCombineRgnµÄ·µ»ØÖµ£¬ÀàÐÍLong
 Const COMPLEXREGION = 3                   '/* ÇøÓòÓÐ»¥Ïà½»µþµÄ±ß½ç */
 Const SIMPLEREGION = 2                    '/* ÇøÓò±ß½çÃ»ÓÐ»¥Ïà½»µþ */
 Const NULLREGION = 1                      '/* ÇøÓòÎª¿Õ */
 Const ERRORAPI = 0                        '/* ²»ÄÜ´´½¨×éºÏÇøÓò */

' ×éºÏÁ½ÇøÓòµÄ·½·¨£¬º¯ÊýCombineRgnµÄµÄ²ÎÊýnCombineModeËùÊ¹ÓÃµÄ³£Êý
 Const RGN_AND = 1                         '/* hDestRgn±»ÉèÖÃÎªÁ½¸öÔ´ÇøÓòµÄ½»¼¯ */
 Const RGN_COPY = 5                        '/* hDestRgn±»ÉèÖÃÎªhSrcRgn1µÄ¿½±´ */
 Const RGN_DIFF = 4                        '/* hDestRgn±»ÉèÖÃÎªhSrcRgn1ÖÐÓëhSrcRgn2²»Ïà½»µÄ²¿·Ö */
 Const RGN_OR = 2                          '/* hDestRgn±»ÉèÖÃÎªÁ½¸öÇøÓòµÄ²¢¼¯ */
 Const RGN_XOR = 3                         '/* hDestRgn±»ÉèÖÃÎª³ýÁ½¸öÔ´ÇøÓòORÖ®ÍâµÄ²¿·Ö */

' Missing Draw State constants declarations£¬²Î¿´DrawStateº¯Êý
'/* Image type */
 Const DST_COMPLEX = &H0                   '/* »æÍ¼ÔÚÓÉlpDrawStateProc²ÎÊýÖ¸¶¨µÄ»Øµ÷º¯ÊýÆÚ¼äÖ´ÐÐ¡£lParamºÍwParam»á´«µÝ¸ø»Øµ÷ÊÂ¼þ
 Const DST_TEXT = &H1                      '/* lParam´ú±íÎÄ×ÖµÄµØÖ·£¨¿ÉÊ¹ÓÃÒ»¸ö×Ö´®±ðÃû£©£¬wParam´ú±í×Ö´®µÄ³¤¶È
 Const DST_PREFIXTEXT = &H2                '/* ÓëDST_TEXTÀàËÆ£¬Ö»ÊÇ & ×Ö·ûÖ¸³öÎªÏÂ¸÷×Ö·û¼ÓÉÏÏÂ»®Ïß
 Const DST_ICON = &H3                      '/* lParam°üÀ¨Í¼±ê¾ä±ú
 Const DST_BITMAP = &H4                    '/* lParamÖÐµÄ¾ä±ú
' /* State type */
 Const DSS_NORMAL = &H0                    '/* ÆÕÍ¨Í¼Ïó
 Const DSS_UNION = &H10                    '/* Í¼Ïó½øÐÐ¶¶¶¯´¦Àí
 Const DSS_DISABLED = &H20                 '/* Í¼Ïó¾ßÓÐ¸¡µñÐ§¹û
 Const DSS_MONO = &H80                     '/* ÓÃhBrushÃè»æÍ¼Ïó
 Const DSS_RIGHT = &H8000                  '/*

' Built in ImageList drawing methods:
 Const ILD_NORMAL = 0&
 Const ILD_TRANSPARENT = 1&
 Const ILD_BLEND25 = 2&
 Const ILD_SELECTED = 4&
 Const ILD_FOCUS = 4&
 Const ILD_MASK = &H10&
 Const ILD_IMAGE = &H20&
 Const ILD_ROP = &H40&
 Const ILD_OVERLAYMASK = 3840&
 Const ILC_MASK = &H1&
 Const ILCF_MOVE = &H0&
 Const ILCF_SWAP = &H1&

 Const CLR_DEFAULT = -16777216
 Const CLR_HILIGHT = -16777216
 Const CLR_NONE = -1

' General windows messages:
 Const WM_COMMAND = &H111
 Const WM_KEYDOWN = &H100
 Const WM_KEYUP = &H101
 Const WM_CHAR = &H102
 Const WM_SETFOCUS = &H7
 Const WM_KILLFOCUS = &H8
 Const WM_SETFONT = &H30
 Const WM_GETTEXT = &HD
 Const WM_GETTEXTLENGTH = &HE
 Const WM_SETTEXT = &HC
 Const WM_NOTIFY = &H4E&

' Show window styles
 Const SW_SHOWNORMAL = 1
 Const SW_ERASE = &H4
 Const SW_HIDE = 0
 Const SW_INVALIDATE = &H2
 Const SW_MAX = 10
 Const SW_MAXIMIZE = 3
 Const SW_MINIMIZE = 6
 Const SW_NORMAL = 1
 Const SW_OTHERUNZOOM = 4
 Const SW_OTHERZOOM = 2
 Const SW_PARENTCLOSING = 1
 Const SW_RESTORE = 9
 Const SW_PARENTOPENING = 3
 Const SW_SHOW = 5
 Const SW_SCROLLCHILDREN = &H1
 Const SW_SHOWDEFAULT = 10
 Const SW_SHOWMAXIMIZED = 3
 Const SW_SHOWMINIMIZED = 2
 Const SW_SHOWMINNOACTIVE = 7
 Const SW_SHOWNA = 8
 Const SW_SHOWNOACTIVATE = 4

' ³£¼ûµÄ¹âÕ¤²Ù×÷´úÂë
 Const BLACKNESS = &H42                    '/* ±íÊ¾Ê¹ÓÃÓëÎïÀíµ÷É«°åµÄË÷Òý0Ïà¹ØµÄÉ«²ÊÀ´Ìî³äÄ¿±ê¾ØÐÎÇøÓò£¬£¨¶ÔÈ±Ê¡µÄÎïÀíµ÷É«°å¶øÑÔ£¬¸ÃÑÕÉ«ÎªºÚÉ«£©¡£
 Const DSTINVERT = &H550009                '/* ±íÊ¾Ê¹Ä¿±ê¾ØÐÎÇøÓòÑÕÉ«È¡·´¡£
 Const MERGECOPY = &HC000CA                '/* ±íÊ¾Ê¹ÓÃ²¼¶ûÐÍµÄAND£¨Óë£©²Ù×÷·û½«Ô´¾ØÐÎÇøÓòµÄÑÕÉ«ÓëÌØ¶¨Ä£Ê½×éºÏÒ»Æð¡£
 Const MERGEPAINT = &HBB0226               '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄOR£¨»ò£©²Ù×÷·û½«·´ÏòµÄÔ´¾ØÐÎÇøÓòµÄÑÕÉ«ÓëÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
 Const NOTSRCCOPY = &H330008               '/* ½«Ô´¾ØÐÎÇøÓòÑÕÉ«È¡·´£¬ÓÚ¿½±´µ½Ä¿±ê¾ØÐÎÇøÓò¡£
 Const NOTSRCERASE = &H1100A6              '/* Ê¹ÓÃ²¼¶ûÀàÐÍµÄOR£¨»ò£©²Ù×÷·û×éºÏÔ´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«Öµ£¬È»ºó½«ºÏ³ÉµÄÑÕÉ«È¡·´¡£
 Const PATCOPY = &HF00021                  '/* ½«ÌØ¶¨µÄÄ£Ê½¿½±´µ½Ä¿±êÎ»Í¼ÉÏ¡£
 Const PATINVERT = &H5A0049                '/* Í¨¹ýÊ¹ÓÃ²¼¶ûOR£¨»ò£©²Ù×÷·û½«Ô´¾ØÐÎÇøÓòÈ¡·´ºóµÄÑÕÉ«ÖµÓëÌØ¶¨Ä£Ê½µÄÑÕÉ«ºÏ²¢¡£È»ºóÊ¹ÓÃOR£¨»ò£©²Ù×÷·û½«¸Ã²Ù×÷µÄ½á¹ûÓëÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢¡£
 Const PATPAINT = &HFB0A09                 '/* Í¨¹ýÊ¹ÓÃXOR£¨Òì»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢¡£
 Const SRCAND = &H8800C6                   '/* Í¨¹ýÊ¹ÓÃAND£¨Óë£©²Ù×÷·ûÀ´½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢
 Const SRCCOPY = &HCC0020                  '/* ½«Ô´¾ØÐÎÇøÓòÖ±½Ó¿½±´µ½Ä¿±ê¾ØÐÎÇøÓò¡£
 Const SRCERASE = &H440328                 '/* Í¨¹ýÊ¹ÓÃAND£¨Óë£©²Ù×÷·û½«Ä¿±ê¾ØÐÎÇøÓòÑÕÉ«È¡·´ºóÓëÔ´¾ØÐÎÇøÓòµÄÑÕÉ«ÖµºÏ²¢¡£
 Const SRCINVERT = &H660046                '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄXOR£¨Òì»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
 Const SRCPAINT = &HEE0086                 '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄOR£¨»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
 Const WHITENESS = &HFF0062                '/* Ê¹ÓÃÓëÎïÀíµ÷É«°åÖÐË÷Òý1ÓÐ¹ØµÄÑÕÉ«Ìî³äÄ¿±ê¾ØÐÎÇøÓò¡££¨¶ÔÓÚÈ±Ê¡ÎïÀíµ÷É«°åÀ´Ëµ£¬Õâ¸öÑÕÉ«¾ÍÊÇ°×É«£©¡£

'--- for mouse_event
 Const MOUSE_MOVED = &H1
 Const MOUSEEVENTF_ABSOLUTE = &H8000       '/*
 Const MOUSEEVENTF_LEFTDOWN = &H2          '/* Ä£ÄâÊó±ê×ó¼ü°´ÏÂ
Const MOUSEEVENTF_LEFTUP = &H4            '/* Ä£ÄâÊó±ê×ó¼üÌ§Æð
 Const MOUSEEVENTF_MIDDLEDOWN = &H20       '/* Ä£ÄâÊó±êÖÐ¼ü°´ÏÂ
 Const MOUSEEVENTF_MIDDLEUP = &H40         '/* Ä£ÄâÊó±êÖÐ¼ü°´ÏÂ
 Const MOUSEEVENTF_MOVE = &H1              '/* ÒÆ¶¯Êó±ê */
 Const MOUSEEVENTF_RIGHTDOWN = &H8         '/* Ä£ÄâÊó±êÓÒ¼ü°´ÏÂ
 Const MOUSEEVENTF_RIGHTUP = &H10          '/* Ä£ÄâÊó±êÓÒ¼ü°´ÏÂ
Const MOUSETRAILS = 39                    '/*

 Const BMP_MAGIC_COOKIE = 19778            '/* this is equivalent to ascii string "BM" */
' constants for the biCompression field
 Const BI_RGB = 0&
 Const BI_RLE4 = 2&
 Const BI_RLE8 = 1&
 Const BI_BITFIELDS = 3&
' Const BITSPIXEL = 12                     '/* Number of bits per pixel
' DIB color table identifiers
 Const DIB_PAL_COLORS = 1                  '/* ÔÚÑÕÉ«±íÖÐ×°ÔØÒ»¸ö16Î»ËùÒÔÊý×é£¬ËüÃÇÓëµ±Ç°Ñ¡¶¨µÄµ÷É«°åÓÐ¹Ø color table in palette indices
 Const DIB_PAL_INDICES = 2                 '/* No color table indices into surf palette
 Const DIB_PAL_LOGINDICES = 4              '/* No color table indices into DC palette
 Const DIB_PAL_PHYSINDICES = 2             '/* No color table indices into surf palette
 Const DIB_RGB_COLORS = 0                  '/* ÔÚÑÕÉ«±íÖÐ×°ÔØRGBÑÕÉ«

' BLENDFUNCTION AlphaFormat-Konstante
 Const AC_SRC_ALPHA = &H1
' BLENDFUNCTION BlendOp-Konstante
 Const AC_SRC_OVER = &H0

' ======================================================================================
' Methods
' ======================================================================================
' º¯ÊýSetBkModen²ÎÊýBkMode
 Enum KhanBackStyles
    TRANSPARENT = 1                              '/* Í¸Ã÷´¦Àí£¬¼´²»×÷ÉÏÊöÌî³ä */
    OPAQUE = 2                                   '/* ÓÃµ±Ç°µÄ±³¾°É«Ìî³äÐéÏß»­±Ê¡¢ÒõÓ°Ë¢×ÓÒÔ¼°×Ö·ûµÄ¿ÕÏ¶ */
    NEWTRANSPARENT = 3                           '/* NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum

' ¶à±ßÐÎµÄÌî³äÄ£Ê½
 Enum KhanPolyFillModeFalgs
    ALTERNATE = 1                                '/* ½»ÌæÌî³ä
    WINDING = 2                                  '/* ¸ù¾Ý»æÍ¼·½ÏòÌî³ä
End Enum

' DrawIconEx
 Enum KhanDrawIconExFlags
    DI_MASK = &H1                                '/* »æÍ¼Ê±Ê¹ÓÃÍ¼±êµÄMASK²¿·Ö£¨Èçµ¥¶ÀÊ¹ÓÃ£¬¿É»ñµÃÍ¼±êµÄÑÚÄ££©
    DI_IMAGE = &H2                               '/* »æÍ¼Ê±Ê¹ÓÃÍ¼±êµÄXOR²¿·Ö£¨¼´Í¼±êÃ»ÓÐÍ¸Ã÷ÇøÓò£©
    DI_NORMAL = &H3                              '/* ÓÃ³£¹æ·½Ê½»æÍ¼£¨ºÏ²¢ DI_IMAGE ºÍ DI_MASK£©
    DI_COMPAT = &H4                              '/* Ãè»æ±ê×¼µÄÏµÍ³Ö¸Õë£¬¶ø²»ÊÇÖ¸¶¨µÄÍ¼Ïó
    DI_DEFAULTSIZE = &H8                         '/* ºöÂÔcxWidthºÍcyWidthÉèÖÃ£¬²¢²ÉÓÃÔ­Ê¼µÄÍ¼±ê´óÐ¡
End Enum

'Ö¸¶¨±»×°ÔØÍ¼ÏñÀàÐÍ,LoadImage,CopyImage
 Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

 Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* ±íÊ¾´´½¨Ò»¸öÍ¼ÏñµÄ¾«È·¸±±¾£¬¶øºöÂÔ²ÎÊýcxDesiredºÍcyDesired
    LR_COPYDELETEORG = &H8                       '/* ±íÊ¾´´½¨Ò»¸ö¸±±¾ºóÉ¾³ýÔ­Ê¼Í¼Ïñ¡£
    LR_CREATEDIBSECTION = &H2000                 '/* µ±²ÎÊýuTypeÖ¸¶¨ÎªIMAGE_BITMAPÊ±£¬Ê¹µÃº¯Êý·µ»ØÒ»¸öDIB²¿·ÖÎ»Í¼£¬¶ø²»ÊÇÒ»¸ö¼æÈÝµÄÎ»Í¼¡£Õâ¸ö±êÖ¾ÔÚ×°ÔØÒ»¸öÎ»Í¼£¬¶ø²»ÊÇÓ³ÉäËüµÄÑÕÉ«µ½ÏÔÊ¾Éè±¸Ê±·Ç³£ÓÐÓÃ¡£
    LR_DEFAULTCOLOR = &H0                        '/* ÒÔ³£¹æ·½Ê½ÔØÈëÍ¼Ïó
    LR_DEFAULTSIZE = &H40                        '/* Èô cxDesired»òcyDesiredÎ´±»ÉèÎªÁã£¬Ê¹ÓÃÏµÍ³Ö¸¶¨µÄ¹«ÖÆÖµ±êÊ¶¹â±ê»òÍ¼±êµÄ¿íºÍ¸ß¡£Èç¹ûÕâ¸ö²ÎÊý²»±»ÉèÖÃÇÒcxDesired»òcyDesired±»ÉèÎªÁã£¬º¯ÊýÊ¹ÓÃÊµ¼Ê×ÊÔ´³ß´ç¡£Èç¹û×ÊÔ´°üº¬¶à¸öÍ¼Ïñ£¬ÔòÊ¹ÓÃµÚÒ»¸öÍ¼ÏñµÄ´óÐ¡¡£
    LR_LOADFROMFILE = &H10                       '/* ¸ù¾Ý²ÎÊýlpszNameµÄÖµ×°ÔØÍ¼Ïñ¡£Èô±ê¼ÇÎ´±»¸ø¶¨£¬lpszNameµÄÖµÎª×ÊÔ´Ãû³Æ¡£
    LR_LOADMAP3DCOLORS = &H1000                  '/* ½«Í¼ÏóÖÐµÄÉî»Ò(Dk Gray RGB£¨128£¬128£¬128£©)¡¢»Ò(Gray RGB£¨192£¬192£¬192£©)¡¢ÒÔ¼°Ç³»Ò(Gray RGB£¨223£¬223£¬223£©)ÏñËØ¶¼Ìæ»»³ÉCOLOR_3DSHADOW£¬COLOR_3DFACEÒÔ¼°COLOR_3DLIGHTµÄµ±Ç°ÉèÖÃ
    LR_LOADTRANSPARENT = &H20                    '/* ÈôfuLoad°üÀ¨LR_LOADTRANSPARENTºÍLR_LOADMAP3DCOLORSÁ½¸öÖµ£¬ÔòLRLOADTRANSPARENTÓÅÏÈ¡£µ«ÊÇ£¬ÑÕÉ«±í½Ó¿ÚÓÉCOLOR_3DFACEÌæ´ú£¬¶ø²»ÊÇCOLOR_WINDOW¡£
    LR_MONOCHROME = &H1                          '/* ½«Í¼Ïó×ª»»³Éµ¥É«
    LR_SHARED = &H8000                           '/* ÈôÍ¼Ïñ½«±»¶à´Î×°ÔØÔò¹²Ïí¡£Èç¹ûLR_SHAREDÎ´±»ÉèÖÃ£¬ÔòÔÙÏòÍ¬Ò»¸ö×ÊÔ´µÚ¶þ´Îµ÷ÓÃÕâ¸öÍ¼ÏñÊÇ¾Í»áÔÙ×°ÔØÒÔ±ãÕâ¸öÍ¼ÏñÇÒ·µ»Ø²»Í¬µÄ¾ä±ú¡£
    LR_COPYFROMRESOURCE = &H4000                 '/*
End Enum

 Enum KhanDrawTextStyles
    DT_BOTTOM = &H8&                             '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄµ×±ß
    DT_CALCRECT = &H400&                         '/* ÏóÏÂÃæÕâÑù¼ÆËã¸ñÊ½»¯¾ØÐÎ£º¶àÐÐ»æÍ¼Ê±¾ØÐÎµÄµ×±ß¸ù¾ÝÐèÒª½øÐÐÑÓÕ¹£¬ÒÔ±ãÈÝÏÂËùÓÐÎÄ×Ö£»µ¥ÐÐ»æÍ¼Ê±£¬ÑÓÕ¹¾ØÐÎµÄÓÒ²à¡£²»Ãè»æÎÄ×Ö¡£ÓÉlpRect²ÎÊýÖ¸¶¨µÄ¾ØÐÎ»áÔØÈë¼ÆËã³öÀ´µÄÖµ
    DT_CENTER = &H1&                             '/* ÎÄ±¾´¹Ö±¾ÓÖÐ
    DT_EXPANDTABS = &H40&                        '/* Ãè»æÎÄ×ÖµÄÊ±ºò£¬¶ÔÖÆ±íÕ¾½øÐÐÀ©Õ¹¡£Ä¬ÈÏµÄÖÆ±íÕ¾¼ä¾àÊÇ8¸ö×Ö·û¡£µ«ÊÇ£¬¿ÉÓÃDT_TABSTOP±êÖ¾¸Ä±äÕâÏîÉè¶¨
    DT_EXTERNALLEADING = &H200&                  '/* ¼ÆËãÎÄ±¾ÐÐ¸ß¶ÈµÄÊ±ºò£¬Ê¹ÓÃµ±Ç°×ÖÌåµÄÍâ²¿¼ä¾àÊôÐÔ£¨the external leading attribute£©
    DT_INTERNAL = &H1000&                        '/* Uses the system font to calculate text metrics
    DT_LEFT = &H0&                               '/* ÎÄ±¾×ó¶ÔÆë
    DT_NOCLIP = &H100&                           '/* Ãè»æÎÄ×ÖÊ±²»¼ôÇÐµ½Ö¸¶¨µÄ¾ØÐÎ£¬DrawTextEx is somewhat faster when DT_NOCLIP is used.
    DT_NOPREFIX = &H800&                         '/* Í¨³££¬º¯ÊýÈÏÎª & ×Ö·û±íÊ¾Ó¦ÎªÏÂÒ»¸ö×Ö·û¼ÓÉÏÏÂ»®Ïß¡£¸Ã±êÖ¾½ûÖ¹ÕâÖÖÐÐÎª
    DT_RIGHT = &H2&                              '/* ÎÄ±¾ÓÒ¶ÔÆë
    DT_SINGLELINE = &H20&                        '/* Ö»»­µ¥ÐÐ
    DT_TABSTOP = &H80&                           '/* Ö¸¶¨ÐÂµÄÖÆ±íÕ¾¼ä¾à£¬²ÉÓÃÕâ¸öÕûÊýµÄ¸ß8Î»
    DT_TOP = &H0&                                '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄµ×±ß
    DT_VCENTER = &H4&                            '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄÖÐ²¿
    DT_WORDBREAK = &H10&                         '/* ½øÐÐ×Ô¶¯»»ÐÐ¡£ÈçÓÃSetTextAlignº¯ÊýÉèÖÃÁËTA_UPDATECP±êÖ¾£¬ÕâÀïµÄÉèÖÃÔòÎÞÐ§
' #if(WINVER >= =&H0400)
    DT_EDITCONTROL = &H2000&                     '/* ¶ÔÒ»¸ö¶àÐÐ±à¼­¿Ø¼þ½øÐÐÄ£Äâ¡£²»ÏÔÊ¾²¿·Ö¿É¼ûµÄÐÐ
    DT_END_ELLIPSIS = &H8000&                    '/* ÌÈÈô×Ö´®²»ÄÜÔÚ¾ØÐÎÀïÈ«²¿ÈÝÏÂ£¬¾ÍÔÚÄ©Î²ÏÔÊ¾Ê¡ÂÔºÅ
    DT_PATH_ELLIPSIS = &H4000&                   '/* Èç×Ö´®°üº¬ÁË \ ×Ö·û£¬¾ÍÓÃÊ¡ÂÔºÅÌæ»»×Ö´®ÄÚÈÝ£¬Ê¹ÆäÄÜÔÚ¾ØÐÎÖÐÈ«²¿ÈÝÏÂ¡£ÀýÈç£¬Ò»¸öºÜ³¤µÄÂ·¾¶Ãû¿ÉÄÜ»»³ÉÕâÑùÏÔÊ¾¡ª¡ªc:\windows\...\doc\readme.txt
    DT_MODIFYSTRING = &H10000                    '/* ÈçÖ¸¶¨ÁËDT_ENDELLIPSES »ò DT_PATHELLIPSES£¬¾Í»á¶Ô×Ö´®½øÐÐÐÞ¸Ä£¬Ê¹ÆäÓëÊµ¼ÊÏÔÊ¾µÄ×Ö´®Ïà·û
    DT_RTLREADING = &H20000                      '/* ÈçÑ¡ÈëÉè±¸³¡¾°µÄ×ÖÌåÊôÓÚÏ£²®À´»ò°¢À­²®ÓïÏµ£¬¾Í´ÓÓÒµ½×óÃè»æÎÄ×Ö
    DT_WORD_ELLIPSIS = &H40000                   '/* Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

 Enum KhanDrawFrameControlType
    DFC_CAPTION = 1                              '/* Title bar.
    DFC_MENU = 2                                 '/* Menu bar.
    DFC_SCROLL = 3                               '/* Scroll bar.
    DFC_BUTTON = 4                               '/* Standard button.
    DFC_POPUPMENU = 5                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Popup menu item.
End Enum

 Enum KhanDrawFrameControlStyle
    DFCS_BUTTONCHECK = &H0                       '/* Check box.
    DFCS_BUTTONRADIOIMAGE = &H1                  '/* Image for radio button (nonsquare needs image).
    DFCS_BUTTONRADIOMASK = &H2                   '/* Mask for radio button (nonsquare needs mask).
    DFCS_BUTTONRADIO = &H4                       '/* Radio button.
    DFCS_BUTTON3STATE = &H8                      '/* Three-state button.
    DFCS_BUTTONPUSH = &H10                       '/* Push button.
    DFCS_CAPTIONCLOSE = &H0                      '/* <b>Close</b> button.
    DFCS_CAPTIONMIN = &H1                        '/* <b>Minimize</b> button.
    DFCS_CAPTIONMAX = &H2                        '/* <b>Maximize</b> button.
    DFCS_CAPTIONRESTORE = &H3                    '/* <b>Restore</b> button.
    DFCS_CAPTIONHELP = &H4                       '/* <b>Help</b> button.
    DFCS_MENUARROW = &H0                         '/* Submenu arrow.
    DFCS_MENUCHECK = &H1                         '/* Check mark.
    DFCS_MENUBULLET = &H2                        '/* Bullet.
    DFCS_MENUARROWRIGHT = &H4                    '/* Submenu arrow pointing left. This is used for the right-to-left cascading menus used with right-to-left languages such as Arabic or Hebrew.
    DFCS_SCROLLUP = &H0                          '/* Up arrow of scroll bar.
    DFCS_SCROLLDOWN = &H1                        '/* Down arrow of scroll bar.
    DFCS_SCROLLLEFT = &H2                        '/* Left arrow of scroll bar.
    DFCS_SCROLLRIGHT = &H3                       '/* Right arrow of scroll bar.
    DFCS_SCROLLCOMBOBOX = &H5                    '/* Combo box scroll bar.
    DFCS_SCROLLSIZEGRIP = &H8                    '/* Size grip in bottom-right corner of window.
    DFCS_SCROLLSIZEGRIPRIGHT = &H10              '/* Size grip in bottom-left corner of window. This is used with right-to-left languages such as Arabic or Hebrew.
    DFCS_INACTIVE = &H100                        '/* Button is inactive (grayed).
    DFCS_PUSHED = &H200                          '/* Button is pushed.
    DFCS_CHECKED = &H400                         '/* Button is checked.
    DFCS_TRANSPARENT = &H800                     '/* <b>Windows 98/Me, Windows 2000 or later:</b> The background remains untouched.
    DFCS_HOT = &H1000                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Button is hot-tracked.
    DFCS_ADJUSTRECT = &H2000                     '/* Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
    DFCS_FLAT = &H4000                           '/* Button has a flat border.
    DFCS_MONO = &H8000                           '/* Button has a monochrome border.
End Enum

' Ö¸¶¨»­±ÊÑùÊ½£¬º¯ÊýCreatePenµÄ²ÎÊýCreatePenËùÊ¹ÓÃµÄ³£Êý
 Enum KhanPenStyles
    ' CreatePen£¬ExtCreatePen
    ' »­±ÊµÄÑùÊ½
    PS_SOLID = 0                                 '/* »­±Ê»­³öµÄÊÇÊµÏß */
    PS_DASH = 1                                  '/* »­±Ê»­³öµÄÊÇÐéÏß£¨nWidth±ØÐëÊÇ1£© */
    PS_DOT = 2                                   '/* »­±Ê»­³öµÄÊÇµãÏß£¨nWidth±ØÐëÊÇ1£© */
    PS_DASHDOT = 3                               '/* »­±Ê»­³öµÄÊÇµã»®Ïß£¨nWidth±ØÐëÊÇ1£© */
    PS_DASHDOTDOT = 4                            '/* »­±Ê»­³öµÄÊÇµã-µã-»®Ïß£¨nWidth±ØÐëÊÇ1£© */
    PS_NULL = 5                                  '/* »­±Ê²»ÄÜ»­Í¼ */
    PS_INSIDEFRAME = 6                           '/* »­±ÊÔÚÓÉÍÖÔ²¡¢¾ØÐÎ¡¢Ô²½Ç¾ØÐÎ¡¢±ýÍ¼ÒÔ¼°ÏÒµÈÉú³ÉµÄ·â±Õ¶ÔÏó¿òÖÐ»­Í¼¡£ÈçÖ¸¶¨µÄ×¼È·RGBÑÕÉ«²»´æÔÚ£¬¾Í½øÐÐ¶¶¶¯´¦Àí */
    ' ExtCreatePen
    ' »­±ÊµÄÑùÊ½
    PS_USERSTYLE = 7                             '/* <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8                             '/* <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    ' »­±ÊµÄ±Ê¼â
    PS_ENDCAP_ROUND = &H0                        '/* End caps are round.
    PS_ENDCAP_SQUARE = &H100                     '/* End caps are square.
    PS_ENDCAP_FLAT = &H200                       '/* End caps are flat.
    PS_ENDCAP_MASK = &HF00                       '/* Mask for previous PS_ENDCAP_XXX values.
    ' ÔÚÍ¼ÐÎÖÐÁ¬½ÓÏß¶Î»òÔÚÂ·¾¶ÖÐÁ¬½ÓÖ±ÏßµÄ·½Ê½
    PS_JOIN_ROUND = &H0                          '/* Joins are beveled.
    PS_JOIN_BEVEL = &H1000                       '/* Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                       '/* Joins are round.
    PS_JOIN_MASK = &HF000                        '/* Mask for previous PS_JOIN_XXX values.
    ' »­±ÊµÄÀàÐÍ
    PS_COSMETIC = &H0                            '/* The pen is cosmetic.
    PS_GEOMETRIC = &H10000                       '/* The pen is geometric.
    '
    PS_STYLE_MASK = &HF                          '/* Mask for previous PS_XXX values.
    PS_TYPE_MASK = &HF0000                       '/* Mask for previous PS_XXX (pen type).
End Enum

 Enum KhanBrushStyle
    BS_SOLID = 0                                 '/* Solid brush.
    BS_HOLLOW = 1                                '/* Hollow brush.
    BS_NULL = 1                                  '/* Same as BS_HOLLOW.
    BS_HATCHED = 2                               '/* Hatched brush.
    BS_PATTERN = 3                               '/* Pattern brush defined by a memory bitmap.
    BS_INDEXED = 4                               '/*
    BS_DIBPATTERN = 5                            '/* A pattern brush defined by a device-independent bitmap (DIB) specification.
    BS_DIBPATTERNPT = 6                          '/* A pattern brush defined by a device-independent bitmap (DIB) specification. If <b>lbStyle</b> is BS_DIBPATTERNPT, the <b>lbHatch</b> member contains a pointer to a packed DIB.
    BS_PATTERN8X8 = 7                            '/* Same as BS_PATTERN.
    BS_DIBPATTERN8X8 = 8                         '/* Same as BS_DIBPATTERN.
    BS_MONOPATTERN = 9                           '/* The brush is a monochrome (black & white) bitmap.
End Enum

 Enum KhanHatchStyles
    HS_HORIZONTAL = 0                            '/* Horizontal hatch.
    HS_VERTICAL = 1                              '/* Vertical hatch.
    HS_FDIAGONAL = 2                             '/* A 45-degree downward, left-to-right hatch.
    HS_BDIAGONAL = 3                             '/* A 45-degree upward, left-to-right hatch.
    HS_CROSS = 4                                 '/* Horizontal and vertical cross-hatch.
    HS_DIAGCROSS = 5                             '/* A 45-degree crosshatch.
End Enum

' DrawEdge
 Enum KhanBorderStyles
    BDR_RAISEDOUTER = &H1                        '/* Raised outer edge.
    BDR_SUNKENOUTER = &H2                        '/* Sunken outer edge.
    BDR_RAISEDINNER = &H4                        '/* Raised inner edge.
    BDR_SUNKENINNER = &H8                        '/* Sunken inner edge.
    BDR_OUTER = &H3                              '/* (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = &HC                              '/* (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

 Enum KhanBorderFlags
    BF_LEFT = &H1                                '/* Left side of border rectangle.
    BF_TOP = &H2                                 '/* Top of border rectangle.
    BF_RIGHT = &H4                               '/* Right side of border rectangle.
    BF_BOTTOM = &H8                              '/* Bottom of border rectangle.
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10                           '/* Diagonal border.
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800                            '/* Fill in the middle.
    BF_SOFT = &H1000                             '/* Use for softer buttons.
    BF_ADJUST = &H2000                           '/* Calculate the space left over.
    BF_FLAT = &H4000                             '/* For flat rather than 3-D borders.
    BF_MONO = &H8000&                            '/* For monochrome borders
End Enum

' ´°¿ÚÖ¸¶¨Ò»¸öÐÂÎ»ÖÃºÍ×´Ì¬£¬ÓÃÓÚSetWindowPosº¯Êý
 Enum KhanSetWindowPosStyles
    HWND_BOTTOM = 1                              '/* ½«´°¿ÚÖÃÓÚ´°¿ÚÁÐ±íµ×²¿ */
    HWND_NOTOPMOST = -2                          '/* ½«´°¿ÚÖÃÓÚÁÐ±í¶¥²¿£¬²¢Î»ÓÚÈÎºÎ×î¶¥²¿´°¿ÚµÄºóÃæ */
    HWND_TOP = 0                                 '/* ½«´°¿ÚÖÃÓÚZÐòÁÐµÄ¶¥²¿£»ZÐòÁÐ´ú±íÔÚ·Ö¼¶½á¹¹ÖÐ£¬´°¿ÚÕë¶ÔÒ»¸ö¸ø¶¨¼¶±ðµÄ´°¿ÚÏÔÊ¾µÄË³Ðò */
    HWND_TOPMOST = -1                            '/* ½«´°¿ÚÖÃÓÚÁÐ±í¶¥²¿£¬²¢Î»ÓÚÈÎºÎ×î¶¥²¿´°¿ÚµÄÇ°Ãæ */
    SWP_SHOWWINDOW = &H40                        '/* ÏÔÊ¾´°¿Ú */
    SWP_HIDEWINDOW = &H80                        '/* Òþ²Ø´°¿Ú */
    SWP_FRAMECHANGED = &H20                      '/* Ç¿ÆÈÒ»ÌõWM_NCCALCSIZEÏûÏ¢½øÈë´°¿Ú£¬¼´Ê¹´°¿ÚµÄ´óÐ¡Ã»ÓÐ¸Ä±ä */
    SWP_NOACTIVATE = &H10                        '/* ²»¼¤»î´°¿Ú */
    SWP_NOCOPYBITS = &H100                       '
    SWP_NOMOVE = &H2                             '/* ±£³Öµ±Ç°Î»ÖÃ£¨xºÍyÉè¶¨½«±»ºöÂÔ£© */
    SWP_NOOWNERZORDER = &H200                    '/* Don't do owner Z ordering */
    SWP_NOREDRAW = &H8                           '/* ´°¿Ú²»×Ô¶¯ÖØ»­ */
    SWP_NOREPOSITION = SWP_NOOWNERZORDER         '
    SWP_NOSIZE = &H1                             '/* ±£³Öµ±Ç°´óÐ¡£¨cxºÍcy»á±»ºöÂÔ£© */
    SWP_NOZORDER = &H4                           '/* ±£³Ö´°¿ÚÔÚÁÐ±íµÄµ±Ç°Î»ÖÃ£¨hWndInsertAfter½«±»ºöÂÔ£© */
    SWP_DRAWFRAME = SWP_FRAMECHANGED             '/* Î§ÈÆ´°¿Ú»­Ò»¸ö¿ò */
'    HWND_BROADCAST = &HFFFF&
'    HWND_DESKTOP = 0
End Enum

' Ö¸¶¨´´½¨´°¿ÚµÄ·ç¸ñ
 Enum KhanCreateWindowSytles
    ' CreateWindow
    WS_BORDER = &H800000                         '/* ´´½¨Ò»¸öµ¥±ß¿òµÄ´°¿Ú¡£
    WS_CAPTION = &HC00000                        '/* ´´½¨Ò»¸öÓÐ±êÌâ¿òµÄ´°¿Ú£¨°üÀ¨WS_BODER·ç¸ñ£©¡£
    WS_CHILD = &H40000000                        '/* ´´½¨Ò»¸ö×Ó´°¿Ú¡£Õâ¸ö·ç¸ñ²»ÄÜÓëWS_POPVP·ç¸ñºÏÓÃ¡£
    WS_CHILDWINDOW = (WS_CHILD)                  '/* ÓëWS_CHILDÏàÍ¬¡£
    WS_CLIPCHILDREN = &H2000000                  '/* µ±ÔÚ¸¸´°¿ÚÄÚ»æÍ¼Ê±£¬ÅÅ³ý×Ó´°¿ÚÇøÓò¡£ÔÚ´´½¨¸¸´°¿ÚÊ±Ê¹ÓÃÕâ¸ö·ç¸ñ¡£
    WS_CLIPSIBLINGS = &H4000000                  '/* ÅÅ³ý×Ó´°¿ÚÖ®¼äµÄÏà¶ÔÇøÓò£¬Ò²¾ÍÊÇ£¬µ±Ò»¸öÌØ¶¨µÄ´°¿Ú½ÓÊÕµ½WM_PAINTÏûÏ¢Ê±£¬WS_CLIPSIBLINGS ·ç¸ñ½«ËùÓÐ²ãµþ´°¿ÚÅÅ³ýÔÚ»æÍ¼Ö®Íâ£¬Ö»ÖØ»æÖ¸¶¨µÄ×Ó´°¿Ú¡£Èç¹ûÎ´Ö¸¶¨WS_CLIPSIBLINGS·ç¸ñ£¬²¢ÇÒ×Ó´°¿ÚÊÇ²ãµþµÄ£¬ÔòÔÚÖØ»æ×Ó´°¿ÚµÄ¿Í»§ÇøÊ±£¬¾Í»áÖØ»æÁÚ½üµÄ×Ó´°¿Ú¡£
    WS_DISABLED = &H8000000                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª½ûÖ¹µÄ×Ó´°¿Ú¡£Ò»¸ö½ûÖ¹×´Ì¬µÄ´°ÈÕ²»ÄÜ½ÓÊÜÀ´×ÔÓÃ»§µÄÊäÈËÐÅÏ¢¡£
    WS_DLGFRAME = &H400000                       '/* ´´½¨Ò»¸ö´ø¶Ô»°¿ò±ß¿ò·ç¸ñµÄ´°¿Ú¡£ÕâÖÖ·ç¸ñµÄ´°¿Ú²»ÄÜ´ø±êÌâÌõ¡£
    WS_GROUP = &H20000                           '/* Ö¸¶¨Ò»×é¿ØÖÆµÄµÚÒ»¸ö¿ØÖÆ¡£Õâ¸ö¿ØÖÆ×éÓÉµÚÒ»¸ö¿ØÖÆºÍËæºó¶¨ÒåµÄ¿ØÖÆ×é³É£¬×ÔµÚ¶þ¸ö¿ØÖÆ¿ªÊ¼Ã¿¸ö¿ØÖÆ£¬¾ßÓÐWS_GROUP·ç¸ñ£¬Ã¿¸ö×éµÄµÚÒ»¸ö¿ØÖÆ´øÓÐWS_TABSTOP·ç¸ñ£¬´Ó¶øÊ¹ÓÃ»§¿ÉÒÔÔÚ×é¼äÒÆ¶¯¡£ÓÃ»§Ëæºó¿ÉÒÔÊ¹ÓÃ¹â±êÔÚ×éÄÚµÄ¿ØÖÆ¼ä¸Ä±ä¼üÅÌ½¹µã¡£
    WS_HSCROLL = &H100000                        '/* ´´½¨Ò»¸öÓÐË®Æ½¹ö¶¯ÌõµÄ´°¿Ú¡£
    WS_MAXIMIZE = &H1000000                      '/* ´´½¨Ò»¸ö¾ßÓÐ×î´ó»¯°´Å¥µÄ´°¿Ú¡£¸Ã·ç¸ñ²»ÄÜÓëWS_EX_CONTEXTHELP·ç¸ñÍ¬Ê±³öÏÖ£¬Í¬Ê±±ØÐëÖ¸¶¨WS_SYSMENU·ç¸ñ¡£
    WS_MAXIMIZEBOX = &H10000                     '/*
    WS_MINIMIZE = &H20000000                     '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª×îÐ¡»¯×´Ì¬µÄ´°¿Ú¡£
    WS_ICONIC = WS_MINIMIZE                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª×îÐ¡»¯×´Ì¬µÄ´°¿Ú¡£ÓëWS_MINIMIZE·ç¸ñÏàÍ¬¡£
    WS_MINIMIZEBOX = &H20000                     '/*
    WS_OVERLAPPED = &H0&                         '/* ²úÉúÒ»¸ö²ãµþµÄ´°¿Ú¡£Ò»¸ö²ãµþµÄ´°¿ÚÓÐÒ»¸ö±êÌâÌõºÍÒ»¸ö±ß¿ò¡£ÓëWS_TILED·ç¸ñÏàÍ¬
    WS_POPUP = &H80000000                        '/* ´´½¨Ò»¸öµ¯³öÊ½´°¿Ú¡£¸Ã·ç¸ñ²»ÄÜÓëWS_CHLD·ç¸ñÍ¬Ê±Ê¹ÓÃ¡£
    WS_SYSMENU = &H80000                         '/* ´´½¨Ò»¸öÔÚ±êÌâÌõÉÏ´øÓÐ´°¿Ú²Ëµ¥µÄ´°¿Ú£¬±ØÐëÍ¬Ê±Éè¶¨WS_CAPTION·ç¸ñ¡£
    WS_TABSTOP = &H10000                         '/* ´´½¨Ò»¸ö¿ØÖÆ£¬Õâ¸ö¿ØÖÆÔÚÓÃ»§°´ÏÂTab¼üÊ±¿ÉÒÔ»ñµÃ¼üÅÌ½¹µã¡£°´ÏÂTab¼üºóÊ¹¼üÅÌ½¹µã×ªÒÆµ½ÏÂÒ»¾ßÓÐWS_TABSTOP·ç¸ñµÄ¿ØÖÆ¡£
    WS_THICKFRAME = &H40000                      '/* ´´½¨Ò»¸ö¾ßÓÐ¿Éµ÷±ß¿òµÄ´°¿Ú¡£
    WS_SIZEBOX = WS_THICKFRAME                   '/* ÓëWS_THICKFRAME·ç¸ñÏàÍ¬
    WS_TILED = WS_OVERLAPPED                     '/* ²úÉúÒ»¸ö²ãµþµÄ´°¿Ú¡£Ò»¸ö²ãµþµÄ´°¿ÚÓÐÒ»¸ö±êÌâºÍÒ»¸ö±ß¿ò¡£ÓëWS_OVERLAPPED·ç¸ñÏàÍ¬¡£
    WS_VISIBLE = &H10000000                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª¿É¼ûµÄ´°¿Ú¡£
    WS_VSCROLL = &H200000                        '/* ´´½¨Ò»¸öÓÐ´¹Ö±¹ö¶¯ÌõµÄ´°¿Ú¡£
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW         '/* ´´½¨Ò»¸ö¾ßÓÐWS_OVERLAPPED£¬WS_CAPTION£¬WS_SYSMENU MS_THICKFRAME£®
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '/* ´´½¨Ò»¸ö¾ßÓÐWS_BORDER£¬WS_POPUP,WS_SYSMENU·ç¸ñµÄ´°¿Ú£¬WS_CAPTIONºÍWS_POPUPWINDOW±ØÐëÍ¬Ê±Éè¶¨²ÅÄÜÊ¹´°¿ÚÄ³µ¥¿É¼û¡£
    ' CreateWindowEx
    WS_EX_ACCEPTFILES = &H10&                    '/* Ö¸¶¨ÒÔ¸Ã·ç¸ñ´´½¨µÄ´°¿Ú½ÓÊÜÒ»¸öÍÏ×§ÎÄ¼þ¡£
    WS_EX_APPWINDOW = &H40000                    '/* µ±´°¿Ú¿É¼ûÊ±£¬½«Ò»¸ö¶¥²ã´°¿Ú·ÅÖÃµ½ÈÎÎñÌõÉÏ¡£
    WS_EX_CLIENTEDGE = &H200                     '/* Ö¸¶¨´°¿ÚÓÐÒ»¸ö´øÒõÓ°µÄ±ß½ç¡£
    WS_EX_CONTEXTHELP = &H400                    '/* ÔÚ´°¿ÚµÄ±êÌâÌõ°üº¬Ò»¸öÎÊºÅ±êÖ¾¡£µ±ÓÃ»§µã»÷ÁËÎÊºÅÊ±£¬Êó±ê¹â±ê±äÎªÒ»¸öÎÊºÅµÄÖ¸Õë¡¢Èç¹ûµã»÷ÁËÒ»¸ö×Ó´°¿Ú£¬Ôò×Ó´°ÈÕ½ÓÊÕµ½WM_HELPÏûÏ¢¡£×Ó´°¿ÚÓ¦¸Ã½«Õâ¸öÏûÏ¢´«µÝ¸ø¸¸´°¿Ú¹ý³Ì£¬¸¸´°¿ÚÔÙÍ¨¹ýHELP_WM_HELPÃüÁîµ÷ÓÃWinHelpº¯Êý¡£Õâ¸öHelpÓ¦ÓÃ³ÌÐòÏÔÊ¾Ò»¸ö°üº¬×Ó´°¿Ú°ïÖúÐÅÏ¢µÄµ¯³öÊ½´°¿Ú¡£ WS_EX_CONTEXTHELP²»ÄÜÓëWS_MAXIMIZEBOXºÍWS_MINIMIZEBOXÍ¬Ê±Ê¹ÓÃ¡£
    WS_EX_CONTROLPARENT = &H10000                '/* ÔÊÐíÓÃ»§Ê¹ÓÃTab¼üÔÚ´°¿ÚµÄ×Ó´°¿Ú¼äËÑË÷¡£
    WS_EX_DLGMODALFRAME = &H1&                   '/* ´´½¨Ò»¸ö´øË«±ßµÄ´°¿Ú£»¸Ã´°¿Ú¿ÉÒÔÔÚdwStyleÖÐÖ¸¶¨WS_CAPTION·ç¸ñÀ´´´½¨Ò»¸ö±êÌâÀ¸¡£
    WS_EX_LEFT = &H0                             '/* ´°¿Ú¾ßÓÐ×ó¶ÔÆëÊôÐÔ£¬ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_LEFTSCROLLBAR = &H4000                 '/* Èç¹ûÍâ¿ÇÓïÑÔÊÇÈçHebrew£¬Arabic£¬»òÆäËûÖ§³Öreading order alignmentµÄÓïÑÔ£¬Ôò±êÌâÌõ£¨Èç¹û´æÔÚ£©ÔòÔÚ¿Í»§ÇøµÄ×ó²¿·Ö¡£ÈôÊÇÆäËûÓïÑÔ£¬ÔÚ¸Ã·ç¸ñ±»ºöÂÔ²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_LTRREADING = &H0                       '/* ´°¿ÚÎÄ±¾ÒÔLEFTµ½RIGHT£¨×Ô×óÏòÓÒ£©ÊôÐÔµÄË³ÐòÏÔÊ¾¡£ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_MDICHILD = &H40                        '/* ´´½¨Ò»¸öMDI×Ó´°¿Ú¡£
    WS_EX_NOACTIVATE = &H8000000                 '/*
    WS_EX_NOPATARENTNOTIFY = &H4&                '/* Ö¸Ã÷ÒÔÕâ¸ö·ç¸ñ´´½¨µÄ´°¿ÚÔÚ±»´´½¨ºÍÏú»ÙÊ±²»Ïò¸¸´°¿Ú·¢ËÍWM_PARENTNOTFYÏûÏ¢¡£
    WS_EX_OVERLAPPEDWINDOW = &H300               '/*
    WS_EX_PALETTEWINDOW = &H188                  '/* WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOWºÍWS_WX_TOPMOST·ç¸ñµÄ×éºÏWS_EX_RIGHT:´°¿Ú¾ßÓÐÆÕÍ¨µÄÓÒ¶ÔÆëÊôÐÔ£¬ÕâÒÀÀµÓÚ´°¿ÚÀà¡£Ö»ÓÐÔÚÍâ¿ÇÓïÑÔÊÇÈçHebrew,Arabic»òÆäËûÖ§³Ö¶ÁË³Ðò¶ÔÆë£¨reading order alignment£©µÄÓïÑÔÊ±¸Ã·ç¸ñ²ÅÓÐÐ§£¬·ñÔò£¬ºöÂÔ¸Ã±êÖ¾²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_RIGHT = &H1000                         '/*
    WS_EX_RIGHTSCROLLBAR = &H0                   '/* ´¹Ö±¹ö¶¯ÌõÔÚ´°¿ÚµÄÓÒ±ß½ç¡£ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_RTLREADING = &H2000                    '/* Èç¹ûÍâ¿ÇÓïÑÔÊÇÈçHebrew£¬Arabic£¬»òÆäËûÖ§³Ö¶ÁË³Ðò¶ÔÆë£¨reading order alignment£©µÄÓïÑÔ£¬Ôò´°¿ÚÎÄ±¾ÊÇÒ»×Ô×óÏòÓÒ£©RIGHTµ½LEFTË³ÐòµÄ¶Á³öË³Ðò¡£ÈôÊÇÆäËûÓïÑÔ£¬ÔÚ¸Ã·ç¸ñ±»ºöÂÔ²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_STATICEDGE = &H20000                   '/* Îª²»½ÓÊÜÓÃ»§ÊäÈëµÄÏî´´½¨Ò»¸ö3Ò»Î¬±ß½ç·ç¸ñ¡£
    WS_EX_TOOLWINDOW = &H80                      '/*
    WS_EX_TOPMOST = &H8&                         '/* Ö¸Ã÷ÒÔ¸Ã·ç¸ñ´´½¨µÄ´°¿ÚÓ¦·ÅÖÃÔÚËùÓÐ·Ç×î¸ß²ã´°¿ÚµÄÉÏÃæ²¢ÇÒÍ£ÁôÔÚÆäL£¬¼´Ê¹´°¿ÚÎ´±»¼¤»î¡£Ê¹ÓÃº¯ÊýSetWindowPosÀ´ÉèÖÃºÍÒÆÈ¥Õâ¸ö·ç¸ñ¡£
    WS_EX_TRANSPARENT = &H20&                    '/* Ö¸¶¨ÒÔÕâ¸ö·ç¸ñ´´½¨µÄ´°¿ÚÔÚ´°¿ÚÏÂµÄÍ¬Êô´°¿ÚÒÑÖØ»­Ê±£¬¸Ã´°¿Ú²Å¿ÉÒÔÖØ»­¡£
    WS_EX_WINDOWEDGE = &H100
End Enum

' Windows»·¾³ÓÐ¹ØµÄÐÅÏ¢£¬ÓÃÓÚGetSystemMetricsº¯Êý
 Enum KhanSystemMetricsFlags
    SM_CXSCREEN = 0                              '/* ÆÁÄ»´óÐ¡ */
    SM_CYSCREEN = 1                              '/* ÆÁÄ»´óÐ¡ */
    SM_CXVSCROLL = 2                             '/* ´¹Ö±¹ö¶¯ÌõÖÐµÄ¼ýÍ·°´Å¥µÄ´óÐ¡ */
    SM_CYHSCROLL = 3                             '/* Ë®Æ½¹ö¶¯ÌõÉÏµÄ¼ýÍ·´óÐ¡ */
    SM_CYCAPTION = 4                             '/* ´°¿Ú±êÌâµÄ¸ß¶È */
    SM_CXBORDER = 5                              '/* ³ß´ç²»¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CYBORDER = 6                              '/* ³ß´ç²»¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CXDLGFRAME = 7                            '/* ¶Ô»°¿ò±ß¿òµÄ´óÐ¡ */
    SM_CYDLGFRAME = 8                            '/* ¶Ô»°¿ò±ß¿òµÄ´óÐ¡ */
    SM_CYVTHUMB = 9                              '/* ¹ö¶¯¿éÔÚË®Æ½¹ö¶¯ÌõÉÏµÄ´óÐ¡ */
    SM_CXHTHUMB = 10                             '/* ¹ö¶¯¿éÔÚË®Æ½¹ö¶¯ÌõÉÏµÄ´óÐ¡ */
    SM_CXICON = 11                               '/* ±ê×¼Í¼±êµÄ´óÐ¡ */
    SM_CYICON = 12                               '/* ±ê×¼Í¼±êµÄ´óÐ¡ */
    SM_CXCURSOR = 13                             '/* ±ê×¼Ö¸Õë´óÐ¡ */
    SM_CYCURSOR = 14                             '/* ±ê×¼Ö¸Õë´óÐ¡ */
    SM_CYMENU = 15                               '/* ²Ëµ¥¸ß¶È */
    SM_CXFULLSCREEN = 16                         '/* ×î´ó»¯´°¿Ú¿Í»§ÇøµÄ´óÐ¡ */
    SM_CYFULLSCREEN = 17                         '/* ×î´ó»¯´°¿Ú¿Í»§ÇøµÄ´óÐ¡ */
    SM_CYKANJIWINDOW = 18                        '/* Kanji´°¿ÚµÄ´óÐ¡£¨Height of Kanji window£© */
    SM_MOUSEPRESENT = 19                         '/* Èç°²×°ÁËÊó±êÔòÎªTRUE */
    SM_CYVSCROLL = 20                            '/* ´¹Ö±¹ö¶¯ÌõÖÐµÄ¼ýÍ·°´Å¥µÄ´óÐ¡ */
    SM_CXHSCROLL = 21                            '/* Ë®Æ½¹ö¶¯ÌõÉÏµÄ¼ýÍ·´óÐ¡ */
    SM_DEBUG = 22                                '/* ÈçwindowsµÄµ÷ÊÔ°æÕýÔÚÔËÐÐ£¬ÔòÎªTRUE */
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28                                '/* ´°¿ÚµÄ×îÐ¡³ß´ç */
    SM_CYMIN = 29                                '/* ´°¿ÚµÄ×îÐ¡³ß´ç */
    SM_CXSIZE = 30                               '/* ±êÌâÀ¸Î»Í¼µÄ´óÐ¡ */
    SM_CYSIZE = 31                               '/* ±êÌâÀ¸Î»Í¼µÄ´óÐ¡ */
    SM_CXFRAME = 32                              '/* ³ß´ç¿É±ä±ß¿òµÄ´óÐ¡£¨ÔÚwin95ºÍnt 4.0ÖÐÊ¹ÓÃSM_C?FIXEDFRAME£© */
    SM_CYFRAME = 33                              '/* ³ß´ç¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CXMINTRACK = 34                           '/* ´°¿ÚµÄ×îÐ¡¹ì¼£¿í¶È */
    SM_CYMINTRACK = 35                           '/* ´°¿ÚµÄ×îÐ¡¹ì¼£¿í¶È */
    SM_CXDOUBLECLK = 36                          '/* Ë«»÷ÇøÓòµÄ´óÐ¡£¨Ö¸¶¨ÆÁÄ»ÉÏÒ»¸öÌØ¶¨µÄÏÔÊ¾ÇøÓò£¬Ö»ÓÐÔÚÕâ¸öÇøÓòÄÚÁ¬Ðø½øÐÐÁ½´ÎÊó±êµ¥»÷£¬²ÅÓÐ¿ÉÄÜ±»µ±×÷Ë«»÷ÊÂ¼þ´¦Àí£© */
    SM_CYDOUBLECLK = 37                          '/* Ë«»÷ÇøÓòµÄ´óÐ¡ */
    SM_CXICONSPACING = 38                        '/* ×ÀÃæÍ¼±êÖ®¼äµÄ¼ä¸ô¾àÀë¡£ÔÚwin95ºÍnt 4.0ÖÐÊÇÖ¸´óÍ¼±êµÄ¼ä¾à */
    SM_CYICONSPACING = 39                        '/* ×ÀÃæÍ¼±êÖ®¼äµÄ¼ä¸ô¾àÀë¡£ÔÚwin95ºÍnt 4.0ÖÐÊÇÖ¸´óÍ¼±êµÄ¼ä¾à */
    SM_MENUDROPALIGNMENT = 40                    '/* Èçµ¯³öÊ½²Ëµ¥¶ÔÆë²Ëµ¥À¸ÏîÄ¿µÄ×ó²à£¬ÔòÎªÁã */
    SM_PENWINDOWS = 41                           '/* Èç×°ÔØÁËÖ§³Ö±Ê´°¿ÚµÄDLL£¬Ôò±íÊ¾±Ê´°¿ÚµÄ¾ä±ú */
    SM_DBCSENABLED = 42                          '/* ÈçÖ§³ÖË«×Ö½ÚÔòÎªTRUE */
    SM_CMOUSEBUTTONS = 43                        '/* Êó±ê°´Å¥£¨°´¼ü£©µÄÊýÁ¿¡£ÈçÃ»ÓÐÊó±ê£¬¾ÍÎªÁã */
    SM_CMETRICS = 44                             '/* ¿ÉÓÃÏµÍ³»·¾³µÄÊýÁ¿ */
End Enum

' SetMapMode
 Enum KhanMapModeStyles
    MM_ANISOTROPIC = 8                           '/* Âß¼­µ¥Î»×ª»»³É¾ßÓÐÈÎÒâ±ÈÀýÖáµÄÈÎÒâµ¥Î»£¬ÓÃSetWindowExtExºÍSetViewportExtExº¯Êý¿ÉÖ¸¶¨µ¥Î»¡¢·½ÏòºÍ±ÈÀý¡£
    MM_HIENGLISH = 5                             '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª0.001inch(Ó¢´ç)£¬XµÄÕý·½ÃæÏòÓÒ£¬YµÄÕý·½ÏòÏòÉÏ
    MM_HIMETRIC = 3                              '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª0.01millimeter(ºÁÃ×)£¬XÕý·½ÏòÏòÓÒ£¬YµÄÕý·½ÏòÏòÉÏ¡£
    MM_ISOTROPIC = 7                             '/* ÊÓ¿ÚºÍ´°¿Ú·¶Î§ÈÎÒâ£¬Ö»ÊÇxºÍyÂß¼­µ¥Ôª³ß´çÒªÏàÍ¬
    MM_LOENGLISH = 4                             '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªÓ¢´ç£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÉÏ¡£
    MM_LOMETRIC = 2                              '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªºÁÃ×£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÉÏ¡£
    MM_TEXT = 1                                  '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªÒ»¸öÉèÖÃ±¸ËØ£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÏÂ¡£
    MM_TWIPS = 6                                 '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª1 twip (1/1440 inch)£¬XÕý·½ÏòÏòÓÒ£¬Y·½ÏòÏòÉÏ¡£
End Enum

' GetROP2,SetROP2
 Enum EnumDrawModeFlags
    R2_BLACK = 1                                 '/* ºÚÉ«
    R2_COPYPEN = 13                              '/* »­±ÊÑÕÉ«
    R2_LAST = 16
    R2_MASKNOTPEN = 3                            '/* »­±ÊÑÕÉ«µÄ·´É«ÓëÏÔÊ¾ÑÕÉ«½øÐÐANDÔËËã
    R2_MASKPEN = 9                               '/* ÏÔÊ¾ÑÕÉ«Óë»­±ÊÑÕÉ«½øÐÐANDÔËËã
    R2_MASKPENNOT = 5                            '/* ÏÔÊ¾ÑÕÉ«µÄ·´É«Óë»­±ÊÑÕÉ«½øÐÐANDÔËËã
    R2_MERGENOTPEN = 12                          '/* »­±ÊÑÕÉ«µÄ·´É«ÓëÏÔÊ¾ÑÕÉ«½øÐÐORÔËËã
    R2_MERGEPEN = 15                             '/* »­±ÊÑÕÉ«ÓëÏÔÊ¾ÑÕÉ«½øÐÐORÔËËã
    R2_MERGEPENNOT = 14                          '/* ÏÔÊ¾ÑÕÉ«µÄ·´É«Óë»­±ÊÑÕÉ«½øÐÐORÔËËã
    R2_NOP = 11                                  '/* ²»±ä
    R2_NOT = 6                                   '/* µ±Ç°ÏÔÊ¾ÑÕÉ«µÄ·´É«
    R2_NOTCOPYPEN = 4                            '/* R2_COPYPENµÄ·´É«
    R2_NOTMASKPEN = 8                            '/* R2_MASKPENµÄ·´É«
    R2_NOTMERGEPEN = 2                           '/* R2_MERGEPENµÄ·´É«
    R2_NOTXORPEN = 10                            '/* R2_XORPENµÄ·´É«
    R2_WHITE = 16                                '/* °×É«
    R2_XORPEN = 7                                '/* ÏÔÊ¾ÑÕÉ«Óë»­±ÊÑÕÉ«½øÐÐÒì»òÔËËã
End Enum

' ======================================================================================
' Types
' ======================================================================================

Private Type tagINITCOMMONCONTROLSEX              '/* icc
   dwSize                   As Long              '/* size of this structure
   dwICC                    As Long              '/* flags indicating which classes to be initialized.
End Type

Private Type POINTAPI
   x                        As Long
   y                        As Long
End Type

 Private Type RECT
   Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Private Type LOGPEN
    lopnStyle               As Long
    lopnWidth               As POINTAPI
    lopnColor               As Long
End Type

Private Type LOGBRUSH
   lbStyle                  As Long
   lbColor                  As Long
   lbHatch                  As Long
End Type

' Õâ¸ö½á¹¹°üº¬ÁË¸½¼ÓµÄ»æÍ¼²ÎÊý£¬º¯ÊýDrawTextEx
Private Type DRAWTEXTPARAMS
    cbSize                  As Long              '/* Specifies the structure size, in bytes */
    iTabLength              As Long              '/* Specifies the size of each tab stop, in units equal to the average character width */
    iLeftMargin             As Long              '/* Specifies the left margin, in units equal to the average character width */
    iRightMargin            As Long              '/* Specifies the right margin, in units equal to the average character width */
    uiLengthDrawn           As Long              '/* Receives the number of characters processed by DrawTextEx, including white-space characters. */
                                                 '/* The number can be the length of the string or the index of the first line that falls below the drawing area. */
                                                 '/* Note that DrawTextEx always processes the entire string if the DT_NOCLIP formatting flag is specified */
End Type

Private Const LF_FACESIZE   As Long = 32
 Private Type LOGFONT
   lfHeight                 As Long              '/* The font size (see below) */
   lfWidth                  As Long              '/* Normally you don't set this, just let Windows create the Default */
   lfEscapement             As Long              '/* The angle, in 0.1 degrees, of the font */
   lfOrientation            As Long              '/* Leave as default */
   lfWeight                 As Long              '/* Bold, Extra Bold, Normal etc */
   lfItalic                 As Byte              '/* As it says */
   lfUnderline              As Byte              '/* As it says */
   lfStrikeOut              As Byte              '/* As it says */
   lfCharSet                As Byte              '/* As it says */
   lfOutPrecision           As Byte              '/* Leave for default */
   lfClipPrecision          As Byte              '/* Leave for defaultv
   lfQuality                As Byte              '/* Leave for default */
   lfPitchAndFamily         As Byte              '/* Leave for default */
   lfFaceName(LF_FACESIZE)  As Byte              '/* The font name converted to a byte array */
End Type

Private Type ICONINFO
   fIcon                    As Long
   xHotspot                 As Long
   yHotspot                 As Long
   hbmMask                  As Long
   hbmColor                 As Long
End Type

Private Type IMAGEINFO
    hBitmapImage            As Long
    hBitmapMask             As Long
    cPlanes                 As Long
    cBitsPerPixel           As Long
    rcImage                 As RECT
End Type

'/* DIB µÄÎÄ¼þ´óÐ¡¼°¼Ü¹¹Ñ¶Ï¢ */
Private Type BITMAPFILEHEADER
    bfType                  As Integer           '/* Ö¸¶¨ÎÄ¼þÀàÐÍ£¬±ØÐë BM("magic cookie" - must be "BM" (19778)) */
    bfSize                  As Long              '/* Ö¸¶¨Î»Í¼ÎÄ¼þ´óÐ¡£¬ÒÔÎ»Ôª×éÎªµ¥Î» */
    bfReserved1             As Integer           '/* ±£Áô£¬±ØÐëÉèÎª0 */
    bfReserved2             As Integer           '/* Í¬ÉÏ */
    bfOffBits               As Long              '/* ´Ó´Ë¼Ü¹¹µ½Î»Í¼Êý¾ÝÎ»µÄÎ»Ôª×éÆ«ÒÆÁ¿ */
End Type

'/* Éè±¸ÎÞ¹ØÎ»Í¼ (DIB)µÄ´óÐ¡¼°ÑÕÉ«ÐÅÏ¢  (ËüÎ»ÓÚ bmp ÎÄ¼þµÄ¿ªÍ·´¦) 40 bytes */
 Private Type BITMAPINFOHEADER
    biSize                  As Long              '/* ½á¹¹³¤¶È */
    biWidth                 As Long              '/* Ö¸¶¨Î»Í¼µÄ¿í¶È£¬ÒÔÏñËØÎªµ¥Î» */
    biHeight                As Long              '/* Ö¸¶¨Î»Í¼µÄ¸ß¶È£¬ÒÔÏñËØÎªµ¥Î» */
    biPlanes                As Integer           '/* Ö¸¶¨Ä¿±êÉè±¸µÄ¼¶Êý(±ØÐëÎª 1 ) */
    biBitCount              As Integer           '/* Î»Í¼µÄÑÕÉ«Î»Êý,Ã¿Ò»¸öÏñËØµÄÎ»(1£¬4£¬8£¬16£¬24£¬32) */
    biCompression           As Long              '/* Ö¸¶¨Ñ¹ËõÀàÐÍ(BI_RGB Îª²»Ñ¹Ëõ) */
    biSizeImage             As Long              '/* Í¼ÏóµÄ´óÐ¡,ÒÔ×Ö½ÚÎªµ¥Î»,µ±ÓÃBI_RGB¸ñÊ½ÊÇ,¿ÉÉèÖÃÎª0 */
    biXPelsPerMeter         As Long              '/* Ö¸¶¨Éè±¸Ë®×¼·Ö±æÂÊ£¬ÒÔÃ¿Ã×µÄÏñËØÎªµ¥Î» */
    biYPelsPerMeter         As Long              '/* ´¹Ö±·Ö±æÂÊ£¬ÆäËûÍ¬ÉÏ */
    biClrUsed               As Long              '/* ËµÃ÷Î»Í¼Êµ¼ÊÊ¹ÓÃµÄ²ÊÉ«±íÖÐµÄÑÕÉ«Ë÷ÒýÊý,ÉèÎª0µÄ»°,ËµÃ÷Ê¹ÓÃËùÓÐµ÷É«°åÏî */
    biClrImportant          As Long              '/* ËµÃ÷¶ÔÍ¼ÏóÏÔÊ¾ÓÐÖØÒªÓ°ÏìµÄÑÕÉ«Ë÷ÒýµÄÊýÄ¿£¬Èç¹ûÊÇ0£¬±íÊ¾¶¼ÖØÒª */
End Type

'/* ÃèÊöÁËÓÉºì¡¢ÂÌ¡¢À¶×é³ÉµÄÑÕÉ«×éºÏ */
 Private Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte              '/* '±£Áô£¬±ØÐëÎª 0 */
End Type

Private Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

 Private Type BITMAPINFO_1BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 1)        As RGBQUAD
End Type

 Private Type BITMAPINFO_4BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 15)       As RGBQUAD
End Type

 Private Type BITMAPINFO_8BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 255)      As RGBQUAD
End Type

 Private Type BITMAPINFO_ABOVE8
   bmiHeader                As BITMAPINFOHEADER
End Type

 Private Type BITMAP
    bmType                  As Long              '/* Type of bitmap */
    bmWidth                 As Long              '/* Pixel width */
    bmHeight                As Long              '/* Pixel height */
    bmWidthBytes            As Long              '/* Byte width = 3 x Pixel width */
    bmPlanes                As Integer           '/* Color depth of bitmap */
    bmBitsPixel             As Integer           '/* Bits per pixel, must be 16 or 24 */
    bmBits                  As Long              '/* This is the pointer to the bitmap data */
End Type

' AlphaBlend
 Private Type BLENDFUNCTION
   BlendOp                  As Byte
   BlendFlags               As Byte
   SourceConstantAlpha      As Byte
   AlphaFormat              As Byte
End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

Private Const MAX_PATH = 260

' ======================================================================================
' Types:
' ======================================================================================

Private Type PICTDESC
    cbSizeofStruct  As Long
    picType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

Private Type Guid
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

' ======================================================================================
' API declares:
' ======================================================================================

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
' ÉèÖÃÖ¸¶¨Éè±¸³¡¾°µÄ»æÍ¼Ä£Ê½¡£ÓëvbµÄDrawModeÊôÐÔÍêÈ«Ò»ÖÂ
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

' ======================================================================================
' API declares:
' ======================================================================================

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------ÏûÏ¢º¯ÊýºÍÏûÏ¢ÁÐ¶Óº¯Êý---------------------------------©§
'©§                                                                                    ©§
'
' µ÷ÓÃÒ»¸ö´°¿ÚµÄ´°¿Úº¯Êý£¬½«Ò»ÌõÏûÏ¢·¢¸øÄÇ¸ö´°¿Ú¡£³ý·ÇÏûÏ¢´¦ÀíÍê±Ï£¬·ñÔò¸Ãº¯Êý²»»á·µ»Ø¡£
' SendMessageBynum£¬ SendMessageByStringÊÇ¸Ãº¯ÊýµÄ¡°ÀàÐÍ°²È«¡±ÉùÃ÷ÐÎÊ½
 Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Private Declare Function SendMessageByString Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
 Private Declare Function SendMessageByLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ½«Ò»ÌõÏûÏ¢Í¶µÝµ½Ö¸¶¨´°¿ÚµÄÏûÏ¢¶ÓÁÐ¡£Í¶µÝµÄÏûÏ¢»áÔÚWindowsÊÂ¼þ´¦Àí¹ý³ÌÖÐµÃµ½´¦Àí¡£
' ÔÚÄÇ¸öÊ±ºò£¬»áËæÍ¬Í¶µÝµÄÏûÏ¢µ÷ÓÃÖ¸¶¨´°¿ÚµÄ´°¿Úº¯Êý¡£ÌØ±ðÊÊºÏÄÇÐ©²»ÐèÒªÁ¢¼´´¦ÀíµÄ´°¿ÚÏûÏ¢µÄ·¢ËÍ
 Private Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§--------------------------------´°¿Úº¯Êý(Window)------------------------------------©§
'©§                                                                                    ©§
'
' Creating new windows:
 Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' ×îÐ¡»¯Ö¸¶¨µÄ´°¿Ú¡£´°¿Ú²»»á´ÓÄÚ´æÖÐÇå³ý
 Private Declare Function CloseWindow Lib "USER32" (ByVal hWnd As Long) As Long
' ÆÆ»µ£¨¼´Çå³ý£©Ö¸¶¨µÄ´°¿ÚÒÔ¼°ËüµÄËùÓÐ×Ó´°¿Ú
 Private Declare Function DestroyWindow Lib "USER32" (ByVal hWnd As Long) As Long
' ÔÚÖ¸¶¨µÄ´°¿ÚÀïÔÊÐí»ò½ûÖ¹ËùÓÐÊó±ê¼°¼üÅÌÊäÈë
 Private Declare Function EnableWindow Lib "USER32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
' ÔÚ´°¿ÚÁÐ±íÖÐÑ°ÕÒÓëÖ¸¶¨Ìõ¼þÏà·ûµÄµÚÒ»¸ö×Ó´°¿Ú
 Private Declare Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' ÅÐ¶ÏÖ¸¶¨´°¿ÚµÄ¸¸´°¿Ú
 Private Declare Function GetParent Lib "USER32" (ByVal hWnd As Long) As Long
' Ö¸¶¨Ò»¸ö´°¿ÚµÄÐÂ¸¸£¨ÔÚvbÀïÊ¹ÓÃ£ºÀûÓÃÕâ¸öº¯Êý£¬vb¿ÉÒÔ¶àÖÖÐÎÊ½Ö§³Ö×Ó´°¿Ú¡£
' ÀýÈç£¬¿É½«¿Ø¼þ´ÓÒ»¸öÈÝÆ÷ÒÆÖÁ´°ÌåÖÐµÄÁíÒ»¸ö¡£ÓÃÕâ¸öº¯ÊýÔÚ´°Ìå¼äÒÆ¶¯¿Ø¼þÊÇÏàµ±Ã°ÏÕµÄ£¬
' µ«È´²»Ê§ÎªÒ»¸öÓÐÐ§µÄ°ì·¨¡£ÈçÕæµÄÕâÑù×ö£¬ÇëÔÚ¹Ø±ÕÈÎºÎÒ»¸ö´°ÌåÖ®Ç°£¬×¢ÒâÓÃSetParent½«¿Ø¼þµÄ¸¸Éè»ØÔ­À´µÄÄÇ¸ö£©
 Private Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Ëø¶¨Ö¸¶¨´°¿Ú£¬½ûÖ¹Ëü¸üÐÂ¡£Í¬Ê±Ö»ÄÜÓÐÒ»¸ö´°¿Ú´¦ÓÚËø¶¨×´Ì¬
 Private Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long
' Ç¿ÖÆÁ¢¼´¸üÐÂ´°¿Ú£¬´°¿ÚÖÐÒÔÇ°ÆÁ±ÎµÄËùÓÐÇøÓò¶¼»áÖØ»­
' ÔÚvbÀïÊ¹ÓÃ£ºÈçvb´°Ìå»ò¿Ø¼þµÄÈÎºÎ²¿·ÖÐèÒª¸üÐÂ£¬¿É¿¼ÂÇÖ±½ÓÊ¹ÓÃrefresh·½·¨
 Private Declare Function UpdateWindow Lib "USER32" (ByVal hWnd As Long) As Long
' ÅÐ¶ÏÒ»¸ö´°¿Ú¾ä±úÊÇ·ñÓÐÐ§
 Private Declare Function IsWindow Lib "USER32" (ByVal hWnd As Long) As Long
' ¿ØÖÆ´°¿ÚµÄ¿É¼ûÐÔ
 Private Declare Function ShowWindow Lib "USER32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
' ¸Ä±äÖ¸¶¨´°¿ÚµÄÎ»ÖÃºÍ´óÐ¡¡£¶¥¼¶´°¿Ú¿ÉÄÜÊÜ×î´ó»ò×îÐ¡³ß´çµÄÏÞÖÆ£¬ÄÇÐ©³ß´çÓÅÏÈÓÚÕâÀïÉèÖÃµÄ²ÎÊý
 Private Declare Function MoveWindow Lib "USER32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' Õâ¸öº¯ÊýÄÜÎª´°¿ÚÖ¸¶¨Ò»¸öÐÂÎ»ÖÃºÍ×´Ì¬¡£ËüÒ²¿É¸Ä±ä´°¿ÚÔÚÄÚ²¿´°¿ÚÁÐ±íÖÐµÄÎ»ÖÃ¡£
' ¸Ãº¯ÊýÓëDeferWindowPosº¯ÊýÏàËÆ£¬Ö»ÊÇËüµÄ×÷ÓÃÊÇÁ¢¼´±íÏÖ³öÀ´µÄ
' ÔÚvbÀïÊ¹ÓÃ£ºÕë¶Ôvb´°Ìå£¬ÈçËüÃÇÔÚwin32ÏÂÆÁ±Î»ò×îÐ¡»¯£¬ÔòÐèÖØÉè×î¶¥²¿×´Ì¬¡£
' ÈçÓÐ±ØÒª£¬ÇëÓÃÒ»¸ö×ÓÀà´¦ÀíÄ£¿éÀ´ÖØÉè×î¶¥²¿×´Ì¬)
' ²ÎÊý
' hwnd             Óû¶¨Î»µÄ´°¿Ú
' hWndInsertAfter  ´°¿Ú¾ä±ú¡£ÔÚ´°¿ÚÁÐ±íÖÐ£¬´°¿Úhwnd»áÖÃÓÚÕâ¸ö´°¿Ú¾ä±úµÄºóÃæ£¬²Î¿´±¾Ä£¿éÃ¶¾ÙKhanSetWindowPosStyles
' x                ´°¿ÚÐÂµÄx×ø±ê¡£ÈçhwndÊÇÒ»¸ö×Ó´°¿Ú£¬ÔòxÓÃ¸¸´°¿ÚµÄ¿Í»§Çø×ø±ê±íÊ¾
' y                ´°¿ÚÐÂµÄy×ø±ê¡£ÈçhwndÊÇÒ»¸ö×Ó´°¿Ú£¬ÔòyÓÃ¸¸´°¿ÚµÄ¿Í»§Çø×ø±ê±íÊ¾
' cx               Ö¸¶¨ÐÂµÄ´°¿Ú¿í¶È
' cy               Ö¸¶¨ÐÂµÄ´°¿Ú¸ß¶È
' wFlags           °üº¬ÁËÆì±êµÄÒ»¸öÕûÊý£¬²Î¿´±¾Ä£¿éÃ¶¾ÙKhanSetWindowPosStyles
 Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
' ´ÓÖ¸¶¨´°¿ÚµÄ½á¹¹ÖÐÈ¡µÃÐÅÏ¢£¬nIndex²ÎÊý²Î¿´±¾Ä£¿é³£Á¿ÉùÃ÷
 Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' ÔÚ´°¿Ú½á¹¹ÖÐÎªÖ¸¶¨µÄ´°¿ÚÉèÖÃÐÅÏ¢£¬nIndex²ÎÊý²Î¿´±¾Ä£¿é³£Á¿ÉùÃ÷
 Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§------------------------------´°¿ÚÀàº¯Êý(Window Class)------------------------------©§
'©§                                                                                    ©§
'
' ÎªÖ¸¶¨µÄ´°¿ÚÈ¡µÃÀàÃû
 Private Declare Function GetClassName Lib "USER32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------Êó±êÊäÈëº¯Êý(Mouse Input)------------------------------©§
'
' »ñµÃÒ»¸ö´°¿ÚµÄ¾ä±ú£¬Õâ¸ö´°¿ÚÎ»ÓÚµ±Ç°ÊäÈëÏß³Ì£¬ÇÒÓµÓÐÊó±ê²¶»ñ£¨Êó±ê»î¶¯ÓÉËü½ÓÊÕ£©
 Private Declare Function GetCapture Lib "USER32" () As Long
' ½«Êó±ê²¶»ñÉèÖÃµ½Ö¸¶¨µÄ´°¿Ú¡£ÔÚÊó±ê°´Å¥°´ÏÂµÄÊ±ºò£¬Õâ¸ö´°¿Ú»áÎªµ±Ç°Ó¦ÓÃ³ÌÐò»òÕû¸öÏµÍ³½ÓÊÕËùÓÐÊó±êÊäÈë
 Private Declare Function SetCapture Lib "USER32" (ByVal hWnd As Long) As Long
' Îªµ±Ç°µÄÓ¦ÓÃ³ÌÐòÊÍ·ÅÊó±ê²¶»ñ
 Private Declare Function ReleaseCapture Lib "USER32" () As Long
' ¿ÉÒÔÄ£ÄâÒ»´ÎÊó±êÊÂ¼þ£¬±ÈÈç×ó¼üµ¥»÷¡¢Ë«»÷ºÍÓÒ¼üµ¥»÷µÈ
 Private Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
' Õâ¸öº¯ÊýÅÐ¶ÏÖ¸¶¨µÄµãÊÇ·ñÎ»ÓÚ¾ØÐÎlpRectÄÚ²¿
' Private Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Private Declare Function PtInRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------¼üÅÌÊäÈëº¯Êý(Mouse Input)------------------------------©§
'
' »ñµÃÓµÓÐÊäÈë½¹µãµÄ´°¿ÚµÄ¾ä±ú
 Private Declare Function GetFocus Lib "USER32" () As Long
' ÊäÈë½¹µãÉèµ½Ö¸¶¨µÄ´°¿Ú
 Private Declare Function SetFocus Lib "USER32" (ByVal hWnd As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§----------------×ø±ê¿Õ¼äÓë±ä»»º¯Êý(Coordinate Space Transtormation)-----------------©§
'
' ÅÐ¶Ï´°¿ÚÄÚÒÔ¿Í»§Çø×ø±ê±íÊ¾µÄÒ»¸öµãµÄÆÁÄ»×ø±ê
Private Declare Function ClientToScreen Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
' ÅÐ¶ÏÆÁÄ»ÉÏÒ»¸öÖ¸¶¨µãµÄ¿Í»§Çø×ø±ê
Private Declare Function ScreenToClient Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------Éè±¸³¡¾°º¯Êý(Device Context)-----------------------------©§
'
' ´´½¨Ò»¸öÓëÌØ¶¨Éè±¸³¡¾°Ò»ÖÂµÄÄÚ´æÉè±¸³¡¾°¡£ÔÚ»æÖÆÖ®Ç°£¬ÏÈÒªÎª¸ÃÉè±¸³¡¾°Ñ¡¶¨Ò»¸öÎ»Í¼¡£
' ²»ÔÙÐèÒªÊ±£¬¸ÃÉè±¸³¡¾°¿ÉÓÃDeleteDCº¯ÊýÉ¾³ý¡£É¾³ýÇ°£¬ÆäËùÓÐ¶ÔÏóÓ¦»Ø¸´³õÊ¼×´Ì¬
 Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
' Îª×¨ÃÅÉè±¸´´½¨Éè±¸³¡¾°
 Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
' »ñÈ¡Ö¸¶¨´°¿ÚµÄÉè±¸³¡¾°£¬ÓÃ±¾º¯Êý»ñÈ¡µÄÉè±¸³¡¾°Ò»¶¨ÒªÓÃReleaseDCº¯ÊýÊÍ·Å£¬²»ÄÜÓÃDeleteDC
 Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
' ÊÍ·ÅÓÉµ÷ÓÃGetDC»òGetWindowDCº¯Êý»ñÈ¡µÄÖ¸¶¨Éè±¸³¡¾°¡£Ëü¶ÔÀà»òË½ÓÐÉè±¸³¡¾°ÎÞÐ§£¨µ«ÕâÑùµÄµ÷ÓÃ²»»áÔì³ÉËðº¦£©
 Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
' É¾³ý×¨ÓÃÉè±¸³¡¾°»òÐÅÏ¢³¡¾°£¬ÊÍ·ÅËùÓÐÏà¹Ø´°¿Ú×ÊÔ´¡£²»Òª½«ËüÓÃÓÚGetDCº¯ÊýÈ¡»ØµÄÉè±¸³¡¾°
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
' Ã¿¸öÉè±¸³¡¾°¶¼¿ÉÄÜÓÐÑ¡ÈëÆäÖÐµÄÍ¼ÐÎ¶ÔÏó¡£ÆäÖÐ°üÀ¨Î»Í¼¡¢Ë¢×Ó¡¢×ÖÌå¡¢»­±ÊÒÔ¼°ÇøÓòµÈµÈ¡£
' Ò»´ÎÑ¡ÈëÉè±¸³¡¾°µÄÖ»ÄÜÓÐÒ»¸ö¶ÔÏó¡£Ñ¡¶¨µÄ¶ÔÏó»áÔÚÉè±¸³¡¾°µÄ»æÍ¼²Ù×÷ÖÐÊ¹ÓÃ¡£
' ÀýÈç£¬µ±Ç°Ñ¡¶¨µÄ»­±Ê¾ö¶¨ÁËÔÚÉè±¸³¡¾°ÖÐÃè»æµÄÏß¶ÎÑÕÉ«¼°ÑùÊ½
' ·µ»ØÖµÍ¨³£ÓÃÓÚ»ñµÃÑ¡ÈëDCµÄ¶ÔÏóµÄÔ­Ê¼Öµ¡£
' »æÍ¼²Ù×÷Íê³Éºó£¬Ô­Ê¼µÄ¶ÔÏóÍ¨³£Ñ¡»ØÉè±¸³¡¾°¡£ÔÚÇå³ýÒ»¸öÉè±¸³¡¾°Ç°£¬Îñ±Ø×¢Òâ»Ö¸´Ô­Ê¼µÄ¶ÔÏó
 Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' ÓÃÕâ¸öº¯ÊýÉ¾³ýGDI¶ÔÏó£¬±ÈÈç»­±Ê¡¢Ë¢×Ó¡¢×ÖÌå¡¢Î»Í¼¡¢ÇøÓòÒÔ¼°µ÷É«°åµÈµÈ¡£¶ÔÏóÊ¹ÓÃµÄËùÓÐÏµÍ³×ÊÔ´¶¼»á±»ÊÍ·Å
' ²»ÒªÉ¾³ýÒ»¸öÒÑÑ¡ÈëÉè±¸³¡¾°µÄ»­±Ê¡¢Ë¢×Ó»òÎ»Í¼¡£ÈçÉ¾³ýÒÔÎ»Í¼Îª»ù´¡µÄÒõÓ°£¨Í¼°¸£©Ë¢×Ó£¬
' Î»Í¼²»»áÓÉÕâ¸öº¯ÊýÉ¾³ý¡ª¡ªÖ»ÓÐË¢×Ó±»É¾µô
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'¸ù¾ÝÖ¸¶¨Éè±¸³¡¾°´ú±íµÄÉè±¸µÄ¹¦ÄÜ·µ»ØÐÅÏ¢
 Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
' È¡µÃ¶ÔÖ¸¶¨¶ÔÏó½øÐÐËµÃ÷µÄÒ»¸ö½á¹¹
' lpObject ÈÎºÎÀàÐÍ£¬ÓÃÓÚÈÝÄÉ¶ÔÏóÊý¾ÝµÄ½á¹¹¡£
' Õë¶Ô»­±Ê£¬Í¨³£ÊÇÒ»¸öLOGPEN½á¹¹£»Õë¶ÔÀ©Õ¹»­±Ê£¬Í¨³£ÊÇEXTLOGPEN£»
' Õë¶Ô×ÖÌåÊÇLOGBRUSH£»Õë¶ÔÎ»Í¼ÊÇBITMAP£»Õë¶ÔDIBSectionÎ»Í¼ÊÇDIBSECTION£»
' Õë¶Ôµ÷É«°å£¬Ó¦Ö¸ÏòÒ»¸öÕûÐÍ±äÁ¿£¬´ú±íµ÷É«°åÖÐµÄÌõÄ¿ÊýÁ¿
 Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' ÔÚ´°¿Ú£¨ÓÉÉè±¸³¡¾°´ú±í£©ÖÐË®Æ½ºÍ£¨»ò£©´¹Ö±¹ö¶¯¾ØÐÎ
Private Declare Function ScrollDC Lib "USER32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
' ½«Á½¸öÇøÓò×éºÏÎªÒ»¸öÐÂÇøÓò
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
' ´´½¨Ò»¸öÓÉµãX1£¬Y1ºÍX2£¬Y2ÃèÊöµÄ¾ØÐÎÇøÓò£¬²»ÓÃÊ±Ò»¶¨ÒªÓÃDeleteObjectº¯ÊýÉ¾³ý¸ÃÇøÓò
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ´´½¨Ò»¸öÓÉlpRectÈ·¶¨µÄ¾ØÐÎÇøÓò
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
' ´´½¨Ò»¸öÔ²½Ç¾ØÐÎ£¬¸Ã¾ØÐÎÓÉX1£¬Y1-X2£¬Y2È·¶¨£¬²¢ÓÉX3£¬Y3È·¶¨µÄÍÖÔ²ÃèÊöÔ²½Ç»¡¶È
' ÓÃ¸Ãº¯Êý´´½¨µÄÇøÓòÓëÓÃRoundRect APIº¯Êý»­µÄÔ²½Ç¾ØÐÎ²»ÍêÈ«ÏàÍ¬£¬ÒòÎª±¾¾ØÐÎµÄÓÒ±ßºÍÏÂ±ß²»°üÀ¨ÔÚÇøÓòÖ®ÄÚ
 Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' ÓÃÖ¸¶¨Ë¢×ÓÌî³äÖ¸¶¨ÇøÓò
 Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
' ÓÃÖ¸¶¨Ë¢×ÓÎ§ÈÆÖ¸¶¨ÇøÓò»­Ò»¸öÍâ¿ò
 Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
' ÕâÊÇÄÇÐ©ºÜÄÑÓÐÈË×¢Òâµ½µÄ¶Ô±à³ÌÕßÀ´ËµÊÇ¸ö¾Þ´óµÄ±¦²ØµÄÒþº¬µÄAPIº¯ÊýÖÐµÄÒ»¸ö¡£±¾º¯ÊýÔÊÐíÄú¸Ä±ä´°¿ÚµÄÇøÓò¡£
' Í¨³£ËùÓÐ´°¿Ú¶¼ÊÇ¾ØÐÎµÄ¡ª¡ª´°¿ÚÒ»µ©´æÔÚ¾Íº¬ÓÐÒ»¸ö¾ØÐÎÇøÓò¡£±¾º¯ÊýÔÊÐíÄú·ÅÆú¸ÃÇøÓò¡£
' ÕâÒâÎ¶×ÅÄú¿ÉÒÔ´´½¨Ô²µÄ¡¢ÐÇÐÎµÄ´°¿Ú£¬Ò²¿ÉÒÔ½«Ëü·ÖÎªÁ½¸ö»òÐí¶à²¿·Ö¡ª¡ªÊµ¼ÊÉÏ¿ÉÒÔÊÇÈÎºÎÐÎ×´
 Private Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
' ¸Ãº¯ÊýÑ¡ÔñÒ»¸öÇøÓò×÷ÎªÖ¸¶¨Éè±¸»·¾³µÄµ±Ç°¼ôÇÐÇøÓò
 Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------------Î»Í¼º¯Êý(Bitmap)-----------------------------------©§
'
' ¸Ãº¯ÊýÓÃÀ´ÏÔÊ¾Í¸Ã÷»ò°ëÍ¸Ã÷ÏñËØµÄÎ»Í¼¡£
 Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
' ½«Ò»·ùÎ»Í¼´ÓÒ»¸öÉè±¸³¡¾°¸´ÖÆµ½ÁíÒ»¸ö¡£Ô´ºÍÄ¿±êDCÏà»¥¼ä±ØÐë¼æÈÝ
' ÔÚNT»·¾³ÏÂ£¬ÈçÔÚÒ»´ÎÊÀ½ç´«ÊäÖÐÒªÇóÔÚÔ´Éè±¸³¡¾°ÖÐ½øÐÐ¼ôÇÐ»òÐý×ª´¦Àí£¬Õâ¸öº¯ÊýµÄÖ´ÐÐ»áÊ§°Ü
' ÈçÄ¿±êºÍÔ´DCµÄÓ³Éä¹ØÏµÒªÇó¾ØÐÎÖÐÏñËØµÄ´óÐ¡±ØÐëÔÚ´«Êä¹ý³ÌÖÐ¸Ä±ä£¬
' ÄÇÃ´Õâ¸öº¯Êý»á¸ù¾ÝÐèÒª×Ô¶¯ÉìËõ¡¢Ðý×ª¡¢ÕÛµþ¡¢»òÇÐ¶Ï£¬ÒÔ±ãÍê³É×îÖÕµÄ´«Êä¹ý³Ì
' dwRop£ºÖ¸¶¨¹âÕ¤²Ù×÷´úÂë¡£ÕâÐ©´úÂë½«¶¨ÒåÔ´¾ØÐÎÇøÓòµÄÑÕÉ«Êý¾Ý£¬ÈçºÎÓëÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«Êý¾Ý×éºÏÒÔÍê³É×îºóµÄÑÕÉ«¡£
 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' ´´½¨Ò»·ùÓëÉè±¸ÓÐ¹ØÎ»Í¼
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
' ´´½¨Ò»·ùÓëÉè±¸ÓÐ¹ØÎ»Í¼£¬ËüÓëÖ¸¶¨µÄÉè±¸³¡¾°¼æÈÝ
' ÄÚ´æÉè±¸³¡¾°¼´Óë²ÊÉ«Î»Í¼¼æÈÝ£¬Ò²Óëµ¥É«Î»Í¼¼æÈÝ¡£Õâ¸öº¯ÊýµÄ×÷ÓÃÊÇ´´½¨Ò»·ùÓëµ±Ç°Ñ¡ÈëhdcÖÐµÄ³¡¾°¼æÈÝ¡£
' ¶ÔÒ»¸öÄÚ´æ³¡¾°À´Ëµ£¬Ä¬ÈÏµÄÎ»Í¼ÊÇµ¥É«µÄ¡£ÌÈÈôÄÚ´æÉè±¸³¡¾°ÓÐÒ»¸öDIBSectionÑ¡ÈëÆäÖÐ£¬
' Õâ¸öº¯Êý¾Í»á·µ»ØDIBSectionµÄÒ»¸ö¾ä±ú¡£ÈçhdcÊÇÒ»·ùÉè±¸Î»Í¼£¬
' ÄÇÃ´½á¹ûÉú³ÉµÄÎ»Í¼¾Í¿Ï¶¨¼æÈÝÓÚÉè±¸£¨Ò²¾ÍÊÇËµ£¬²ÊÉ«Éè±¸Éú³ÉµÄ¿Ï¶¨ÊÇ²ÊÉ«Î»Í¼£©
' Èç¹ûnWidthºÍnHeightÎªÁã£¬·µ»ØµÄÎ»Í¼¾ÍÊÇÒ»¸ö1¡Á1µÄµ¥É«Î»Í¼
' Ò»µ©Î»Í¼²»ÔÙÐèÒª£¬Ò»¶¨ÓÃDeleteObjectº¯ÊýÊÍ·ÅËüÕ¼ÓÃµÄÄÚ´æ¼°×ÊÔ´
 Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
' ¸Ãº¯ÊýÓÉÓëÉè±¸ÎÞ¹ØµÄÎ»Í¼£¨DIB£©´´½¨ÓëÉè±¸ÓÐ¹ØµÄÎ»Í¼£¨DDB£©£¬²¢ÇÒÓÐÑ¡ÔñµØÎªÎ»Í¼ÖÃÎ»¡£
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
' ¸Ãº¯Êý´´½¨Ó¦ÓÃ³ÌÐò¿ÉÒÔÖ±½ÓÐ´ÈëµÄ¡¢ÓëÉè±¸ÎÞ¹ØµÄÎ»Í¼£¨DIB£©¡£
' ¸Ãº¯ÊýÌá¹©Ò»¸öÖ¸Õë£¬¸ÃÖ¸ÕëÖ¸ÏòÎ»Í¼Î»Êý¾ÝÖµµÄµØ·½¡£
' ¿ÉÒÔ¸øÎÄ¼þÓ³Éä¶ÔÏóÌá¹©¾ä±ú£¬º¯ÊýÊ¹ÓÃÎÄ¼þÓ³Éä¶ÔÏóÀ´´´½¨Î»Í¼£¬»òÕßÈÃÏµÍ³ÎªÎ»Í¼·ÖÅäÄÚ´æ¡£
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
' ¸´ÖÆÎ»Í¼¡¢Í¼±ê»òÖ¸Õë£¬Í¬Ê±ÔÚ¸´ÖÆ¹ý³ÌÖÐ½øÐÐÒ»Ð©×ª»»¹¤×÷
 Private Declare Function CopyImage Lib "USER32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
' ÔØÈëÒ»¸öÎ»Í¼¡¢Í¼±ê»òÖ¸Õë
 Private Declare Function LoadImage Lib "USER32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 Private Declare Function LoadImageLong Lib "USER32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§----------------------------------Í¼±êº¯Êý(Icon)------------------------------------©§
'
' ÖÆ×÷Ö¸¶¨Í¼±ê»òÊó±êÖ¸ÕëµÄÒ»¸ö¸±±¾¡£Õâ¸ö¸±±¾´ÓÊôÓÚ·¢³öµ÷ÓÃµÄÓ¦ÓÃ³ÌÐò
 Private Declare Function CopyIcon Lib "USER32" (ByVal hIcon As Long) As Long
' ´´½¨Ò»¸öÍ¼±ê
Private Declare Function CreateIconIndirect Lib "USER32" (piconinfo As ICONINFO) As Long
' ¸Ãº¯ÊýÇå³ýÍ¼±êºÍÊÍ·ÅÈÎºÎ±»Í¼±êÕ¼ÓÃµÄ´æ´¢¿Õ¼ä¡£
 Private Declare Function DestroyIcon Lib "USER32" (ByVal hIcon As Long) As Long
' ¸Ãº¯ÊýÔÚÏÞ¶¨µÄÉè±¸ÉÏÏÂÎÄ´°¿ÚµÄ¿Í»§ÇøÓò»æÖÆÍ¼±ê
 Private Declare Function DrawIcon Lib "USER32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' ¸Ãº¯ÊýÔÚÏÞ¶¨µÄÉè±¸ÉÏÏÂÎÄ´°¿ÚµÄ¿Í»§ÇøÓò»æÖÆÍ¼±ê£¬Ö´ÐÐÏÞ¶¨µÄ¹âÕ¤²Ù×÷£¬²¢°´ÌØ¶¨ÒªÇóÉì³¤»òÑ¹ËõÍ¼±ê»ò¹â±ê¡£
 Private Declare Function DrawIconEx Lib "USER32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
' È¡µÃÓëÍ¼±êÓÐ¹ØµÄÐÅÏ¢
Private Declare Function GetIconInfo Lib "USER32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------------¹â±êº¯Êý(Cursor)-----------------------------------©§
'
 Private Declare Function CopyCursor Lib "USER32" (ByVal hcur As Long) As Long
' ´ÓÖ¸¶¨µÄÄ£¿é»òÓ¦ÓÃ³ÌÐòÊµÀýÖÐÔØÈëÒ»¸öÊó±êÖ¸Õë¡£LoadCursorBynumÊÇLoadCursorº¯ÊýµÄÀàÐÍ°²È«ÉùÃ÷
 Private Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
' ¸Ãº¯ÊýÏú»ÙÒ»¸ö¹â±ê²¢ÊÍ·ÅËüÕ¼ÓÃµÄÈÎºÎÄÚ´æ£¬²»ÒªÊ¹ÓÃ¸Ãº¯ÊýÈ¥Ïû»ÙÒ»¸ö¹²Ïí¹â±ê¡£
 Private Declare Function DestroyCursor Lib "USER32" (ByVal hCursor As Long) As Long
' »ñÈ¡Êó±êÖ¸ÕëµÄµ±Ç°Î»ÖÃ
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
' ¸Ãº¯Êý°Ñ¹â±êÒÆµ½ÆÁÄ»µÄÖ¸¶¨Î»ÖÃ¡£Èç¹ûÐÂÎ»ÖÃ²»ÔÚÓÉ ClipCursorº¯ÊýÉèÖÃµÄÆÁÄ»¾ØÐÎÇøÓòÖ®ÄÚ£¬
' ÔòÏµÍ³×Ô¶¯µ÷Õû×ø±ê£¬Ê¹µÃ¹â±êÔÚ¾ØÐÎÖ®ÄÚ¡£
 Private Declare Function SetCursorPos Lib "USER32" (ByVal x As Long, ByVal y As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------±ÊË¢º¯Êý(Pen and Brush)---------------------------------©§
'
' ÓÃÖ¸¶¨µÄÑùÊ½¡¢¿í¶ÈºÍÑÕÉ«´´½¨Ò»¸ö»­±Ê£¬ÓÃDeleteObjectº¯Êý½«ÆäÉ¾³ý
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' ¸ù¾ÝÖ¸¶¨µÄLOGPEN½á¹¹´´½¨Ò»¸ö»­±Ê
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
' ´´½¨Ò»¸öÀ©Õ¹»­±Ê£¨×°ÊÎ»ò¼¸ºÎ£©
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
' ÔÚÒ»¸öLOGBRUSHÊý¾Ý½á¹¹µÄ»ù´¡ÉÏ´´½¨Ò»¸öË¢×Ó
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
' ¸Ãº¯Êý¿ÉÒÔ´´½¨Ò»¸ö¾ßÓÐÖ¸¶¨ÒõÓ°Ä£Ê½ºÍÑÕÉ«µÄÂß¼­Ë¢×Ó¡£
 Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
' ¸Ãº¯Êý¿ÉÒÔ´´½¨¾ßÓÐÖ¸¶¨Î»Í¼Ä£Ê½µÄÂß¼­Ë¢×Ó£¬¸ÃÎ»Í¼²»ÄÜÊÇDIBÀàÐÍµÄÎ»Í¼£¬DIBÎ»Í¼ÊÇÓÉCreateDIBSectionº¯Êý´´½¨µÄ¡£
 Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
' ÓÃ´¿É«´´½¨Ò»¸öË¢×Ó£¬Ò»µ©Ë¢×Ó²»ÔÙÐèÒª£¬¾ÍÓÃDeleteObjectº¯Êý½«ÆäÉ¾³ý
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' ÎªÈÎºÎÒ»ÖÖ±ê×¼ÏµÍ³ÑÕÉ«È¡µÃÒ»¸öË¢×Ó£¬²»ÒªÓÃDeleteObjectº¯ÊýÉ¾³ýÕâÐ©Ë¢×Ó¡£
' ËüÃÇÊÇÓÉÏµÍ³ÓµÓÐµÄ¹ÌÓÐ¶ÔÏó¡£²»Òª½«ÕâÐ©Ë¢×ÓÖ¸¶¨³ÉÒ»ÖÖ´°¿ÚÀàµÄÄ¬ÈÏË¢×Ó
 Private Declare Function GetSysColorBrush Lib "USER32" (ByVal nIndex As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------×ÖÌåºÍÕýÎÄº¯Êý(Font and Text)-----------------------------©§
'
' ÓÃÖ¸¶¨µÄÊôÐÔ´´½¨Ò»ÖÖÂß¼­×ÖÌå£¬VBµÄ×ÖÌåÊôÐÔÔÚÑ¡Ôñ×ÖÌåµÄÊ±ºòÏÔµÃ¸üÓÐÐ§
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
' ½«ÎÄ±¾Ãè»æµ½Ö¸¶¨µÄ¾ØÐÎÖÐ£¬wFormat±êÖ¾³£Êý²Î¿´KhanDrawTextStyles
Private Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "USER32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
' ¸Ãº¯ÊýÈ¡µÃÖ¸¶¨Éè±¸»·¾³µÄµ±Ç°ÕýÎÄÑÕÉ«¡£
 Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
' ÉèÖÃµ±Ç°ÎÄ±¾ÑÕÉ«¡£ÕâÖÖÑÕÉ«Ò²³ÆÎª¡°Ç°¾°É«¡±£¬Èç¸Ä±äÁËÕâ¸öÉèÖÃ£¬×¢Òâ»Ö¸´VB´°Ìå»ò¿Ø¼þÔ­Ê¼µÄÎÄ±¾ÑÕÉ«
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§------------------------------------»æÍ¼º¯Êý----------------------------------------©§
'
' ¸Ãº¯Êý»­Ò»¶ÎÔ²»¡£¬Ô²»¡ÊÇÓÉÒ»¸öÍÖÔ²ºÍÒ»ÌõÏß¶Î£¨³ÆÖ®Îª¸îÏß£©Ïà½»ÏÞ¶¨µÄ±ÕºÏÇøÓò¡£
' ´Ë»¡ÓÉµ±Ç°µÄ»­±Ê»­ÂÖÀª£¬ÓÉµ±Ç°µÄ»­Ë¢Ìî³ä¡£
 Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ÓÃÖ¸¶¨µÄÑùÊ½Ãè»æÒ»¸ö¾ØÐÎµÄ±ß¿ò¡£ÀûÓÃÕâ¸öº¯Êý£¬ÎÒÃÇÃ»ÓÐ±ØÒªÔÙÊ¹ÓÃÐí¶à3D±ß¿òºÍÃæ°å¡£
' ËùÒÔ¾Í×ÊÔ´ºÍÄÚ´æµÄÕ¼ÓÃÂÊÀ´Ëµ£¬Õâ¸öº¯ÊýµÄÐ§ÂÊÒª¸ßµÃ¶à¡£Ëü¿ÉÔÚÒ»¶¨³Ì¶ÈÉÏÌáÉýÐÔÄÜ
' hdc      ÒªÔÚÆäÖÐ»æÍ¼µÄÉè±¸³¡¾°
' qrc      ÒªÎªÆäÃè»æ±ß¿òµÄ¾ØÐÎ
' edge     ´øÓÐÇ°×ºBDR_µÄÁ½¸ö³£ÊýµÄ×éºÏ¡£Ò»¸öÖ¸¶¨ÄÚ²¿±ß¿òÊÇÉÏÍ¹»¹ÊÇÏÂ°¼£»ÁíÒ»¸öÔòÖ¸¶¨Íâ²¿±ß¿ò¡£ÓÐÊ±ÄÜ»»ÓÃ´øEDGE_Ç°×ºµÄ³£Êý¡£
' grfFlags ´øÓÐBF_Ç°×ºµÄ³£ÊýµÄ×éºÏ
Private Declare Function DrawEdge Lib "USER32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
' »­Ò»¸ö½¹µã¾ØÐÎ¡£Õâ¸ö¾ØÐÎÊÇÔÚ±êÖ¾½¹µãµÄÑùÊ½ÖÐÍ¨¹ýÒì»òÔËËãÍê³ÉµÄ£¨½¹µãÍ¨³£ÓÃÒ»¸öµãÏß±íÊ¾£©
' ÈçÓÃÍ¬ÑùµÄ²ÎÊýÔÙ´Îµ÷ÓÃÕâ¸öº¯Êý£¬¾Í±íÊ¾É¾³ý½¹µã¾ØÐÎ
Private Declare Function DrawFocusRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT) As Long
' Õâ¸öº¯ÊýÓÃÓÚÃè»æÒ»¸ö±ê×¼¿Ø¼þ
Private Declare Function DrawFrameControl Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
' Õâ¸öº¯Êý¿ÉÎªÒ»·ùÍ¼Ïó»ò»æÍ¼²Ù×÷Ó¦ÓÃ¸÷Ê½¸÷ÑùµÄÐ§¹û
 Private Declare Function DrawState Lib "USER32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal fuFlags As Long) As Long
' ¸Ãº¯ÊýÓÃÓÚ»­Ò»¸öÍÖÔ²£¬ÍÖÔ²µÄÖÐÐÄÊÇÏÞ¶¨¾ØÐÎµÄÖÐÐÄ£¬Ê¹ÓÃµ±Ç°»­±Ê»­ÍÖÔ²£¬ÓÃµ±Ç°µÄ»­Ë¢Ìî³äÍÖÔ²¡£
 Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ÓÃÖ¸¶¨µÄË¢×ÓÌî³äÒ»¸ö¾ØÐÎ£¬¾ØÐÎµÄÓÒ±ßºÍµ×±ß²»»áÃè»æ
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ÓÃÖ¸¶¨µÄË¢×ÓÎ§ÈÆÒ»¸ö¾ØÐÎ»­Ò»¸ö±ß¿ò£¨×é³ÉÒ»¸öÖ¡£©£¬±ß¿òµÄ¿í¶ÈÊÇÒ»¸öÂß¼­µ¥Î»
 Private Declare Function FrameRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' È¡µÃÖ¸¶¨Éè±¸³¡¾°µ±Ç°µÄ±³¾°ÑÕÉ«
 Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬È¡µÃµ±Ç°µÄ±³¾°Ìî³äÄ£Ê½
 Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
' ÎªÖ¸¶¨µÄÉè±¸³¡¾°ÉèÖÃ±³¾°ÑÕÉ«¡£±³¾°ÑÕÉ«ÓÃÓÚÌî³äÒõÓ°Ë¢×Ó¡¢ÐéÏß»­±ÊÒÔ¼°×Ö·û£¨Èç±³¾°Ä£Ê½ÎªOPAQUE£©ÖÐµÄ¿ÕÏ¶¡£
' Ò²ÔÚÎ»Í¼ÑÕÉ«×ª»»ÆÚ¼äÊ¹ÓÃ¡£±³¾°Êµ¼ÊÊÇÉè±¸ÄÜ¹»ÏÔÊ¾µÄ×î½Ó½üÓÚ crColor µÄÑÕÉ«
 Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
' Ö¸¶¨ÒõÓ°Ë¢×Ó¡¢ÐéÏß»­±ÊÒÔ¼°×Ö·ûÖÐµÄ¿ÕÏ¶µÄÌî³ä·½Ê½£¬±³¾°Ä£Ê½²»»áÓ°ÏìÓÃÀ©Õ¹»­±ÊÃè»æµÄÏßÌõ
 Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
' ÔÚÖ¸¶¨µÄÉè±¸³¡¾°ÖÐÈ¡µÃÒ»¸öÏñËØµÄRGBÖµ
 Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
' ÔÚÖ¸¶¨µÄÉè±¸³¡¾°ÖÐÉèÖÃÒ»¸öÏñËØµÄRGBÖµ
 Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
' ½«À´×ÔÒ»·ùÎ»Í¼µÄ¶þ½øÖÆÎ»¸´ÖÆµ½Ò»·ùÓëÉè±¸ÎÞ¹ØµÄÎ»Í¼Àï
' Private Declare Function GetDIBits Lib "gdi32" ( ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
 Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ½«À´×ÔÓëÉè±¸ÎÞ¹ØÎ»Í¼µÄ¶þ½øÖÆÎ»¸´ÖÆµ½Ò»·ùÓëÉè±¸ÓÐ¹ØµÄÎ»Í¼Àï
 Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬»ñµÃ¶à±ßÐÎÌî³äÄ£Ê½¡£
 Private Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
' ÉèÖÃ¶à±ßÐÎµÄÌî³äÄ£Ê½
 Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬È¡µÃµ±Ç°µÄ»æÍ¼Ä£Ê½¡£ÕâÑù¿É¶¨Òå»æÍ¼²Ù×÷ÈçºÎÓëÕýÔÚÏÔÊ¾µÄÍ¼ÏóºÏ²¢ÆðÀ´
' Õâ¸öº¯ÊýÖ»¶Ô¹âÕ¤Éè±¸ÓÐÐ§
 Private Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
' ÉèÖÃÖ¸¶¨Éè±¸³¡¾°µÄ»æÍ¼Ä£Ê½¡£

' ÓÃµ±Ç°»­±Ê»­Ò»ÌõÏß£¬´Óµ±Ç°Î»ÖÃÁ¬µ½Ò»¸öÖ¸¶¨µÄµã¡£Õâ¸öº¯Êýµ÷ÓÃÍê±Ï£¬µ±Ç°Î»ÖÃ±ä³Éx,yµã
 Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
' ÎªÖ¸¶¨µÄÉè±¸³¡¾°Ö¸¶¨Ò»¸öÐÂµÄµ±Ç°»­±ÊÎ»ÖÃ¡£Ç°Ò»¸öÎ»ÖÃ±£´æÔÚlpPointÖÐ
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
' ¸Ãº¯Êý»­Ò»¸öÓÉÍÖÔ²ºÍÁ½Ìõ°ë¾¶Ïà½»±ÕºÏ¶ø³ÉµÄ±ý×´Ð¨ÐÎÍ¼£¬´Ë±ýÍ¼ÓÉµ±Ç°»­±Ê»­ÂÖÀª£¬ÓÉµ±Ç°»­Ë¢Ìî³ä¡£
 Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ¸Ãº¯Êý»­Ò»¸öÓÉÖ±ÏßÏàÎÅµÄÁ½¸öÒÔÉÏ¶¥µã×é³ÉµÄ¶à±ßÐÎ£¬ÓÃµ±Ç°»­±Ê»­¶à±ßÐÎÂÖÀª£¬
' ÓÃµ±Ç°»­Ë¢ºÍ¶à±ßÐÎÌî³äÄ£Ê½Ìî³ä¶à±ßÐÎ¡£
 Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
' ÓÃµ±Ç°»­±ÊÃè»æÒ»ÏµÁÐÏß¶Î¡£Ê¹ÓÃPolylineToº¯ÊýÊ±£¬µ±Ç°Î»ÖÃ»áÉèÎª×îºóÒ»ÌõÏß¶ÎµÄÖÕµã¡£
' Ëü²»»áÓÉPolylineº¯Êý¸Ä¶¯
 Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
 Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
 Private Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
' ¸Ãº¯Êý»­Ò»¸ö¾ØÐÎ£¬ÓÃµ±Ç°µÄ»­±Ê»­¾ØÐÎÂÖÀª£¬ÓÃµ±Ç°»­Ë¢½øÐÐÌî³ä¡£
 Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' º¯Êý»­Ò»¸ö´øÔ²½ÇµÄ¾ØÐÎ£¬´Ë¾ØÐÎÓÉµ±Ç°»­±Ê»­ÂÖÀÈ£¬ÓÉµ±Ç°»­Ë¢Ìî³ä¡£
 Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Õâ¸öº¯ÊýÓÃÓÚÔö´ó»ò¼õÐ¡Ò»¸ö¾ØÐÎµÄ´óÐ¡¡£
' x¼ÓÔÚÓÒ²àÇøÓò£¬²¢´Ó×ó²àÇøÓò¼õÈ¥£»ÈçxÎªÕý£¬ÔòÄÜÔö´ó¾ØÐÎµÄ¿í¶È£»ÈçxÎª¸º£¬ÔòÄÜ¼õÐ¡Ëü¡£
' y¶Ô¶¥²¿Óëµ×²¿ÇøÓò²úÉúµÄÓ°ÏìÊÇÊÇÀàËÆµÄ
 Private Declare Function InflateRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' ¸Ãº¯ÊýÍ¨¹ýÓ¦ÓÃÒ»¸öÖ¸¶¨µÄÆ«ÒÆ£¬´Ó¶øÈÃ¾ØÐÎÒÆ¶¯ÆðÀ´¡£
' x»áÌí¼Óµ½ÓÒ²àºÍ×ó²àÇøÓò¡£yÌí¼Óµ½¶¥²¿ºÍµ×²¿ÇøÓò¡£
' Æ«ÒÆ·½ÏòÔòÈ¡¾öÓÚ²ÎÊýÊÇÕýÊý»¹ÊÇ¸ºÊý£¬ÒÔ¼°²ÉÓÃµÄÊÇÊ²Ã´×ø±êÏµÍ³
 Private Declare Function OffsetRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' ·µ»ØÓëwindows»·¾³ÓÐ¹ØµÄÐÅÏ¢£¬nIndexÖµ²Î¿´±¾Ä£¿éµÄ³£Á¿ÉùÃ÷
 Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
' »ñµÃÕû¸ö´°¿ÚµÄ·¶Î§¾ØÐÎ£¬´°¿ÚµÄ±ß¿ò¡¢±êÌâÀ¸¡¢¹ö¶¯Ìõ¼°²Ëµ¥µÈ¶¼ÔÚÕâ¸ö¾ØÐÎÄÚ
 Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
' ·µ»ØÖ¸¶¨´°¿Ú¿Í»§Çø¾ØÐÎµÄ´óÐ¡
 Private Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
' Õâ¸öº¯ÊýÆÁ±ÎÒ»¸ö´°¿Ú¿Í»§ÇøµÄÈ«²¿»ò²¿·ÖÇøÓò¡£Õâ»áµ¼ÖÂ´°¿ÚÔÚÊÂ¼þÆÚ¼ä²¿·ÖÖØ»­
 Private Declare Function InvalidateRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
' ÅÐ¶ÏÖ¸¶¨windowsÏÔÊ¾¶ÔÏóµÄÑÕÉ«£¬ÑÕÉ«¶ÔÏó¿´±¾Ä£¿éÉùÃ÷
 Private Declare Function GetSysColor Lib "USER32" (ByVal nIndex As Long) As Long


 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
 Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

'Initializes the entire common control dynamic-link library.
'Exported by all versions of Comctl32.dll.
 Private Declare Sub InitCommonControls Lib "Comctl32" ()
'Initializes specific common controls classes from the common
'control dynamic-link library.
'Returns TRUE (non-zero) if successful, or FALSE otherwise.
'Began being exported with Comctl32.dll version 4.7 (IE3.0 & later).
 Private Declare Function InitCommonControlsEx Lib "Comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
 Private Declare Function ImageList_GetBkColor Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_ReplaceIcon Lib "Comctl32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
 Private Declare Function ImageList_Convert Lib "Comctl32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal Flags As Long) As Long
 Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
 Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
 Private Declare Function ImageList_Replace Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
 Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, hbmMask As Long) As Long
 Private Declare Function ImageList_Remove Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long
 Private Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Long
 Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
 Private Declare Function ImageList_GetIcon Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
 Private Declare Function ImageList_SetImageCount Lib "Comctl32" (ByVal hImageList As Long, uNewCount As Long)
 Private Declare Function ImageList_GetImageCount Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hImageList As Long, Cx As Long, Cy As Long) As Long
 Private Declare Function ImageList_SetIconSize Lib "Comctl32" (ByVal hImageList As Long, Cx As Long, Cy As Long) As Long
 Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
' Draw an item in an ImageList with more control over positioning and colour:
 Private Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
 Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, prcImage As RECT) As Long
 Private Declare Function ImageList_LoadImage Lib "Comctl32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal Cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)
 Private Declare Function ImageList_SetBkColor Lib "Comctl32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
 Private Declare Function ImageList_Copy Lib "Comctl32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long

' ======================================================================================
' Enums
' ======================================================================================

' ±ß¿òÑùÊ½
Public Enum GPTAB_BORDERSTYLE_METHOD
   GpTabBorderStyleNone = 0               ' Ã»ÓÐ±ß¿ò
   GpTabBorderStyle3D = 1                 ' 3D
   GpTabBorderStyle3DThin = 2             ' 3DThin
End Enum

' ÑùÊ½
Public Enum GPTAB_STYLE_METHOD
    GpTabStyleStandard = 0                '/* Win32 ·ç¸ñ
    GpTabStyleWinXP = 1                   '/* XP ·ç¸ñ
End Enum

' Ñ¡Ïî¿¨²¼¾Ö
Public Enum GPTAB_PLACEMENT_METHOD
    GpTabPlacementTopleft = 0
    GpTabPlacementTopRight = 1
    GpTabPlacementBottomLeft = 2
    GpTabPlacementBottomRight = 3
    GpTabPlacementLeftTop = 4
    GpTabPlacementLeftBottom = 5
    GpTabPlacementRightTop = 6
    GpTabPlacementRightBottom = 7
End Enum

' Ñ¡Ïî¿¨ÑùÊ½
Public Enum GPTAB_TABSTYLE_METHOD
    GpTabRectangle = 0                 '
    GpTabRoundRect = 1                 '
    GpTabTrapezoid = 2                 '
End Enum

' Ñ¡Ïî¿¨¿í¶È
Public Enum GPTAB_TABWIDTHSTYLE_METHOD
    GpTabJustified = 0                 '
    GpTabnonJustified = 1              '
    GpTabFixed = 2                     '
End Enum

Public Enum GPTAB_XPCOLORSCHEME_METHOD
    GpTabUseWindows = 0
    GpTabCustom = 1
End Enum

' ======================================================================================
' Types
' ======================================================================================

' ±ãÓÚ±êÖ¾À©Õ¹
Private Type TabState
    Index As Long       ' ÓÃÓÚ´æ·ÅTabµÄIndex
End Type

Private Type TabListType
    list() As TabState  ' È¡µÃÃ¿ÐÐÖÐTabµÄIndex
    Count As Long       ' Ò»ÐÐÖÐTabµÄ¸öÊý
End Type

' ======================================================================================
' Private variables:
' ======================================================================================

' Icons:
Private m_hIml                    As Long
Private m_lIconSizeX              As Long
Private m_lIconSizeY              As Long
Private m_lngFontHeight           As Long
Private m_lngDefaultTabHeight     As Long

Private m_lngXPFaceColor         As Long
Private m_oleBackColor As OLE_COLOR      ' ¿Ø¼þµÄ±³¾°ÑÕÉ«
Private m_oleTabColor As OLE_COLOR       ' Ñ¡Ïî¿¨²»¿ÉÑ¡Ê±ÑÕÉ«
Private m_oleTabColorActive As OLE_COLOR ' Ñ¡Ïî¿¨¼¤»îÊ±ÑÕÉ«
Private m_oleTabColorHover As OLE_COLOR  ' Ñ¡Ïî¿¨ÈÈ¸ú×ÙÊ±ÑÕÉ«
Private m_oleTabBorderColor As OLE_COLOR ' XP·ç¸ñ,GpTabBorderStyleNone¿Ø¼þ±ß¿òµÄÑÕÉ«
Private m_blnAutoBackColor As Boolean ' ÅÐ¶Ï¿Ø¼þµÄ±³¾°ÑÕÉ«ÊÇ·ñËæ¸¸´°ÌåµÄ±³¾°ÑÕÉ«¸Ä±ä¶ø¸Ä±ä
Private m_blnUserMode As Boolean ' ¿Ø¼þÔËÐÐÔÚÉè¼Æ½×¶Î?ÔËÐÐ½×¶Î?
Private m_blnEnabled As Boolean ' Enable
Private m_blnHotTracking As Boolean ' ÈÈ¸ú×Ù
Private m_blnMultiRow As Boolean

Private m_lngTabFixedHeight As Long     ' ¶¨ÖÆTabµÄ¸ß¶È
Private m_lngTabFixedWidth As Long      ' ¶¨ÖÆTabµÄ¿í¶È

Private m_udtMainRect         As RECT     ' Ö÷ÇøÓò
Private m_lngCurrentList        As Long  ' µ±Ç°TabÁÐ±íµÄË÷Òý
Private m_lngListCount         As Long  ' TabÓÐ¼¸ÐÐ
Private m_aryTabList()       As TabListType

Private m_udtBorderStyle As GPTAB_BORDERSTYLE_METHOD
Private m_udtXPColorScheme As GPTAB_XPCOLORSCHEME_METHOD
Private m_udtPlacement As GPTAB_PLACEMENT_METHOD
Private m_udtStyle As GPTAB_STYLE_METHOD
Private m_udtTabStyle As GPTAB_TABSTYLE_METHOD
Private m_udtTabWidthStyle As GPTAB_TABWIDTHSTYLE_METHOD
Private m_udtDrawTextParams           As DRAWTEXTPARAMS
Private m_clsSelectTab As cTabItem
Private m_clsHoverTab As cTabItem
Private WithEvents m_clsTabs As cTabItems
Attribute m_clsTabs.VB_VarHelpID = -1

' ======================================================================================
' Events
' ======================================================================================
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event TabClick()

Public Property Get AutoBackColor() As Boolean
    AutoBackColor = m_blnAutoBackColor
End Property

Public Property Let AutoBackColor(ByVal New_AutoBackColor As Boolean)
    m_blnAutoBackColor = New_AutoBackColor
    PropertyChanged "AutoBackColor"
    Call pvDraw
End Property

'/* ±³¾°ÑÕÉ«£¨Ä¬ÈÏÎª-1£¬Ëæ¸¸´°ÌåµÄÑÕÉ«¶ø¸Ä±ä£© */
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oleBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_oleBackColor = VerifyColor(New_BackColor)
    If Not m_blnAutoBackColor Then UserControl.BackColor = m_oleBackColor
    PropertyChanged "BackColor"
    Call pvDraw
End Property

Public Property Get BorderStyle() As GPTAB_BORDERSTYLE_METHOD
    BorderStyle = m_udtBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As GPTAB_BORDERSTYLE_METHOD)
    m_udtBorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get Enable() As Boolean
    Enable = m_blnEnabled
End Property

Public Property Let Enable(ByVal New_Enable As Boolean)
    m_blnEnabled = New_Enable
    PropertyChanged "Enable"
    Call pvDraw
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
    Dim udtFont As LOGFONT
     
    Set UserControl.Font = New_Font
    Set lblFont.Font = New_Font
    '/* È¡µÃµ±Ç°×ÖÌåÏÂÎÄ×ÖµÄ¸ß¶È */
    m_lngFontHeight = lblFont.Height + 1
    If m_lngFontHeight > m_lIconSizeY Then
       m_lngDefaultTabHeight = m_lngFontHeight + InflateFontHeight
    Else
       m_lngDefaultTabHeight = m_lIconSizeY + InflateIconHeight
    End If
'    If m_lngFontHeight > m_lIconSizeY Then
'       m_cListItems.DefaultListitemHeight = m_lngFontHeight
'    Else
'       m_cListItems.DefaultListitemHeight = m_lIconSizeY
'    End If
    
'    If m_lngFontHeight > m_lngColumnIconHeight Then
'       If m_lngFontHeight > m_cHeader.Height Then
'          m_cHeader.Height = m_lngFontHeight
'       End If
'    Else
'       If m_lngColumnIconHeight > m_cHeader.Height Then
'          m_cHeader.Height = m_lngColumnIconHeight
'       End If
'    End If
    PropertyChanged "Font"
    Call pvDraw
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   UserControl.ForeColor = VerifyColor(New_ForeColor)
   PropertyChanged "ForeColor"
    Call pvDraw
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Function HitTest(ByVal x As Single, ByVal y As Single) As cTabItem
    Dim lngI As Long
    Dim lngY As Long
    Dim tR As RECT
    Const ProcName = "HitTest"
    
    On Error GoTo ErrorHandle
    
    ' ÅÐ¶ÏÊó±êÊÇ·ñ³ö½ç
    If x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
       Set HitTest = Nothing
       Exit Function
    End If
    ' ÅÐ¶ÏÊó±êÊÇ·ñÔÚÖ÷ÇøÓòÄÚ
    If y >= m_udtMainRect.Top Then
       Set HitTest = Nothing
       Exit Function
    End If
    If m_clsTabs Is Nothing Or m_clsTabs.Count <= 0 Then
       Set HitTest = Nothing
       Exit Function
    End If
    Select Case m_udtTabStyle
           Case GpTabRectangle, GpTabRoundRect
             For lngI = 1 To m_clsTabs.Count
                 With m_clsTabs.Item(lngI)
                      tR.Top = .Top
                      tR.Left = .Left
                      tR.Right = .Left + .Width
                      tR.Bottom = .Top + .Height
                 End With
                 If PtInRect(tR, x, y) <> 0 Then
                    Set HitTest = m_clsTabs.Item(lngI)
                    Exit Function
                 End If
             Next lngI
           Case GpTabTrapezoid
             For lngI = 1 To m_clsTabs.Count
                 With m_clsTabs.Item(lngI)
                      tR.Top = .Top
                      tR.Left = .Left + .Height
                      tR.Right = .Left + .Width
                      tR.Bottom = .Top + .Height
                 End With
                 If PtInRect(tR, x, y) <> 0 Then
                    Set HitTest = m_clsTabs.Item(lngI)
                    Exit Function
                 End If
                 
'                 With m_clsTabs.Item(lngI)
'                      For lngY = 1 To .Height
'                          If Y = lngY And X >= .Left - .Height - 1 Then
'                             Set HitTest = m_clsTabs.Item(lngI)
'                             Exit Function
'                          End If
'                      Next lngY
'                 End With
             Next lngI
    End Select
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Function

Public Property Get HotTracking() As Boolean
    HotTracking = m_blnHotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    m_blnHotTracking = New_HotTracking
    PropertyChanged "HotTracking"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Let ImageList(New_ImageList As Variant)
    Dim hIml As Long
    
    ' Set the ImageList handle property either from a VB
    ' image list or directly:
    If VarType(New_ImageList) = vbObject Then
       ' Assume VB ImageList control.  Note that unless
       ' some call has been made to an object within a
       ' VB ImageList the image list itself is not
       ' created.  Therefore hImageList returns error. So
       ' ensure that the ImageList has been initialised by
       ' drawing into nowhere:
       On Error Resume Next
       ' Get the image list initialised..
       New_ImageList.ListImages(1).Draw 0, 0, 0, 1
       hIml = New_ImageList.hImageList
       If (Err.Number <> 0) Then
           hIml = 0
       End If
       On Error GoTo 0
    ElseIf VarType(New_ImageList) = vbLong Then
       ' Assume ImageList handle:
       hIml = New_ImageList
    Else
       Err.Raise vbObjectError + 1049, "GpTabs." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
    End If
    
    ' If we have a valid image list, then associate it with the control:
    If (hIml <> 0) Then
       m_hIml = hIml
       Call ImageList_GetIconSize(m_hIml, m_lIconSizeX, m_lIconSizeY)
       m_lIconSizeY = m_lIconSizeY + 2
       If m_lngFontHeight > m_lIconSizeY Then
          m_lngDefaultTabHeight = m_lngFontHeight + InflateFontHeight
       Else
          m_lngDefaultTabHeight = m_lIconSizeY + InflateIconHeight
       End If
    End If
End Property

' Êó±êIcon
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

' Êó±êÑùÊ½
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

' ¶àÁÐÏÔÊ¾
Public Property Get MultiRow() As Boolean
    MultiRow = m_blnMultiRow
End Property

Public Property Let MultiRow(ByVal New_MultiRow As Boolean)
    m_blnMultiRow = New_MultiRow
    PropertyChanged "MultiRow"
    Call pvDraw
End Property

' Ñ¡Ïî¿¨²¼¾Ö
Public Property Get Placement() As GPTAB_PLACEMENT_METHOD
    Placement = m_udtPlacement
End Property

Public Property Let Placement(ByVal New_Placement As GPTAB_PLACEMENT_METHOD)
    m_udtPlacement = New_Placement
    PropertyChanged "Placement"
    Call pvDraw
End Property

' ¼ÆËã¹¹³ÉÖ÷ÇøÓò¸÷¸öµãµÄ×ø±ê
Private Sub pvCalculateRect(ByRef DstPoint() As POINTAPI, _
                            ByVal Count As Long, _
                            ByVal Top As Long, _
                            ByVal Left As Long, _
                            ByVal Right As Long, _
                            ByVal Bottom As Long, _
                            ByVal LeftTop As Boolean, _
                            ByVal LeftBottom As Boolean, _
                            ByVal RightBottom As Boolean, _
                            ByVal RightTop As Boolean)
    Dim lngindex            As Long
    
    ReDim DstPoint(Count - 1)
    lngindex = 0
    If LeftTop Then
       With DstPoint(lngindex)
            .x = Left + RoundRectSize
            .y = Top
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .x = Left
            .y = Top + RoundRectSize
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .x = Left
            .y = Top
       End With
       lngindex = lngindex + 1
    End If
    
    If LeftBottom Then
       With DstPoint(lngindex)
            .x = Left
            .y = Bottom - RoundRectSize
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .x = Left + RoundRectSize
            .y = Bottom
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .x = Left
            .y = Bottom
       End With
       lngindex = lngindex + 1
    End If
    If RightBottom Then
       With DstPoint(lngindex)
            .x = Right - RoundRectSize
            .y = Bottom
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .x = Right
            .y = Bottom - RoundRectSize
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .x = Right
            .y = Bottom
       End With
       lngindex = lngindex + 1
    End If
    If RightTop Then
       With DstPoint(lngindex)
            .x = Right
            .y = Top + RoundRectSize
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .x = Right - RoundRectSize
            .y = Top
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .x = Right
            .y = Top
       End With
       lngindex = lngindex + 1
    End If
    
    If LeftTop Then
       With DstPoint(lngindex)
            .x = Left + RoundRectSize
            .y = Top
       End With
    Else
       With DstPoint(lngindex)
            .x = Left
            .y = Top
       End With
    End If
End Sub

Private Sub pvCalculateRoundPoint(ByRef DstPoint() As POINTAPI, _
                                  ByVal Count As Long, _
                                  ByVal Top As Long, _
                                  ByVal Left As Long, _
                                  ByVal Right As Long, _
                                  ByVal Bottom As Long)
    Dim lngindex As Long
    
    lngindex = 0
    ReDim DstPoint(Count - 1)
    With DstPoint(lngindex)
         .x = Left
         .y = Bottom
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .x = Left
         .y = Top + RoundRectSize
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .x = Left + RoundRectSize
         .y = Top
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .x = Right - RoundRectSize
         .y = Top
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .x = Right
         .y = Top + RoundRectSize
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .x = Right
         .y = Bottom
    End With
End Sub

Private Sub pvCalculateSize()
    Dim lngI                 As Long  ' Ñ­»·¼ÇÊý
    Dim lngY                 As Long
    Dim lngTabCount          As Long  ' TabµÄ×Ü¸öÊý
    Dim lngAllWidth          As Long  ' ËùÓÐµÄTabµÄ¿í¶È
    Dim lngListIndex         As Long  ' Ã»ÐÐTabµÄË÷Òý
    Dim lngListTabIndex      As Long  ' Ò»ÐÐTabµÄË÷Òý
    Dim lngListWidth         As Long  ' ÀÛ¼ÓÒ»ÐÐTabµÄ¿í¶È
    Dim lngManualWidth       As Long  ' ¶¨ÖÆTabµÄ¿í¶È
    Dim lngManualHeight      As Long  ' ¶¨ÖÆTabµÄ¸ß¶È
    Dim lngWidth             As Long  ' ¿Ø¼þµÄ¿í¶È
    Dim lngHeight            As Long  ' ¿Ø¼þµÄ¸ß¶È
    Dim lngDiscrepancy       As Long
    Dim tR                   As RECT
    
    If m_udtBorderStyle = GpTabBorderStyleNone Then
       lngDiscrepancy = 0
    Else
       lngDiscrepancy = DiscrepancyHeight
    End If
    Const ProcName = "pvCalculateSize"
    
    On Error GoTo ErrorHandle
    
    lngWidth = UserControl.ScaleWidth - 1
    lngHeight = UserControl.ScaleHeight - 1
    With m_udtMainRect
         .Top = 0
         .Left = 0
         .Right = lngWidth
         .Bottom = lngHeight
    End With
    If m_clsTabs Is Nothing Then
       m_udtMainRect.Top = m_lngDefaultTabHeight
       Exit Sub
    End If
    
    m_lngListCount = 0
    Erase m_aryTabList
    lngManualWidth = m_lngTabFixedWidth \ Screen.TwipsPerPixelX
    lngManualHeight = m_lngTabFixedHeight \ Screen.TwipsPerPixelY
    With m_clsTabs
         lngTabCount = m_clsTabs.Count
         ' ¼ÆËãÃ¿¸öTabµÄÊµ¼Ê×îÐ¡¿í¶È
         For lngI = 1 To lngTabCount
             Call DrawTextEx(UserControl.hdc, .Item(lngI).Caption & vbNullChar, -1, tR, _
                             DT_CALCRECT Or DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, _
                             m_udtDrawTextParams)
             If m_udtTabStyle = GpTabTrapezoid Then
                .Item(lngI).DefaultWidth = tR.Right - tR.Left + InflateFontWidth + m_lngDefaultTabHeight
             Else
                .Item(lngI).DefaultWidth = tR.Right - tR.Left + InflateFontWidth
             End If
             lngAllWidth = lngAllWidth + .Item(lngI).DefaultWidth
         Next lngI
         Select Case m_udtPlacement
                Case GpTabPlacementTopleft
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = m_lngDefaultTabHeight
                               .Item(lngI).Width = .Item(lngI).DefaultWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ³õÊÔÊý×é
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(0)
                              ' ·ÖÐÐ,²¢ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              m_lngListCount = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                         Case GpTabnonJustified
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = m_lngDefaultTabHeight
                               .Item(lngI).Width = .Item(lngI).DefaultWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ³õÊÔÊý×é
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(0)
                              ' ·ÖÐÐ,²¢ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              m_lngListCount = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                              For lngI = 1 To lngTabCount
                                  ReDim Preserve m_aryTabList(m_lngListCount).list(lngListTabIndex)
                                  With m_aryTabList(m_lngListCount).list(lngListTabIndex)
                                       .Index = lngI
                                  End With
                                  If lngListTabIndex > 0 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                                  If lngI + 1 <= lngTabCount Then
                                     lngListWidth = lngListWidth + .Item(lngI + 1).Width
                                  End If
                                  lngListTabIndex = lngListTabIndex + 1
                                  If lngListWidth > lngWidth Then
                                     ' ´æ´¢Ã¿ÐÐÖÐTabµÄ¸öÊý
                                     m_aryTabList(m_lngListCount).Count = lngListTabIndex
                                     lngListTabIndex = 0
                                     m_lngListCount = m_lngListCount + 1
                                     ReDim Preserve m_aryTabList(m_lngListCount)
                                     ReDim Preserve m_aryTabList(m_lngListCount).list(0)
                                     m_aryTabList(m_lngListCount).Count = 1
                                     lngListWidth = .Item(lngI + 1).Width
                                  End If
                              Next lngI
                              ' ÉèÖÃÃ¿ÐÐµÄ¸ß¶È
                              For lngI = m_lngListCount To 0 Step -1
                                  For lngY = 0 To m_aryTabList(lngI).Count - 1
                                      m_clsTabs.Item(m_aryTabList(lngI).list(lngY).Index).Top = (m_lngListCount - lngI) * m_lngDefaultTabHeight + lngDiscrepancy
                                  Next lngY
                              Next lngI
                              m_lngListCount = m_lngListCount + 1
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥µã
                              m_udtMainRect.Top = m_lngListCount * m_lngDefaultTabHeight + lngDiscrepancy
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ¼ÆËãTabµÄÐÐÊý
                              m_lngListCount = CLng(lngAllWidth \ lngWidth)
                              If lngAllWidth Mod lngWidth > 0 Then m_lngListCount = m_lngListCount + 1
                              m_udtMainRect.Top = m_lngListCount * lngManualHeight + lngDiscrepancy
                              ' ³õÊÔÊý×é
                              ReDim m_aryTabList(m_lngListCount - 1)
                              For lngI = 0 To m_lngListCount - 1
                                  ReDim m_aryTabList(lngI).list(0)
                              Next lngI
                              ' ·ÖÐÐ,²¢ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              lngListIndex = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                              For lngI = 1 To lngTabCount
                                  ReDim Preserve m_aryTabList(lngListIndex).list(lngListTabIndex)
                                  With m_aryTabList(lngListIndex).list(lngListTabIndex)
                                       .Index = lngI
                                  End With
                                  .Item(lngI).Top = (m_lngListCount - lngListIndex - 1) * m_lngDefaultTabHeight + 1
                                  If lngListTabIndex > 0 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                                  If lngI + 1 <= lngTabCount Then
                                     lngListWidth = lngListWidth + .Item(lngI + 1).Width
                                  End If
                                  If lngListWidth > lngWidth Then
                                     lngListIndex = lngListIndex + 1
                                     lngListTabIndex = 0
                                  Else
                                     lngListTabIndex = lngListTabIndex + 1
                                  End If
                                  ' ´æ´¢Ã¿ÐÐÖÐTabµÄ¸öÊý
                                  If lngListIndex <= m_lngListCount - 1 Then m_aryTabList(lngListIndex).Count = lngListTabIndex
                              Next lngI
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Top = lngManualHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = (lngManualWidth + TabsInterval) * (lngI - 1)
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementTopRight
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Top = lngManualHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = m_udtMainRect.Right - 10 - (lngManualWidth + TabsInterval) * lngI
                                  Else
                                     .Item(lngI).Left = m_udtMainRect.Right - 10 - .Item(lngI).Width
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementLeftTop
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Left = lngManualWidth
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Left = lngManualWidth
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Left = 0
                                  If lngI > 1 Then
                                     .Item(lngI).Top = m_udtMainRect.Top + lngManualHeight * (lngI - 1)
                                  Else
                                     .Item(lngI).Top = 0
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementLeftBottom
                Case GpTabPlacementBottomLeft
                Case GpTabPlacementBottomRight
                Case GpTabPlacementRightTop
                Case GpTabPlacementRightBottom
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Right = m_udtMainRect.Right - lngManualWidth
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ÉèÖÃÃ¿¸öTabµÄ¸ß¶ÈºÍ¿í¶È
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ÉèÖÃÖ÷ÇøÓòµÄ¶¥²¿
                              m_udtMainRect.Left = lngManualWidth
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ÉèÖÃÃ¿¸öTabµÄ×ø±ê
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Left = m_udtMainRect.Right
                                  If lngI > 1 Then
                                     .Item(lngI).Top = m_udtMainRect.Top + lngManualHeight * (lngI - 1)
                                  Else
                                     .Item(lngI).Top = 0
                                  End If
                              Next lngI
                           End If
                  End Select
         End Select
    End With
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub pvCalculateTrapezoidPoint(ByRef DstPoint() As POINTAPI, _
                                      ByVal Count As Long, _
                                      ByVal Top As Long, _
                                      ByVal Left As Long, _
                                      ByVal Right As Long, _
                                      ByVal Bottom As Long)
    Dim lngindex As Long
    
    Select Case m_udtPlacement
           Case GpTabPlacementTopleft
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .x = Left
                  .y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Left
                  .y = Top + RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Left + RoundRectSize
                  .y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right - (Bottom - Top)
                  .y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right
                  .y = Bottom
             End With
           Case GpTabPlacementTopRight
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .x = Left
                  .y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Left + Bottom - Top
                  .y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right - RoundRectSize
                  .y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right
                  .y = Top + RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right
                  .y = Bottom
             End With
           Case GpTabPlacementBottomLeft
           Case GpTabPlacementBottomRight
           Case GpTabPlacementLeftTop
           Case GpTabPlacementLeftBottom
           Case GpTabPlacementRightTop
           Case GpTabPlacementRightBottom
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .x = Left
                  .y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right
                  .y = Top + Right - Left
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right
                  .y = Bottom - RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Right - RoundRectSize
                  .y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .x = Left
                  .y = Bottom
             End With
    End Select
End Sub

' »æÖÆ¿Ø¼þ½çÃæ
Private Sub pvDraw()
    Dim lngI                As Long
    Dim lngTop1             As Long
    Dim lngTop2             As Long
    Dim lngBottom1          As Long
    Dim lngBottom2          As Long
    Dim lngHeight           As Long
    Dim lngPointCount       As Long
    Dim lngBrush            As Long
    Dim lngPen              As Long
    Dim lngXPColor          As Long
    Dim lngStepXP           As Single
    Dim blnHover            As Boolean
    Dim blnLeftTop          As Boolean
    Dim blnLeftBottom       As Boolean
    Dim blnRightBottom      As Boolean
    Dim blnRightTop         As Boolean
    Dim tBR                 As RECT
    Dim udtPointA()         As POINTAPI
    Dim udtPointB()         As POINTAPI
    Const ProcName = "pvDraw"
    
    On Error GoTo ErrorHandle
    If m_lngDefaultTabHeight <= 0 Then Exit Sub
    With UserControl
         ' ÉèÖÃ¿Ø¼þ±³¾°ÑÕÉ«
         If m_blnAutoBackColor Then
            .BackColor = .Ambient.BackColor
         Else
            .BackColor = m_oleBackColor
         End If
         ' Çå³ý
         .Cls
    End With
    
    ' ¿Ø¼þWin32·ç¸ñ
    If m_udtStyle = GpTabStyleStandard Then
       ' ´´½¨Ë¢×Ó
       lngBrush = CreateSolidBrush(TranslateColor(m_oleTabColorActive))
       ' Ìî³ä
       Call FillRect(UserControl.hdc, m_udtMainRect, lngBrush)
       ' É¾³ý»­Ë¢
       Call DeleteObject(lngBrush): lngBrush = 0
       If m_udtBorderStyle = GpTabBorderStyle3D Then
          Call DrawEdge(UserControl.hdc, m_udtMainRect, EDGE_RAISED, BF_RECT)
       ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
          Call DrawEdge(UserControl.hdc, m_udtMainRect, BDR_RAISEDINNER, BF_RECT)
       End If
    ' ¿Ø¼þWinXP·ç¸ñ
    ElseIf m_udtStyle = GpTabStyleWinXP Then
       If m_udtBorderStyle = GpTabBorderStyleNone Then
          ' ´´½¨Ë¢×Ó
          lngBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorActive))
          ' Ìî³ä
          Call FillRect(UserControl.hdc, m_udtMainRect, lngBrush)
          ' É¾³ý»­Ë¢
          Call DeleteObject(lngBrush): lngBrush = 0
          With m_udtMainRect
               tBR.Top = .Top
               tBR.Left = .Left
               tBR.Right = .Right
               tBR.Bottom = .Top + 3
          End With
          ' ´´½¨Ë¢×Ó
          lngBrush = CreateSolidBrush(TranslateColor(XPFlatBorderColor))
          ' Ìî³ä
          Call FillRect(UserControl.hdc, tBR, lngBrush)
          ' É¾³ý»­Ë¢
          Call DeleteObject(lngBrush): lngBrush = 0
       Else
          blnLeftTop = False
          blnLeftBottom = True
          blnRightBottom = True
          blnRightTop = True
          lngPointCount = 8
          ' ÉèÖÃµ±Ç°»­±Ê,±ß¿òÑÕÉ«
          'lngPen = CreatePen(PS_SOLID, 1, TranslateColor(XPBorderColor))
          'Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
          With m_udtMainRect
               Call pvCalculateRect(udtPointA, lngPointCount, .Top, .Left, .Right, .Bottom, _
                                    blnLeftTop, blnLeftBottom, blnRightBottom, blnRightTop)
               lngHeight = .Bottom - .Top
               lngStepXP = 25 / lngHeight
               lngTop1 = 1
               lngTop2 = 2
               lngBottom1 = lngHeight - 1
               lngBottom2 = lngHeight
               For lngI = 1 To lngHeight
                   Select Case lngI
                            'HERE
                          Case lngTop1
                            lngXPColor = TranslateColor(m_oleBackColor)
                            Call DrawLine(UserControl.hdc, IIf(blnLeftTop, .Left + 2, .Left), .Top, _
                                          IIf(blnRightTop, .Right - 2, .Right), .Top, _
                                          lngXPColor)
                          Case lngTop2
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftTop, .Left + 1, .Left), .Top + 1, _
                                          IIf(blnRightTop, .Right - 1, .Right), .Top + 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case lngBottom1
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftBottom, .Left + 2, .Left), .Bottom - 1, _
                                          IIf(blnRightBottom, .Right - 2, .Right), .Bottom - 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case lngBottom2
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftBottom, .Left + 1, .Left), .Bottom - 2, _
                                          IIf(blnRightBottom, .Right - 1, .Right), .Bottom - 2, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case Else
                            'Call DrawLine(UserControl.hdc, .Left, lngI + .Top - 1, .Right, lngI + .Top - 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                   End Select
               Next lngI
               Call Polyline(UserControl.hdc, udtPointA(0), lngPointCount)
          End With
       End If
    End If
    If m_clsTabs.Count <= 0 Then
       Call pvDrawTab("", 0, DiscrepancyHeight, 60, m_lngDefaultTabHeight, m_oleTabColorActive, lngXPColor, True, False)
    Else
       For lngI = m_clsTabs.Count To 1 Step -1
           blnHover = False
           With m_clsTabs.Item(lngI)
                If Not (m_clsHoverTab Is Nothing) Then
                   If m_clsHoverTab.Index = lngI Then blnHover = True
                End If
                Call pvDrawTab(.Caption, .Left, .Top, .Width, .Height, m_oleTabColorActive, lngXPColor, .Selected, blnHover)
           End With
       Next lngI
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub pvDrawTab(ByVal Caption As String, _
                      ByVal Left As Long, _
                      ByVal Top As Long, _
                      ByVal Width As Long, _
                      ByVal Height As Long, _
                      ByVal TabColor As Long, _
                      ByVal TranXPColor As Long, _
                      ByVal Selected As Boolean, _
                      ByVal Hover As Boolean)
    Dim lngTop1             As Long
    Dim lngTop2             As Long
    Dim lngEdgeStyle        As Long
    Dim lngEdgeFlag         As Long
    Dim lngPen              As Long
    Dim lngOldBrush         As Long
    Dim lngXPBorderBrush    As Long
    Dim lngTabBorderBrush   As Long
    
    Dim lngFlatBrush        As Long
    Dim lngFlatActiveBrush  As Long
    Dim lngFlatHoverBrush   As Long
    Dim lngFlatBorderBrush  As Long
    
    Dim lngI                As Long
    Dim lngLightColor       As Long
    Dim lngHighLightColor   As Long
    Dim lngShadowColor      As Long
    Dim lngDarkShadowColor  As Long
    Dim lngXPColor          As Long
    Dim lngStepXP           As Single
    Dim udtPointA()         As POINTAPI
    Dim udtPointB()         As POINTAPI
    Dim udtTabRect          As RECT
    Dim udtCaptionRect      As RECT
    Const ProcName = "pvDrawTab"
    
    On Error GoTo ErrorHandle
    
    ' È¡µÃÏµÍ³ÏÔÊ¾¶ÔÏóµÄÑÕÉ«
    lngShadowColor = GetSysColor(COLOR_BTNSHADOW)
    lngLightColor = GetSysColor(COLOR_BTNLIGHT)
    lngDarkShadowColor = GetSysColor(COLOR_BTNDKSHADOW)
    lngHighLightColor = GetSysColor(COLOR_BTNHIGHLIGHT)
    
    ' ½¨Á¢»­Ë¢
    lngXPBorderBrush = CreateSolidBrush(TranslateColor(XPBorderColor))
    
    lngFlatBrush = CreateSolidBrush(TranslateColor(XPFlatTabColor))
    lngFlatActiveBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorActive))
    lngFlatHoverBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorHover))
    lngFlatBorderBrush = CreateSolidBrush(TranslateColor(XPFlatBorderColor))
    
    lngTabBorderBrush = CreateSolidBrush(TranslateColor(TabColor))
    lngOldBrush = SelectObject(UserControl.hdc, lngTabBorderBrush)
    With udtTabRect
         .Left = Left
         .Top = Top
         .Right = Left + Width
         .Bottom = Top + Height
    End With
    
    ' ¿Ø¼þWin32·ç¸ñ
    If m_udtStyle = GpTabStyleStandard Then
       Select Case m_udtTabStyle
              Case GpTabRectangle
                If m_udtBorderStyle = GpTabBorderStyle3D Then
                   lngEdgeStyle = EDGE_RAISED
                   'lngEdgeStyle = BDR_RAISEDINNER
                   If Selected Then udtTabRect.Bottom = Top + Height + 2
                ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                   lngEdgeStyle = BDR_RAISEDINNER
                   If Selected Then udtTabRect.Bottom = Top + Height + 1
                End If
                Call FillRect(UserControl.hdc, udtTabRect, lngTabBorderBrush)
                If m_udtBorderStyle <> GpTabBorderStyleNone And Selected = True Then Call DrawEdge(UserControl.hdc, udtTabRect, lngEdgeStyle, BF_LEFT Or BF_TOP Or BF_RIGHT)
              Case GpTabRoundRect
                ReDim udtPointB(5)
                With udtTabRect
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        .Left = .Left + 1
                        .Bottom = .Bottom + 2
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        .Bottom = .Bottom + 1
                     End If
                     ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                     Call pvCalculateRoundPoint(udtPointA, 6, .Top, _
                                                .Left, .Right, _
                                                .Bottom)
                     ' Ìî³ä
                     lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                     Call Polygon(UserControl.hdc, udtPointA(0), 6)
                     lngPen = CreatePen(PS_SOLID, 1, lngLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call pvInflatePoint(udtPointB(0), udtPointA(0), 1, 0)
                     Call pvInflatePoint(udtPointB(1), udtPointA(1), 1, 1)
                     Call pvInflatePoint(udtPointB(2), udtPointA(2), 1, 1)
                     Call pvInflatePoint(udtPointB(3), udtPointA(3), -1, 1)
                     Call pvInflatePoint(udtPointB(4), udtPointA(4), -1, 1)
                     Call pvInflatePoint(udtPointB(5), udtPointA(5), -1, 0)
                     Call Polyline(UserControl.hdc, udtPointB(0), 6)
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        ' È¥µ×±ß
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ¼ÓÒõÓ°
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 3, Top, .Right - 1, .Top + 3, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 1, TranslateColor(lngShadowColor))
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        ' È¥µ×±ß
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ¼ÓÒõÓ°
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngShadowColor))
                     End If
                End With
              Case GpTabTrapezoid
                ReDim udtPointB(4)
                With udtTabRect
                     If m_udtBorderStyle = GpTabBorderStyleNone Then
                        ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                        ' Ìî³ä
                        lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                        Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                        Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                        Call Polygon(UserControl.hdc, udtPointA(0), 5)
                     Else
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        .Left = .Left + 1
                        .Bottom = .Bottom + 2
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        .Bottom = .Bottom + 1
                     End If
                     ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                     Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                     ' Ìî³ä
                     lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                     Call Polygon(UserControl.hdc, udtPointA(0), 5)
                     lngPen = CreatePen(PS_SOLID, 1, lngLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call pvInflatePoint(udtPointB(0), udtPointA(0), 1, 0)
                     Call pvInflatePoint(udtPointB(1), udtPointA(1), 1, 1)
                     Call pvInflatePoint(udtPointB(2), udtPointA(2), -1, 1)
                     Call pvInflatePoint(udtPointB(3), udtPointA(3), -1, 1)
                     Call pvInflatePoint(udtPointB(4), udtPointA(4), -1, 0)
                     Call Polyline(UserControl.hdc, udtPointB(0), 5)
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        ' È¥µ×±ß
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ¼ÓÒõÓ°
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 3, .Top, .Right - 1, .Top + 3, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 1, TranslateColor(lngShadowColor))
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        ' È¥µ×±ß
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ¼ÓÒõÓ°
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngShadowColor))
                     End If
                     End If
                End With
       End Select
    ' ¿Ø¼þWinXP·ç¸ñ
    'LAST HERE -> LOOK FOR THE BLUE EDGE ON TAB HOW TO REMOVE
    ElseIf m_udtStyle = GpTabStyleWinXP Then
       Select Case m_udtTabStyle
              Case GpTabRectangle
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                Else
                   With udtTabRect
                        ' Ìî³ä
                        lngStepXP = 25 / Height
                        For lngI = Height To 1 Step -1
                            Call DrawLine(UserControl.hdc, .Left + 1, lngI + .Top, .Right, lngI + .Top, _
                                          BrightnessColor(TranXPColor, lngStepXP * lngI))
                        Next lngI
                        lngXPColor = BrightnessColor(TranXPColor, lngStepXP * 1)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        .Right = .Right + 1  ' Ê¹¼ä¾à±äÐ¡
                        .Bottom = Top + Height + 1
                        ' »­±ß¿ò
                        Call FrameRect(UserControl.hdc, udtTabRect, lngXPBorderBrush)
                        If Selected Then
                           ' È¥¶¥Ïß
                           'Call DrawLine(UserControl.hdc, .Left, .Top, .Right, .Top, lngXPColor)
                           ' »­½¹µãÏß
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 2, Top - 2, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right - 1, Top - 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top, .Right, Top, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right - 1, Top + 1, &HFFC5B2)
                           ' µ±Ñ¡ÖÐÊ±È¥µ×±ß
                           'Call DrawLine(UserControl.hdc, .Left + 1, .Bottom - 1, .Right - 1, .Bottom - 1, TranXPColor)
                        Else
                           If Hover Then
                              ' È¥¶¥Ïß
                              'Call DrawLine(UserControl.hdc, .Left, .Top, .Right, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 2, Top - 2, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right - 1, Top - 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top, .Right, Top, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right - 1, Top + 1, &H9DDBFF)
                           End If
                        End If
                   End With
                End If
              Case GpTabRoundRect
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                Else
                   lngStepXP = 25 / Height
                   With udtTabRect
                        ' Ìî³ä
                        lngTop1 = .Top
                        lngTop2 = .Top + 1
                        For lngI = Height To 1 Step -1
                            If lngI = lngTop1 Then
                               Call DrawLine(UserControl.hdc, .Left + 2, lngI + .Top, .Right - 2, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            ElseIf lngI = lngTop2 Then
                               Call DrawLine(UserControl.hdc, .Left + 1, lngI + .Top, .Right - 1, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            Else
                               Call DrawLine(UserControl.hdc, .Left, lngI + .Top, .Right, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            End If
                        Next lngI
                        lngXPColor = BrightnessColor(TranXPColor, lngStepXP * .Top)
                        ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                        Call pvCalculateRoundPoint(udtPointA, 6, .Top, .Left, .Right, .Bottom)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        ' »­±ß¿ò
                        Call Polyline(UserControl.hdc, udtPointA(0), 6)
                        If Selected Then
                           ' È¥¶¥Ïß
                           'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                           ' »­½¹µãÏß
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 1, Top - 2, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right, Top - 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top, .Right + 1, Top, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right, Top + 1, &HFFC5B2)
                        Else
                           If Hover Then
                              ' È¥¶¥Ïß
                              'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 1, Top - 2, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right, Top - 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top, .Right + 1, Top, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &H9DDBFF)
                           End If
                           ' Ã»ÓÐÑ¡ÖÐÊ±¼Óµ×±ß
                           Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, XPBorderColor)
                        End If
                   End With
                End If
              Case GpTabTrapezoid
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                   ReDim udtPointB(4)
                   With udtTabRect
                        ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom - 1)
                        ' Ìî³ä
                        If Selected Then
                          
                           lngPen = CreatePen(PS_SOLID, 1, vbWhite)
                           Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                           Call SelectObject(UserControl.hdc, vbWhite)
                           Call pvInflatePoint(udtPointB(0), udtPointA(0), 2, 2)
                           Call pvInflatePoint(udtPointB(1), udtPointA(1), 2, 2)
                           Call pvInflatePoint(udtPointB(2), udtPointA(2), -2, 2)
                           Call pvInflatePoint(udtPointB(3), udtPointA(3), -2, 2)
                           Call pvInflatePoint(udtPointB(4), udtPointA(4), -2, 2)
                           Call Polygon(UserControl.hdc, udtPointB(0), 5)
                           
                           'Call DrawLine(UserControl.hdc, .Left, .Bottom + 2, .Right, .Bottom + 2, XPFlatTabColorActive)
                        Else
                           lngPen = CreatePen(PS_SOLID, 0, &H808080)
                           Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                           Call SelectObject(UserControl.hdc, &H808080)
                           Call Polygon(UserControl.hdc, udtPointA(0), 5)
                        End If
                   End With
                Else
                   lngStepXP = 25 / Height
                   With udtTabRect
                        ' Ìî³ä
                        lngTop1 = .Top
                        lngTop2 = .Top + 1
                        For lngI = Height To 1 Step -1
                            If lngI = lngTop1 Then
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + Height, lngI + .Top, .Right - 2, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            ElseIf lngI = lngTop2 Then
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + Height - 1, lngI + .Top, .Right - 1, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            Else
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + (.Bottom - lngI), lngI + .Top, .Right, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            End If
                        Next lngI
                        ' È¡µÃÔ²½ÇµÄ¸÷¸öµãµÄ×ø±ê
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        ' »­±ß¿ò
                        Call Polyline(UserControl.hdc, udtPointA(0), 5)
                        If Selected Then
                           ' È¥¶¥Ïß
                           'HERE
                           'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                           ' »­½¹µãÏß
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top, .Right - 1, Top, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top + 2, .Right + 1, Top + 2, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 3, .Right, Top + 3, &HFFC5B2)
                        Else
                           If Hover Then
                              ' È¥¶¥Ïß
                              'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top, .Right - 1, Top, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top + 2, .Right + 1, Top + 2, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 3, .Right, Top + 3, &H9DDBFF)
                           End If
                           ' Ã»ÓÐÑ¡ÖÐÊ±¼Óµ×±ß
                           Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, XPBorderColor)
                        End If
                   End With
                End If
       End Select
    End If
    
    If m_udtTabStyle = GpTabTrapezoid Then
       With udtTabRect
            .Left = .Left + m_lngDefaultTabHeight
       End With
    End If
    ' Draw Caption
    Call DrawTextEx(UserControl.hdc, Caption & vbNullString, -1, udtTabRect, _
                    pvGetCaptionFlags(GpTabCaptionRight), m_udtDrawTextParams)
    If lngOldBrush <> 0 Then Call SelectObject(UserControl.hdc, lngOldBrush): lngOldBrush = 0
    Call DeleteObject(lngXPBorderBrush)
    Call DeleteObject(lngTabBorderBrush)
    Call DeleteObject(lngFlatBrush)
    Call DeleteObject(lngFlatActiveBrush)
    Call DeleteObject(lngFlatHoverBrush)
    Call DeleteObject(lngFlatBorderBrush)
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             If lngOldBrush <> 0 Then Call SelectObject(UserControl.hdc, lngOldBrush)
             Call DeleteObject(lngXPBorderBrush)
             Call DeleteObject(lngTabBorderBrush)
             Call DeleteObject(lngFlatBrush)
             Call DeleteObject(lngFlatActiveBrush)
             Call DeleteObject(lngFlatHoverBrush)
             Call DeleteObject(lngFlatBorderBrush)
    End Select
End Sub

Private Function pvGetCaptionFlags(ByVal Alignment As GPTAB_ALIGNMENT_METHOD) As Long
    Select Case Alignment
           Case GpTabCaptionLeft
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_LEFT Or DT_VCENTER
           Case GpTabCaptionRight
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_RIGHT Or DT_VCENTER
           Case GpTabCaptionCenter
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CENTER Or DT_VCENTER
    End Select
End Function

Private Sub pvInflatePoint(ByRef DstPoint As POINTAPI, ByRef SrcPoint As POINTAPI, ByVal x As Long, ByVal y As Long)
    With DstPoint
         .x = SrcPoint.x + x
         .y = SrcPoint.y + y
    End With
End Sub

Private Sub pvSetColor()
    If m_udtXPColorScheme = GpTabUseWindows Then
       m_lngXPFaceColor = BrightnessColor(GetSysColor(COLOR_BTNFACE), &H30)
    Else
       m_lngXPFaceColor = BrightnessColor(m_oleTabColorActive, &H30)
    End If
End Sub

' Ë¢ÐÂ
Public Sub Refresh()
    Call pvDraw
End Sub

Public Property Get SelectTabItem() As cTabItem
    If m_clsSelectTab Is Nothing Then
       Set SelectTabItem = Nothing
    Else
       Set SelectTabItem = m_clsSelectTab
    End If
End Property

Public Property Get Style() As GPTAB_STYLE_METHOD
    Style = m_udtStyle
End Property

Public Property Let Style(ByVal New_Style As GPTAB_STYLE_METHOD)
    m_udtStyle = New_Style
    PropertyChanged "Style"
    Call pvDraw
End Property

Public Property Get TabBorderColor() As OLE_COLOR
    TabBorderColor = m_oleTabBorderColor
End Property

Public Property Let TabBorderColor(ByVal New_TabBorderColor As OLE_COLOR)
    m_oleTabBorderColor = New_TabBorderColor
    PropertyChanged "TabBorderColor"
End Property

Public Property Get TabColor() As OLE_COLOR
    TabColor = m_oleTabColor
End Property

Public Property Let TabColor(ByVal New_TabColor As OLE_COLOR)
    m_oleTabColor = New_TabColor
    PropertyChanged "TabColor"
    Call pvDraw
End Property

Public Property Get TabColorActive() As OLE_COLOR
    TabColorActive = m_oleTabColorActive
End Property

Public Property Let TabColorActive(ByVal New_TabColorActive As OLE_COLOR)
    m_oleTabColorActive = New_TabColorActive
    PropertyChanged "TabColorActive"
    Call pvSetColor
    Call pvDraw
End Property

Public Property Get TabColorHover() As OLE_COLOR
    TabColorHover = m_oleTabColorHover
End Property

Public Property Let TabColorHover(ByVal New_TabColorHover As OLE_COLOR)
    m_oleTabColorHover = New_TabColorHover
    PropertyChanged "TabColorHover"
End Property

Public Property Get TabFixedHeight() As Long
    TabFixedHeight = m_lngTabFixedHeight
End Property

Public Property Let TabFixedHeight(ByVal New_TabFixedHeight As Long)
    m_lngTabFixedHeight = New_TabFixedHeight
    PropertyChanged "TabFixedHeight"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get TabFixedWidth() As Long
    TabFixedWidth = m_lngTabFixedWidth
End Property

Public Property Let TabFixedWidth(ByVal New_TabFixedWidth As Long)
    m_lngTabFixedWidth = New_TabFixedWidth
    PropertyChanged "TabFixedWidth"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get Tabs() As cTabItems
    Set Tabs = m_clsTabs
End Property

Public Property Get TabStyle() As GPTAB_TABSTYLE_METHOD
    TabStyle = m_udtTabStyle
End Property

Public Property Let TabStyle(ByVal New_TabStyle As GPTAB_TABSTYLE_METHOD)
    m_udtTabStyle = New_TabStyle
    PropertyChanged "TabStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get TabWidthStyle() As GPTAB_TABWIDTHSTYLE_METHOD
    TabWidthStyle = m_udtTabWidthStyle
End Property

Public Property Let TabWidthStyle(ByVal New_TabWidthStyle As GPTAB_TABWIDTHSTYLE_METHOD)
    m_udtTabWidthStyle = New_TabWidthStyle
    PropertyChanged "TabWidthStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get XPColorScheme() As GPTAB_XPCOLORSCHEME_METHOD
    XPColorScheme = m_udtXPColorScheme
End Property

Public Property Let XPColorScheme(ByVal New_XPColorScheme As GPTAB_XPCOLORSCHEME_METHOD)
    m_udtXPColorScheme = New_XPColorScheme
    PropertyChanged "XPColorScheme"
    Call pvSetColor
    Call pvDraw
End Property

Private Sub m_clsTabs_TabAddNew()
    If m_clsSelectTab Is Nothing Then
       Set m_clsSelectTab = m_clsTabs.Item(1)
       m_clsSelectTab.Selected = True
    End If
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabAlignmentChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabCaptionChanged(ByVal Index As Long)
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabIconAlignChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabIconChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabRemove()
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabSelectedChanged(ByVal Index As Long)
    '
End Sub

Private Sub UserControl_Click()
    If m_blnEnabled Then RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Set m_clsTabs = New cTabItems
    With m_udtDrawTextParams
         .iLeftMargin = 1
         .iRightMargin = 1
         .iTabLength = 1
         .cbSize = Len(m_udtDrawTextParams)
    End With
    m_blnAutoBackColor = True
    m_blnEnabled = True
    m_blnHotTracking = True
    m_blnMultiRow = True
    m_lngTabFixedHeight = 0
    m_lngTabFixedWidth = 0
    m_udtXPColorScheme = GpTabUseWindows
    m_udtBorderStyle = GpTabBorderStyle3D
    m_udtPlacement = GpTabPlacementTopleft
    m_udtStyle = GpTabStyleStandard
    m_oleTabBorderColor = vbWhite
    m_oleTabColor = vbButtonShadow
    m_oleTabColorActive = vbButtonFace
    m_oleTabColorHover = vbHighlight
    m_udtTabStyle = GpTabRectangle
    m_udtTabWidthStyle = GpTabJustified
    Call pvSetColor
End Sub

Private Sub UserControl_InitProperties()
'    Me.AutoBackColor = True
    Call pvCalculateSize
    Me.BackColor = UserControl.Ambient.BackColor
'    Me.BorderStyle = GpTabBorderStyle3D
    Set Me.Font = UserControl.Ambient.Font
'    Me.ForeColor = vbWindowText
'    Me.Enable = True
'    Me.MultiRow = True
'    Me.Placement = GpTabPlacementTopLeft
'    Me.Style = GpTabStyleStandard
'    Me.TabFixedHeight = 0
'    Me.TabFixedWidth = 0
'    Me.TabStyle = GpTabStandard
'    Me.TabWidthStyle = GpTabJustified
End Sub

Public Sub SelectTab(Index As Integer)
    Dim lngI As Long
    Dim clsTemp As cTabItem
    
    'On Error GoTo ErrorHandle
    
        '
        For lngI = 1 To m_clsTabs.Count
            m_clsTabs.Item(lngI).Selected = False
        Next lngI
        
        m_clsTabs.Item(Index).Selected = True
        'clsTemp = m_clsTabs.Item(Index)
        Set m_clsSelectTab = m_clsTabs.Item(Index)
        m_clsSelectTab.Selected = True
        Call pvDraw
        'Set clsTemp = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError("SelectTab", MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort: Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngI As Long
    Dim clsTemp As cTabItem
    Const ProcName = "UserControl_MouseDown"
    On Error GoTo ErrorHandle
    DoEvents
    If m_blnEnabled Then RaiseEvent MouseDown(Button, Shift, x, y)
    If Button = vbLeftButton Then
       Set clsTemp = HitTest(x, y)
       If Not (clsTemp Is Nothing) Then
          If Not clsTemp Is m_clsSelectTab Then
             For lngI = 1 To m_clsTabs.Count
             DoEvents
                 m_clsTabs.Item(lngI).Selected = False
             Next lngI
             Set m_clsSelectTab = clsTemp
             m_clsSelectTab.Selected = True
             Call pvDraw
             RaiseEvent TabClick
          End If
       End If
    End If
    Set clsTemp = Nothing
   Exit Sub
ErrorHandle:
DoEvents
   'Select Case ShowError(ProcName, MODULE_NAME)
'           Case vbRetry: Resume
'           Case vbIgnore: Resume Next
'           Case vbAbort: Set clsTemp = Nothing
'    End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim clsTemp As cTabItem
    Const ProcName = "UserControl_MouseMove"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent MouseMove(Button, Shift, x, y)
    Set clsTemp = HitTest(x, y)
    If Not clsTemp Is m_clsHoverTab Then
       Set m_clsHoverTab = clsTemp
       If m_blnHotTracking Then Call pvDraw
    End If
    ' Êó±ê²¶»ñ
    If x >= 0 And x < UserControl.ScaleWidth And y >= 0 And y < UserControl.ScaleHeight And Button = 0 Then
        'If GetCapture() <> UserControl.hWnd Then Call SetCapture(UserControl.hWnd)
    Else
        'If GetCapture() = UserControl.hWnd And Button = 0 Then Call ReleaseCapture
    End If
    'If GetCapture() = UserControl.hWnd And Button = 0 Then Call ReleaseCapture
    Set clsTemp = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort: Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_blnEnabled Then RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    UserControl.AutoRedraw = True
    Call pvDraw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim clsFont As New StdFont
    
    Call pvCalculateSize
    m_blnUserMode = UserControl.Ambient.UserMode
    With PropBag
         Me.AutoBackColor = .ReadProperty("AutoBackColor", True)
         Me.BackColor = .ReadProperty("BackColor", UserControl.Ambient.BackColor)
         Me.BorderStyle = .ReadProperty("BorderStyle", GpTabBorderStyle3D)
         Me.Enable = .ReadProperty("Enable", True)
         Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
         Me.HotTracking = .ReadProperty("HotTracking", True)
         Set Me.MouseIcon = .ReadProperty("MouseIcon", Nothing)
         Me.MousePointer = .ReadProperty("MousePointer", 0)
         Me.MultiRow = .ReadProperty("MultiRow", True)
         Me.Placement = .ReadProperty("Placement", GpTabPlacementTopleft)
         Me.Style = .ReadProperty("Style", GpTabStyleStandard)
         Me.TabBorderColor = .ReadProperty("TabBorderColor", vbWhite)
         Me.TabColor = .ReadProperty("TabColor", vbButtonShadow)
         Me.TabColorActive = .ReadProperty("TabColorActive", vbButtonFace)
         Me.TabColorHover = .ReadProperty("TabColorHover", vbHighlight)
         Me.TabFixedHeight = .ReadProperty("TabFixedHeight", 0)
         Me.TabFixedWidth = .ReadProperty("TabFixedWidth", 0)
         Me.TabStyle = .ReadProperty("TabStyle", GpTabRectangle)
         Me.TabWidthStyle = .ReadProperty("TabWidthStyle", GpTabJustified)
         Me.XPColorScheme = .ReadProperty("XPColorScheme", GpTabUseWindows)
         With clsFont
              .Name = "MS Sans Serif"
              .Size = 8
         End With
         Set Me.Font = .ReadProperty("Font", clsFont)
    End With
    Set clsFont = Nothing
    Call pvSetColor
End Sub

Private Sub UserControl_Resize()
    Call pvCalculateSize
    Call pvDraw
End Sub

Public Sub Redraw()
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub UserControl_Terminate()
    Set m_clsSelectTab = Nothing
    Set m_clsHoverTab = Nothing
    Set m_clsTabs = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim clsFont As New StdFont
    
    With PropBag
         .WriteProperty "AutoBackColor", m_blnAutoBackColor, True
         .WriteProperty "BackColor", m_oleBackColor, UserControl.Ambient.BackColor
         .WriteProperty "BorderStyle", m_udtBorderStyle, GpTabBorderStyle3D
         .WriteProperty "Enable", m_blnEnabled, True
         .WriteProperty "ForeColor", Me.ForeColor, vbWindowText
         .WriteProperty "HotTracking", m_blnHotTracking, True
         .WriteProperty "MouseIcon", Me.MouseIcon, Nothing
         .WriteProperty "MousePointer", UserControl.MousePointer, 0
         .WriteProperty "MultiRow", m_blnMultiRow, True
         .WriteProperty "Placement", m_udtPlacement, GpTabPlacementTopleft
         .WriteProperty "Style", m_udtStyle, GpTabStyleStandard
         .WriteProperty "TabBorderColor", m_oleTabBorderColor, vbWhite
         .WriteProperty "TabColor", Me.TabColor, vbButtonShadow
         .WriteProperty "TabColorActive", Me.TabColorActive, vbButtonFace
         .WriteProperty "TabColorHover", Me.TabColorHover, vbHighlight
         .WriteProperty "TabFixedHeight", m_lngTabFixedHeight, 0
         .WriteProperty "TabFixedWidth", m_lngTabFixedWidth, 0
         .WriteProperty "TabStyle", m_udtTabStyle, GpTabRectangle
         .WriteProperty "TabWidthStyle", m_udtTabWidthStyle, GpTabJustified
         .WriteProperty "XPColorScheme", m_udtXPColorScheme, GpTabUseWindows
         With clsFont
              .Name = "MS Sans Serif"
              .Size = 8
         End With
         .WriteProperty "Font", Me.Font, clsFont
    End With
    Set clsFont = Nothing
End Sub

Private Function ShowError(ByVal strFunc As String, ByVal strModule As String) As VbMsgBoxResult
    Dim lngErrNumber             As Long
    Dim strErrDescription        As String
    Dim strErrSource             As String
    
    lngErrNumber = Err.Number
    strErrDescription = Err.Description
    strErrSource = IIf(Len(strModule) > 0, _
                            "[\\" & ErrComputerName() & "] " & _
                            App.EXEName & "." & _
                            strModule & "." & _
                            strFunc & _
                            IIf(Erl <> 0, "(" & Erl & ")", ""), "") & "--" & Err.Source
    ShowError = MsgBox( _
            strErrDescription & vbCrLf & vbCrLf & _
            "Error: 0x" & Hex(lngErrNumber) & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & _
            strErrSource, vbCritical Or vbAbortRetryIgnore, "Error")
End Function

Private Function ErrComputerName() As String
    Static sName        As String
        
    If Len(sName) = 0 Then
        sName = String(256, 0)
        GetComputerName sName, Len(sName)
        sName = Left$(sName, InStr(sName, Chr(0)) - 1)
    End If
    ErrComputerName = sName
End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture
    Dim IGuid    As Guid
    Dim NewPic   As Picture
    Dim tPicConv As PICTDESC
    
    If (hBmp = 0) Then Exit Function
   
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
         .cbSizeofStruct = Len(tPicConv)
         .picType = vbPicTypeBitmap
         .hImage = hBmp
    End With
    
    ' Fill in IDispatch Interface ID
    With IGuid
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
    End With
   
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
    ' Return it:
    Set BitmapToPicture = NewPic
End Function

Public Function BrightnessColor(ByVal ColorValue As Long, ByVal Increment As Long) As Long
    Dim R, g, b As Long
    
    b = ((ColorValue \ &H10000) Mod &H100): b = b + ((b * Increment) \ &HC0)
    g = ((ColorValue \ &H100) Mod &H100) + Increment
    R = (ColorValue And &HFF) + Increment
    If R < 0 Then R = 0
    If R > 255 Then R = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    BrightnessColor = RGB(R, g, b)
End Function

Private Sub DrawDragImage(ByRef rcNew As RECT, _
                         ByVal bFirst As Boolean, _
                         ByVal bLast As Boolean)
    Static rcCurrent     As RECT
    Dim hdc              As Long
    Dim lngReturn        As Long
    
    On Error Resume Next
    ' First get the Desktop DC:
    hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ' Set the draw mode to XOR:
    lngReturn = SetROP2(hdc, R2_NOTXORPEN)
    '// Draw over and erase the old rectangle
    If Not (bFirst) Then
       lngReturn = Rectangle(hdc, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
    End If
    If Not (bLast) Then
       '// Draw the new rectangle
       lngReturn = Rectangle(hdc, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom)
    End If
    ' Store this position so we can erase it next time:
    LSet rcCurrent = rcNew
    ' Free the reference to the Desktop DC we got (make sure you do this!)
    lngReturn = DeleteDC(hdc)
End Sub

Public Sub DrawImage(ByVal hIml As Long, _
                     ByVal iIndex As Long, _
                     ByVal hdc As Long, _
                     ByVal xPixels As Integer, _
                     ByVal yPixels As Integer, _
                     ByVal lIconSizeX As Long, _
                     ByVal lIconSizeY As Long, _
                     Optional ByVal bSelected = False, _
                     Optional ByVal bCut = False, _
                     Optional ByVal bDisabled = False, _
                     Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
                     Optional ByVal hExternalIml As Long = 0)
    Dim hIcon        As Long
    Dim lFlags       As Long
    Dim lhIml        As Long
    Dim lColor       As Long
    Dim iImgIndex    As Long
    Dim lngReturn    As Long
    
    ' Draw the image at 1 based index or key supplied in vKey.
    ' on the hDC at xPixels,yPixels with the supplied options.
    ' You can even draw an ImageList from another ImageList control
    ' if you supply the handle to hExternalIml with this function.
    On Error Resume Next
    iImgIndex = iIndex
    If (iImgIndex > -1) Then
       If (hExternalIml <> 0) Then
          lhIml = hExternalIml
       Else
          lhIml = hIml
       End If
       lFlags = ILD_TRANSPARENT
       If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
       End If
       If (bCut) Then
          ' Draw dithered:
          lColor = TranslateColor(oCutDitherColour)
          If (lColor = -1) Then lColor = TranslateColor(vbWindowBackground)
          lngReturn = ImageList_DrawEx(lhIml, iImgIndex, hdc, xPixels, yPixels, 0, 0, CLR_NONE, lColor, lFlags)
       ElseIf (bDisabled) Then
          ' extract a copy of the icon:
          hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
          ' Draw it disabled at x,y:
          lngReturn = DrawState(hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED)
          ' Clear up the icon:
          lngReturn = DestroyIcon(hIcon)
       Else
          ' Standard draw:
          lngReturn = ImageList_Draw(lhIml, iImgIndex, hdc, xPixels, yPixels, lFlags)
       End If
    End If
End Sub

Public Sub DrawLine(ByVal hdc As Long, _
                    ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    ByVal Color As Long, _
                    Optional Width As Long = 1)
    Dim lngPen      As Long
    Dim lngPenOld   As Long
    Dim pt          As POINTAPI
    Const FuncName = "DrawLine"
    
    On Error GoTo ErrorHandle
    
    '/* ´´½¨Ò»¸ö»­±Ê */
    lngPen = CreatePen(PS_SOLID, Width, Color)
    If lngPen <> 0 Then lngPenOld = SelectObject(hdc, lngPen)
    '/* Ö¸¶¨Ò»¸öÐÂµÄµ±Ç°»­±ÊÎ»ÖÃX1,Y1¡£Ç°Ò»¸öÎ»ÖÃ±£´æÔÚptÖÐ */
    MoveToEx hdc, X1, Y1, pt
    '/* »­Ò»ÌõÏß */
    LineTo hdc, X2, Y2
    If lngPenOld <> 0 Then SelectObject hdc, lngPenOld
    lngPenOld = 0
    If lngPen <> 0 Then DeleteObject lngPen
    Exit Sub
ErrorHandle:
    Select Case ShowError(FuncName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             If lngPenOld <> 0 Then SelectObject hdc, lngPenOld
             lngPenOld = 0
             If lngPen <> 0 Then DeleteObject lngPen
    End Select
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_NONE
End Function

' Returns Color as long, accepts SystemColorConstants
Public Function VerifyColor(ByVal ColorVal As Long) As Long
    VerifyColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then VerifyColor = GetSysColor(ColorVal And &HFFFFFF)
End Function


