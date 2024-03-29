VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRgnShaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' BRIEF HISTORY
' none of the following versions below ever used CombineRgn as a primary method
' \/\/\/\/\/\/\
' *** VERSION 1 - used GetPixel
' *** VERSION 2 - use of DIBs  vs GetPixel (Jun 04)
' UPDATED to include using DIB sections. The increase in speed is
' truly amazing over earlier version.....
' Over a 300% increase in speed noted on smallish bitmaps (96x96) &
' Over a 425% increase noted on mid-sized bitmaps (265x265)

' This approach should be somewhat faster than typical approaches that
' create a region by using CreateRectRgn, CombineRgn & DeleteObject.
' That is 'cause this approach does not use any of those functions to
' create regions. Like typical approaches a rectangle format is used to
' "add to a region", however, this approach directly creates the region header
' & region structure and passes that to a single API to finish the job of
' creating the end result. For those that play around with such things,
' I think you will recognize the difference.

' EDITED: The ExtCreateRegion seems to have an undocumented restriction: it
' won't create regions comprising of more than 4K rectangles (win98). So to
' get around this for very complex bitmaps, I rewrote the function to create
' regions of 2K rectangles at a time if needed. This is still extremely
' fast. I compared the window shaping code from vbAccelerator with the
' SandStone.bmp (15,000 rects) in Windows directory. vbAccelerator's code
' averaged 4,900 ms. My routines averaged 77 ms & that's not a typo!

' EDITED: I was allowing default error trapping on UBound() to resize the
' rectangle array: when trying to update array element beyond UBound, error
' would occur & be redirected to resize the array. However, thanx to
' Robert Rayment, if the UBound checks are disabled in compile optimizations,
' then we get a crash. Therefore, checks made appropriately & a tiny loss of
' speed is the trade-off for safety.

' *** VERSION 3.1 - Anti Regions & speed modifications (10 Jan 05)
' Function has an optional parameter to return the anti-Region.
' That is the region of only transparent pixels. This could be used, for
' example, with APIs like FillRgn to replace the "transparent" color with
' another color.

' **** VERSION 4 - Accessed bitmap bytes directly within memory.
' This version was scrapped. Although the time saved on very large bitmaps was
' somewhat significant, the routines were not robust enough to handle all the
' various exceptions within bitmap structures (4bpp, 8pbb, compression, etc)
' Something on the backburner to possibly attack at a later time.

' -- Converted from Module to Class and added other methods
'    Note: This class does not store any regions or pointers. Regions that
'          you create from this class must be destroyed by you. There is only
'          one exception: Should you apply/assign a region to a window, then
'          you must not destroy that region; Windows owns it thereafter

'   About some functions. Not all Region-related functions are available here.
'   Feel free to add others as you find you want them. Some of the functions
'   are one-liners, meaning all they do is call an API and no other processing.
'   Those functions may not be of common-knowledge or not often used. Adding
'   them to this class was just a matter of courtesy vs necessity.

' *** VERSION 4.1 - Included option to exclude a rectangular area of the bitmap
' By excluding a portion of your bitmap that will never have "transparent" pixels
' in it, you can help speed up the process even more. This is simply because the
' routines will just blindly add that rectangular area to the shaped region without
' the need to compare any pixel in that area. The time savings is directly
' proportional to the area of the rectangle & size of the bitmap. The larger
' the bitmap and the larger the exclusion rectangle; the larger the time savings.
' Example: a 200x200 pixel rectangular exception area: 120,000 fewer bytes processed

' GDI32 APIs
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function RectInRgn Lib "gdi32.dll" Alias "RectInRegion" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRectRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function EqualRgn Lib "gdi32.dll" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long

' Kernel32 APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' User32 APIs
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long 'also used in clsBarColors
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 3) As Byte ' used vs RGBQUAD structure
End Type
Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

    
' File attribute enumeration
Public Enum FileAttributes
    attr_Normal = vbNormal
    attr_ReadOnly = vbReadOnly
    attr_Hidden = vbHidden
    attr_System = vbSystem
End Enum
' Options for creating/retrieving a region
Public Enum RegionImportLoc
    rgn_FromWindow = 0
    rgn_FromFile = 1
    rgn_FromResource = 2
    rgn_FromByteArray = 3
End Enum

' Constants used
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const RGN_OR As Long = 2
Private Const OBJ_REGION As Long = 8

' these are erased immediately after use:
Private rgnRects() As RECT  ' array of rectangles comprising region
Private bDib() As Byte      ' the DIB byte array

Public Function AreReqionsEqual(hRgn1 As Long, hRgn2 As Long) As Long

    ' compares 2 regions & returns non-zero if they are the same
    ' Regions are the same if they are the same size & shape
    
    AreReqionsEqual = EqualRgn(hRgn1, hRgn2)
    
End Function

Public Function ConvertRgnToRectRgn(hRgn As Long, Left As Long, Top As Long, Right As Long, Bottom As Long) As Long
    
    ' function converts any existing region to a normal rectangular region
    ' non-zero indicates success.
    
    'The region does not include the lower and right boundaries of the rectangle
    
    If c_isRegion(hRgn) Then
        ConvertRgnToRectRgn = SetRectRgn(hRgn, Top, Left, Right, Bottom)
    End If
    
End Function

Public Function DestroyRegion(hRgn As Long)
    
    ' Simply destroy an existing region
    If c_isRegion(hRgn) Then DeleteObject hRgn
    
End Function

Public Function ExtractRegionRectangles(hRgn As Long, Optional nrRectangles As Long) As Byte()

' This function extracts the rectangles within the region
' Should you need the function, you will need to use a bit of memory shifting
' since classes don't allow returning of UDTs (i.e., RECT structures)
' Use the following when function returns to your control

' Create 2 arrays first:
' Dim rgnRects() As RECT, rgnBytes() As Byte
' Create a long variable (Optional)
' Dim nrRects As Long
' Call function:
' rgnBytes = ExtractRegionRectangles(hRgn)
' Ensure array returned ok:
' If UBound(rgnBytes) > -1 Then
'       Resize your Rect array...
'       If you used the optional nrRects variable then
'           ReDim rgnRects(0 To nrRects - 1)
'       Otherwise, do a little math...
'           ReDim rgnRects(0 To (UBound(rgnBytes) + 1) \ 16)
'       Now copy the bytes into your rectangle array
'       CopyMemory rgnRects(0), rgnBytes(0), UBound(rgnBytes) + 1
'       Erase rgnBytes()
' End If


    Dim b() As Byte
    Dim nBytes As Long
    
    ' get nr bytes in the array
    b() = GetRegionBytes(hRgn)
    nBytes = UBound(b) + 1
    If nBytes > 0 Then
        nrRectangles = (nBytes - 32) \ 16
        CopyMemory b(0), b(32), nBytes - 32
        ReDim Preserve b(0 To nBytes - 33)
    End If
    
    ExtractRegionRectangles = b()

End Function

Public Function GetRegionBytes(hRgn As Long, Optional bDestroyRegion As Boolean) As Byte()

' function returns actual bytes of a region structure
' if UBound() of function < 1 then function failed

    Dim nBytes As Long
    Dim b() As Byte
    
    If c_isRegion(hRgn) Then
        ' get nr bytes in the array
        nBytes = GetRegionData(hRgn, 0, ByVal 0&)
        If nBytes > 16 Then
            ' size our array to retrieve the bytes
             ReDim b(0 To nBytes - 1) As Byte
             ' call function again, filling the array
             If nBytes <> GetRegionData(hRgn, nBytes, b(0)) Then ReDim b(-1 To -1)
        Else
            ReDim b(-1 To -1)
        End If
        If bDestroyRegion Then DeleteObject hRgn
    Else
        ReDim b(-1 To -1)
    End If
    
    GetRegionBytes = b()
    
End Function

Public Function ImportRegion(RgnType As RegionImportLoc, RgnLocation As Variant, Optional ResourceID As Long) As Long

' This function retrieves/creates a region from one of 4 specific locations
' If function returns zero, then no region was created

' 1. From another window
'    RgnType = rgn_FromWindow
'    RgnLocation = hWnd of the window containing the region
'    ResourceID not used

' 2. From a file that contains region data
'    RgnType = rgn_FromFile
'    RgnLocation = the existing filename containing the data
'    ResourceID not used

' 3. From a .RES file's resource
'    RgnType = rgn_FromResource
'    RgnLocation = name of resource section, usually "Custom"
'    ResourceID = The numerical ID of the data in that custom resource section

' 4. From a byte array you pass
'    RgnType = rgn_FromByteArray
'    RgnLocation = the byte array
'    ResourceID not used

Dim hRgn As Long
Dim rgnBytes() As Byte

On Error Resume Next

Select Case RgnType

Case rgn_FromWindow
    Dim hWnd As Long
    hWnd = Val(RgnLocation)
    If hWnd Then
        If IsWindow(hWnd) Then
            hRgn = CreateRectRgn(0, 0, 0, 0)
            If GetWindowRgn(hWnd, hRgn) = 0 Then
                DeleteObject hRgn
                hRgn = 0
            End If
        End If
    End If
    
Case rgn_FromResource
    rgnBytes() = LoadResData(ResourceID, RgnLocation)
    If Not Err Then hRgn = ExtCreateRegion(ByVal 0&, UBound(rgnBytes) + 1, rgnBytes(0))
    
Case rgn_FromByteArray
    If (VarType(RgnLocation) And (vbArray Or vbByte)) > 0 Then
        rgnBytes = RgnLocation
        If Not Err Then hRgn = ExtCreateRegion(ByVal 0&, UBound(rgnBytes) + 1, rgnBytes(0))
    End If

Case rgn_FromFile
    Dim fnr As Integer
    If Len(Dir$(RgnLocation, vbHidden Or vbArchive Or vbReadOnly Or vbSystem)) > 0 Then
        If Not Err Then
            fnr = FreeFile()
            Open RgnLocation For Binary Access Read As #fnr
            If Not Err Then
                ReDim rgnBytes(0 To LOF(fnr) - 1)
                Get #fnr, , rgnBytes()
                hRgn = ExtCreateRegion(ByVal 0&, UBound(rgnBytes) + 1, rgnBytes(0))
            End If
            Close #fnr
        End If
    End If
    
Case Else
    Exit Function
    
End Select

If Err Then
    Err.Clear
    If hRgn Then DeleteObject hRgn
    hRgn = 0
End If

ImportRegion = hRgn

End Function

Public Function MoveRegion(hRgn As Long, x As Long, y As Long) As Boolean

    ' this function moves a region to another X,Y coordinate
    ' return value indicates success
    Dim rgnRect As RECT
    If c_isRegion(hRgn) Then
        GetRgnBox hRgn, rgnRect
        OffsetRgn hRgn, -rgnRect.Left + x, -rgnRect.Top + y
        MoveRegion = True
    End If
    
End Function

Public Function PointInRegion(hRgn As Long, x As Long, y As Long) As Long
    ' here we are not calling c_isRegion first. If this function were called
    ' generally, a very fast return is needed. We'll let windows verify hRgn
    
    PointInRegion = PtInRegion(hRgn, x, y)

End Function

Public Function RectInRegion(hRgn As Long, Left As Long, Top As Long, Right As Long, Bottom As Long) As Long
    ' here we are not calling c_isRegion first. If this function were called
    ' generally, a fast return is needed. We'll let windows verify hRgn
    
    Dim tRect As RECT
    SetRect tRect, Left, Top, Right, Bottom
    RectInRegion = RectInRgn(hRgn, tRect)

End Function

Public Sub RegionBounds(hRgn As Long, Left As Long, Top As Long, Right As Long, Bottom As Long)

    ' this function returns the bounding rectangle of the passed region
    ' by updating the passed parameters
    
    Dim rgnRect As RECT
    If c_isRegion(hRgn) Then GetRgnBox hRgn, rgnRect
    Left = rgnRect.Left
    Top = rgnRect.Top
    Right = rgnRect.Right
    Bottom = rgnRect.Bottom
    
End Sub

Public Function RegionFromBitmap(ByVal hBitmap As Long, _
        Optional hWndToApply As Long, _
        Optional ByVal transColor As Long = -1, _
        Optional returnAntiRegion As Boolean, _
        Optional ByVal ExcludeL As Long, Optional ByVal ExcludeT As Long, _
        Optional ByVal ExcludeR As Long, Optional ByVal ExcludeB As Long) As Long

'*******************************************************
' FUNCTION RETURNS.
' 1. If hWndToApply is zero, the shaped region handle is returned
' 2. Otherwise, a non-zero value indicates region created & applied
'    to the window handle
'*******************************************************

' PARAMETERS
'=============
' hBitmap (Required) : handle to a bitmap to be used to create the region
' Optional hWndToApply : hWnd to assign the shaped region to
' Optional transColor : the transparent color, if -1, then top/left corner is used
' Optional returnAntiRegion : If False (default) then the region excluding transparent
'       pixels will be used/returned.  If True, then the region including only
'       transparent pixels will be used/returned
' Optional ExcludeL, ExcludeT, ExcludeR, ExcludeB are exclusion rectangle coordinates
'   If provided, the pixels in those coordinates will not be examined and
'   that rectangle's area will be added to the shaped region.


' test for required variable first
If hBitmap = 0 Then Exit Function
' if applying to a window, ensure the value passed is a window
If hWndToApply <> 0 Then
    If IsWindow(hWndToApply) = 0 Then Exit Function
End If

' now ensure hBitmap handle passed is a usable bitmap
Dim bmpInfo As BITMAPINFO
If GetGDIObject(hBitmap, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

' declare bunch of variables...
Dim rectCount As Long ' number of rectangles & used to increment above array
Dim lScanLines As Long ' used to size the DIB bit array
Dim rtnRegion As Long ' region handle returned by this function if appropriate
Dim dibDC As Long ' DC to use for GetDIBits

On Error GoTo CleanUp
  
' Scans must align on dword boundaries:
lScanLines = (bmpInfo.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
ReDim bDib(0 To lScanLines - 1, 0 To bmpInfo.bmiHeader.biHeight - 1)

With bmpInfo.bmiHeader
    
    ' validate passed exclusion rectangle dimensions
    If ExcludeL < 0 Then ExcludeL = 0
    If ExcludeT < 0 Then ExcludeT = 0
    If ExcludeB > .biHeight Then ExcludeB = .biHeight
    If ExcludeR > .biWidth Then ExcludeT = .biWidth
    If ExcludeB <= ExcludeT Or ExcludeR <= ExcludeL Then
        ExcludeR = ExcludeL
        ExcludeB = ExcludeT
    End If
   
    ' build the DIB header
   .biSize = Len(bmpInfo.bmiHeader)
   .biBitCount = 24
   .biPlanes = 1
   .biCompression = BI_RGB
   .biHeight = -Abs(.biHeight) ' per msdn the .biHeight may be negative already
    ' ^^ most DIBs are bottom:top, by using negative Height it will load top:bottom
End With

' get the image into DIB bits,
dibDC = GetDC(0)
' note that biHeight above was changed to negative so we reverse it form here on
Call GetDIBits(dibDC, hBitmap, 0, -bmpInfo.bmiHeader.biHeight, bDib(0, 0), bmpInfo, 0)
ReleaseDC 0, dibDC

    
' now calculate the transparent color if needed
With bmpInfo
    If transColor = -1 Then
        ' when negative value passed, use top left corner pixel color
        CopyMemory .bmiColors(0), bDib(0, 0), &H3
    Else
        ' convert vbSystemColor if possible
        If transColor < 0 Then transColor = GetSysColor(transColor And &HFF&)
        ' 24bit DIBs are stored as BGR vs RGB
        ' convert to BGR vs converting each bitmap pixel to RGB
        .bmiColors(2) = (transColor And &HFF&)
        .bmiColors(1) = (transColor And &HFF00&) \ &H100&
        .bmiColors(0) = (transColor And &HFF0000) \ &H10000
    End If
    ' copy bytes to long
    CopyMemory transColor, .bmiColors(0), &H4
End With
    
' Process the bitmap bytes
With bmpInfo.bmiHeader
 
     ' start with an arbritray number of rectangles
    ReDim rgnRects(0 To .biWidth * 3)
        
    If ExcludeR > ExcludeL Then
        ' when an exclusion rectangle is provided, we process the bitmap like so:
        
        ' Do from the top of bitmap to top of exclusion rectangle
        c_ScanSection 0, .biWidth, .biWidth, .biWidth, ExcludeT, transColor, returnAntiRegion, rectCount
        
        ' Do the exclusion rectangle, add estimated number of unexamined rows
        ReDim Preserve rgnRects(0 To UBound(rgnRects) + (ExcludeB - ExcludeT))
        c_ScanSection ExcludeT, ExcludeL, ExcludeR, .biWidth, ExcludeB, transColor, returnAntiRegion, rectCount
        
        ' Do from bottom of the exclusion rectangle to the bottom of the bitmap
        c_ScanSection ExcludeB, .biWidth, .biWidth, .biWidth, Abs(.biHeight), transColor, returnAntiRegion, rectCount

    Else
        ' No exclusion rectangle, process from top to bottom of the bitmap
        c_ScanSection 0, .biWidth, 0, .biWidth, Abs(.biHeight), transColor, returnAntiRegion, rectCount
    End If
End With

Erase bDib  ' no longer needed; we can purge it now
        
On Error Resume Next
' check for failure & engage backup plan if needed
If rectCount Then
    ' there were rectangles identified, try to create the region
    rtnRegion = c_CreatePartialRegion(2, rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
    
    ' ok, now to test whether or not we are good to go...
    ' if less than 2000 rectangles, function should have worked & if it didn't
    ' it wasn't due O/S restrictions -- failure
    
    If rtnRegion = 0 And rectCount > 2000 Then
        rtnRegion = c_CreateWin9xRegion(rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
        ' ^^ if rtnRegion is zero; windows could not create the region
    End If

End If

CleanUp:

Erase rgnRects() ' no longer needed; we can purge it now

If Err Then
    If rtnRegion Then DeleteObject rtnRegion
    Err.Clear
    Debug.Print "Shaped Region failed. Windows could not create the region."
    ' return value for function is zero
Else
    If hWndToApply Then
        RegionFromBitmap = SetWindowRgn(hWndToApply, rtnRegion, True)
        ' ^^ if above API call fails, return value will be zero
    Else
        RegionFromBitmap = rtnRegion
        ' ^^ return handle to the shaped region
    End If
End If


End Function

Public Function RegionSizeBytes(hRgn As Long) As Long

    ' function simply returns the number of bytes in the region structure
    ' The number of individual rectangles is an easy calculation:
    
    ' NrRects = (NrBytes - 32)\16
    ' ^^ The 1st 32 bytes is the region header, where 1st RECT starts on 33d byte
    
    If c_isRegion(hRgn) Then RegionSizeBytes = GetRegionData(hRgn, 0, ByVal 0&)

End Function

Public Function SaveRegionToFile(hRgn As Long, sFileName As String, _
        Optional ByVal fileAttrs As FileAttributes, Optional bDestroyRegion As Boolean) As Long

' function saves a region structure to a filename of your choice.
' Zero return value indicates success, non-zero indicates failure

' hRgn :: region handle to save
' sFileName :: full path & filename of file; path must exist
'   - if file already exists it will be overwritten
' fileAttrs :: optional file attributes for the file after it is written
' bDestroyRegion :: remove region from memory when function exits

Dim iFile As Integer
Dim fAttr As Long
Dim b() As Byte
Dim rtnVal As Long
      
If Not c_isRegion(hRgn) Then
    SaveRegionToFile = -4
    Exit Function
End If
If Len(sFileName) = 0 Then
    SaveRegionToFile = -3
    If hRgn <> 0 And bDestroyRegion = True Then DeleteObject hRgn
    Exit Function
End If
      
' prevent bogus values passed by user
For iFile = 1 To 4
    fAttr = fAttr Or (fileAttrs And iFile)
Next
      
On Error Resume Next
' any errors here & we bug out

If Len(Dir$(sFileName, fAttr)) > 0 Then
    ' drive & path exists?  if not can't continue
    If Err Then
        Err.Clear
        rtnVal = Err.Number
    Else
        ' try to kill existing file; if not can't continue
        Kill sFileName
        If Err Then
            rtnVal = Err.Number
            Err.Clear
        End If
    End If
End If

' if no errors to this point, continue on
If rtnVal = 0 Then
    ' call helper function to get nr of bytes in the region
    b() = GetRegionBytes(hRgn, bDestroyRegion)
    If UBound(b) > -1 Then
        iFile = FreeFile
        Open sFileName For Binary Access Write Lock Read As #iFile
        If Not Err Then
            ' write the region data to the file; hopefully w/o errors
            Put #iFile, , b()
            Close #iFile
        End If
        If Err Then
            rtnVal = Err.Number
            Err.Clear
        End If
        Erase b()
    Else
        rtnVal = -1
    End If
Else
    If bDestroyRegion Then DestroyRegion hRgn
End If


SaveRegionToFile = rtnVal
End Function

Public Function SetClipRegionToDC(hRgn As Long, hdc As Long, Optional bDestroyRgn As Boolean) As Long
    ' this function applies a clipping region to a DC
    ' Passing zero as the hRgn, removes any custom clipping region from the DC
    ' Non-zero return value indicates success
    '   see MSDN:SelectClipRgn for possible return values
    If hdc <> 0 Then
        SetClipRegionToDC = SelectClipRgn(hdc, hRgn)
    End If
    If bDestroyRgn Then DestroyRegion hRgn
    
End Function

Public Function SetRegionToWindow(hRgn As Long, hWnd As Long) As Long
    ' this function applies a region to a window
    ' Passing zero as the hRgn, removes any custom region from the window
    ' non-zero return value indicates success
    If hWnd <> 0 Then
        If IsWindow(hWnd) Then SetRegionToWindow = SetWindowRgn(hWnd, hRgn, True)
    End If
End Function

Public Function StretchRegion(hRgn As Long, GrowCx As Single, GrowCy As Single, bDestroyOriginalRgn As Boolean) As Long
    
    ' function returns a stretched region if successful
    
    ' To stretch a region, supply the GrowCx and GrowCy with
    ' decimal values where 1 is the same size as current
    ' values btwn .1 and .99 reduce size and
    ' values btwn 1.1 and > increase size
    
    If c_isRegion(hRgn) Then
        Dim nBytes As Long, rtnRgn As Long
        Dim xFrm As XFORM
        
        ' apply the XForm values (these are only XForm members compatible with all O/S)
        xFrm.eM11 = GrowCx
        xFrm.eM22 = GrowCy
        
        ' thru trial and error it is actually faster to break the region apart
        ' and do it in sections than to send a large region to be done at once.
        ' The time savings can be literally several seconds...
        
        ' Get the number of bytes in the region
        nBytes = GetRegionData(hRgn, 0, ByVal 0)
        
        If nBytes > 32 Then
            ' resize the rectangle array to receive the bytes
            ReDim rgnRects(0 To nBytes \ 16 - 1)
            ' get the bytes
            If nBytes = GetRegionData(hRgn, nBytes, rgnRects(0)) Then
                ' now step thru the stretching 2K at a time; very fast
                rtnRgn = c_CreateWin9xRegion(UBound(rgnRects), rgnRects(0).Left, rgnRects(0).Right, VarPtr(xFrm))
                StretchRegion = rtnRgn
            End If
        End If
    End If
    
    If bDestroyOriginalRgn Then DestroyRegion hRgn
    
End Function



' /////////////////////////  CLASS ONLY FUNCTIONS \\\\\\\\\\\\\\\\\\\\
' all non-public class functions/routines are prefixed with c_ in their name

Private Function c_CreatePartialRegion(lIndex As Long, uIndex As Long, leftOffset As Long, Cx As Long, Optional ByVal xFrmPtr As Long) As Long
' Called when large region fails (can be the case with Win98) and also called
' when rotation a region 90 or 270 degrees (see RotateSimpleRegion)

On Error Resume Next
' Note: Ideally contiguous rectangles of equal height & width should be combined
' into one larger rectangle. However, thru trial & error I found that Windows
' does this for us and taking the extra time to do it ourselves
' is to cumbersome & slows down the results.

' the first 32 bytes of a region is the header describing the region.
' Well 32 bytes equates to 2 rectangles (16 bytes each), so I'll
' cheat a little & use rectangles to store the header
With rgnRects(lIndex - 2) ' bytes 0-15
    .Left = 32                      ' length of region header in bytes
    .Top = 1                        ' required cannot be anything else
    .Right = uIndex - lIndex + 1    ' number of rectangles for the region
    .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
End With
With rgnRects(lIndex - 1) ' bytes 16-31 bounding rectangle identification
    .Left = leftOffset                  ' left
    .Top = rgnRects(lIndex).Top         ' top
    .Right = leftOffset + Cx            ' right
    .Bottom = rgnRects(uIndex).Bottom   ' bottom
End With
' call function to create region from our byte (RECT) array
c_CreatePartialRegion = ExtCreateRegion(ByVal xFrmPtr, (rgnRects(lIndex - 2).Right + 2) * 16, rgnRects(lIndex - 2))
If Err Then Err.Clear
End Function

Private Function c_CreateWin9xRegion(rectCount As Long, leftOffset As Long, Cx As Long, Optional ByVal xFrmPtr As Long) As Long
' Pulled out of main routine 'cause now two routines use the same logic
' and we will simply share this part of the code

' Win98 has problems with regional rectangles over 4000
' So, we'll try again in case this is the prob with other systems too.
' We'll step it at 2000 at a time which is stil very quick

Dim x As Long, y As Long ' loop counters
Dim win9xRgn As Long     ' partial region
Dim rtnRegion As Long    ' combined region & return value of this function

' we start with 2 'cause first 2 RECTs is the header
For x = 2 To rectCount Step 2000

    If x + 2000 > rectCount Then
        y = rectCount
    Else
        y = x + 2000
    End If
    
    ' attempt to create partial region
    win9xRgn = c_CreatePartialRegion(x, y, leftOffset, Cx, xFrmPtr)
    
    If win9xRgn = 0 Then    ' failure
        ' cleaup combined region if needed
        If rtnRegion Then DeleteObject rtnRegion
        rtnRegion = 0
        Exit For ' abort
    Else
        If rtnRegion Then ' already started
            ' use combineRgn, but only every 2000th time
            CombineRgn rtnRegion, rtnRegion, win9xRgn, RGN_OR
            DeleteObject win9xRgn
        Else    ' first time thru
            rtnRegion = win9xRgn
        End If
    End If
Next
' done; return result
c_CreateWin9xRegion = rtnRegion
End Function

Private Function c_isRegion(hRgn As Long) As Boolean
    ' returns whether or not the passed pointer is a region
    If hRgn Then c_isRegion = (GetObjectType(hRgn) = OBJ_REGION)

End Function

Private Sub c_ScanSection(y As Long, Cx As Long, Cx2 As Long, Cx3 As Long, Cy As Long, _
        transColor As Long, returnAntiRegion As Boolean, rectCount As Long)

' function added as an extension of the main routine; therefore, many of the
' user passed variables are also passed here too

' This function will scan a specific number of bitmap lines and can skip over
' a continguous section of pixels identified by the exclusion rectangle.
' See the CreateShapedRegion2 remarks.

    Dim scanX As Long, scanY As Long, x As Long ' simple loop variables
    Dim tgtColor As Long                ' a DIB pixel color (BGR format)
    Dim rStart As Long                  ' rectangle tracking started
    Dim bExludeDone As Boolean          ' exclusion rectangle processed
    Dim endX As Long                    ' last byte on line to scan

    ' reset flag
    rStart = -1
    endX = Cx
    
    ' begin pixel by pixel comparisons
    For scanY = y To Cy - 1
        For scanX = x To endX - 1
            ' my hack continued: we already saved a long as BGR, now
            ' get the current DIB pixel into a long (BGR also) & compare
            CopyMemory tgtColor, bDib(scanX * 3, scanY), &H3
            
            ' test to see if next pixel is a target color
            If transColor = tgtColor Xor returnAntiRegion Then
                
                If rStart > -1 Then ' we're currently tracking a rectangle,
                                    ' so let's close it
                    ' see if array needs to be resized
                   If rectCount + 1 = UBound(rgnRects) Then _
                       ReDim Preserve rgnRects(0 To UBound(rgnRects) + Cx3)
                    
                    ' add the rectangle to our array
                    SetRect rgnRects(rectCount + 2), rStart, scanY, scanX, scanY + 1
                    rStart = -1 ' reset flag
                    rectCount = rectCount + 1     ' keep track of nr in use
                End If
            
            Else
                ' not a target color
                If rStart = -1 Then rStart = scanX ' set start point
            
            End If
        Next scanX
        If rStart > -1 Then
            ' got to end of section without hitting another transparent pixel
            ' but we're tracking so we'll close rectangle now
           
                ' see if array needs to be resized
           If rectCount + 1 = UBound(rgnRects) Then _
               ReDim Preserve rgnRects(0 To UBound(rgnRects) + Cx3)
                ' add the rectangle to our array
            SetRect rgnRects(rectCount + 2), rStart, scanY, scanX, scanY + 1
            rStart = -1 ' reset flag
            rectCount = rectCount + 1     ' keep track of nr in use
        End If
        
        ' exclusion rectangle checks here
        If bExludeDone Then
            ' when the left & right side of exclusion rect finished, reset variables
            bExludeDone = False
            x = 0
            endX = Cx
        Else
            ' see if we have an exclusion rectangle being processed
            If Cx2 > 0 Then             ' no exclusion rectangle passed
                If Cx2 <> Cx3 Then      ' no exclusion rectangle passed
                    
                    ' add the excluded portion of the bitmap to our region
                    SetRect rgnRects(rectCount + 2), Cx, scanY, Cx2 + 1, scanY + 1
                    rectCount = rectCount + 1
                    ' set the next row of pixels to be checked (right side of rectangle)
                    x = Cx2
                    ' set the last row to be checked (left edge of bitmap)
                    endX = Cx3
                    ' decrement the loop variable so it will run same row again
                    scanY = scanY - 1
                    ' toggle flag so X & endX can be restored after row is processed
                    bExludeDone = True
                End If
            End If
        End If
        
    Next scanY

End Sub



