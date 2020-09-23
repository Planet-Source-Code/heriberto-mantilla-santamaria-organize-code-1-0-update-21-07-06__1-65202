Attribute VB_Name = "vbunzip"
Option Explicit

 '-- Please Do Not Remove These Comment Lines!
 '----------------------------------------------------------------
 '-- Sample VB 5 code to drive unzip32.dll
 '-- Contributed to the Info-ZIP project by Mike Le Voi
 '--
 '-- Contact me at: mlevoi@modemss.brisnet.org.au
 '--
 '-- Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
 '--
 '-- Use this code at your own risk. Nothing implied or warranted
 '-- to work on your machine :-)
 '----------------------------------------------------------------
 '--
 '-- This Source Code Is Freely Available From The Info-ZIP Project
 '-- Web Server At:
 '-- ftp://ftp.info-zip.org/pub/infozip/infozip.html
 '--
 '-- A Very Special Thanks To Mr. Mike Le Voi
 '-- And Mr. Mike White
 '-- And The Fine People Of The Info-ZIP Group
 '-- For Letting Me Use And Modify Their Original
 '-- Visual Basic 5.0 Code! Thank You Mike Le Voi.
 '-- For Your Hard Work In Helping Me Get This To Work!!!
 '---------------------------------------------------------------
 '--
 '-- Contributed To The Info-ZIP Project By Raymond L. King.
 '-- Modified June 21, 1998
 '-- By Raymond L. King
 '-- Custom Software Designers
 '--
 '-- Contact Me At: king@ntplx.net
 '-- ICQ 434355
 '-- Or Visit Our Home Page At: http://www.ntplx.net/~king
 '--
 '---------------------------------------------------------------
 '--
 '-- Modified August 17, 1998
 '-- by Christian Spieler
 '-- (implemented sort of a "real" user interface)
 '--
 '---------------------------------------------------------------
 
 '-- C Style argv
 Private Type UNZIPnames
  uzFiles(0 To 4000) As String
 End Type
 
 '-- Callback Large "String"
 Private Type UNZIPCBChar
  ch(32800) As Byte
 End Type
 
 '-- Callback Small "String"
 Private Type UNZIPCBCh
  ch(256) As Byte
 End Type
 
 '-- UNZIP32.DLL DCL Structure
 Private Type DCLIST
  ExtractOnlyNewer  As Long    ' 1 = Extract Only Newer/New, Else 0
  SpaceToUnderscore As Long    ' 1 = Convert Space To Underscore, Else 0
  PromptToOverwrite As Long    ' 1 = Prompt To Overwrite Required, Else 0
  fQuiet            As Long    ' 2 = No Messages, 1 = Less, 0 = All
  ncflag            As Long    ' 1 = Write To Stdout, Else 0
  ntflag            As Long    ' 1 = Test Zip File, Else 0
  nvflag            As Long    ' 0 = Extract, 1 = List Zip Contents
  nfflag            As Long    ' 1 = Extract Only Newer Over Existing, Else 0
  nzflag            As Long    ' 1 = Display Zip File Comment, Else 0
  ndflag            As Long    ' 1 = Honor Directories, Else 0
  noflag            As Long    ' 1 = Overwrite Files, Else 0
  naflag            As Long    ' 1 = Convert CR To CRLF, Else 0
  nZIflag           As Long    ' 1 = Zip Info Verbose, Else 0
  C_flag            As Long    ' 1 = Case Insensitivity, 0 = Case Sensitivity
  fPrivilege        As Long    ' 1 = ACL, 2 = Privileges
  Zip               As String  ' The Zip Filename To Extract Files
  ExtractDir        As String  ' The Extraction Directory, NULL If Extracting To Current Dir
 End Type
 
 '-- UNZIP32.DLL Userfunctions Structure
 Private Type USERFUNCTION
  UZDLLPrnt     As Long     ' Pointer To Apps Print Function
  UZDLLSND      As Long     ' Pointer To Apps Sound Function
  UZDLLREPLACE  As Long     ' Pointer To Apps Replace Function
  UZDLLPASSWORD As Long     ' Pointer To Apps Password Function
  UZDLLMESSAGE  As Long     ' Pointer To Apps Message Function
  UZDLLSERVICE  As Long     ' Pointer To Apps Service Function (Not Coded!)
  TotalSizeComp As Long     ' Total Size Of Zip Archive
  TotalSize     As Long     ' Total Size Of All Files In Archive
  CompFactor    As Long     ' Compression Factor
  NumMembers    As Long     ' Total Number Of All Files In The Archive
  cchComment    As Integer  ' Flag If Archive Has A Comment!
 End Type

 '-- UNZIP32.DLL Version Structure
 Private Type UZPVER
  StructLen       As Long         ' Length Of The Structure Being Passed
  Flag            As Long         ' Bit 0: is_beta  bit 1: uses_zlib
  Beta            As String * 10  ' e.g., "g BETA" or ""
  Date            As String * 20  ' e.g., "4 Sep 95" (beta) or "4 September 1995"
  zLib            As String * 10  ' e.g., "1.0.5" or NULL
  UnZip(1 To 4)   As Byte         ' Version Type Unzip
  ZipInfo(1 To 4) As Byte         ' Version Type Zip Info
  os2dll          As Long         ' Version Type OS2 DLL
  Windll(1 To 4)  As Byte         ' Version Type Windows DLL
 End Type

 '-- This Assumes UNZIP32.DLL Is In Your \Windows\System Directory!
 Private Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, ByVal xfnc As Long, ByRef xfnv As UNZIPnames, dcll As DCLIST, Userf As USERFUNCTION) As Long

 Private Declare Sub UzpVersion2 Lib "unzip32.dll" (uzpv As UZPVER)

 '-- Private Variables For Structure Access
 Private UZDCL  As DCLIST
 Private UZUSER As USERFUNCTION
 Private UZVER  As UZPVER

 '-- Public Variables For Setting The
 '-- UNZIP32.DLL DCLIST Structure
 '-- These Must Be Set Before The Actual Call To VBUnZip32
 Public uExtractOnlyNewer As Integer  ' 1 = Extract Only Newer/New, Else 0
 Public uSpaceUnderScore  As Integer  ' 1 = Convert Space To Underscore, Else 0
 Public uPromptOverWrite  As Integer  ' 1 = Prompt To Overwrite Required, Else 0
 Public uQuiet            As Integer  ' 2 = No Messages, 1 = Less, 0 = All
 Public uWriteStdOut      As Integer  ' 1 = Write To Stdout, Else 0
 Public uTestZip          As Integer  ' 1 = Test Zip File, Else 0
 Public uExtractList      As Integer  ' 0 = Extract, 1 = List Contents
 Public uFreshenExisting  As Integer  ' 1 = Update Existing by Newer, Else 0
 Public uDisplayComment   As Integer  ' 1 = Display Zip File Comment, Else 0
 Public uHonorDirectories As Integer  ' 1 = Honor Directories, Else 0
 Public uOverWriteFiles   As Integer  ' 1 = Overwrite Files, Else 0
 Public uConvertCR_CRLF   As Integer  ' 1 = Convert CR To CRLF, Else 0
 Public uVerbose          As Integer  ' 1 = Zip Info Verbose
 Public uCaseSensitivity  As Integer  ' 1 = Case Insensitivity, 0 = Case Sensitivity
 Public uPrivilege        As Integer  ' 1 = ACL, 2 = Privileges, Else 0
 Public uZipFileName      As String   ' The Zip File Name
 Public uExtractDir       As String   ' Extraction Directory, Null If Current Directory

 '-- Public Program Variables
 Public uZipNumber        As Long       ' Zip File Number
 Public uNumberFiles      As Long       ' Number Of Files
 Public uNumberXFiles     As Long       ' Number Of Extracted Files
 Public uZipMessage       As String     ' For Zip Message
 Public uZipInfo          As String     ' For Zip Information
 Public uZipNames         As UNZIPnames ' Names Of Files To Unzip
 Public uExcludeNames     As UNZIPnames ' Names Of Zip Files To Exclude
 Public uVbSkip           As Integer    ' For DLL Password Function
 
 Public CompressedSize(4000)     As String  ' Size of a file compressed
 Public UncompressedSize(4000)   As String  ' Size of a file uncompressed
 Public CompressedDateTime(4000) As String  ' Date and time of a file when compressed
 Public CompressedRatio(4000)    As String  ' Percent ratio of a file that has been compressed
 Public CompressedFileName(4000) As String  ' Name of a compressed file in a zip
 Public CompressedTotal          As Integer ' Total number of files in zip
 Public CompressedPath(4000)     As String  ' Path of a file in a zip
 Public CompressedFileType(4000) As String  ' Type of file
 Public TotalUncompressedZipSize As String  ' Total uncompressed size of all files in a zip
 Public TotalCompressedZipSize   As String  ' Total compressed size of zip file itself
 Public ZipCompressFactor        As String
 Public RetCode                  As Long
 
 Public Const HKEY_CLASSES_ROOT = &H80000000
 
 Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
 Private Const KEY_CREATE_LINK As Long = &H20
 Private Const KEY_CREATE_SUB_KEY As Long = &H4
 Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
 Private Const KEY_NOTIFY As Long = &H10
 Private Const KEY_QUERY_VALUE As Long = &H1
 Private Const KEY_SET_VALUE As Long = &H2
 Private Const SYNCHRONIZE As Long = &H100000
 Private Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
 Private Const ERROR_SUCCESS = 0&
 Private Const REG_SZ = 1
 
 '* ADVAPI32.
 Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
 Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
 Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

'-- Puts A Function Pointer In A Structure
'-- For Callbacks.
Public Function FnPtr(ByVal lp As Long) As Long
 FnPtr = lp
End Function

'-- Callback For UNZIP32.DLL - Receive Message Function
Public Sub UZReceiveDLLMessage(ByVal ucSize As Long, ByVal cSiz As Long, ByVal cFactor As Integer, ByVal mo As Integer, ByVal dy As Integer, ByVal yr As Integer, ByVal hh As Integer, ByVal mm As Integer, ByVal c As Byte, ByRef fname As UNZIPCBCh, ByRef meth As UNZIPCBCh, ByVal crc As Long, ByVal fCrypt As Byte)
 Dim s0     As String, xx      As Long, A As Integer
 Dim strOut As String * 80, ab As Integer

 '-- Always Put This In Callback Routines!
On Error Resume Next
 '------------------------------------------------
 '-- This Is Where The Received Messages Are
 '-- Printed Out And Displayed.
 '-- You Can Modify Below!
 '------------------------------------------------
 strOut = Space$(80)
 '-- For Zip Message Printing
 CompressedTotal = CompressedTotal + 1
 For xx = 0 To 255
  If (fname.ch(xx) = 0) Then Exit For
  s0 = s0 & Chr$(fname.ch(xx))
 Next
 'Takes a full file specification and returns the filename
 For ab = Len(s0) To 1 Step -1
  If (Mid$(s0, ab, 1) = "\" Or Mid$(s0, ab, 1) = "/") Then
   CompressedFileName(CompressedTotal) = Mid$(s0, ab + 1)
   Exit For
  Else
   ' fixs filenames when there is no directory included in zip
   CompressedFileName(CompressedTotal) = s0
   Exit For
  End If
 Next
 CompressedSize(CompressedTotal) = Format$(cSiz, "###,###,###,###")
 UncompressedSize(CompressedTotal) = Format$(ucSize, "###,###,###,###")
 CompressedDateTime(CompressedTotal) = Right$("0" & Trim$(CStr(mo)), 2) & "/" & Right$("0" & Trim$(CStr(dy)), 2) & "/" & Right$("0" & Trim$(CStr(yr)), 2) & " - " & Right$(Str$(hh), 2) & ":" & Right$("0" & Trim$(CStr(mm)), 2)
 If (ucSize <> 0) Then
  CompressedRatio(CompressedTotal) = Format$(CInt((1 - (cSiz / ucSize)) * 100)) & "%"
 Else
  CompressedRatio(CompressedTotal) = "0%"
 End If
 'Takes a full file specification and returns the path
 For A = Len(s0) To 1 Step -1
  If (Mid$(s0, A, 1) = "\") Or (Mid$(s0, A, 1) = "/") Then
   'Add the correct path separator for the input
   If (Mid$(s0, A, 1) = "\") Then
    CompressedPath(CompressedTotal) = LCase$(Left$(s0, A - 1) & "\")
    Exit For
   Else
    CompressedPath(CompressedTotal) = LCase$(Left$(s0, A - 1) & "/")
    Exit For
   End If
  End If
 Next
 '-- Do Not Modify Below!!!
 uZipMessage = uZipMessage & strOut & vbNewLine
 uZipNumber = uZipNumber + 1
End Sub

'-- Callback For UNZIP32.DLL - Print Message Function
Public Function UZDLLPrnt(ByRef fname As UNZIPCBChar, ByVal X As Long) As Long
 Dim s0 As String, xx As Long

 '-- Always Put This In Callback Routines!
On Error Resume Next
 s0 = ""
 '-- Gets The UNZIP32.DLL Message For Displaying.
 For xx = 0 To X - 1
  If (fname.ch(xx) = 0) Then Exit For
  s0 = s0 & Chr$(fname.ch(xx))
 Next
 '-- Assign Zip Information
 If (Mid$(s0, 1, 1) = vbLf) Then s0 = vbNewLine  ' Damn UNIX :-)
 uZipInfo = uZipInfo & s0
 UZDLLPrnt = 0
End Function

'-- Callback For UNZIP32.DLL - DLL Service Function
Public Function UZDLLServ(ByRef mname As UNZIPCBChar, ByVal X As Long) As Long
 Dim s0 As String, xx As Long

 '-- Always Put This In Callback Routines!
On Error Resume Next
 s0 = ""
 '-- Get Zip32.DLL Message For processing
 For xx = 0 To X - 1
  If (mname.ch(xx) = 0) Then Exit For
  s0 = s0 & Chr$(mname.ch(xx))
 Next
 ' At this point, s0 contains the message passed from the DLL
 ' It is up to the developer to code something useful here :)
 UZDLLServ = 0 ' Setting this to 1 will abort the zip!
End Function

'-- Callback For UNZIP32.DLL - Password Function
Public Function UZDLLPass(ByRef p As UNZIPCBCh, ByVal n As Long, ByRef m As UNZIPCBCh, ByRef Name As UNZIPCBCh) As Integer
 Dim Prompt     As String, xx As Integer
 Dim SzPassword As String

 '-- Always Put This In Callback Routines!
On Error Resume Next
 UZDLLPass = 1
 If (uVbSkip = 1) Then Exit Function
 '-- Get The Zip File Password
 SzPassword = InputBox$("Please Enter The Password!")
 '-- No Password So Exit The Function
 If (Len(SzPassword) = 0) Then
  uVbSkip = 1
  Exit Function
 End If
 '-- Zip File Password So Process It
 For xx = 0 To 255
  If (m.ch(xx) = 0) Then
   Exit For
  Else
   Prompt = Prompt & Chr$(m.ch(xx))
  End If
 Next
 For xx = 0 To n - 1
  p.ch(xx) = 0
 Next
 For xx = 0 To Len(SzPassword) - 1
  p.ch(xx) = Asc(Mid$(SzPassword, xx + 1, 1))
 Next
 p.ch(xx) = 0 ' Put Null Terminator For C
 UZDLLPass = 0
End Function

'-- Callback For UNZIP32.DLL - Report Function To Overwrite Files.
'-- This Function Will Display A MsgBox Asking The User
'-- If They Would Like To Overwrite The Files.
Public Function UZDLLRep(ByRef fname As UNZIPCBChar) As Long
 Dim s0 As String, xx As Long

 '-- Always Put This In Callback Routines!
On Error Resume Next
 UZDLLRep = 100 ' 100 = Do Not Overwrite - Keep Asking User
 s0 = ""
 For xx = 0 To 255
  If (fname.ch(xx) = 0) Then Exit For
  s0 = s0 & Chr$(fname.ch(xx))
 Next
 '-- This Is The MsgBox Code
 xx = MsgBox("Overwrite " & s0 & "?", vbYesNoCancel, "VBUnZip32 - File Already Exists!")
 If (xx = vbNo) Then Exit Function
 If (xx = vbCancel) Then
  UZDLLRep = 104       ' 104 = Overwrite None
  Exit Function
 End If
 UZDLLRep = 102         ' 102 = Overwrite, 103 = Overwrite All
End Function

'-- ASCIIZ To String Function
Public Function szTrim(szString As String) As String
 Dim Pos As Long

 Pos = InStr(szString, vbNullChar)
 Select Case Pos
  Case Is > 1
   szTrim = Trim$(Left$(szString, Pos - 1))
  Case 1
   szTrim = ""
  Case Else
   szTrim = Trim$(szString)
 End Select
End Function

'-- Main UNZIP32.DLL UnZip32 Subroutine
'-- (WARNING!) Do Not Change!
Public Sub VBUnZip32()
 Dim MsgStr As String

 '-- Set The UNZIP32.DLL Options
 '-- (WARNING!) Do Not Change
 UZDCL.ExtractOnlyNewer = uExtractOnlyNewer ' 1 = Extract Only Newer/New
 UZDCL.SpaceToUnderscore = uSpaceUnderScore ' 1 = Convert Space To Underscore
 UZDCL.PromptToOverwrite = uPromptOverWrite ' 1 = Prompt To Overwrite Required
 UZDCL.fQuiet = uQuiet                      ' 2 = No Messages 1 = Less 0 = All
 UZDCL.ncflag = uWriteStdOut                ' 1 = Write To Stdout
 UZDCL.ntflag = uTestZip                    ' 1 = Test Zip File
 UZDCL.nvflag = uExtractList                ' 0 = Extract 1 = List Contents
 UZDCL.nfflag = uFreshenExisting            ' 1 = Update Existing by Newer
 UZDCL.nzflag = uDisplayComment             ' 1 = Display Zip File Comment
 UZDCL.ndflag = uHonorDirectories           ' 1 = Honour Directories
 UZDCL.noflag = uOverWriteFiles             ' 1 = Overwrite Files
 UZDCL.naflag = uConvertCR_CRLF             ' 1 = Convert CR To CRLF
 UZDCL.nZIflag = uVerbose                   ' 1 = Zip Info Verbose
 UZDCL.C_flag = uCaseSensitivity            ' 1 = Case insensitivity, 0 = Case Sensitivity
 UZDCL.fPrivilege = uPrivilege              ' 1 = ACL 2 = Priv
 UZDCL.Zip = uZipFileName                   ' ZIP Filename
 UZDCL.ExtractDir = uExtractDir             ' Extraction Directory, NULL If Extracting
 ' To Current Directory
 '-- Set Callback Addresses
 '-- (WARNING!!!) Do Not Change
 UZUSER.UZDLLPrnt = FnPtr(AddressOf UZDLLPrnt)
 UZUSER.UZDLLSND = 0&    '-- Not Supported
 UZUSER.UZDLLREPLACE = FnPtr(AddressOf UZDLLRep)
 UZUSER.UZDLLPASSWORD = FnPtr(AddressOf UZDLLPass)
 UZUSER.UZDLLMESSAGE = FnPtr(AddressOf UZReceiveDLLMessage)
 UZUSER.UZDLLSERVICE = FnPtr(AddressOf UZDLLServ)
 '-- Set UNZIP32.DLL Version Space
 '-- (WARNING!!!) Do Not Change
 With UZVER
  .StructLen = Len(UZVER)
  .Beta = Space$(9) & vbNullChar
  .Date = Space$(19) & vbNullChar
  .zLib = Space$(9) & vbNullChar
 End With
 '-- Get Version
 Call UzpVersion2(UZVER)
 '-- Go UnZip The Files! (Do Not Change Below!!!)
 '-- This Is The Actual UnZip Routine
 RetCode = Wiz_SingleEntryUnzip(uNumberFiles, uZipNames, uNumberXFiles, uExcludeNames, UZDCL, UZUSER)
 '---------------------------------------------------------------
 TotalUncompressedZipSize = Format$(UZUSER.TotalSize, "###,###,###,###")
 TotalCompressedZipSize = Format$(UZUSER.TotalSizeComp, "###,###,###,###")
 ZipCompressFactor = UZUSER.CompFactor & "%"
End Sub

Public Function GetKey(hKey As Long, ByVal sKeyName, ByVal lValueName)
 Dim retval$, hSubKey As Long, dwType As Long, SZ As Long, v$, r As Long
 
 retval$ = ""
 r = RegOpenKeyEx(hKey, sKeyName, 0, KEY_ALL_ACCESS, hSubKey)
 If (r <> ERROR_SUCCESS) Then GoTo Quit_Now
 SZ = 256
 v$ = String$(SZ, 0)
 r = RegQueryValueEx(hSubKey, lValueName, 0, dwType, ByVal v$, SZ)
 If (r = ERROR_SUCCESS) And (dwType = REG_SZ) Then
  retval$ = Left$(v$, SZ - 1)
 Else
  retval$ = ""
 End If
 If (hKey = 0) Then r = RegCloseKey(hSubKey)
Quit_Now:
 GetKey = retval$
End Function

Public Sub ErrorHandler()
 Select Case RetCode
  Case 2
   Call MsgBox("Error Code: 2 - Unexpected End of Zip File Error.", vbCritical + vbOKOnly, Ttl)
  Case 3
   Call MsgBox("Error Code: 3 - Zip File Structure Error.", vbCritical + vbOKOnly, Ttl)
  Case 4
   Call MsgBox("Error Code: 4 - Out of Memory Error.", vbCritical + vbOKOnly, Ttl)
  Case 5
   Call MsgBox("Error Code: 5 - Internal Logic Error in Zip dll.", vbCritical + vbOKOnly, Ttl)
  Case 6
   Call MsgBox("Error Code: 6 - Entry Too Large to Split Error.", vbCritical + vbOKOnly, Ttl)
  Case 7
   Call MsgBox("Error Code: 7 - Invalid Comment Format Error.", vbCritical + vbOKOnly, Ttl)
  Case 8
   Call MsgBox("Error Code: 8 - Zip Test Failed or Out of Memory Error.", vbCritical + vbOKOnly, Ttl)
  Case 9
   Call MsgBox("Error Code: 9 - User Interrupted or Termination Error.", vbCritical + vbOKOnly, Ttl)
  Case 10
   Call MsgBox("Error Code: 10 - Error Using a Temp File.", vbCritical + vbOKOnly, Ttl)
  Case 11
   Call MsgBox("Error Code: 11 - Read or Seek Error.", vbCritical + vbOKOnly, Ttl)
  Case 12
   Call MsgBox("Error Code: 12 - Nothing to do Error.", vbCritical + vbOKOnly, Ttl)
  Case 13
   Call MsgBox("Error Code: 13 - Missing or Empty Zip File Error.", vbCritical + vbOKOnly, Ttl)
  Case 14
   Call MsgBox("Error Code: 14 - Error Writing to a File.", vbCritical + vbOKOnly, Ttl)
  Case 15
   Call MsgBox("Error Code: 15 - Couldn't Open to Write Error.", vbCritical + vbOKOnly, Ttl)
  Case 16
   Call MsgBox("Error Code: 16 - Bad Command Line Argument Error.", vbCritical + vbOKOnly, Ttl)
  Case 18
   Call MsgBox("Error Code: 18 - Could Not Open a Specified File.", vbCritical + vbOKOnly, Ttl)
 End Select
End Sub
