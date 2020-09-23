Attribute VB_Name = "modVar"
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit
 
 Public Const Ttl = "Organize Code 1.0"
 
 Public Datos     As ADODB.Connection '* Abre la Base de Datos.
 Public isEdit    As Boolean          '* Permite editar un registro.
 Public lData     As String           '* Database Name.
 Public LineCode  As String           '* Line of Code.
 Public SQL       As String           '* Guarda cualquier instrucción SQL.
 Public Tabla     As ADODB.Recordset  '* Abre una tabla de la Base de Datos.
 Public Tebla     As ADODB.Recordset  '* Abre una tabla de la Base de Datos temporalmente.
 Public ToLine    As Integer          '* Number of File.
 Public xSQL      As String           '* Guarda la SQL principal para poder ser edita si se requiere.
 Public xText     As String           '* Valores para campos de Tabla.
 Public yText     As String           '* Valores para datos de campos de Tabla.
 
 Private AntRemoveComm  As String  '* Save the comment of the line.
   
Public Function AppDir() As String
 '* Devuelve el Path real del archivo.
 If (Right$(App.Path, 1) <> "\") Then AppDir = App.Path & "\" Else AppDir = App.Path
End Function

Public Function CargarBD() As Boolean
 Dim AppDB     As String '* Ruta de la Base de Datos.
 Dim tDatabase As String
 
 '* Permite cargar la Base de Datos.
On Error GoTo myErr
 '* Crea y devuelve una referencia a un objeto ActiveX.
 ToLine = FreeFile
 Open AppDir & "Config/Config.txt" For Input As ToLine
  '* Read till End-Of-File.
  Do While Not (EOF(ToLine))
   Line Input #ToLine, LineCode
   '* Read a Text line.
   LineCode = Trim$(RemoveComment(LineCode, "#"))
   If (LineCode <> "") Then
    '* Set database name in the var.
    tDatabase = " " & Trim$(LineCode)
   End If
  Loop
 Close ToLine
 AppDB = Trim$(Mid$(tDatabase, 13))
 lData = AppDB
 AppDB = AppDir & "Database\" & lData
 If (FileExits(AppDB) = True) Then
  Set Datos = CreateObject("ADODB.Connection")
  With Datos
   '* Contiene la información que se utiliza para establecer una conexión a un origen de datos.
   .Provider = "Microsoft.Jet.OLEDB.4.0"
   .Properties("Data Source") = AppDB
   .Properties("Jet OLEDB:Database Password") = "HeryId1304"
   .CursorLocation = adUseClient '* Establece o devuelve la posición de un servicio de cursores.
   .Open '* Abre una conexión a un origen de datos.
  End With
  CargarBD = True
 Else
  CargarBD = False
 End If
 Exit Function
myErr:
 CargarBD = False
End Function

Public Function CargarTabla(ByVal SQL As String, Optional ByVal Closed As Boolean = False, Optional ByVal TempTable As Boolean = False) As Long
 '* Permite cargar temporalmente una tabla de la B.D.
On Error Resume Next
 '* Libero Tabla de la memoria.
 If (TempTable = False) Then
  Set Tabla = Nothing
  '* Crea y devuelve una referencia a unobjeto ActiveX.
  Set Tabla = CreateObject("ADODB.RecordSet")
  With Tabla
   '* Averiguo si objeto esta abierto. Si entonces lo cierro primero.
   If (.State = adStateOpen) Then .Close
   .ActiveConnection = Datos     '* Indica a qué objeto Connection pertenece actualmente el objeto Command o el objeto Recordset especificado.
   .CursorLocation = adUseClient '* Establece o devuelve la posición de un servicio de cursores.
   .LockType = adLockOptimistic  '* Indica el tipo de bloqueo que se pone en los registros durante el proceso de edición.
   .CursorType = adOpenKeyset    '* Indica el tipo de cursor que se usa en un objeto Recordset.
   .Source = SQL                 '* Indica el origen de los datos contenidos en un objeto Recordset (un objeto Command, una instrucción SQL, un nombre de tabla o un procedimiento almacenado).
   Call .Open                    '* Abre un cursor.
   CargarTabla = .RecordCount    '* Almacena el Nro de registros de la tabla.
   .MoveLast                     '* Pasa al primer, último, siguiente o anterior registro de un objeto Recordset especificado y lo convierte en el registro actual.
   .MoveFirst                    '* Pasa al primer, último, siguiente o anterior registro de un objeto Recordset especificado y lo convierte en el registro actual.
  End With
  If (Closed = True) Then CerrarTabla
 Else
  Set Tebla = Nothing
  '* Crea y devuelve una referencia a unobjeto ActiveX.
  Set Tebla = CreateObject("ADODB.RecordSet")
  With Tebla
   '* Averiguo si objeto esta abierto. Si entonces lo cierro primero.
   If (.State = adStateOpen) Then .Close
   .ActiveConnection = Datos     '* Indica a qué objeto Connection pertenece actualmente el objeto Command o el objeto Recordset especificado.
   .CursorLocation = adUseClient '* Establece o devuelve la posición de un servicio de cursores.
   .LockType = adLockOptimistic  '* Indica el tipo de bloqueo que se pone en los registros durante el proceso de edición.
   .CursorType = adOpenKeyset    '* Indica el tipo de cursor que se usa en un objeto Recordset.
   .Source = SQL                 '* Indica el origen de los datos contenidos en un objeto Recordset (un objeto Command, una instrucción SQL, un nombre de tabla o un procedimiento almacenado).
   Call .Open                    '* Abre un cursor.
   CargarTabla = .RecordCount    '* Almacena el Nro de registros de la tabla.
   .MoveLast                     '* Pasa al primer, último, siguiente o anterior registro de un objeto Recordset especificado y lo convierte en el registro actual.
   .MoveFirst                    '* Pasa al primer, último, siguiente o anterior registro de un objeto Recordset especificado y lo convierte en el registro actual.
  End With
  If (Closed = True) Then
   Tebla.Close
   Set Tebla = Nothing
  End If
 End If
On Error GoTo 0
End Function

Public Sub CerrarTabla()
 '* Cierro y libero Tabla de la memoria.
On Error Resume Next
 Tabla.Close
 Set Tabla = Nothing
End Sub

'* Count the number of occurrences of one string within _
   another string.
Public Function CountCharacters(ByVal TheString As String, ByVal CharToCheck As String) As Integer
 Dim mPos As Long, ReturnAgain As Boolean
 Dim Char As String

 CountCharacters = 0
 For mPos = 1 To Len(TheString)
  If (mPos < (Len(TheString) + 1 - Len(CharToCheck))) Then
   Char = Mid$(TheString, mPos, Len(CharToCheck))
   ReturnAgain = True
  Else
   Char = Mid$(TheString, mPos)
   ReturnAgain = False
  End If
  If (Char = CharToCheck) Then CountCharacters = CountCharacters + 1
  If (ReturnAgain = False) Then Exit For
 Next
End Function

Public Function FileExits(ByVal Exists As String) As Boolean
 Dim Gol As String
 
 '* Verifica si existe un archivo.
 FileExits = False
On Error GoTo FileError
 Gol = Dir(Exists)
 If (Gol <> "") Then
  FileExits = True  '* Encontró el Archivo.
  Exit Function
 End If
FileError:
 FileExits = False
End Function

Public Function IsConvertNullEmpty(ByVal isText As Variant, Optional ByVal isNullText As String = "") As String
 '* Comprueba si el campo es NULL ó vacío.
On Error GoTo ErrNullEmpty
 isText = Trim$(isText)
 If (Trim$(isText) = "") Or (IsNull(isText) = True) Or (IsEmpty(isText) = True) Then
  IsConvertNullEmpty = isNullText
 Else
  IsConvertNullEmpty = isText
 End If
 Exit Function
ErrNullEmpty:
 IsConvertNullEmpty = isNullText
End Function

Public Sub Main()
 '* Procedimiento principal del programa.
 modSQL.Delimiter = "|"
 modSQL.LikeOperator = "%"
 If (CargarBD = False) Then
  End
 Else
  frmSplash.Show
 End If
End Sub

'* Remove the comments of a string.
Public Function RemoveComment(ByVal pLine As String, Optional ByVal Token As String = "'") As String
 Dim pComa As Long, nCarac    As String, initPos As Long
 Dim AllOk As Boolean, pCount As Long
 
 pComa = -1
 initPos = 1
 AllOk = False
 AntRemoveComm = ""
 Do While (AllOk = False) And (Len(pLine) > 0)
  '* Search the position of the simple quotation marks.
  pComa = InStr(initPos, pLine, Token)
  If (pComa = 0) Then Exit Do
  '* We take the text until the position of the simple _
     quotation marks.
  nCarac = RTrim$(Mid$(pLine, 1, pComa))
  pCount = CountCharacters(nCarac, Chr$(34)) Mod 2
  If (pCount = 1) Then
   initPos = pComa + 1
   AllOk = False
  Else
   AllOk = True
   Exit Do
  End If
 Loop
 If (AllOk = True) Then '* Return the string without comment.
  RemoveComment = Mid$(pLine, 1, pComa - 1)
  If (pComa - 1 = 0) Then
   AntRemoveComm = ""
  Else
   AntRemoveComm = Mid$(pLine, pComa - 1, Len(pLine)) '* Return the Comment.
  End If
 Else
  RemoveComment = pLine
  AntRemoveComm = ""
 End If
End Function
