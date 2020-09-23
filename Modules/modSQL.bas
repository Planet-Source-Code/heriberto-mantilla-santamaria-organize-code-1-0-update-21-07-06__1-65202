Attribute VB_Name = "modSQL"
'**********************************************'
'* Programmed by HACKPRO TM © Copyright 2005  *'
'* Programado por HACKPRO TM © Copyright 2005 *'
'**********************************************'
Option Explicit

 Private m_sDelimiter As String, m_bLikeOperator As String

'* SELECT: Permite hacer una selección de registros.
Public Function Get_Select(ByVal Table As String, ByVal Columns As String, Optional ByVal Where As String = "", Optional ByVal Order As String = "", Optional ByVal GroupBy As String = "", Optional ByVal PutTable As Boolean = False) As String
 Dim tmp As String, tmp1 As String
 
 tmp1 = ""
 If (PutTable = True) Then tmp1 = Table & "."
 tmp = "SELECT " & tmp1 & Trim$(Columns) & " FROM " & Trim$(Table)
 If (Where <> "") Then tmp = tmp & " WHERE " & tmp1 & Trim$(Where)
 If (Order <> "") Then tmp = tmp & " ORDER BY " & tmp1 & Trim$(Order)
 If (GroupBy <> "") Then tmp = tmp & " GROUP BY " & tmp1 & Trim$(GroupBy)
 Get_Select = tmp & ";"
End Function

'* SELECT: Entre 2 tablas por 2 indices comunes.
Public Function Get_Select_Join(ByVal Table_A As String, ByVal Table_B As String, ByVal Id_A As String, ByVal Id_B As String, ByVal Columns As String, Optional ByVal Where As String = "", Optional ByVal Order As String = "") As String
 Dim W As String, Table As String
 
 Table = Trim$(Table_A) & "." & Trim$(Id_A) & ", " & Trim$(Table_B) & "." & Trim$(Id_B)
 W = Trim$(Table_A) & "." & Trim$(Id_A) & " = " & Trim$(Table_B) & "." & Trim$(Id_B)
 If (Where <> "") Then W = W & " AND (" & Trim$(Where) & ")"
 Get_Select_Join = Get_Select(Table_A & " INNER JOIN " & Table_B, Table, W, Trim$(Order))
End Function

'* Get_Simp_Set: Devuelve una asignación (por defecto) simple entre comillas X = '1'.
Public Function Get_Simp_Set(ByVal Col As String, ByVal Bal As String, Optional ByVal Sign As String = "=", Optional ByVal Comillas As Boolean = True) As String
 Dim Cm As String
 
 Cm = "'"
 Sign = " " & UCase$(Trim$(Sign)) & " "
 If (Comillas = False) Then Cm = ""
 If (Mid$(Bal, 1, 2) = "!!") Then
  '* Si empieza por !! no le pongo comas...
  Bal = Mid$(Bal, 3)
  Cm = ""
 ElseIf (Mid$(Bal, 1, 1) = "'") And (Mid$(Bal, Len(Bal), 1) = "'") Then
  Cm = ""
 End If
 If (Trim$(Sign) = "LIKE") Then
  Get_Simp_Set = Col & Sign & Cm & LikeOperator & Trim$(Bal) & LikeOperator & Cm
 Else
  Get_Simp_Set = Col & Sign & Cm & Trim$(Bal) & Cm
 End If
End Function
 
'* Get_Mult_Set: Devuelve asignaciones múltiples comúnmente utilizadas en sentencias SQL.
Public Function Get_Mult_Set(ByVal A_Cols As String, ByVal A_Vals As String, Optional ByVal Simb As String = ",", Optional ByVal Sign As String = "=", Optional ByVal Comillas As Boolean = True, Optional ByVal Equal As Boolean = False) As String
 Dim Simbol As String, tmp  As Variant, Col As String, X   As Long, tmp2 As String
 Dim Temp   As String, tmp1 As Variant, Cm  As String, Bal As String
 
 tmp = Split(A_Cols, Delimiter)
 tmp1 = Split(A_Vals, Delimiter)
 Sign = " " & UCase$(Trim$(Sign)) & " "
 If (Trim$(Simb) = ",") Then
  Simb = UCase$(Trim$(Simb)) & " "
 Else
  Simb = " " & UCase$(Trim$(Simb)) & " "
 End If
 For X = 0 To UBound(tmp)
  If (Temp <> "") Then Temp = Temp & Simb
  Cm = "'"
  If (Trim$(Sign) = "LIKE") Then
   Simbol = LikeOperator
  Else
   Simbol = ""
  End If
  If (Equal = False) Then
   tmp2 = Trim$(tmp1(X))
   If (Mid$(tmp2, 1, 1) = "'") Then
    Bal = Mid$(tmp2, 1, 1) & Simbol & Mid$(tmp2, 2, Len(tmp2) - 2) & Simbol & "'"
   Else
    Bal = Simbol & tmp2 & Simbol
   End If
  Else
   tmp2 = Trim$(tmp1(0))
   If (Mid$(tmp2, 1, 1) = "'") Then
    Bal = Mid$(tmp2, 1, 1) & Simbol & Mid$(tmp2, 2, Len(tmp2) - 2) & Simbol & "'"
   Else
    Bal = Simbol & tmp2 & Simbol
   End If
  End If
  Col = Trim$(tmp(X))
  If (Comillas = False) Then Cm = ""
  If (Trim$(Bal) = "") Then Bal = " "
 On Error Resume Next
  If (Mid$(Bal, 1, 2) = "!!") Then
   '* Si empieza por !! no le pongo comas...
   Bal = Mid$(Bal, 3)
   Cm = ""
  ElseIf (Mid$(Bal, 1, 1) = "'") And (Mid$(Bal, Len(Bal), 1) = "'") Then
   Cm = ""
  End If
  Temp = Temp & Col & Sign & Cm & Trim$(Bal) & Cm
 Next
 Get_Mult_Set = Temp
End Function
    
'* Get_Commas: (True|False, 1, 2, 4...) True pone comillas => '1','2','4'...
Public Function Get_Commas(ParamArray Arr_In()) As String
 Dim A As Variant, Com As Long

 For Com = 1 To UBound(Arr_In)
  A = A & Arr_In(Com) & IIf(Com < UBound(Arr_In), Delimiter, "")
 Next
 Get_Commas = Get_CommasA(A, Arr_In(0))
End Function
    
'* Get_CommasA: Como la anterior pero devuelve entre comas el Array pasado.
Public Function Get_CommasA(ByVal Arr_In As String, Optional ByVal Comillas As Boolean = True) As String
 Dim Coma  As String, Temp As String
 Dim Filas As Variant, i   As Long
 
 Coma = "'"
 Filas = Split(Arr_In, Delimiter)
 If (Comillas = False) Then Coma = "" '* El 1er param = true, metemos comas.
 For i = 0 To UBound(Filas)
  If (Temp <> "") Then Temp = Temp & ", "
  If (Mid$(Filas(i), 1, 2) = "!!") Then
   '* Si empieza por !! no le pongo comas...
   Temp = Temp & Trim$(Mid$(Filas(i), 3))
  Else
   Temp = Temp & Coma & Trim$(Filas(i)) & Coma
  End If
 Next
 Get_CommasA = Temp
End Function
 
'* INSERT: Inserta valores en la B.D.
Public Function Get_Insert(ByVal Table As String, ByVal Columns As String, ByVal Values As String) As String
 Get_Insert = "INSERT INTO " & Trim$(Table) & " (" & Trim$(Columns) & ") VALUES(" & Values & ");"
End Function

'* UPDATE: Actualiza valores en la B.D.
Public Function Get_Update(ByVal Table As String, ByVal Values As String, ByVal Where As String) As String
 Get_Update = "UPDATE " & Trim$(Table) & " SET " & Trim$(Values) & " WHERE " & Trim$(Where) & ";"
End Function

'* BETWEEN: Actualiza una tabla con valores de otra.
Public Function Get_Between(ByVal Value_A As String, ByVal Value_B As String) As String
 Get_Between = " BETWEEN " & Value_A & " AND " & Value_B
End Function

'* DELETE: Borra valores en la B.D.
Public Function Get_Delete(ByVal Table As String, Optional ByVal Values As String = "*", Optional ByVal Where As String = "") As String
 Dim tmp As String
 
 tmp = "DELETE " & Trim$(Values) & " FROM " & Trim$(Table)
 If (Where <> "") Then tmp = tmp & " WHERE " & Trim$(Where)
 Get_Delete = tmp & ";"
End Function

Public Property Get Delimiter() As String
 Delimiter = m_sDelimiter
End Property

Public Property Let Delimiter(ByVal sDelimiter As String)
 m_sDelimiter = sDelimiter
End Property

Public Property Get LikeOperator() As String
 LikeOperator = m_bLikeOperator
End Property

Public Property Let LikeOperator(ByVal bLikeOperator As String)
 m_bLikeOperator = bLikeOperator
End Property
