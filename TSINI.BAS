Attribute VB_Name = "tsini"

' **********************************************
' * PROGRAMA CERTIFICADO PARA EL AÑO 2000
' * AUTOR: GRENVILLE FRANCIS TRYON PERA
' * FECHA: 05/10/1998
' **********************************************

Public Function tsleeini(ByVal ARCHIVO As String, Optional Completo As Boolean) As Variant
Dim Rpta() As Variant, texto As Variant, Actual As Integer
Actual = 0
Open ARCHIVO For Input As #1
Do While Not EOF(1)
    Line Input #1, texto
    If Mid(texto, 1, 1) <> ";" Then
        ReDim Preserve Rpta(Actual)
        Rpta(Actual) = texto
        Actual = Actual + 1
    Else
        If Completo Then
             ReDim Preserve Rpta(Actual)
             Rpta(Actual) = texto
            Actual = Actual + 1
        End If
    End If
Loop
tsleeini = Rpta
Close #1
End Function

'OBTIENE UNA VARIABLE DEL ARRAY INI
Public Function tsgetini(ByVal Arr As Variant, ByVal Variable As String, ByVal DEFAULT As String) As String
Dim Contador As Integer, Indice As Integer, Rpta As String, Actual As String, Buffer As String
Rpta = DEFAULT
For Contador = 0 To UBound(Arr)
    Actual = RTrim(UCase(Arr(Contador)))
    Buffer = RTrim(Arr(Contador))
    Variable = RTrim(UCase(Variable))
    Indice = InStr(1, Actual, Variable)
    If Indice = 1 And Mid(Buffer, Len(Variable) + 1, 1) = "=" Then
        Rpta = Mid(Buffer, Len(Variable) + 2)
        Exit For
    End If
Next
tsgetini = Rpta
End Function

'OBTIENE UN RANGO DE UN ARRAY
Public Function tsrangoini(ByVal Arr As Variant, ByVal Desde As String, ByVal Hasta As String) As Variant
Dim Contador As Integer, Indice As Integer, Rpta() As Variant, Actual As String, Numero As Integer, OK As Boolean
Dim Buffer As String
OK = False
Numero = 0
For Contador = 0 To UBound(Arr)
    Actual = RTrim(UCase(Arr(Contador)))
    Desde = RTrim(UCase(Desde))
    Hasta = RTrim(UCase(Hasta))
    Indice = InStr(1, Actual, Desde)
    If Indice = 1 Then
        OK = True
        Contador = Contador + 1
        Actual = RTrim(UCase(Arr(Contador)))
        While InStr(1, Actual, Hasta) <> 1
         Buffer = Arr(Contador)
'         If Len(Buffer) <> 0 Then
          ReDim Preserve Rpta(Numero)
          Rpta(Numero) = Arr(Contador)
          Numero = Numero + 1
'         End If
         Contador = Contador + 1
         Actual = RTrim(UCase(Arr(Contador)))
        Wend
    End If
Next
If OK = True Then
     tsrangoini = Rpta
Else
     tsrangoini = Array(0)
End If
End Function

'CONVIERTE UNA CADENA EN ARRAY
Public Function tsstrarr(ByVal Cadena As String, ByVal delimitador As String)
Dim Rpta() As Variant, Contador As String, Numero As Integer, Buffer As String
Numero = 0
If IsEmpty(delimitador) Then
 delimitador = "|"
End If
Cadena = Cadena + delimitador
While InStr(1, Cadena, delimitador) <> 0
    Buffer = Mid(Cadena, 1, InStr(1, Cadena, delimitador) - 1)
    ReDim Preserve Rpta(Numero)
    Rpta(Numero) = Buffer
    Numero = Numero + 1
    Cadena = Mid(Cadena, InStr(1, Cadena, delimitador) + 1)
Wend
tsstrarr = Rpta
End Function

'GRABA EL ARCHIVO INI
Public Function tsgraini(ARCHIVO As String, Variable As String, valor As String)
Dim Arr As Variant, Indice As Integer, Cadena As String, Rpta As Variant
If Len(Trim(Dir(ARCHIVO))) > 0 Then
  Arr = tsleeini(ARCHIVO, True)
  Variable = UCase(RTrim(Variable))
  For Contador = 0 To UBound(Arr)
      Cadena = UCase(RTrim(Arr(Contador)))
      If InStr(1, Cadena, Variable) = 1 Then
        Indice = Contador
      End If
  Next
  Open ARCHIVO For Output As #1
  If Indice = 0 Then
     For Contador = 0 To UBound(Arr)
         Print #1, Arr(Contador)
     Next
     Print #1, Variable + "=" + valor
  Else
     For Contador = 0 To Indice - 1
         Print #1, Arr(Contador)
     Next
     Print #1, LCase(Variable) + "=" + valor
     For Contador = Indice + 1 To UBound(Arr)
         Print #1, Arr(Contador)
     Next
  End If
  Close #1
Else
  Open ARCHIVO For Output As #1
   Print #1, Variable + "=" + valor
  Close #1
End If
tsgraini = tsleeini(ARCHIVO)
End Function

'DEVUELVE UN COLOR DE UN GETINI TIPICO
Function RGBColor(Colores As String) As Long
Dim ArrColores As Variant
If Len(Colores) > 0 Then
    ArrColores = tsstrarr(Colores, "|")
    RGBColor = RGB(Val(ArrColores(0)), Val(ArrColores(1)), Val(ArrColores(2)))
End If
End Function

'CARGA LOS DOS ARRAYS EN UNO.
Public Function IniMasIni(ByVal FilePrimero As String, ByVal FileSegundo As String) As Variant
Dim Arr1 As Variant, Arr2 As Variant, Arr3 As Variant, Contador As Integer, Cuenta As Integer
On Error Resume Next
If Len(Dir(FilePrimero)) > 0 And Len(FilePrimero) > 0 Then
     Arr1 = tsleeini(FilePrimero)
     Cuenta = UBound(Arr1, 1) + 1
End If
If Len(Dir(FileSegundo)) > 0 And Len(FileSegundo) > 0 Then
     Arr2 = tsleeini(FileSegundo)
     Cuenta = Cuenta + UBound(Arr2, 1) + 1
End If
ReDim Arr3(Cuenta)
Cuenta = 0
If Not IsEmpty(Arr1) Then
     For Contador = 0 To UBound(Arr1, 1)
          Arr3(Cuenta) = Arr1(Contador)
          Cuenta = Cuenta + 1
     Next
End If
If Not IsEmpty(Arr2) Then
     For Contador = 0 To UBound(Arr2, 1)
          Arr3(Cuenta) = Arr2(Contador)
          Cuenta = Cuenta + 1
     Next
End If
IniMasIni = Arr3
On Error GoTo 0
End Function

