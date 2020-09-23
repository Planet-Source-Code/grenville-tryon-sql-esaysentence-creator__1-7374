Attribute VB_Name = "Module1"

' **********************************************
' * PROGRAMA CERTIFICADO PARA EL AÑO 2000
' * AUTOR: GRENVILLE FRANCIS TRYON PERA
' * FECHA: 05/10/1998
' **********************************************

Private Declare Function CopyFile Lib "Kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'RETORNA EL CODIGO DEL VALOR ACTIVO DEL COMBO
Public Function NCodigoCmb(cmb As ComboBox)
Dim Buffer As String, Arr As Variant
Arr = tsstrarr(cmb.Tag, "|")
If cmb.ListIndex >= 0 Then
     NCodigoCmb = Arr(cmb.ListIndex)
End If
End Function

'BUSCA UN ELEMENTO EN EL COMBO POR SU CODIGO
Public Function NHallaIndiceCombo(ByVal Combo As ComboBox, ByVal Codigo As String)
Dim Contador As Integer, Buffer As String, Arr As Variant
Arr = tsstrarr(Combo.Tag, "|")
NHallaIndiceCombo = 0
For Contador = 0 To UBound(Arr, 1) - 1
  If Codigo = CStr(Arr(Contador)) Then
    NHallaIndiceCombo = Contador
    Exit For
  End If
Next
End Function

'LLENA UN COMBO DE VALORES. DEBE RECIBIR UN ARRAY EN FORMA: CODIGO,NOMBRE.
Public Sub NCargaCombo(ByVal Arr As Variant, ByVal Combo As ComboBox, Optional DEFAULT As String)
Dim Contador As Integer, Buffer As String, Parte1 As String, Parte2 As String, Arr2 As Variant
Buffer = ""
If Not IsEmpty(Arr) Then
     For Contador = 0 To UBound(Arr) - 1
          Parte1 = Arr(Contador, 1)
          Parte2 = Arr(Contador, 0)
          Buffer = Buffer + Parte2 + "|"
          Combo.AddItem Parte1
     Next
     Combo.Tag = Buffer
     If Len(DEFAULT) = 0 Then
          DEFAULT = Arr(0, 1)
          If Len(Trim(DEFAULT)) > 0 Then
               Combo.Text = DEFAULT
          End If
     Else
          Arr2 = tsstrarr(Combo.Tag, "|")
          For Contador = 0 To UBound(Arr2, 1) - 1
               If Arr2(Contador) = DEFAULT Then
                    Combo.ListIndex = Contador
                    Exit For
               End If
          Next
     End If
End If
End Sub

'NUEVO. CARGA ARREGLO DE 1 DIMENSION TABULADA
Public Function nCargArray(ByVal rs As rdoResultset) As Variant
Dim Ejex As Double
Dim Arr As Variant
ReDim Arr(RowCount(rs))
Ejex = 0
Do Until rs.EOF
    Arr(Ejex) = rs.GetClipString(1, Chr(9))
    Ejex = Ejex + 1
Loop
nCargArray = Arr
End Function

'CARGA EN UN ARRAY UN RESULTSET
Public Function CargArray(ByVal rs As rdoResultset) As Variant
Dim Ejex As Integer
Dim Ejey As Integer
Dim Arr As Variant
ReDim Arr(RowCount(rs), rs.rdoColumns.Count)
Ejex = 0
Ejey = 0
Do
    Do Until rs.EOF
        For Each Columna In rs.rdoColumns
            Arr(Ejex, Ejey) = Columna.Value
            Ejey = Ejey + 1
        Next
        Ejey = 0
        Ejex = Ejex + 1
        rs.MoveNext
    Loop
Loop Until rs.MoreResults = False
CargArray = Arr
End Function

'RETORNA LA FECHA Y HORA DE TRABAJO
Public Function FechaHora()
If Mid(Time, 10, 1) = "P" Then
 FechaHora = Mid(Date, 7, 4) + Mid(Date, 4, 2) + Mid(Date, 1, 2) + Trim(Str(Val(Mid(Time, 1, 2) + 12))) + Mid(Time, 4, 2)
Else
 FechaHora = Mid(Date, 7, 4) + Mid(Date, 4, 2) + Mid(Date, 1, 2) + Mid(Time, 1, 2) + Mid(Time, 4, 2)
End If
End Function

'ARRAY DE UN DIRECTORIO SEGUN MASCARA
Public Function Adir(mascara)
Dim mypath As String, myname As String, Contador As Integer, ultimo As Integer, madir() As Variant, Actual As Integer
ultimo = 0
Actual = 0
For Contador = 1 To Len(mascara)
 If Mid(mascara, Contador, 1) = "\" Then
  ultimo = Contador
 End If
Next
If ultimo = 0 Then
 mypath = ""
 mascara = "\" + mascara
 ultimo = 1
Else
 mypath = Mid(mascara, 1, ultimo)
End If
myname = Mid(mascara, ultimo + 1)
myname = Dir(mascara, vbNormal)
Do While myname <> ""
 If myname <> "." And myname <> ".." Then
  ReDim Preserve madir(Actual)
  madir(Actual) = mypath + myname
  Actual = Actual + 1
 End If
 myname = Dir
Loop
If Actual <> 0 Then
 Adir = madir
Else
 Adir = Null
End If
End Function

'COMPACTA UNA MDB
Public Sub CompactarMdb(nombreorigen As String)
Dim nombrebackup As String
If Dir(Path + "temp.mdb") <> "" Then Kill (Path + "temp.mdb")
On Error GoTo errormdb
DBEngine.CompactDatabase nombreorigen, Path + "temp.mdb", dbLangGeneral
nombrebackup = Mid(nombreorigen, 1, InStr(nombreorigen, ".")) + "bak"
If Dir(nombrebackup) <> "" Then Kill nombrebackup
Name nombreorigen As nombrebackup
Name Path + "temp.mdb" As nombreorigen
On Error Resume Next
On Error GoTo 0
Exit Sub
'MANEJADOR DE ERROR DE BORRADO
errormdb:
MsgBox "No se puede abrir archivo " + nombreorigen + " para optimizar tiempo. (Intentar posteriormente) : " + Str(Err.Number), vbInformation, "AVISO"
End Sub

'DETERMINA NUMERO DE FILAS EN UN RESULTSET
Public Function RowCount(rs As rdoResultset)
Dim Contador As Double
Contador = 0
If rs.RowCount > 0 Then
 Contador = rs.RowCount
Else
 If rs.EOF And rs.BOF Then
 Else
     rs.MoveFirst
     While Not rs.EOF
         Contador = Contador + 1
         rs.MoveNext
     Wend
     rs.MoveFirst
 End If
End If
RowCount = Contador
End Function

'PADL
Public Function Padl(ByVal Cadena As String, ByVal Longitud As Integer, Optional Caracter As String)
Dim ActualLongitud As Integer
Cadena = Trim(Cadena)
ActualLongitud = Len(Cadena)
If Len(Caracter) <> 1 Then
     Caracter = " "
End If
If ActualLongitud > Longitud Then
  Padl = Mid(Cadena, 1, Longitud)
  If InStr(Padl, ".") Then
    Padl = "0" + Mid(Padl, 1, 1)
  End If
Else
  Padl = String(Longitud - ActualLongitud, Caracter) + Cadena
End If
End Function

'PADR
Public Function Padr(ByVal Cadena As String, ByVal Longitud As Integer, Optional Caracter As String)
Dim ActualLongitud As Integer
Cadena = Trim(Cadena)
ActualLongitud = Len(Cadena)
If Len(Caracter) <> 1 Then
     Caracter = " "
End If
If Longitud - ActualLongitud > 0 Then
     Padr = Cadena + String(Longitud - ActualLongitud, Caracter)
Else
     Padr = Mid(Cadena, 1, Longitud)
End If
End Function

'ENCRIPTACION DE DATOS
Public Function Cript(ByVal Cadena As String, Factor As Integer)
Dim Buffer As String, Contador As Integer
Buffer = ""
For Contador = 1 To Len(Cadena)
   Buffer = Buffer + Chr(Asc(Mid(Cadena, Contador, 1)) + Factor)
Next
Cript = Buffer
End Function

'LLENA UN COMBO DE VALORES. DEBE RECIBIR UN ARRAY EN FORMA: CODIGO,NOMBRE.
Public Sub CargaCombo(ByVal Arr As Variant, ByVal Combo As ComboBox, Optional DEFAULT As String)
Dim Contador As Integer, Buffer As String, Parte1 As String, Parte2 As String
For Contador = 0 To UBound(Arr) - 1
     Parte1 = Arr(Contador, 1)
     Parte2 = Arr(Contador, 0)
     Buffer = Padr(Parte1, 100, " ") + Parte2
     Combo.AddItem Buffer
Next
If Len(DEFAULT) = 0 Then
     Parte1 = Arr(0, 1)
     Parte2 = Arr(0, 0)
     DEFAULT = Padr(Parte1, 100, " ") + Parte2
     If Len(Trim(DEFAULT)) > 0 Then
          Combo.Text = DEFAULT
     End If
Else
    Combo.ListIndex = 0
    For Contador = 0 To Combo.ListCount - 1
        If Mid(Combo.List(Contador), 101) = DEFAULT Then
            Combo.ListIndex = Contador
            Exit For
        End If
    Next
End If
End Sub

'BUSCA UN ELEMENTO EN EL COMBO POR SU CODIGO
Public Function HallaIndiceCombo(ByVal Combo As ComboBox, ByVal Codigo As String)
Dim Contador As Integer, Buffer As String
HallaIndiceCombo = 0
For Contador = 0 To Combo.ListCount
  Buffer = Mid(Combo.List(Contador), 101, 6)
  If Buffer = Codigo Then
    HallaIndiceCombo = Contador
    Exit For
  End If
Next
End Function

'RETORNA UNA FECHA EN FORMATO AAAAMMDD
Public Function AAAAMMDD()
Dim Fecha As String
Fecha = Format(Now, "yyyymmdd")
AAAAMMDD = Val(Fecha)
End Function

'convierte tiempo en segundos. Forma HHH:MM:SS
Public Function TsTimSec(ByVal Tiempo As String) As Double
If Len(Tiempo) = 5 Then
     Tiempo = Tiempo + ":00"
End If
If Mid(Tiempo, 2, 1) = ":" Then
 Tiempo = "00" + Tiempo
End If
If Mid(Tiempo, 3, 1) = ":" Then
 Tiempo = "0" + Tiempo
End If
TsTimSec = Val(Mid(Tiempo, 1, 3)) * 3600 + Val(Mid(Tiempo, 5, 2)) * 60 + Val(Mid(Tiempo, 8, 2))
End Function

'CONVIERT SEGUNDOS EN TIEMPO Forma HHH:MM:SS
Public Function TsSecTim(ByVal Tiempo As Double) As String
Dim Horas As String, Minutos As String, Segundos As String, Buffer As Double
Buffer = Int(Tiempo / 3600)
Tiempo = Tiempo - Buffer * 3600
Horas = Padl(Trim(Str(Buffer)), 3, "0")
Buffer = Int(Tiempo / 60)
Tiempo = Tiempo - Buffer * 60
Minutos = Padl(Trim(Str(Buffer)), 2, "0")
Buffer = Int(Tiempo)
Segundos = Padl(Trim(Str(Buffer)), 2, "0")
TsSecTim = Horas + ":" + Minutos + ":" + Segundos
End Function

'RETORNA UNA FECHA DE FORMATO DD/MM/AAAA EN FORMATO AAAAMMDD
Public Function SAAAAMMDD(ByVal Fecha As String, Optional Separador As String) As String
If Len(Separador) <> 0 Then
    Caracter = Separador
Else
    Caracter = "/"
End If
Fecha = Trim(Fecha)
If InStr(Fecha, Caracter) <> 0 Then
   If Len(Trim(Fecha)) <> 10 Then
      SAAAAMMDD = "00000000"
   Else
      SAAAAMMDD = Mid(Fecha, 7, 4) + Mid(Fecha, 4, 2) + Mid(Fecha, 1, 2)
   End If
Else
   If Len(Trim(Fecha)) <> 8 Then
      SAAAAMMDD = "  " + Caracter + "  " + Caracter + "    "
   Else
      SAAAAMMDD = Mid(Fecha, 7, 2) + Caracter + Mid(Fecha, 5, 2) + Caracter + Mid(Fecha, 1, 4)
   End If
End If
End Function

'CAMBIA UNA CDADENA POR OTRA
Public Function StrTran(ByVal Cadena As String, ByVal Inicial As String, ByVal Final As String) As String
Dim Contador As Integer
Do While InStr(Cadena, Inicial) <> 0
     Contador = InStr(Cadena, Inicial) - 1
     Cadena = Mid(Cadena, 1, Contador) + Final + Mid(Cadena, Contador + Len(Inicial) + 1)
Loop
StrTran = Cadena
End Function

'RETORNA EL CODIGO DEL VALOR ACTIVO DEL COMBO
Public Function CodigoCmb(cmb As ComboBox)
Dim Buffer As String
Buffer = cmb.List(cmb.ListIndex)
CodigoCmb = Mid(Buffer, 101)
End Function

'CARGA NOMBRE DEL USUARIO
Public Function NetUserName() As String
   Dim i As Long
   Dim UserName As String * 255
   i = WNetGetUser("", UserName, 255)
   If i = 0 Then
      NetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
   Else
      NetUserName = ""
   End If
End Function

'DEVUELVE UNA CADENA REPETIDA
Public Function Replicate(ByVal Caracter As String, ByVal Veces As Integer)
Dim Contador As Integer
Replicate = ""
For Contador = 1 To Veces
     Replicate = Replicate + Caracter
Next
End Function

'RESTA 2 FECHAS NUMERICAS, FORMATO AAAAMMDD
Public Function RestaFecha(F2 As Long, F1 As Long) As Long
Dim Fe1 As String, Fe2 As String
Fe1 = Trim(Str(F1))
Fe2 = Trim(Str(F2))
Fe1 = Mid(Fe1, 7, 2) + "/" + Mid(Fe1, 5, 2) + "/" + Mid(Fe1, 1, 4)
Fe2 = Mid(Fe2, 7, 2) + "/" + Mid(Fe2, 5, 2) + "/" + Mid(Fe2, 1, 4)
RestaFecha = CDate(Fe2) - CDate(Fe1) + 1
End Function

'CONVIERTE UN VALOR A RGB
Public Function LongtoRGB(Valor As Long) As Variant
Dim X As Long, R As Integer, G As Integer, B As Integer
X& = Valor
R = X& And &HFF&
G = (X& And &HFF00&) / &H100&
B = (X& And &HFF0000) / &H10000
LongtoRGB = Array(R, G, B)
End Function

'CONVIERTE UN TEXTYO EN UN ARREGLO JUSTIFICADO
Public Function Ajusta(Cadena As String, Cualfont As Font, Ancho As Double, lblFantasma As Label) As Variant
Dim Conta As Integer, Contador As Integer, Buffer As String, Rpta As Variant, Arreglo() As Variant, Final() As Variant
Dim Buffer2 As Variant, Cambio As Integer
Rpta = tsstrarr(Cadena, " ")
lblFantasma.Font = Cualfont
lblFantasma.AutoSize = True
lblFantasma.Caption = ""
Buffer = ""
Numero = 1
For Contador = 0 To UBound(Rpta, 1)
     lblFantasma.Caption = lblFantasma.Caption + CStr(Rpta(Contador)) + " "
     If lblFantasma.Width > Ancho Then
          lblFantasma.Caption = Trim(Buffer)
          Do While Trim(lblFantasma.Width) < Ancho
               For Conta = Len(lblFantasma.Caption) - 1 To 1 Step -1
                    If Mid(lblFantasma.Caption, Conta, 1) = " " Then
                         lblFantasma.Caption = Mid(lblFantasma.Caption, 1, Conta) + Mid(lblFantasma.Caption, Conta)
                         If lblFantasma.Width >= Ancho Then
                              Exit For
                         End If
                    End If
               Next
          Loop
          ReDim Preserve Arreglo(Numero)
          Arreglo(Numero - 1) = lblFantasma.Caption
          Numero = Numero + 1
          lblFantasma.Caption = ""
          Buffer = ""
          Contador = Contador - 1
     Else
          Buffer = Buffer + CStr(Rpta(Contador)) + " "
     End If
Next
ReDim Preserve Arreglo(Numero)
Arreglo(Numero - 1) = lblFantasma.Caption
Numero = 1
For Contador = 0 To UBound(Arreglo, 1) - 1
     Cambio = InStr(Arreglo(Contador), vbCrLf)
     Do While Cambio <> 0
          ReDim Preserve Final(Numero)
          Final(Numero - 1) = Mid(CStr(Arreglo(Contador)), 1, Cambio - 1)
          Arreglo(Contador) = Mid(CStr(Arreglo(Contador)), Cambio + 2)
          Numero = Numero + 1
          Cambio = InStr(Arreglo(Contador), vbCrLf)
     Loop
     ReDim Preserve Final(Numero)
     Final(Numero - 1) = Arreglo(Contador)
     Numero = Numero + 1
Next
Ajusta = Final
End Function

'AJUSTA EL TEXTO A LA DERECHA
Public Function Derecha(Cadena As String, Longitud As Double, Font As Font, lbl As Label) As String
lbl.AutoSize = True
lbl.Caption = Cadena
Do While lbl.Width < Longitud
     lbl.Caption = " " + lbl.Caption
Loop
Derecha = lbl.Caption
End Function

'SUMA UN NUMERO DE DIAS A UNA FECHA
Public Function AumentaFecha(Fecha As Long, NumeroDias As Integer) As Long
Dim Contador As Long, ActualFecha As Date, Buffer As String
Buffer = CStr(Fecha)
ActualFecha = DateSerial(Val(Mid(Buffer, 1, 4)), Val(Mid(Buffer, 5, 2)), Val(Mid(Buffer, 7, 2)))
For Contador = 0 To NumeroDias - 1
     ActualFecha = ActualFecha + 1
Next
AumentaFecha = Val(Format(ActualFecha, "yyyymmdd"))
End Function

'RESTA UN NUMERO DE DIAS A UNA FECHA
Public Function DisminuyeFecha(Fecha As Long, NumeroDias As Integer) As Long
Dim Contador As Long, ActualFecha As Date, Buffer As String
Buffer = CStr(Fecha)
ActualFecha = DateSerial(Mid(Buffer, 1, 4), Mid(Buffer, 5, 2), Mid(Buffer, 7, 2))
For Contador = 0 To NumeroDias - 1
     ActualFecha = ActualFecha - 1
Next
DisminuyeFecha = Val(Format(ActualFecha, "yyyymmdd"))
End Function

' Verifica la existencia de archivos segun Parámetro sNombreArchivo: Ubicación, nombre y extensión de archivo a localizar.
Public Function ExisteArchivo(ByVal sNombreArchivo As String) As Boolean
ExisteArchivo = False
Dim AttrDev As Integer
On Error Resume Next
AttrDev = GetAttr(sNombreArchivo)
If Err.Number Then
    Err.Clear
Else
    ExisteArchivo = True
End If
On Error GoTo 0
End Function

' HALLA CADENA DE DERECHA DEL DELIMITADOR
Public Function Rat(ByVal Cadena As String, ByVal delimitador As String) As Integer
Dim Contador As Integer, Buffer As String
Rat = 0
For Contador = Len(Cadena) To 1 Step -1
     If Mid(Cadena, Contador, 1) = delimitador Then
          Rat = Contador
          Exit For
     End If
Next
End Function

' EXISTE ARCHIVO?
Public Function File(Nombre As String) As Boolean
File = False
If LCase(Dir(Nombre)) = LCase(Mid(Nombre, Rat(Nombre, "\") + 1)) Then
     File = True
End If
End Function
