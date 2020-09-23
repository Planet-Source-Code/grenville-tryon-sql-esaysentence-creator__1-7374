VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ESTABLECER CONEXION A BASES DE DATOS VIA SQL"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Matricular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5970
      TabIndex        =   2
      ToolTipText     =   "Utilize este boton para crear una nueva conexion a la lista"
      Top             =   3180
      Width           =   915
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Escriba aqui los parametros de conexion al SQL (Doble click invoca al activo en la lista para editarlo)"
      Top             =   3180
      Width           =   5685
   End
   Begin MSComDlg.CommonDialog cmm 
      Left            =   2580
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ingresar (Enter)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   90
      TabIndex        =   3
      ToolTipText     =   "Ingrese al constructor de consultas SQL"
      Top             =   3600
      Width           =   6825
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      ItemData        =   "frm1.frx":08CA
      Left            =   90
      List            =   "frm1.frx":08CC
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Presione Del para eliminar el elemento activo en la lista"
      Top             =   90
      Width           =   6795
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Arrini As Variant

Private Sub Form_Load()
cnActual = 0
Set cn = Nothing
CargaRegistro
End Sub

Private Sub Command1_Click()
CreaRegistro
CargaRegistro
End Sub

Private Sub cmd_Click(Index As Integer)
Dim Contador As Integer, OK As Boolean
On Error GoTo HELL
OK = False
Screen.MousePointer = vbHourglass
For Contador = 0 To lst.ListCount - 1
    If lst.Selected(Contador) Then
        Me.Caption = "Enlazando a : " + lst.List(Contador)
        DoEvents
        OK = True
        Set cn = Nothing
        Set cn = New rdoConnection
        cn.CursorDriver = rdUseOdbc
        cn.Connect = lst.List(Contador)
        cn.EstablishConnection
        CreaRegistro
        frm.CargaTree
    End If
Next
If Not OK Then
    MsgBox "Debe definir alguna cadena de conexion!", vbOKOnly, "QueryBuilder"
Else
    Unload Me
    frm.Show
End If
SIGUE:
On Error GoTo 0
Screen.MousePointer = vbDefault
Exit Sub
HELL:
    MsgBox "No es posible conectarse a esta BD > " + Err.Description, vbOKOnly, "QueryBuilder"
    GoTo SIGUE
End Sub

Private Sub CreaRegistro()
Dim Contador As Integer, Actual As String, Existe As Boolean
Actual = txt.Text
Existe = False
For Contador = 0 To lst.ListCount
    If Actual = lst.List(Contador) Then
        Existe = True
    End If
Next
If Not Existe Then
    tsgraini App.Path + "\sele.ini", "f" + Format(Now, "yyyymmddhhmmss"), Actual
End If
End Sub

Private Sub CargaRegistro()
Dim Contador As Integer, Buffer As String
Arrini = tsleeini(App.Path + "\sele.ini")
lst.Clear
For Contador = 0 To UBound(Arrini, 1)
    Buffer = CStr(Arrini(Contador))
    Buffer = Mid(Buffer, InStr(Buffer, "=") + 1)
    If Len(Trim(Buffer)) > 0 Then
        lst.AddItem Buffer
    End If
Next
lst.ListIndex = 0
End Sub


Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Contador As Integer, Parte2 As String, CU As Variant
If KeyCode = 46 Then
    For Contador = 0 To UBound(Arrini, 1)
        CU = tsstrarr(CStr(Arrini(Contador)), "=")
        Parte2 = Mid(CStr(Arrini(Contador)), InStr(CStr(Arrini(Contador)), "=") + 1)
        If Parte2 = lst.Text Then
            Arrini = tsgraini(App.Path + "\sele.ini", CStr(CU(0)), "")
            Exit For
        End If
    Next
    Arrini = tsleeini(App.Path + "\sele.ini")
    CargaRegistro
    lst.SetFocus
End If
End Sub

Private Sub txt_DblClick()
If Len(Trim(txt.Text)) = 0 Then
    txt.Text = lst.Text
End If
End Sub
