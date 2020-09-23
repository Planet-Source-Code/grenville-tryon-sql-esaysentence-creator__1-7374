VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QueryBuilder"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Too 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   847
      ButtonWidth     =   1244
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Permite re-establecer la conexion con otra BD"
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ejecuta el comando"
            Object.Tag             =   "2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cancela ejecucion del comando"
            Object.Tag             =   "3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Pone en el ClipBoard el texto ingresado"
            Object.Tag             =   "4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Recupera del ClipBoard el texto ingresado"
            Object.Tag             =   "5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Pone en el Buffer interno (Propio) el texto ingresado"
            Object.Tag             =   "6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Recupera del Buffer interno (Propio) el texto ingresado"
            Object.Tag             =   "7"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Crea una variable Visual Basic con el texto ingresado"
            Object.Tag             =   "8"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Graba el texto"
            Object.Tag             =   "9"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Recupera un texto grabado"
            Object.Tag             =   "10"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton Command1 
      DragMode        =   1  'Automatic
      Height          =   5865
      Left            =   2070
      MousePointer    =   9  'Size W E
      TabIndex        =   17
      Top             =   480
      Width           =   105
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5925
      Left            =   -30
      TabIndex        =   14
      Top             =   450
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   10451
      _Version        =   327682
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   405
      Left            =   2190
      TabIndex        =   15
      Top             =   6000
      Width           =   6915
      Begin VB.Label Label1 
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   60
         Width           =   3795
      End
   End
   Begin MSComDlg.CommonDialog com 
      Left            =   5070
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5550
      Left            =   2160
      TabIndex        =   0
      Top             =   450
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   9790
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Accion"
      TabPicture(0)   =   "frm.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Respuestas"
      TabPicture(1)   =   "frm.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Configuracion"
      TabPicture(2)   =   "frm.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         Caption         =   "Buffer Propio"
         Height          =   4215
         Left            =   -74910
         TabIndex        =   10
         Top             =   1200
         Width           =   6405
         Begin VB.TextBox Text5 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3885
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   11
            Top             =   240
            Width           =   6225
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Caption         =   "TimeOut"
         Height          =   1035
         Left            =   -70500
         TabIndex        =   7
         Top             =   120
         Width           =   1965
         Begin VB.TextBox Text4 
            Height          =   330
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "10"
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar en:"
            Height          =   225
            Left            =   120
            TabIndex        =   9
            Top             =   390
            Width           =   1050
         End
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5385
         Left            =   -74940
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   90
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5400
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frm.frx":091E
         Top             =   90
         Width           =   6405
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Recuperacion de datos:"
         Height          =   1035
         Left            =   -74940
         TabIndex        =   3
         Top             =   120
         Width           =   4395
         Begin VB.TextBox Text2 
            Height          =   330
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   210
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   370
            _Version        =   327682
            LargeChange     =   10
            Min             =   10
            Max             =   300
            SelStart        =   10
            TickFrequency   =   10
            Value           =   10
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por vez."
            Height          =   225
            Left            =   1950
            TabIndex        =   12
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grupos de:"
            Height          =   225
            Left            =   210
            TabIndex        =   6
            Top             =   660
            Width           =   915
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8430
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":0955
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":10C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":1839
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":1FAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":271D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":2E8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":3601
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":3D73
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":44E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":4C57
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":53C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu o 
         Caption         =   "Copiar"
         Index           =   0
      End
      Begin VB.Menu o 
         Caption         =   "Pegar"
         Index           =   1
      End
      Begin VB.Menu o 
         Caption         =   "Eliminar"
         Index           =   2
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu o 
         Caption         =   "Mascara Inicial"
         Index           =   4
      End
      Begin VB.Menu o 
         Caption         =   "Quitar ultima coma"
         Index           =   5
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu o 
         Caption         =   "Insert"
         Index           =   7
      End
      Begin VB.Menu o 
         Caption         =   "Update"
         Index           =   8
      End
      Begin VB.Menu o 
         Caption         =   "Select"
         Index           =   9
      End
      Begin VB.Menu o 
         Caption         =   "Delete"
         Index           =   10
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu o 
         Caption         =   "where"
         Index           =   12
      End
      Begin VB.Menu o 
         Caption         =   "=''"
         Index           =   13
      End
      Begin VB.Menu o 
         Caption         =   "Enter"
         Index           =   14
      End
      Begin VB.Menu o 
         Caption         =   "Order by"
         Index           =   15
      End
      Begin VB.Menu o 
         Caption         =   "Group by"
         Index           =   16
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu o 
         Caption         =   "Anadir Campos"
         Index           =   18
      End
      Begin VB.Menu o 
         Caption         =   "Anadir Campo"
         Index           =   19
      End
      Begin VB.Menu o 
         Caption         =   "Crear Tabla e Indices"
         Index           =   20
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const TVM_SETBKCOLOR = 4381&

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Dim Seguir As Boolean

Private Sub Form_Activate()
Me.WindowState = 0
End Sub

Private Sub Form_Load()
Dim nodX As Node, Contador As Integer, Conta As Integer, Enter As String
Dim hMenu As Long
Const SC_SIZE = &HF000
Const MF_BYCOMMAND = &H0
hMenu = GetSystemMenu(hWnd, 0)
Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
    
Call SendMessage(TreeView1.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(232, 232, 232))
Enter = Chr(13) + Chr(10)
Text1.Text = "select " + Enter + Enter + "from " + Enter + Enter + "where " + Enter + Enter + "order by " + Enter + Enter + "group by"
Me.Left = 0
End Sub

Private Sub Recon()
frm1.Show
End Sub

Private Sub ABuffer()
If MsgBox("Desea reemplazar al buffer " + Text5.Text, vbYesNo, "QueryBuilder") = vbYes Then
    Text5.Text = Text1.Text
End If
End Sub

Private Sub HazSelect()
Dim Buffer As String
If TreeView1.SelectedItem.Child Is Nothing Then
Else
    If MsgBox("Confirme reemplazo de sentencia en pantalla", vbYesNo, "QueryBuilder") = vbYes Then
        TreeView1.SetFocus
        Buffer = StrTran(TreeView1.SelectedItem.FullPath, "\", ".")
        Text1.Text = "select " + Chr(13) + Chr(10)
        DoEvents
        SendKeys "{DOWN}", True
        Do While TreeView1.SelectedItem.Child Is Nothing
            Text1.Text = Text1.Text + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + ", " + Chr(13) + Chr(10)
            SendKeys "{DOWN}", True
        Loop
        Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 4) + Chr(13) + Chr(10)
        Text1.Text = Text1.Text + "from " + Buffer
    End If
End If
End Sub

Private Sub Variable()
Dim valor As String
valor = Text1.Text
valor = StrTran(valor, Chr(13) + Chr(10), "@@@@@")
valor = StrTran(valor, "@@@@@", Chr(34) + " + _" + Chr(13) + Chr(10) + Chr(34))
valor = "Cadena=" + Chr(34) + valor + Chr(34)
Clipboard.Clear
Clipboard.SetText valor
MsgBox "Se ha creado la variable : " + valor, vbOKOnly, "QueryBuilder"
End Sub

Private Sub AClipBoard()
Select Case SSTab1.Tab
Case 0
    Clipboard.Clear
    Clipboard.SetText Text1.Text
Case 2
    Clipboard.Clear
    Clipboard.SetText Text5.Text
End Select
End Sub

Private Sub Ejecuta()
Dim RowBuf As Variant, Cuantos As Double
Dim RowsReturned As Integer
Dim Mycn As New rdoConnection
Dim qy As New rdoQuery
Dim rs As rdoResultset
Dim i As Integer, J As Integer, Buffer As String
Screen.MousePointer = vbHourglass
Seguir = True
Set Mycn = cn
Set qy = New rdoQuery
qy.Name = "GetRowsQuery"
If Text1.SelLength <> 0 Then
    qy.SQL = Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength + 1)
Else
    qy.SQL = Text1.Text
End If
qy.RowsetSize = 1
Set qy.ActiveConnection = cn
qy.QueryTimeout = Val(Text4.Text)
On Error GoTo MALSQL
Set rs = qy.OpenResultset(rdOpenStatic, rdConcurReadOnly, rdExecDirect)
On Error GoTo 0
Text3.Text = ""
On Error Resume Next
Cuantos = 0
SSTab1.Tab = 1
Do Until rs.EOF
     RowBuf = rs.GetRows(CLng(Text2.Text))
     RowsReturned = UBound(RowBuf, 2) + 1
     For i = 0 To RowsReturned - 1
          Cuantos = Cuantos + 1
          Buffer = ""
          For J = 0 To UBound(RowBuf, 1)
               Buffer = Buffer + CStr(RowBuf(J, i)) & Chr(9)
          Next
          Text3.Text = Text3.Text + Buffer + Chr(13) + Chr(10)
     Next i
     Label1.Caption = "Recuperando " + Str(Cuantos)
     DoEvents
     If Not Seguir Then Exit Do
Loop
On Error GoTo 0
SALEMALSQL:
Screen.MousePointer = vbDefault
Set rs = Nothing
Set qy = Nothing
Set Mycn = Nothing
Label1.Caption = ""
Exit Sub
MALSQL:
     MsgBox "Algun error con el SQL!  >  " + Err.Description, vbOKOnly, "QueryBuilder"
     GoTo SALEMALSQL
End Sub

Private Sub Detener()
Seguir = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
If MsgBox("Desea salir del modulo?", vbYesNo, "QueryBuilder") = vbYes Then
    Cancel = 0
Else
    Me.WindowState = 1
End If
End Sub


Private Sub Slider1_Change()
Text2.Text = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
Text2.Text = Slider1.Value
End Sub

Public Sub CargaTree()
Dim Contador As Integer, Conta As Integer, Buffer, Actual As String
Dim nodX As Node
Buffer = UCase(cn.Connect)
If InStr(Buffer, "DATABASE=") <> 0 Then
    Buffer = Mid(Buffer, InStr(Buffer, "DATABASE=") + 9) + ".dbo."
Else
    Buffer = ""
End If
Screen.MousePointer = vbHourglass
TreeView1.LabelEdit = tvwManual
For Contador = 0 To cn.rdoTables.Count - 1
     Set nodX = TreeView1.Nodes.Add(, , Chr(64 + cnActual) + CStr(Offset) + CStr(Contador), Buffer + cn.rdoTables(Contador).Name)
     For Conta = 0 To cn.rdoTables(Contador).rdoColumns.Count - 1
          Set nodX = TreeView1.Nodes.Add(Chr(64 + cnActual) + CStr(Offset) + CStr(Contador), tvwChild, Chr(64 + cnActual) + Chr(64 + cnActual) + CStr(Offset) + CStr(Contador) + CStr(Conta), cn.rdoTables(Contador).rdoColumns(Conta).Name)
     Next
     Offset = Offset + 1
Next
cnActual = cnActual + 1
Screen.MousePointer = vbDefault
End Sub

Private Sub DeClipBoard()
If MsgBox("Desea cargar informacion del Clipboard:" + Clipboard.GetText, vbYesNo, "QueryBuilder") = vbYes Then
    Select Case SSTab1.Tab
    Case 0
        Text1.Text = Clipboard.GetText
    Case 2
        Text5.Text = Clipboard.GetText
    End Select
End If
End Sub

Private Sub DeBuffer()
If MsgBox("Desea cargar buffer " + Text5.Text, vbYesNo, "QueryBuilder") = vbYes Then
    Text1.Text = Text5.Text
End If
End Sub

Private Sub Text1_DblClick()
Dim Desde As Integer
If Text1.SelLength > 2 Then
    Clipboard.Clear
    Clipboard.SetText Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)
Else
    Text1.Text = Mid(Text1.Text, 1, Text1.SelStart) + Clipboard.GetText + Mid(Text1.Text, Text1.SelStart + 1)
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
End If
If Button = vbRightButton Then
   LockWindowUpdate Text1.hWnd
   Text1.Enabled = False
   DoEvents
   Text1.Enabled = True
   PopupMenu frm.Opciones
   LockWindowUpdate 0&
End If
End Sub

Private Sub Too_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Val(Button.Tag)
Case 1
    Recon
Case 2
    Ejecuta
Case 3
    Detener
Case 4
    AClipBoard
Case 5
    DeClipBoard
Case 6
    ABuffer
Case 7
    DeBuffer
Case 8
    Variable
Case 9
    Graba
Case 10
    Recupera
Case 11
    HazSelect
    SSTab1.Tab = 0
End Select
End Sub

Private Sub Graba()
com.Filter = "MiSQL (*.TS)|*.TS"
com.DialogTitle = "GRABAR COMO..."
com.ShowSave
If com.filename <> "" Then
    Open (com.filename) For Output As #1
    Print #1, Text1.Text
    Close #1
    Beep
End If
End Sub

Private Sub Recupera()
Dim Arr As Variant, Contador As Integer
com.Filter = "MiSQL (*.TS)|*.TS"
com.DialogTitle = "CARGAR ARCHIVO"
com.ShowOpen
If File(com.filename) Then
    Text1.Text = ""
    Arr = tsleeini(com.filename)
    For Contador = 0 To UBound(Arr, 1) - 1
        Text1.Text = Text1.Text + CStr(Arr(Contador)) + Chr(13) + Chr(10)
    Next
    SSTab1.Tab = 0
    Beep
End If
End Sub

Private Sub o_Click(Index As Integer)
Dim Buffer As String, Veces As Integer, Contador As Integer, Buffer2 As String
Dim Parte1 As String, Parte2 As String, Posicion As Integer
Select Case Index
Case 0
    Clipboard.Clear
    If Text1.SelLength > 0 Then
        Clipboard.SetText Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)
    Else
        Clipboard.SetText Text1.Text
    End If
Case 1
    Text1.Text = Text1.Text + Chr(13) + Chr(10) + Clipboard.GetText
Case 2
    If MsgBox("Desea borrar el texto?", vbYesNo, "QueryBuilder") = vbYes Then
        If Text1.SelLength > 0 Then
            Text1.SetFocus
            SendKeys "{DEL}", True
        Else
            Text1.Text = ""
        End If
    End If
Case 4
    Text1.Text = "select" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "from " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "where " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "order by " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "group by " + Chr(13) + Chr(10)
Case 5
    Text1.SetFocus
    Posicion = Text1.SelStart
    Parte1 = Mid(Text1.Text, 1, Text1.SelStart)
    Parte2 = Mid(Text1.Text, Text1.SelStart + 1)
    For Contador = Len(Parte1) To 1 Step -1
        If Mid(Parte1, Contador, 1) = "," Then
            Parte1 = Mid(Parte1, 1, Contador - 1)
            Exit For
        End If
    Next
    Text1.Text = Parte1 + IIf(Parte2 = Chr(13) + Chr(10), "", Parte2)
    If Posicion > 1 Then
        Text1.SelStart = Posicion - 1
    End If
Case 7
    If Not TreeView1.SelectedItem.Child Is Nothing Then
        Veces = 0
        TreeView1.SetFocus
        Buffer = StrTran(TreeView1.SelectedItem.FullPath, "\", ".")
        Buffer2 = ""
        Text1.Text = Text1.Text + "insert into " + Buffer + Chr(13) + Chr(10) + "(" + Chr(13) + Chr(10)
        DoEvents
        SendKeys "{DOWN}", True
        Do While TreeView1.SelectedItem.Child Is Nothing
            Text1.Text = Text1.Text + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + ", " + Chr(13) + Chr(10)
            Buffer2 = Buffer2 + "'" + Chr(34) + "+ " + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + " + " + Chr(34) + "', " + Chr(13) + Chr(10)
            SendKeys "{DOWN}", True
            Veces = Veces + 1
        Loop
        Buffer2 = Mid(Buffer2, 1, Len(Buffer2) - 4) + ")"
        Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 4) + Chr(13) + Chr(10) + ")" + Chr(13) + Chr(10)
        Text1.Text = Text1.Text + " values " + Chr(13) + Chr(10) + "(" + Chr(13) + Chr(10) + Buffer2
    End If
Case 8
    If Not TreeView1.SelectedItem.Child Is Nothing Then
        TreeView1.SetFocus
        Buffer = StrTran(TreeView1.SelectedItem.FullPath, "\", ".")
        Text1.Text = Text1.Text + "Update " + Buffer + " " + Chr(13) + Chr(10) + "set " + Chr(13) + Chr(10)
        DoEvents
        SendKeys "{DOWN}", True
        Do While TreeView1.SelectedItem.Child Is Nothing
            Text1.Text = Text1.Text + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + " = '" + Chr(34) + " + " + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + " + " + Chr(34) + "', " + Chr(13) + Chr(10)
            SendKeys "{DOWN}", True
        Loop
        Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 4) + Chr(13) + Chr(10) + "where " + Chr(13) + Chr(10)
    End If
Case 9
    If Not TreeView1.SelectedItem.Child Is Nothing Then
        TreeView1.SetFocus
        Buffer = StrTran(TreeView1.SelectedItem.FullPath, "\", ".")
        Text1.Text = Text1.Text + "select " + Chr(13) + Chr(10)
        DoEvents
        SendKeys "{DOWN}", True
        Do While TreeView1.SelectedItem.Child Is Nothing
            Text1.Text = Text1.Text + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + ", " + Chr(13) + Chr(10)
            SendKeys "{DOWN}", True
        Loop
        Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 4) + Chr(13) + Chr(10)
        Text1.Text = Text1.Text + "from " + Buffer + Chr(13) + Chr(10) + "where " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "like '%'"
    End If
Case 10
    If Not TreeView1.SelectedItem.Child Is Nothing Then
        Text1.Text = Text1.Text + "delete from " + StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + Chr(13) + Chr(10) + "where    =''"
    End If
Case 12
    Text1.SetFocus
    SendKeys "where " + Chr(vbKeyReturn), True
Case 13
    Text1.SetFocus
    SendKeys "=''{LEFT}", True
Case 14
    Text1.SetFocus
    SendKeys "{ENTER}", True
Case 15
    Text1.SetFocus
    SendKeys "Order by " + Chr(13) + Chr(10), True
Case 16
    Text1.SetFocus
    SendKeys "Group by " + Chr(13) + Chr(10), True
Case 18
    TreeView1.SetFocus
    Do While TreeView1.SelectedItem.Child Is Nothing
        Text1.SetFocus
        SendKeys StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + ", " + Chr(vbKeyReturn), True
        TreeView1.SetFocus
        SendKeys "{DOWN}", True
    Loop
Case 19
    Text1.SetFocus
    SendKeys StrTran(TreeView1.SelectedItem.FullPath, "\", ".") + ", " + Chr(vbKeyReturn), True
    TreeView1.SetFocus
Case 20
    Text1.Text = Text1.Text + "if exists (select * from sysobjects where id = object_id('dbo.tabla1') and sysstat & 0xf = 3) drop table dbo.tabla1" + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + "CREATE TABLE dbo.tabla1 (" + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + "campo1 varchar (100) NOT NULL ," + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + "campo2 numeric(8, 0) NOT NULL ," + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + "campo3 varchar (2) NOT NULL )" + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + "CREATE  UNIQUE  CLUSTERED  INDEX idx_tabla11 ON dbo.tabla1(campo1, campo2)" + Chr(13) + Chr(10)
End Select
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
If X < SSTab1.Left Then
    Source.Left = SSTab1.Left - Source.Width
Else
    Source.Left = X
End If
TreeView1.Width = Source.Left + Source.Width
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
Source.Left = SSTab1.Left + Text1.Left + X
TreeView1.Width = Source.Left + Source.Width
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 5 Then
    Ejecuta
End If
End Sub

Private Sub Form_Resize()
If Me.Visible And Me.WindowState = 0 Then
    Me.Left = 0
    Me.Width = Screen.Width
    Aderecha
End If
End Sub

Private Sub Aderecha()
SSTab1.Width = Screen.Width - TreeView1.Width - Command1.Width
Frame4.Width = Screen.Width - TreeView1.Width - Command1.Width
SSTab1.Tab = 0
Text1.Width = SSTab1.Width - Text1.Left - 500
SSTab1.Tab = 1
Text3.Width = SSTab1.Width - Text3.Left - 500
SSTab1.Tab = 2
Frame3.Width = SSTab1.Width - Frame3.Left - 500
Text5.Width = SSTab1.Width - Text5.Left - 750
Frame2.Left = SSTab1.Width - Frame2.Width - 500
SSTab1.Tab = 0
End Sub


