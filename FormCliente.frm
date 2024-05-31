VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormCliente 
   BackColor       =   &H8000000A&
   Caption         =   "Cliente - AJ POS"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9480
      Top             =   6960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Hernandez\Desktop\ajpos\database\pos_project.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Hernandez\Desktop\ajpos\database\pos_project.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "cliente"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton crearclientecmd 
      Caption         =   "Crear Cliente"
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton buscarBtn 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox buscartxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton ultimocmd 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton primerocmd 
      Caption         =   "Primero"
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton anteriorcmd 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton siguientecmd 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton guardarcmd 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox emailtxt 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox telefonotxt 
      DataField       =   "telefono"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox direcciontxt 
      DataField       =   "direccion"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox cedulatxt 
      DataField       =   "dni"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox nombretxt 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox idtxt 
      DataField       =   "id_cliente"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FormCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub mostrardatos()
Text1.Text = Adodc1.Recordset.Fields(0)
Text2.Text = Adodc1.Recordset.Fields(1)
End Sub

Private Sub anteriorcmd_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
MsgBox "Primer registro", vbOK + vbInformation, "Advetencia"
If vbOK Then
Adodc1.Recordset.MoveFirst
End If
End If
End Sub

Private Sub buscarBtn_Click()
Dim ok As String
Dim encontrado As Boolean
ok = Text4.Text
If ok <> "" Then
 Adodc1.Recordset.MoveFirst
encontrado = False
While (Adodc1.Recordset.EOF = False) And (encontrado = False)
 If Adodc1.Recordset.Fields(0) = ok Then
   encontrado = True
   MsgBox " registro encontrado", vbOKOnly + vbInformation, "titulo"
    mostrardatos
Else
Adodc1.Recordset.MoveNext
End If
Wend
If (Adodc1.Recordset.EOF = True) And (encontrado = False) Then
 MsgBox "registro no encontrado", vbOK + vbCritical, "advertencia"
 Text3.Text = ""
 Text4.SetFocus
End If
End If
End Sub

Private Sub crearclientecmd_Click()
FormNuevoCliente.Show
End Sub

Private Sub Form_Load()
Adodc1.Visible = False
idtxt.Enabled = False
End Sub

Private Sub guardarcmd_Click()
Dim resp As Integer
If Text1.Text = "" Then
MsgBox "Colocar nombre"
Text1.SetFocus
Exit Sub
ElseIf Text2.Text = "" Then
MsgBox "Colocar cédula"
Text2.SetFocus
Exit Sub
ElseIf incluir Then
resp = MsgBox("Desea guardar?", vbOKCancel + vbQuestion, "Advertencia")
If resp = vbOK Then
Adodc1.Recordset.UpdateBatch
Else
Adodc1.Recordset.CancelUpdate
End If
incluir = False
End If
If modificar Then
MsgBox "Desea modificar?", vbOKCancel + vbQuestion, "Advertencia"
If vbOK Then
Adodc1.Recordset.UpdateBatch
Else
Adodc1.Recordset.CancelUpdate
End If
modificar = False
End If
End Sub

Private Sub primerocmd_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub siguientecmd_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "Ultimo registro", vbOK + vbInformation, "Advertencia"
If vbOK Then
Adodc1.Recordset.MoveLast
End If
End If
End Sub

Private Sub ultimocmd_Click()
Adodc1.Recordset.MoveLast
End Sub
