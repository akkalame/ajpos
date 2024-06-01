VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormCliente 
   Caption         =   "Cliente - AJ POS"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormCliente.frx":0000
      Height          =   975
      Left            =   6360
      TabIndex        =   20
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox buscartxt 
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
      Left            =   2640
      TabIndex        =   19
      Top             =   7200
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7200
      Top             =   5520
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
      Left            =   6360
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton buscarBtn 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton ultimocmd 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton primerocmd 
      Caption         =   "Primero"
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton anteriorcmd 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton siguientecmd 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton guardarcmd 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   5760
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
    idtxt.Text = Adodc1.Recordset.Fields(0)
    nombretxt.Text = Adodc1.Recordset.Fields(1)
    cedulatxt.Text = Adodc1.Recordset.Fields(2)
    direcciontxt.Text = Adodc1.Recordset.Fields(3)
    telefonotxt.Text = Adodc1.Recordset.Fields(4)
    emailtxt.Text = Adodc1.Recordset.Fields(5)
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
   Dim encontrado As Boolean
    encontrado = False
    Adodc1.Recordset.MoveFirst
    
    If buscartxt.Text <> "" Then
        While (Adodc1.Recordset.EOF = False) And (encontrado = False)
            If Adodc1.Recordset.Fields(0) = buscartxt.Text Then
                encontrado = True
                MsgBox "Registro encontrado", vbOKOnly + vbInformation, "Notificacion"
                
            Else
                Adodc1.Recordset.MoveNext
            End If
        Wend
        If encontrado = False Then
            MsgBox "Registro no encontrado", vbOK + vbCritical, "Advertencia"
            buscartxt.Text = ""
            buscartxt.SetFocus
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
If nombretxt.Text = "" Then
MsgBox "Colocar Nombre"
nombretxt.SetFocus
Exit Sub
ElseIf direcciontxt.Text = "" Then
MsgBox "Colocar Dirección"
direcciontxt.SetFocus
Exit Sub
ElseIf cedulatxt.Text = "" Then
MsgBox "Colocar Cédula"
cedulatxt.SetFocus
Exit Sub
ElseIf telefonotxt.Text = "" Then
MsgBox "Colocar Teléfono"
telefonotxt.SetFocus
Exit Sub
ElseIf emailtxt.Text = "" Then
MsgBox "Colocar Email"
emailtxt.SetFocus
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

Private Sub Text1_Change()

End Sub

Private Sub ultimocmd_Click()
Adodc1.Recordset.MoveLast
End Sub
