VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormAlmacen 
   Caption         =   "Almacen - AJPOS"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   2040
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   2040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "almacen"
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
   Begin VB.CommandButton crearalmacenBtn 
      Caption         =   "Crear Almacen"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton buscarBtn 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox buscartxt 
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton ultimoBtn 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton pirmeroBtn 
      Caption         =   "Primero"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton siguienteBtn 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton anteriorBtn 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton guardarBtn 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
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
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox idtxt 
      DataField       =   "id_almacen"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
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
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FormAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub anteriorBtn_Click()
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
    Adodc1.Recordset.MoveLast
    
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

Private Sub crearalmacenBtn_Click()
FormNuevoAlmacen.Show
End Sub

Private Sub Form_Load()
idtxt.Enabled = False
Adodc1.Visible = False
Text2.Visible = False
End Sub
Private Sub pirmeroBtn_Click()
Adodc1.Recordset.MoveFirst
End Sub
Private Sub siguienteBtn_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "Ultimo registro", vbOK + vbInformation, "Advertencia"
If vbOK Then
Adodc1.Recordset.MoveLast
End If
End If
End Sub

Private Sub ultimoBtn_Click()
Adodc1.Recordset.MoveLast
End Sub
