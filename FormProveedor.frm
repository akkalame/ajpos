VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormProveedor 
   Caption         =   "Proveedor - AJ POS"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7680
      Top             =   3840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "proveedor"
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
   Begin VB.CommandButton crearproveedorBtn 
      Caption         =   "Crear Proveedor"
      Height          =   495
      Left            =   7560
      TabIndex        =   23
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton buscarBtn 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   6840
      TabIndex        =   22
      Top             =   6840
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
      Left            =   4440
      TabIndex        =   21
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton ultimoBtn 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   8520
      TabIndex        =   20
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton pirmeroBtn 
      Caption         =   "Primero"
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton anteriorBtn 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton siguienteBtn 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton guardarBtn 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox sitiowebtxt 
      DataField       =   "sitio_web"
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
      Left            =   5520
      TabIndex        =   15
      Top             =   5280
      Width           =   1695
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
      Left            =   5520
      TabIndex        =   14
      Top             =   4560
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
      Left            =   5520
      TabIndex        =   13
      Top             =   3840
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
      Left            =   5520
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox contactotxt 
      DataField       =   "contacto"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox dnitxt 
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
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
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
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox idtxt 
      DataField       =   "Id_proveedor"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Sitio Web"
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
      Left            =   3960
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      Left            =   3960
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Contacto"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
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
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Id"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FormProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()
FormNuevoProveedor.Show
End Sub

Private Sub Text5_Change()

End Sub

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

Private Sub crearproveedorBtn_Click()
FormNuevoProveedor.Show
End Sub

Private Sub Form_Load()
Adodc1.Visible = False
idtxt.Enabled = False
End Sub

Private Sub guardarBtn_Click()
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
ElseIf sitiowebtxt.Text = "" Then
MsgBox "Colocar Sitio Web"
sitiowebtxt.SetFocus
Exit Sub
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
