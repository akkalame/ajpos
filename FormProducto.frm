VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormProducto 
   BackColor       =   &H8000000A&
   Caption         =   "Producto - AJ POS"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "rsCategoria"
      Height          =   285
      Left            =   8280
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc rsCategoria 
      Height          =   330
      Left            =   7560
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "categoriaProducto"
      Caption         =   "Rs Categoria"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   5760
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "producto"
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
   Begin VB.CheckBox mantieneStockCheck 
      Caption         =   "Mantiene Stock"
      DataField       =   "mantiene_stock"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   5520
      Width           =   2775
   End
   Begin VB.ComboBox categoriaCmb 
      DataField       =   "categoria_producto"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton nuevoItemBtn 
      Caption         =   "Crear Producto"
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox imagenPic 
      Height          =   2175
      Left            =   5760
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   18
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton buscarBtn 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox buscarTxt 
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
      Left            =   720
      TabIndex        =   16
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton ultimoBtn 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton primeroBtn 
      Caption         =   "Primero"
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton anteriorBtn 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton siguienteBtn 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton guardarBtn 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox costoTxt 
      DataField       =   "costo"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox preciotxt 
      DataField       =   "precio"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox descripcionTxt 
      DataField       =   "descripcion"
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
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox nombreTxt 
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox codigoTxt 
      DataField       =   "Id_producto"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Costo"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Precio"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Categoría"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Descripción"
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
      Left            =   240
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
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Código"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FormProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nombre, descripcion, categoria, costo, precio As String
Dim mantStock As Integer
Sub set_datos_registro()
    nombre = Adodc1.Recordset.Fields(1)
    descripcion = Adodc1.Recordset.Fields(2)
    precio = Adodc1.Recordset.Fields(3)
    costo = Adodc1.Recordset.Fields(4)
    categoria = Adodc1.Recordset.Fields(5)
    mantStock = Abs(Val(Adodc1.Recordset.Fields(6)))
End Sub

Sub cargar_categorias()
    categoriaCmb.Clear
    rsCategoria.Recordset.MoveFirst
    
    While rsCategoria.Recordset.EOF = False
        categoriaCmb.AddItem rsCategoria.Recordset.Fields(1)
        rsCategoria.Recordset.MoveNext
    Wend
End Sub
Sub activar_guardar()
    Dim activar As Boolean
    activar = False
    
    If nombreTxt.Text <> nombre Then
        activar = True
    ElseIf descripcionTxt.Text <> descripcion Then
        activar = True
    ElseIf preciotxt.Text <> precio Then
        activar = True
    ElseIf costoTxt.Text <> costo Then
        activar = True
    ElseIf categoriaCmb.Text <> categoria Then
        activar = True
    ElseIf mantieneStockCheck.Value <> mantStock Then
        activar = True
    End If
    
    guardarBtn.Enabled = activar
End Sub

Private Sub anteriorBtn_Click()
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then
        MsgBox "Ha llegado al primer registro", vbOKOnly + vbInformation, "Notificacion"
        Adodc1.Recordset.MoveFirst
    End If
End Sub

Sub setear_registro()
    codigoTxt.Text = Adodc1.Recordset.Fields(0)
    nombreTxt.Text = Adodc1.Recordset.Fields(1)
    descripcionTxt.Text = Adodc1.Recordset.Fields(2)
    preciotxt.Text = Adodc1.Recordset.Fields(3)
    costoTxt.Text = Adodc1.Recordset.Fields(4)
    categoriaCmb.Text = Adodc1.Recordset.Fields(5)
    mantieneStockCheck.Value = Abs(Val(Adodc1.Recordset.Fields(6)))
End Sub
Private Sub buscarBtn_Click()
    Dim encontrado As Boolean
    encontrado = False
    
    Adodc1.Recordset.MoveFirst
    While encontrado = False And Adodc1.Recordset.EOF = False
        If Adodc1.Recordset.Fields(0) = buscarTxt.Text Then
            encontrado = True
        Else
            Adodc1.Recordset.MoveNext
        End If
    Wend
    
    If encontrado Then
        setear_registro
    End If
End Sub

Private Sub categoriaCmb_Change()
    activar_guardar
End Sub

Private Sub costoTxt_Change()
    activar_guardar
End Sub

Private Sub descripcionTxt_Change()
    activar_guardar
End Sub

Private Sub Form_Load()
    Adodc1.Refresh
    rsCategoria.Refresh
    
    guardarBtn.Enabled = False
    codigoTxt.Enabled = False
    Adodc1.Visible = False
    rsCategoria.Visible = False
    cargar_categorias
    setear_registro
    set_datos_registro
End Sub

Private Sub guardarBtn_Click()
    Adodc1.Recordset.UpdateBatch
End Sub

Private Sub mantieneStockCheck_Click()
    activar_guardar
End Sub

Private Sub nombreTxt_Change()
    activar_guardar
End Sub

Private Sub nuevoItemBtn_Click()
    FormNuevoProducto.Show
End Sub

Private Sub preciotxt_Change()
    activar_guardar
End Sub

Private Sub primeroBtn_Click()
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub siguienteBtn_Click()
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF = True Then
        MsgBox "Ha llegado al ultimo registro", vbOKOnly + vbInformation, "Notificacion"
        Adodc1.Recordset.MoveLast
    End If
End Sub

Private Sub ultimoBtn_Click()
    Adodc1.Recordset.MoveLast
End Sub
