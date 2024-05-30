VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormNuevoProducto 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Producto - AJ POS"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "rsCategoria"
      Height          =   285
      Left            =   8760
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
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
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   3495
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
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
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
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
      Left            =   2640
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox categoriaCmb 
      DataField       =   "categoria_producto"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CheckBox mantieneStockCheck 
      Caption         =   "Mantiene Stock"
      DataField       =   "mantiene_stock"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.PictureBox imagenPic 
      Height          =   2175
      Left            =   6840
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton guardarBtn 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cancelBtn 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc rsCategoria 
      Height          =   330
      Left            =   8040
      Top             =   4200
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
      Connect         =   $"FormNuevoProducto.frx":0000
      OLEDBString     =   $"FormNuevoProducto.frx":0089
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
      Left            =   1080
      TabIndex        =   13
      Top             =   360
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
      Left            =   1080
      TabIndex        =   12
      Top             =   1200
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
      Left            =   1080
      TabIndex        =   11
      Top             =   2040
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "FormNuevoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cargar_categorias()
    categoriaCmb.Clear
    rsCategoria.Recordset.MoveFirst
    
    While rsCategoria.Recordset.EOF = False
        categoriaCmb.AddItem rsCategoria.Recordset.Fields(1)
        rsCategoria.Recordset.MoveNext
    Wend
End Sub

Sub guardar_registro()
    FormProducto.Adodc1.Recordset.AddNew
    FormProducto.Adodc1.Recordset.Fields(1) = nombretxt.Text
    FormProducto.Adodc1.Recordset.Fields(2) = descripcionTxt.Text
    FormProducto.Adodc1.Recordset.Fields(3) = categoriaCmb.Text
    FormProducto.Adodc1.Recordset.Fields(4) = preciotxt.Text
    FormProducto.Adodc1.Recordset.Fields(5) = costoTxt.Text
    FormProducto.Adodc1.Recordset.Fields(6) = mantieneStockCheck.Value
    FormProducto.Adodc1.Recordset.Update
End Sub
Private Sub cancelBtn_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    cargar_categorias
    rsCategoria.Visible = False
    FormProducto.Adodc1.Refresh
End Sub

Private Sub guardarBtn_Click()
    guardar_registro
    
    Me.Hide
    FormProducto.Adodc1.Recordset.MoveLast
End Sub

