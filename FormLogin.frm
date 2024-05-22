VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormLogin 
   Caption         =   "Login - AJ POS"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox dataShowTxt 
      DataField       =   "usuario"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton loginBtn 
      Caption         =   "Iniciar Sesion"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2280
      Top             =   3840
      Width           =   2775
      _ExtentX        =   4895
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
      Connect         =   $"FormLogin.frx":0000
      OLEDBString     =   $"FormLogin.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "usuario"
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
   Begin VB.TextBox claveTxt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox usuarioTxt 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "(c) Akkalame Ereut y Juan Hernandez. devakkalame@gmail.com"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Adodc1.Visible = False
    dataShowTxt.Visible = False
End Sub

Private Sub loginBtn_Click()
    Dim loginExitoso As Boolean
    loginExitoso = False
    
    Adodc1.Recordset.MoveFirst
    While loginExitoso = False And Adodc1.Recordset.EOF = False
        If Adodc1.Recordset.Fields(0) = usuarioTxt.Text And Adodc1.Recordset.Fields(1) = claveTxt.Text Then
            loginExitoso = True
        Else
            Adodc1.Recordset.MoveNext
        End If
    Wend
    
    If loginExitoso = True Then
        Me.Hide
        
    Else
        MsgBox "Credenciales invalidas", vbOKOnly + vbCritical
    End If
End Sub
