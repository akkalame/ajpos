VERSION 5.00
Begin VB.Form FormProveedor 
   BackColor       =   &H8000000A&
   Caption         =   "Proveedor - AJ POS"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Crear Proveedor"
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   6720
      TabIndex        =   22
      Top             =   9240
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
      Left            =   4200
      TabIndex        =   21
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   9120
      TabIndex        =   20
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Primero"
      Height          =   495
      Left            =   7320
      TabIndex        =   19
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox sitiowebtxt 
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
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox emailtxt 
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
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox telefonotxt 
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
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox direcciontxt 
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
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox contactotxt 
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
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox dnitxt 
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
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox nombretxt 
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
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox idtxt 
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
      Top             =   480
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
      Top             =   6600
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
      Top             =   5640
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
      Top             =   4680
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
      Top             =   3840
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
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DNI"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1320
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
      Top             =   480
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
