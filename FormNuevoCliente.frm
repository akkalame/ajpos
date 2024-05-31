VERSION 5.00
Begin VB.Form FormNuevoCliente 
   BackColor       =   &H8000000A&
   Caption         =   "Nuevo Cliente -AJ POS"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
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
      Left            =   6360
      TabIndex        =   9
      Top             =   4200
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
      Left            =   6360
      TabIndex        =   8
      Top             =   3240
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
      Left            =   6360
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox cedulatxt 
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
      Left            =   6360
      TabIndex        =   6
      Top             =   1320
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
      Left            =   6360
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
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
      Left            =   4800
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Left            =   4800
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FormNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

End Sub
