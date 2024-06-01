VERSION 5.00
Begin VB.Form FormMenu 
   Caption         =   "AJ POS"
   ClientHeight    =   6225
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      Picture         =   "FormMenu.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   13155
      TabIndex        =   0
      Top             =   -1200
      Width           =   13215
   End
   Begin VB.Menu opcs 
      Caption         =   "Opciones"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu maestros 
      Caption         =   "Maestros"
      Begin VB.Menu producto 
         Caption         =   "Producto"
      End
      Begin VB.Menu cliente 
         Caption         =   "Cliente"
      End
      Begin VB.Menu proveedor 
         Caption         =   "Proveedor"
      End
      Begin VB.Menu almacen 
         Caption         =   "Almacen"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu factura 
         Caption         =   "Factura"
      End
      Begin VB.Menu pago 
         Caption         =   "Pago"
      End
   End
   Begin VB.Menu informes 
      Caption         =   "Informes"
      Begin VB.Menu resumen_ventas 
         Caption         =   "Resumen de Ventas"
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub almacen_Click()
FormAlmacen.Show
End Sub

Private Sub cliente_Click()
    FormCliente.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub producto_Click()
    FormProducto.Show
End Sub

Private Sub proveedor_Click()
    FormProveedor.Show
End Sub

Private Sub salir_Click()
    End
End Sub
