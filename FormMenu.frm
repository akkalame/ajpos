VERSION 5.00
Begin VB.Form FormMenu 
   Caption         =   "AJ POS"
   ClientHeight    =   4920
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu configuracion 
         Caption         =   "Configuracion"
      End
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
      Begin VB.Menu categoria_producto 
         Caption         =   "Categoria de Producto"
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
      Begin VB.Menu recepcion 
         Caption         =   "Recepcion"
      End
      Begin VB.Menu pago 
         Caption         =   "Pago"
      End
      Begin VB.Menu movimiento_inventario 
         Caption         =   "Movimiento de Inventario"
      End
   End
   Begin VB.Menu informes 
      Caption         =   "Informes"
      Begin VB.Menu detalle_ventas 
         Caption         =   "Detalle de Ventas"
      End
      Begin VB.Menu resumen_ventas 
         Caption         =   "Resumen de Ventas"
      End
      Begin VB.Menu movimiento_inventario_informe 
         Caption         =   "Movimiento de Inventario"
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
