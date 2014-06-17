VERSION 4.00
Begin VB.Form frmOnSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador de Descuentos"
   ClientHeight    =   5265
   ClientLeft      =   3180
   ClientTop       =   2325
   ClientWidth     =   5535
   Height          =   5775
   Icon            =   "FRMONSAL.frx":0000
   Left            =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Top             =   1875
   Width           =   5655
   Begin VB.Data CFAMILIAS 
      Caption         =   "CFAMILIAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\NETCASH.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "CFAMILIAS"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data ADMDESC 
      Caption         =   "ADMDESC"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\NETCASH.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ADMDESC"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Descuento General"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   5295
      Begin VB.CommandButton Command5 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_Todo 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chk_Todo 
         Caption         =   "Aplicar descuento general a todo el inventario."
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descuentos por Grupos"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt_Descuento 
         DataField       =   "DESCUENTO"
         DataSource      =   "ADMDESC"
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   2415
      End
      Begin MSDBCtls.DBCombo cbo_Grupo 
         Bindings        =   "FRMONSAL.frx":0442
         DataField       =   "GRUPO"
         DataSource      =   "ADMDESC"
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   2415
         _version        =   65536
         _extentx        =   4260
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   16777215
         listfield       =   "DESCRIP"
         boundcolumn     =   "DESCRIP"
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4320
         Picture         =   "FRMONSAL.frx":0454
         Top             =   2160
         Width           =   480
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FRMONSAL.frx":0896
         Height          =   1575
         Left            =   240
         OleObjectBlob   =   "FRMONSAL.frx":08A6
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOnSales"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo LocalHandlerNew

    ADMDESC.Refresh
    ADMDESC.Recordset.AddNew
    mg = MsgBox("Seleccione Grupo y Descuento del nuevo registro y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
LocalHandlerNew:

End Sub

Private Sub Command2_Click()

On Error GoTo LocalHandlerSave

    ADMDESC.Recordset.Update
    ADMDESC.Refresh
    DBGrid1.Refresh
    Exit Sub
    
    
LocalHandlerSave:
    ms = MsgBox("No se puede guardar el descuento, ya existe.", vbCritical, PRODUCT_NAME)

End Sub


Private Sub Command3_Click()


On Error GoTo LocalHandlerDel

    ADMDESC.Recordset.Delete
    ADMDESC.Refresh
    DBGrid1.Refresh
    mg = MsgBox("El Descuento fue Eliminado.", vbOKOnly, PRODUCT_NAME)
    cbo_Grupo.Text = ""
    txt_Descuento.Text = "0"
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Descuento.", vbCritical, PRODUCT_NAME)

End Sub

Private Sub Command4_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub Command5_Click()

Dim strCheck1 As String
Dim strCheck2 As String

Dim strValue1 As String
Dim strValue2 As String
        
strCheck1 = "[ONSALE]"
strValue1 = "[VALUE]"
        
If (chk_Todo.Value = 1) Then
    strCheck2 = "ENABLED"
    strValue2 = txt_Todo.Text
Else
    strCheck2 = "DISABLED"
    strValue2 = "0"
End If
        
    Open "C:\CHM\ONSALE.INI" For Output As #1
            Print #1, strCheck1
            Print #1, strCheck2
            Print #1, strValue1
            Print #1, strValue2
    Close #1
    
    
x = MsgBox("Descuento General aplicado.", 64, PRODUCT_NAME)
    
End Sub

Private Sub Form_Load()

Dim strCheck As String
Dim strValue As String
    
    
    ADMDESC.DatabaseName = NetCashFile
    CFAMILIAS.DatabaseName = NetCashFile
    

    Open "C:\CHM\ONSALE.INI" For Input As #1
            Input #1, strCheck
            Input #1, strCheck
            Input #1, strValue
            Input #1, strValue
    Close #1
    
    If (Trim(UCase(strCheck)) = "ENABLED") Then
        chk_Todo.Value = 1
    Else
        chk_Todo.Value = 0
    End If
    
    txt_Todo.Text = strValue
        
End Sub


