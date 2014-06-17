VERSION 4.00
Begin VB.Form frmCProvider 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Proveedores"
   ClientHeight    =   9015
   ClientLeft      =   -4065
   ClientTop       =   9495
   ClientWidth     =   13695
   Height          =   9420
   Icon            =   "CPROV.frx":0000
   Left            =   -4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   Top             =   9150
   Width           =   13815
   Begin VB.Data Data1 
      Caption         =   "Proveedor"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PROVEEDORES"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   13200
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   12840
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   12480
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   12120
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdeleted 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "NOMBRE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_PROVEEDOR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9240
      Picture         =   "CPROV.frx":0442
      Top             =   480
      Width           =   480
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CPROV.frx":0884
      Height          =   7335
      Left            =   120
      OleObjectBlob   =   "CPROV.frx":0892
      TabIndex        =   12
      Top             =   1560
      Width           =   13455
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmCProvider"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCProvider.Hide
    Unload frmCProvider
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    Data1.Recordset.Delete
    Data1.Refresh
    DBGrid1.Refresh
    mg = MsgBox("El Proveedor fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Proveedor.", vbCritical, PRODUCT_NAME)
End Sub

Private Sub cmdFirst_Click()

On Error GoTo LocalHandlerMoveFirst
    
    Data1.Recordset.MoveFirst
    
LocalHandlerMoveFirst:
    
End Sub

Private Sub cmdLast_Click()
    
On Error GoTo LocalHandlerMoveLast

    Data1.Recordset.MoveLast
    
LocalHandlerMoveLast:

End Sub

Private Sub cmdNew_Click()

On Error GoTo LocalHandlerNew

    Data1.Refresh
    Data1.Recordset.AddNew
    mg = MsgBox("Digite la clave y descripción del nuevo Proveedor y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
LocalHandlerNew:

End Sub

Private Sub cmdNext_Click()
    
On Error GoTo LocalHandlerMoveNext
    
    Data1.Recordset.MoveNext
    Exit Sub
    
LocalHandlerMoveNext:
    
End Sub

Private Sub cmdPrevious_Click()

On Error GoTo LocalHandlerMovePrev

    Data1.Recordset.MovePrevious
    Exit Sub
    
LocalHandlerMovePrev:
    
End Sub

Private Sub cmdSave_Click()

On Error GoTo LocalHandlerSave

    Data1.Recordset.Update
    Data1.Refresh
    DBGrid1.Refresh
    Exit Sub
    
    
LocalHandlerSave:
    ms = MsgBox("No se puede guardar el Proveedor, ya existe.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


