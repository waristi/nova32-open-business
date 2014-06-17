VERSION 4.00
Begin VB.Form frmCStocks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Almacenes"
   ClientHeight    =   5505
   ClientLeft      =   1095
   ClientTop       =   1155
   ClientWidth     =   4680
   Height          =   5910
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Top             =   810
   Width           =   4800
   Begin VB.Data Data1 
      Caption         =   "Almacen"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.DOT"
      Exclusive       =   -1  'True
      Height          =   270
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CALMACEN"
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdeleted 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nuevo"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "DESCRIP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_ALMACEN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FRMCPROV.frx":0000
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "FRMCPROV.frx":000E
      TabIndex        =   12
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion de Almacen:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Clave:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmCStocks"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCStocks.Hide
    Unload frmCStocks
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    Data1.Recordset.Delete
    mg = MsgBox("El Almacen fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Almacen.", vbCritical, PRODUCT_NAME)
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

    Data1.Recordset.AddNew
    mg = MsgBox("Digite la clave y descripcion del nuevo Almacen y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
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

    Data1.UpdateRecord
    Exit Sub
    
    
LocalHandlerSave:
    ms = MsgBox("No se puede guardar el Almacen ya existe.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


