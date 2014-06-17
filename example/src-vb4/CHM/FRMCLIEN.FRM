VERSION 4.00
Begin VB.Form frmClientsCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Clientes"
   ClientHeight    =   9045
   ClientLeft      =   -4185
   ClientTop       =   2535
   ClientWidth     =   13755
   Height          =   9450
   Icon            =   "FRMCLIEN.frx":0000
   Left            =   -4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   Top             =   2190
   Width           =   13875
   Begin VB.TextBox Text1 
      DataField       =   "ID_CLIENTE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "NOMBRE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdeleted 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   12120
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   12840
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Cliente"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\NETCASH.MDB"
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CLIENTES"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FRMCLIEN.frx":0442
      Height          =   7335
      Left            =   120
      OleObjectBlob   =   "FRMCLIEN.frx":0450
      TabIndex        =   10
      Top             =   1560
      Width           =   13455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9240
      Picture         =   "FRMCLIEN.frx":0B58
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmClientsCat"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdClosed_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    Data1.Recordset.Delete
    Data1.Refresh
    DBGrid1.Refresh
    mg = MsgBox("El Cliente fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Cliente.", vbCritical, PRODUCT_NAME)

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
    mg = MsgBox("Digite la clave y descripción del nuevo Cliente y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
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
    ms = MsgBox("No se puede guardar el Cliente, ya existe.", vbCritical, PRODUCT_NAME)

End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


