VERSION 4.00
Begin VB.Form frmCFloor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Pasillos"
   ClientHeight    =   7410
   ClientLeft      =   1275
   ClientTop       =   1980
   ClientWidth     =   6270
   Height          =   7920
   Icon            =   "CFLOO.frx":0000
   Left            =   1215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Top             =   1530
   Width           =   6390
   Begin VB.Data Data1 
      Caption         =   "Pasillos"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CPASILLO"
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdeleted 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "DESCRIP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_PASILLO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "CFLOO.frx":0442
      Top             =   1440
      Width           =   480
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CFLOO.frx":0884
      Height          =   4500
      Left            =   120
      OleObjectBlob   =   "CFLOO.frx":0892
      TabIndex        =   12
      Top             =   2400
      Width           =   6000
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción de Pasillos:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Clave:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmCFloor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCFloor.Hide
    Unload frmCFloor
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    Data1.Recordset.Delete
    Data1.Refresh
    DBGrid1.Refresh
    mg = MsgBox("El Pasillo fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Pasillo.", vbCritical, PRODUCT_NAME)
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
    mg = MsgBox("Digite la clave y descripción del nuevo Pasillo y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
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
    ms = MsgBox("No se puede guardar el Pasillo, ya existe.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


