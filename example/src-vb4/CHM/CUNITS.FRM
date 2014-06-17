VERSION 4.00
Begin VB.Form frmCUnits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Unidades"
   ClientHeight    =   7035
   ClientLeft      =   1560
   ClientTop       =   2460
   ClientWidth     =   6270
   Height          =   7545
   Icon            =   "CUNITS.frx":0000
   Left            =   1500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Top             =   2010
   Width           =   6390
   Begin VB.Data Data1 
      Caption         =   "Unidades"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.DOT"
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CUNIDADES"
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
      DataField       =   "DESCRIP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_UNIDAD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "CUNITS.frx":0442
      Top             =   600
      Width           =   480
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CUNITS.frx":0884
      Height          =   4500
      Left            =   120
      OleObjectBlob   =   "CUNITS.frx":0892
      TabIndex        =   12
      Top             =   2400
      Width           =   6000
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción de Unidades:"
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
Attribute VB_Name = "frmCUnits"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCUnits.Hide
    Unload frmCUnits
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    Data1.Recordset.Delete
    Data1.Refresh
    DBGrid1.Refresh
    mg = MsgBox("La Unidad fue Eliminada.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar la Unidad.", vbCritical, PRODUCT_NAME)
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
    mg = MsgBox("Digite la clave y descripción de la nueva Unidad y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
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
    ms = MsgBox("No se puede guardar la Unidad ya existe.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


