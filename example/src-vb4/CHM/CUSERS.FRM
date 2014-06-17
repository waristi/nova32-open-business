VERSION 4.00
Begin VB.Form frmCUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Usuarios"
   ClientHeight    =   3765
   ClientLeft      =   1305
   ClientTop       =   1710
   ClientWidth     =   5340
   Height          =   4275
   Icon            =   "CUSERS.frx":0000
   Left            =   1245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Top             =   1260
   Width           =   5460
   Begin VB.TextBox Text4 
      DataField       =   "MODULO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "PERFIL"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Top             =   1440
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Usuarios"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   -1  'True
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "USUARIOS"
      Top             =   240
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
      DataField       =   "CLAVE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "LOGIN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3840
      Picture         =   "CUSERS.frx":0442
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Módulo"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Perfil"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CUSERS.frx":0884
      Height          =   1215
      Left            =   120
      OleObjectBlob   =   "CUSERS.frx":0892
      TabIndex        =   12
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmCUsers"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCUsers.Hide
    Unload frmCUsers
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    If Text1.Text = "ADMINISTRADOR" Then
        mg = MsgBox("No se puede eliminar la Cuenta Administrador, Cuenta del Sistema privilegiada, operación anulada. ", vbCritical, PRODUCT_NAME)
        Exit Sub
    End If


    Data1.Recordset.Delete
    Data1.Refresh
    DBGrid1.Refresh
        
    mg = MsgBox("El Usuario fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Usuario.", vbCritical, PRODUCT_NAME)
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
    mg = MsgBox("Digite la Cuenta del nuevo Usuario y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
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
    ms = MsgBox("No se puede guardar el Usuario, ya existe.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


