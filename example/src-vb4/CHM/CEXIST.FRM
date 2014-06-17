VERSION 4.00
Begin VB.Form frmCExist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias de Almacén"
   ClientHeight    =   8760
   ClientLeft      =   2130
   ClientTop       =   1860
   ClientWidth     =   8040
   Height          =   9270
   Icon            =   "CEXIST.frx":0000
   Left            =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Top             =   1410
   Width           =   8160
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7815
      _version        =   65536
      _extentx        =   13785
      _extenty        =   4260
      _stockprops     =   15
      caption         =   "Detalle de Artículo"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   1
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Detalle de Artículo"
      tab(0).controlcount=   17
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label9"
      tab(0).control(1)=   "Label6"
      tab(0).control(2)=   "Label5"
      tab(0).control(3)=   "Label1"
      tab(0).control(4)=   "Label2"
      tab(0).control(5)=   "Text9"
      tab(0).control(6)=   "Text6"
      tab(0).control(7)=   "Text5"
      tab(0).control(8)=   "Text1"
      tab(0).control(9)=   "Text2"
      tab(0).control(10)=   "cmdLast"
      tab(0).control(11)=   "cmdNext"
      tab(0).control(12)=   "cmdPrevious"
      tab(0).control(13)=   "cmdFirst"
      tab(0).control(14)=   "cmdView"
      tab(0).control(15)=   "cmdSeek"
      tab(0).control(16)=   "cmdClosed"
      Begin VB.CommandButton cmdClosed 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   6720
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdSeek 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "Ver Listado"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<<"
         Height          =   375
         Left            =   6480
         TabIndex        =   13
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   375
         Left            =   6840
         TabIndex        =   12
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text2 
         DataField       =   "CODIGO_BARRAS"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0000FFFF&
         DataField       =   "CLAVE"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         DataField       =   "ARTICULO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox Text6 
         DataField       =   "MARCA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         DataField       =   "EXISTENCIAS"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Barras"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Marca"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Existencias"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Data OALMACEN 
      Caption         =   "ALMACEN"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "ALMACEN"
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CEXIST.frx":0442
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "CEXIST.frx":0453
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "frmCExist"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCExist.Hide
    Unload frmCExist
End Sub




Private Sub cmdFirst_Click()

On Error GoTo LocalHandlerMoveFirst
    
    OALMACEN.Recordset.MoveFirst
    
LocalHandlerMoveFirst:
    
End Sub

Private Sub cmdLast_Click()
    
On Error GoTo LocalHandlerMoveLast

    OALMACEN.Recordset.MoveLast
    
LocalHandlerMoveLast:

End Sub



Private Sub cmdNext_Click()
    
On Error GoTo LocalHandlerMoveNext
    
    OALMACEN.Recordset.MoveNext
    Exit Sub
    
LocalHandlerMoveNext:
    
End Sub

Private Sub cmdPrevious_Click()

On Error GoTo LocalHandlerMovePrev

    OALMACEN.Recordset.MovePrevious
    Exit Sub
    
LocalHandlerMovePrev:
    
End Sub














Private Sub cmdSeek_Click()
    OALMACEN.Recordset.FindFirst ("CLAVE='" + Text1.Text + "'")
End Sub

Private Sub cmdView_Click()

cmdView.Visible = False
DBGrid1.Visible = True

    For x = 0 To 18
        DBGrid1.Columns.Remove (4)
    Next x
    
'Me.Height = 6030
    
End Sub

Private Sub Form_Load()
    OALMACEN.DatabaseName = NetCashFile
End Sub


Private Sub Text1_LostFocus()

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = Text1.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    Text1.Text = sTMP


End Sub


Private Sub Text2_LostFocus()

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = Text2.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    Text2.Text = sTMP


End Sub


