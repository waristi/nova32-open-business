VERSION 4.00
Begin VB.Form frmNetCash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "NetCash 2.0"
   ClientHeight    =   11190
   ClientLeft      =   -2520
   ClientTop       =   2115
   ClientWidth     =   15420
   Height          =   11700
   Icon            =   "CHM.frx":0000
   Left            =   -2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Top             =   1665
   Width           =   15540
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   72
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1785
      Left            =   6720
      TabIndex        =   31
      Top             =   1440
      Width           =   810
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   555
      ItemData        =   "CHM.frx":0442
      Left            =   120
      List            =   "CHM.frx":0444
      TabIndex        =   29
      Text            =   "EFECTIVO"
      Top             =   10800
      Width           =   6375
   End
   Begin VB.TextBox DISPLAY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   72
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1785
      Left            =   7440
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   1440
      Width           =   7695
   End
   Begin VB.CommandButton cmdCAN 
      Caption         =   "F7 CANC"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9840
      TabIndex        =   12
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEFECT 
      Caption         =   "F10 EFECT"
      Enabled         =   0   'False
      Height          =   615
      Left            =   13080
      TabIndex        =   23
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCASHTick 
      Caption         =   "F4 TICKE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   16
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 SALIR"
      Enabled         =   0   'False
      Height          =   615
      Left            =   14160
      TabIndex        =   15
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "F9 CONS"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12000
      TabIndex        =   14
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCLAV 
      Caption         =   "F8 CLAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10920
      TabIndex        =   13
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOpenCash 
      Caption         =   "F6 ABRE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8760
      TabIndex        =   11
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCASH 
      Caption         =   "F5 COBRA"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7680
      TabIndex        =   10
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   6720
      ScaleHeight     =   5295
      ScaleWidth      =   8415
      TabIndex        =   4
      Top             =   3480
      Width           =   8415
      Begin VB.TextBox txt_Quantity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   750
         Left            =   3960
         TabIndex        =   30
         Text            =   "1"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtCHANGE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   750
         Left            =   5280
         TabIndex        =   21
         Text            =   "000.00"
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox txtMONEY 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Left            =   5280
         TabIndex        =   2
         Text            =   "000.00"
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txtDESC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   750
         Left            =   5280
         TabIndex        =   17
         Text            =   "000.00"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtPRODUCT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Left            =   5280
         TabIndex        =   1
         Top             =   120
         Width           =   2775
      End
      Begin VB.TextBox txtSUBTOT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   750
         Left            =   5280
         TabIndex        =   6
         Text            =   "000.00"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   750
         Left            =   5280
         TabIndex        =   5
         Text            =   "000.00"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Cambio"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Efectivo"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Descuento"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Artículo"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Sub-Total"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "IVA"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   27.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
      End
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ARTICULO                          PRECIO/DESC."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   15.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Total a Pagar"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   27.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   8640
      TabIndex        =   26
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   $"CHM.frx":0446
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6600
      TabIndex        =   25
      Top             =   9600
      Width           =   8535
   End
   Begin MSGrid.Grid Grid2 
      Height          =   1095
      Left            =   11040
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
      _version        =   65536
      _extentx        =   2778
      _extenty        =   1931
      _stockprops     =   77
      backcolor       =   16777215
      mouseicon       =   "CHM.frx":04C3
   End
   Begin MSGrid.Grid Grid1 
      Height          =   9135
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
      _version        =   65536
      _extentx        =   11245
      _extenty        =   16113
      _stockprops     =   77
      forecolor       =   65280
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderstyle     =   0
      scrollbars      =   0
      gridlines       =   0   'False
      highlight       =   0   'False
      mouseicon       =   "CHM.frx":04DF
   End
   Begin VB.Label lblEmployee 
      BackColor       =   &H00000000&
      Caption         =   "Cajero: ANA HERNANDEZ PEREZ"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CAJA"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   48
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmNetCash"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public sARTICULO As String
Public sCLAVE As String
Public mPRECIO As Double
Public mIVAS As Double
Public mDESCS As Double
Public mGRUPO As String

Public sCOBAR As String
Public mCOSTO As String

Public ICursor As Integer

Public MDB As Database
Public REC As Recordset

Public mFolios As Integer
Public MSeries As String
Public mCopys As Boolean

Public mDescGrupo As Boolean
Public mDescGrupoValor As Integer

Public mDescATodo As Boolean
Public mDescATodoValor As Integer

Public Leyenda1 As String
Public Leyenda2 As String
Public Leyenda3 As String




Private Declare Function GetModuleUsage% Lib "kernel" (ByVal hModule%)

Private Sub Bascula()

Dim STR_BASCULA_ARTICULO As String
Dim STR_BASCULA_CANTIDAD As String
Dim STR_BASCULA_IMPORTE As String


    Rem --------- INVOCAR AL PROGRAMA PESADOR Y ESPERAR A QUE TERMINE ------------

    lRet& = Shell("C:\CHM\NETW86.EXE /run", vbNormal)

    While GetModuleUsage(lRet&) > 0
        DoEvents
    Wend
    

    Rem ------- IMPORTAR Articulos PESADOS PARA NetCash (CHM FORM)
    
    Open "C:\CHM\RS232.DAT" For Input As #1
    
            Input #1, STR_BASCULA_ARTICULO
            Input #1, STR_BASCULA_CANTIDAD
            Input #1, STR_BASCULA_IMPORTE
                
    Close #1
    
    
    Rem --------- VALIDAR SI EL USUARIO NO CANCELO EL PESAJE --------------------
        
        If Trim(STR_BASCULA_ARTICULO) = "0" And Trim(STR_BASCULA_CANTIDAD) = "0" And Trim(STR_BASCULA_IMPORTE) = "0" Then
            Exit Sub
        End If
    
    
    Rem -------------------------------------------------------------------------
    
    
    
    
        Rem ---------------- UPDATE GRID1 --------------------------
        Grid1.Col = 0
        Grid1.Row = ICursor
        Grid1.Text = "BASCULA"
        
        Grid1.Col = 1
        Grid1.Row = ICursor
        Grid1.Text = STR_BASCULA_ARTICULO + " CANT." + STR_BASCULA_CANTIDAD + "GRS."
                
        Grid1.Col = 2
        Grid1.Row = ICursor
        Grid1.Text = STR_BASCULA_IMPORTE
        
        Grid1.Col = 3
        Grid1.Row = ICursor
        Grid1.Text = "0"
        
        Grid1.Col = 4
        Grid1.Row = ICursor
        Grid1.Text = "0"
                
        Rem ---------------- UPDATE GRID2 --------------------------
        Grid2.Col = 0
        Grid2.Row = ICursor
        Grid2.Text = "BASCULA"
        
        Grid2.Col = 1
        Grid2.Row = ICursor
        Grid2.Text = "BASCULA"
        
        Grid2.Col = 2
        Grid2.Row = ICursor
        Grid2.Text = "1"
        
        Grid2.Col = 3
        Grid2.Row = ICursor
        Grid2.Text = "VENTA"
        
        Grid2.Col = 4
        Grid2.Row = ICursor
        Grid2.Text = Date
        
        Grid2.Col = 5
        Grid2.Row = ICursor
        Grid2.Text = STR_BASCULA_IMPORTE
        
        Grid2.Col = 6
        Grid2.Row = ICursor
        Grid2.Text = STR_BASCULA_ARTICULO + " CANT." + STR_BASCULA_CANTIDAD + "GRS."
        
        Rem ----------------------------------------------
        
        ICursor = ICursor + 1
        
        'Dim mCalIVa As Double
        'Dim mCalDesc As Double
        'mCalDesc = (REC.Fields(20).Value * CInt(txt_Quantity.Text)) * (REC.Fields(25).Value / 100)
        'mCalIVa = ((REC.Fields(20).Value * CInt(txt_Quantity.Text)) - mCalDesc) * (REC.Fields(24).Value / 100)
        
        txtSUBTOT.Text = Format(CDbl(txtSUBTOT.Text) + CDbl(STR_BASCULA_IMPORTE), "Currency")
        txtIVA.Text = Format(CDbl(txtIVA.Text) + 0, "Currency")
        txtDESC.Text = Format(CDbl(txtDESC.Text) + 0, "Currency")
        DISPLAY.Text = Format(CDbl(txtSUBTOT.Text) - CDbl(txtDESC.Text) + CDbl(txtIVA.Text), "Currency")
        
        txtPRODUCT.Text = ""
    
End Sub


Private Function TestFun(ByVal lVal As Long) As Integer

    If (lVal And &HFFFF&) = 0 Then
        TestFun = lVal And &HFFFF&
    Else
        TestFun = &H8000 Or (lVal And &H7FFF&)
    End If

End Function

Private Sub cmdCAN_Click()

Rem --------------------- CANCEL GRID1 PRODS ---------------------------

Grid1.Col = 0
Grid1.Text = "000"

Grid1.Col = 1
Grid1.Text = "CANCELADO"

Grid1.Col = 2
txtSUBTOT.Text = Format(CDbl(txtSUBTOT.Text) - CDbl(Grid1.Text), "Currency")
Grid1.Text = "0"

Grid1.Col = 3
txtIVA.Text = Format(CDbl(txtIVA.Text) - CDbl(Grid1.Text), "Currency")
Grid1.Text = "0"

Grid1.Col = 4
txtDESC.Text = Format(CDbl(txtDESC.Text) - CDbl(Grid1.Text), "Currency")
Grid1.Text = "0"

Rem --------------------- CANCEL GRID2 PRODS ---------------------------
Grid2.Row = Grid1.Row
Grid2.Col = 1
Grid2.Text = "CANCELADO"
Rem --------------------------------------------------------------------


DISPLAY.Text = Format(CDbl(txtSUBTOT.Text) - CDbl(txtDESC.Text) + CDbl(txtIVA.Text), "Currency")

End Sub


Private Sub cmdCASH_Click()

        If (txtMONEY.Text = "") Then
            txtMONEY.Text = "0"
        End If
        
        If (CDbl(txtMONEY.Text) < CDbl(DISPLAY.Text)) Then
            M = MsgBox("EL EFECTIVO ES INSUFICIENTE PARA LIQUIDAR LA OPERACION.", vbOKOnly, PRODUCT_NAME)
            Exit Sub
        End If
        
        txtCHANGE.Text = Format(Abs(CDbl(DISPLAY.Text) - CDbl(txtMONEY.Text)), "Currency")


    PRINTER_TICKET = True
    Beep
    Call Label5_DblClick

End Sub

Private Sub cmdCASHTick_Click()

    
        If (txtMONEY.Text = "") Then
            txtMONEY.Text = "0"
        End If
        
        If (CDbl(txtMONEY.Text) < CDbl(DISPLAY.Text)) Then
            M = MsgBox("EL EFECTIVO ES INSUFICIENTE PARA LIQUIDAR LA OPERACION.", vbOKOnly, PRODUCT_NAME)
            Exit Sub
        End If
        
        txtCHANGE.Text = Format(Abs(CDbl(DISPLAY.Text) - CDbl(txtMONEY.Text)), "Currency")


    PRINTER_TICKET = True
    Call Label5_DblClick
    
    
    
End Sub

Private Sub cmdCLAV_Click()
    txtPRODUCT.SetFocus
End Sub

Private Sub cmdEFECT_Click()
    txtMONEY.Text = ""
    txtMONEY.SetFocus
End Sub

Private Sub cmdExit_Click()
    frmNetCash.Hide
    Unload frmNetCash
End Sub




Private Sub Command2_Click()

End Sub

Private Sub cmdOpenCash_Click()
    txt_Quantity.SetFocus
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select

End Sub


Private Sub Command5_Click()
    Load frmSelectArt
    frmSelectArt.Show 1
End Sub

Private Sub DISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select


End Sub


Private Sub Form_Load()

Dim Line As String

    Open "C:\CHM\MONEY.INI" For Input As #1
    
        Do While Not EOF(1)
        
            Input #1, Line
            Combo1.AddItem (Line)
        
        Loop
    
    Close #1

Rem ------- GET CAJA --------------------------------------------

    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("FOLIOS")

    REC.MoveFirst
    Label1.Caption = "CAJA:" + CStr(REC.Fields(1).Value)
    MSeries = CStr(REC.Fields(2).Value)

REC.Close
MDB.Close



    lblEmployee.Caption = "CAJERO: " + UserName
    
    Rem ------------ GRID1 --------------------------------------
    
    Grid1.Cols = 5
    Grid1.Rows = 100
    
    Grid1.Col = 0
    Grid1.RowHeight(0) = 10
    
    Grid1.Row = 0
    Grid1.Col = 0
    Grid1.Text = "CLAV"
    Grid1.ColWidth(0) = 10
        
    Grid1.Row = 0
    Grid1.Col = 1
    Grid1.Text = "ARTICULO"
    Grid1.ColWidth(1) = 4000
    
    Grid1.Row = 0
    Grid1.Col = 2
    Grid1.Text = "PRE"
    Grid1.ColWidth(2) = 1500

    Grid1.Row = 0
    Grid1.Col = 3
    Grid1.Text = "IVA"
    Grid1.ColWidth(3) = 1500

    Grid1.Row = 0
    Grid1.Col = 4
    Grid1.Text = "DESC"
    Grid1.ColWidth(4) = 4000


    ICursor = 1

    Rem ------------ GRID2 HIDDEN --------------------------------------
    
    Grid2.Cols = 7
    Grid2.Rows = 100
    
    Grid2.Row = 0
    Grid2.Col = 0
    Grid2.Text = "CLAVE"
    Grid2.ColWidth(1) = 2400
        
    Grid2.Row = 0
    Grid2.Col = 1
    Grid2.Text = "CODIGO DE BARRAS"
    Grid2.ColWidth(1) = 2400
    
    Grid2.Row = 0
    Grid2.Col = 2
    Grid2.Text = "CANTIDAD"
    Grid2.ColWidth(2) = 900

    Grid2.Row = 0
    Grid2.Col = 3
    Grid2.Text = "CONCEPTO"
    Grid2.ColWidth(2) = 900

    Grid2.Row = 0
    Grid2.Col = 4
    Grid2.Text = "FECHA"
    Grid2.ColWidth(2) = 900

    Grid2.Row = 0
    Grid2.Col = 5
    Grid2.Text = "COSTO"
    Grid2.ColWidth(2) = 900

    Grid2.Row = 0
    Grid2.Col = 6
    Grid2.Text = "ARTICULO"
    Grid2.ColWidth(2) = 900


    Open "C:\CHM\CONFIG.INI" For Input As #1
            Input #1, Line
            Input #1, Line
    Close #1

    If (UCase(Trim(Line)) = "ENABLED") Then
        mCopys = True
    Else
        mCopys = False
    End If


Dim strCheck As String
Dim strValue As String
    
    
    Open "C:\CHM\ONSALE.INI" For Input As #1
            Input #1, strCheck
            Input #1, strCheck
            Input #1, strValue
            Input #1, strValue
    Close #1
    
    If (Trim(UCase(strCheck)) = "ENABLED") Then
        mDescATodo = True
    Else
        mDescATodo = False
    End If
    
    mDescATodoValor = strValue


    Open "C:\CHM\TICKET.INI" For Input As #1
            Input #1, Leyenda1
            Input #1, Leyenda2
            Input #1, Leyenda3
    Close #1

End Sub




Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select

End Sub

Private Sub Label5_DblClick()

On Error GoTo LOCAL_CASH_HANDLER


If ICursor = 1 Then
    M = MsgBox("NO PUEDE REALIZAR UN COBRO DE CAJA SIN ARTICULOS.", vbOKOnly, PRODUCT_NAME)
    Exit Sub
End If



Rem ------- VALIDAR CIERRE --------------------------------------------
Dim mMDB As Database
Dim mREC As Recordset
Dim mFOUND As Boolean


    Set mMDB = OpenDatabase(NetCashFile)
    Set mREC = mMDB.OpenRecordset("CIERRES")
        
    mFOUND = False
    mREC.MoveFirst
        
    Do While (mREC.EOF <> True)
    
        If mREC.Fields(0).Value = Date Then
            mFOUND = True
            Exit Do
        End If
        
    mREC.MoveNext
    Loop
        
    If mFOUND Then
        x = MsgBox("El día contable " + CStr(Date) + " ya fue cerrado, la Caja no puede realizar transacciones con esta fecha por disposiciones fiscales.", 64, PRODUCT_NAME)
        
        Rem ------- REFRESH NETCASH ------------------------------------------------

        ICursor = 1

        sARTICULO = ""
        sCLAVE = ""
        mPRECIO = 0
        mIVAS = 0
        mDESCS = 0
        sCOBAR = ""
        mCOSTO = ""

        Grid1.Rows = 1
        Grid1.Rows = 100
        Grid2.Rows = 1
        Grid2.Rows = 100

        DISPLAY.Text = "0.00"
        txtPRODUCT.Text = ""
        txtSUBTOT.Text = "000.00"
        txtDESC.Text = "000.00"
        txtIVA.Text = "000.00"
        txtMONEY.Text = "000.00"
        txtCHANGE.Text = "000.00"
        txtPRODUCT.SetFocus
        
        Exit Sub
    End If

mREC.Close
mMDB.Close


Rem ------- Update EXISTENCIAS --------------------------------------------

    Dim mSales As Integer
    Dim mPointer As Integer
        
    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("ALMACEN")
        
    mSales = ICursor - 1
    mPointer = 1
    
    
    Do While (mPointer <= mSales)
    
        Grid1.Row = mPointer
        Grid2.Row = mPointer
        Grid1.Col = 0
        
        REC.MoveFirst
        Do While (REC.EOF = False)
            
            If (REC.Fields(0).Value = Grid1.Text) Then
                    REC.Edit
                    Grid2.Col = 2
                    REC.Fields(23).Value = REC.Fields(23).Value - CInt(Grid2.Text)
                    REC.Update
                    Exit Do
            End If
        
        REC.MoveNext
        Loop

    mPointer = mPointer + 1
    Loop

    REC.Close
    MDB.Close


Rem ------- Update FOLIOS --------------------------------------------


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("FOLIOS")

    REC.MoveFirst
    mFolios = REC.Fields(0).Value + 1
    REC.Edit
    REC.Fields(0).Value = mFolios
    REC.Update

    REC.Close
    MDB.Close

Rem ------- Update SALIDAS --------------------------------------------


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("SALIDAS")
       
    mSales = ICursor - 1
    mPointer = 1
    
    
    Do While (mPointer <= mSales)
    
            Grid2.Row = mPointer
            Grid2.Col = 1
            
            If (Grid2.Text <> "CANCELADO") Then
                    
                    REC.AddNew
                    
                        Grid2.Col = 0
                        REC.Fields(0).Value = Grid2.Text               'CLAVE
                        
                        Grid2.Col = 1
                        REC.Fields(1).Value = Grid2.Text               'COBAR
                                                
                        If (Trim(txt_Quantity.Text) = "") Then
                            txt_Quantity.Text = "1"
                        End If
                        
                        Grid2.Col = 2
                        REC.Fields(2).Value = CInt(Grid2.Text)         'CANTIDAD
                                            
                        REC.Fields(3).Value = "PUBLICO EN GENERAL"     'CONCEPTO
                        
                        REC.Fields(4).Value = Date                     'FECHA
                                                
                        REC.Fields(5).Value = CStr(mFolios)            'FACTURA

                        REC.Fields(6).Value = CStr(mFolios)            'REMISION
                    
                        REC.Fields(7).Value = CStr(mFolios)            'REFERENCIA
                    
                        REC.Fields(8).Value = CStr(mFolios)            'ORDEN

                        REC.Fields(9).Value = CStr(mFolios)            'PEDIDO

                        Grid2.Col = 5
                        REC.Fields(10).Value = Grid2.Text              'COSTO
                    
                        Grid2.Col = 6
                        REC.Fields(11).Value = Grid2.Text              'ARTICULO
                    
                    
                    REC.Update
            End If
            
    mPointer = mPointer + 1
    Loop

    REC.Close
    MDB.Close
    

Rem ------- Update INGRESOS $$$ --------------------------------------------

    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("INGRESOS")
       
    mSales = ICursor - 1
    mPointer = 1
    
    Dim sMEMO As String
    
    
    Do While (mPointer <= mSales)
    
            Grid2.Row = mPointer
            Grid2.Col = 1
            
            If (Grid2.Text <> "CANCELADO") Then
                    
                        Grid2.Col = 0
                        sMEMO = sMEMO + " CLAVE:" + Grid2.Text
                        Grid2.Col = 6
                        sMEMO = sMEMO + " CANT." + Grid2.Text
                    
            End If
            
    mPointer = mPointer + 1
    Loop



    REC.AddNew
    
        REC.Fields(0).Value = Date
        REC.Fields(1).Value = Combo1.Text
        REC.Fields(2).Value = CDbl(DISPLAY.Text)
        REC.Fields(3).Value = CDbl(txtSUBTOT.Text)
        REC.Fields(4).Value = CDbl(txtDESC.Text)
        REC.Fields(5).Value = CDbl(txtIVA.Text)
        REC.Fields(6).Value = CDbl(txtMONEY.Text)
        REC.Fields(7).Value = CDbl(txtCHANGE.Text)
        REC.Fields(8).Value = lblEmployee.Caption
        REC.Fields(9).Value = CStr(Time)
        REC.Fields(10).Value = Label1.Caption
        REC.Fields(11).Value = sMEMO
        REC.Fields(12).Value = mFolios
        REC.Fields(13).Value = "PUBLICO EN GENERAL"
        
    
    REC.Update


REC.Close
MDB.Close



Rem ----- TICKET ----------------------------------------------------------------

    Dim sEmpresa As String
    Dim sRFC As String
    Dim sDireccion As String
    Dim sCiudad As String
    Dim sEstado As String
    Dim sTelefono As String
    
    Dim cArt As String
    Dim iRec As Integer
    
    iRec = 1

    Set objDB = OpenDatabase(NetCashFile)
    Set objRecord = objDB.OpenRecordset("EMPRESA", dbOpenTable, dbReadOnly)
   
    objRecord.MoveFirst
    sEmpresa = objRecord.Fields(0).Value
    sRFC = objRecord.Fields(1).Value
    sDireccion = objRecord.Fields(2).Value
    sCiudad = objRecord.Fields(3).Value
    sEstado = objRecord.Fields(4).Value
    sTelefono = objRecord.Fields(8).Value

    
    objRecord.Close
    objDB.Close

Open "C:\TICKET.TMP" For Output As #1

    Print #1, sEmpresa
    Print #1, "-------------------------------------------------------------"
    Print #1, sRFC + " Tel." + sTelefono
    Print #1, sDireccion
    Print #1, sCiudad + " - " + sEstado
    Print #1, "-------------------------------------------------------------"
    Print #1, lblEmployee.Caption
    Print #1, "TICKET  FECHA: " + CStr(Date)
    Print #1, "HORA: " + CStr(Time)
        
    If (Trim(MSeries) <> "OFF") Then
        Print #1, "FOLIO: " + MSeries + CStr(mFolios)
    Else
        Print #1, "FOLIO: " + CStr(mFolios)
    End If

    Print #1, Label1.Caption
    Print #1, "TIPO: " + Combo1.Text
    Print #1, "-------------------------------------------------------------"
    
    Dim strDIG As String
    Do While (iRec < ICursor)
        Grid1.Row = iRec
        Grid1.Col = 1
        If (Trim(Grid1.Text) <> "CANCELADO") Then
            Print #1, Grid1.Text
            Grid1.Col = 2
            strDIG = Format(CDbl(Grid1.Text), "Currency")
                If InStr(strDIG, ",") > 0 Then
                    Mid(strDIG, InStr(strDIG, ","), 1) = "."
                End If
            Print #1, strDIG
        End If
    iRec = iRec + 1
    Loop

    Print #1, "-------------------------------------------------------------"
                
    strDIG = txtSUBTOT.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "SUB-TOTAL:  " + strDIG
    
          
    strDIG = txtDESC.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "DESCUENTOS: " + strDIG
    
        
    strDIG = txtIVA.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "IVA:        " + strDIG
    
    
    strDIG = DISPLAY.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "TOTAL:      " + strDIG
    
    Print #1, "-------------------------------------------------------------"
    
    strDIG = Format(CDbl(txtMONEY.Text), "###,##0.00")
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "PAGO:       " + strDIG
    
    
    strDIG = txtCHANGE.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "SU CAMBIO: " + strDIG
    
    Print #1, "  "
    
    Print #1, Leyenda1
    Print #1, Leyenda2
    Print #1, Leyenda3


    Rem ----- TICKET COPIA ----------------------------------------------------------------
        
    iRec = 1
    
If (mCopys = True) Then
        
    Print #1, ""
    Print #1, " * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * "
    Print #1, ""

    Print #1, sEmpresa
    Print #1, "-------------------------------------------------------------"
    Print #1, sRFC + " Tel." + sTelefono
    Print #1, sDireccion
    Print #1, sCiudad + " - " + sEstado
    Print #1, "-------------------------------------------------------------"
    Print #1, lblEmployee.Caption
    
    Print #1, "TICKET  FECHA: " + CStr(Date)
    Print #1, "HORA: " + CStr(Time)

    If (Trim(MSeries) <> "OFF") Then
        Print #1, "FOLIO: " + MSeries + CStr(mFolios)
    Else
        Print #1, "FOLIO: " + CStr(mFolios)
    End If
    
    Print #1, Label1.Caption
    Print #1, "TIPO: " + Combo1.Text
    Print #1, "-------------------------------------------------------------"
    
    strDIG = ""
    Do While (iRec < ICursor)
        Grid1.Row = iRec
        Grid1.Col = 1
        If (Trim(Grid1.Text) <> "CANCELADO") Then
            Print #1, Grid1.Text
            Grid1.Col = 2
            strDIG = Format(CDbl(Grid1.Text), "Currency")
                If InStr(strDIG, ",") > 0 Then
                    Mid(strDIG, InStr(strDIG, ","), 1) = "."
                End If
            Print #1, strDIG
        End If
    iRec = iRec + 1
    Loop

    Print #1, "-------------------------------------------------------------"
                
    strDIG = txtSUBTOT.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "SUB-TOTAL:  " + strDIG
    
          
    strDIG = txtDESC.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "DESCUENTOS: " + strDIG
    
        
    strDIG = txtIVA.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "IVA:        " + strDIG
    
    
    strDIG = DISPLAY.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "TOTAL:      " + strDIG
    
    Print #1, "-------------------------------------------------------------"
    
    strDIG = Format(CDbl(txtMONEY.Text), "###,##0.00")
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "PAGO:       " + strDIG
    
    
    strDIG = txtCHANGE.Text
    If InStr(strDIG, ",") > 0 Then
       Mid(strDIG, InStr(strDIG, ","), 1) = "."
    End If
    Print #1, "SU CAMBIO: " + strDIG
    
    Print #1, "  "
    
    Print #1, Leyenda1
    Print #1, Leyenda2
    Print #1, Leyenda3


End If

Close #1








Rem ------- REFRESH NETCASH ------------------------------------------------

ICursor = 1

sARTICULO = ""
sCLAVE = ""
mPRECIO = 0
mIVAS = 0
mDESCS = 0
sCOBAR = ""
mCOSTO = ""

M = MsgBox("VENTA REALIZADA, CAMBIO A DEVOLVER AL CLIENTE: " + txtCHANGE.Text, vbOKOnly, PRODUCT_NAME)

Grid1.Rows = 1
Grid1.Rows = 100

Grid2.Rows = 1
Grid2.Rows = 100

DISPLAY.Text = "0.00"
txtPRODUCT.Text = ""
txtSUBTOT.Text = "000.00"
txtDESC.Text = "000.00"
txtIVA.Text = "000.00"
txtMONEY.Text = "000.00"
txtCHANGE.Text = "000.00"

txtPRODUCT.SetFocus

If PRINTER_TICKET Then
    Shell "C:\CHM\PRINT32.EXE /autofix", vbNormalFocus
End If


Exit Sub

LOCAL_CASH_HANDLER:

M = MsgBox("Error de privilegios para acceder a dispositivos Hardware, Favor de cerrar el programa y ejecutarlo con permisos de Administrador, Si el error persiste consulte a soporte técnico de Nova32 Development Software.", vbCritical, PRODUCT_NAME)

End Sub


Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select

End Sub


Private Sub txt_Quantity_Click()
    txt_Quantity.Text = ""
End Sub

Private Sub txt_Quantity_GotFocus()
    txt_Quantity.Text = ""
End Sub


Private Sub txt_Quantity_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select

End Sub




Private Sub txt_Quantity_KeyPress(KeyAscii As Integer)

On Error GoTo LOCAL_KEYS

If KeyAscii = 13 Then
    txtPRODUCT.SetFocus
    Exit Sub
End If

LOCAL_KEYS:

End Sub


Private Sub txtMONEY_Click()
    txtMONEY.Text = ""
End Sub


Private Sub txtMONEY_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If (txtMONEY.Text = "") Then
            txtMONEY.Text = "0"
        End If
        
        If (CDbl(txtMONEY.Text) < CDbl(DISPLAY.Text)) Then
            M = MsgBox("EL EFECTIVO ES INSUFICIENTE PARA LIQUIDAR LA OPERACION.", vbOKOnly, PRODUCT_NAME)
            Exit Sub
        End If
        
        txtCHANGE.Text = Format(Abs(CDbl(DISPLAY.Text) - CDbl(txtMONEY.Text)), "Currency")
        
        Exit Sub
    
    End If


Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case Default

End Select

End Sub




Private Sub txtPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
        
 
    Case 114 '[F3] CERRAR
        Call cmdExit_Click
    Case 115 '[F4] COBRA
        Grid1.Row = 1
        Grid1.Col = 1
        Call cmdCASHTick_Click
    Case 116 '[F5] EFECTIVO
        Call cmdEFECT_Click
    Case 117 '[F6] TIPO PAGO
        Call cmdOpenCash_Click
    Case 118 '[F7] CANCELA
        Call cmdCAN_Click
    Case 119 '[F8] CLAVE
        Call cmdCLAV_Click
    Case 120 '[F9] CONSULTA
        Call Command5_Click
    Case 121 '[F10] BASCULA
        Call Bascula
    Case Default

End Select


End Sub

Private Sub txtPRODUCT_KeyPress(KeyAscii As Integer)


On Error GoTo LOCAL_HANDLER_CASH_NULL

If KeyAscii = 13 Then

    mDESCS = 0
    mIVAS = 0

    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("ALMACEN")

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = txtPRODUCT.Text
    
    If (Trim(txt_Quantity.Text) = "") Then
        txt_Quantity.Text = "1"
    End If
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    Beep
    Dim Found As Boolean
    Found = False

    REC.MoveFirst
    Do While (REC.EOF = False)
        If (REC.Fields(0).Value = sTMP) Then
            sCLAVE = REC.Fields(0).Value
            sARTICULO = txt_Quantity.Text + ": " + REC.Fields(2).Value
            mPRECIO = REC.Fields(20).Value * CInt(txt_Quantity.Text)
            mCOSTO = REC.Fields(7).Value
            sCOBAR = REC.Fields(1).Value
            mGRUPO = REC.Fields("FAMILIA").Value
            
            Rem ------- COMPUTE DESC -----------------------------------
            
            
            'Descuento General
            If (mDescATodo = True) Then
                mDESCS = (REC.Fields(20).Value * CInt(txt_Quantity.Text)) * (mDescATodoValor / 100)
            Else
                'Buscar Descuento por Grupos
                Dim tx_FOUND As Boolean
                Dim tx_MDB As Database
                Dim tx_REC As Recordset
                Dim tx_DESC As Integer
                
                tx_FOUND = False
                
                Set tx_MDB = OpenDatabase(NetCashFile)
                Set tx_REC = MDB.OpenRecordset("ADMDESC")
                
                tx_REC.MoveFirst
                
                Do While (tx_REC.EOF = False)
                    If (mGRUPO = tx_REC.Fields("GRUPO").Value) Then
                        tx_FOUND = True
                        tx_DESC = tx_REC.Fields("DESCUENTO").Value
                        Exit Do
                    End If
                    tx_REC.MoveNext
                Loop
                
                tx_REC.Close
                tx_MDB.Close
                                
                If (tx_FOUND = True) Then
                    Rem Aplicar Decuento de Grupo
                    mDESCS = (REC.Fields(20).Value * CInt(txt_Quantity.Text)) * (tx_DESC / 100)
                Else
                    Rem Aplicar Descuento de producto individual
                    mDESCS = (REC.Fields(20).Value * CInt(txt_Quantity.Text)) * (REC.Fields(25).Value / 100)
                End If
                
            End If
            
            
            Rem ------- COMPUTE IVA ---------------------------
            mIVAS = ((REC.Fields(20).Value * CInt(txt_Quantity.Text)) - mDESCS) * (REC.Fields(24).Value / 100)
            Found = True
            Exit Do
        End If
    REC.MoveNext
    Loop


    If (Found) Then
        
    Dim EvalStore As Integer
    EvalStore = CInt(REC.Fields(23).Value)
    If EvalStore <= 0 Then
        M = MsgBox("EN INVENTARIOS NO HAY EXISTENCIAS DISPONIBLES DE ESTE ARTICULO, NO SE PUEDE AGREGAR A LA VENTA, DEBE ACTUALIZAR INVENTARIO.", vbCritical, "PRECAUCION:")
        txtPRODUCT.Text = ""
        Exit Sub
    End If
                
        Rem ---------------- UPDATE GRID1 --------------------------
        Grid1.Col = 0
        Grid1.Row = ICursor
        Grid1.Text = sCLAVE
        
        Grid1.Col = 1
        Grid1.Row = ICursor
        
        Grid1.Text = sARTICULO
                
        Grid1.Col = 2
        Grid1.Row = ICursor
        Grid1.Text = CStr(mPRECIO)
        
        Grid1.Col = 3
        Grid1.Row = ICursor
        Grid1.Text = CStr(mIVAS)
        
        Grid1.Col = 4
        Grid1.Row = ICursor
        Grid1.Text = CStr(mDESCS)
        
        
        Rem ---------------- UPDATE GRID2 --------------------------
        Grid2.Col = 0
        Grid2.Row = ICursor
        Grid2.Text = sCLAVE
        
        Grid2.Col = 1
        Grid2.Row = ICursor
        Grid2.Text = sCOBAR
        
        Grid2.Col = 2
        Grid2.Row = ICursor
        Grid2.Text = txt_Quantity.Text
        
        Grid2.Col = 3
        Grid2.Row = ICursor
        Grid2.Text = "VENTA"
        
        Grid2.Col = 4
        Grid2.Row = ICursor
        Grid2.Text = Date
        
        Grid2.Col = 5
        Grid2.Row = ICursor
        Grid2.Text = CStr(mCOSTO)
        
        Grid2.Col = 6
        Grid2.Row = ICursor
        Grid2.Text = sARTICULO
        
        Rem ----------------------------------------------
        
        ICursor = ICursor + 1
        
        'Dim mCalIVa As Double
        'Dim mCalDesc As Double
        'mCalDesc = (REC.Fields(20).Value * CInt(txt_Quantity.Text)) * (REC.Fields(25).Value / 100)
        'mCalIVa = ((REC.Fields(20).Value * CInt(txt_Quantity.Text)) - mCalDesc) * (REC.Fields(24).Value / 100)
        
        txtSUBTOT.Text = Format(CDbl(txtSUBTOT.Text) + mPRECIO, "Currency")
        txtIVA.Text = Format(CDbl(txtIVA.Text) + mIVAS, "Currency")
        txtDESC.Text = Format(CDbl(txtDESC.Text) + mDESCS, "Currency")
        DISPLAY.Text = Format(CDbl(txtSUBTOT.Text) - CDbl(txtDESC.Text) + CDbl(txtIVA.Text), "Currency")
        
        txtPRODUCT.Text = ""
        
    Else
        M = MsgBox("NO SE ENCONTRO EL ARTICULO EN INVENTARIOS.", vbCritical, PRODUCT_NAME)
        txtPRODUCT.Text = ""
    End If


REC.Close
MDB.Close

txt_Quantity.Text = "1"

End If


Exit Sub

LOCAL_HANDLER_CASH_NULL:
    x = MsgBox("ERROR DE CAPTURA DE DATOS, CORREGIR Y VOLVER A INTENTAR", 64, PRODUCT_NAME)

End Sub


