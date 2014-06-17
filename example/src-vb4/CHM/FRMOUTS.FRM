VERSION 4.00
Begin VB.Form frmOutputs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas de Almacén"
   ClientHeight    =   5145
   ClientLeft      =   1020
   ClientTop       =   1950
   ClientWidth     =   6360
   Height          =   5655
   Icon            =   "FRMOUTS.frx":0000
   Left            =   960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Top             =   1500
   Width           =   6480
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   4680
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _version        =   65536
      _extentx        =   10821
      _extenty        =   7858
      _stockprops     =   15
      caption         =   "Salida de Producto"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Salida de Producto"
      tab(0).controlcount=   12
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label2"
      tab(0).control(1)=   "Label1"
      tab(0).control(2)=   "Label3"
      tab(0).control(3)=   "Label4"
      tab(0).control(4)=   "cboConcepto"
      tab(0).control(5)=   "Label5"
      tab(0).control(6)=   "txtCodigoBarras"
      tab(0).control(7)=   "txtClave"
      tab(0).control(8)=   "txtCantidad"
      tab(0).control(9)=   "txtFecha"
      tab(0).control(10)=   "cmdSave"
      tab(0).control(11)=   "Data1"
      tabcaption(1)   =   "Detalle de Salida"
      tab(1).controlcount=   13
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "Picture1"
      tab(1).control(1)=   "txtCosto"
      tab(1).control(2)=   "txtPedido"
      tab(1).control(3)=   "txtOrden"
      tab(1).control(4)=   "txtReferencia"
      tab(1).control(5)=   "txtRemision"
      tab(1).control(6)=   "txtFactura"
      tab(1).control(7)=   "Label15"
      tab(1).control(8)=   "Label12"
      tab(1).control(9)=   "Label9"
      tab(1).control(10)=   "Label8"
      tab(1).control(11)=   "Label7"
      tab(1).control(12)=   "Label6"
      tabcaption(2)   =   "Consulta de Salidas"
      tab(2).controlcount=   5
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "CENTRADAS"
      tab(2).control(1)=   "cmdSeek"
      tab(2).control(2)=   "txtClave2"
      tab(2).control(3)=   "DBGrid1"
      tab(2).control(4)=   "Label10"
      Begin VB.PictureBox Picture1 
         Height          =   1575
         Left            =   -70920
         Picture         =   "FRMOUTS.frx":0442
         ScaleHeight     =   1515
         ScaleWidth      =   1635
         TabIndex        =   29
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Data CENTRADAS 
         Caption         =   "Navegar"
         Connect         =   "Access"
         DatabaseName    =   "C:\NOVA32.MDB"
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -71280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "SALIDAS"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdSeek 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   -72360
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtClave2 
         Height          =   285
         Left            =   -74280
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.Data Data1 
         Caption         =   "COPER"
         Connect         =   "Access"
         DatabaseName    =   "C:\NOVA32.MDB"
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "COPERAC"
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtCosto 
         DataField       =   "COSTO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   11
         Text            =   "0"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtPedido 
         DataField       =   "PEDIDO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   10
         Text            =   "0"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtOrden 
         DataField       =   "ORDEN"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   9
         Text            =   "0"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtReferencia 
         DataField       =   "REFERENCIA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   8
         Text            =   "0"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtRemision 
         DataField       =   "REMISION"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   7
         Text            =   "0"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtFactura 
         DataField       =   "FACTURA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   -73320
         TabIndex        =   6
         Text            =   "0"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H0000FFFF&
         DataField       =   "CLAVE"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtCodigoBarras 
         BackColor       =   &H0000FFFF&
         DataField       =   "CODIGO_BARRAS"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FRMOUTS.frx":9044
         Height          =   3375
         Left            =   -74880
         OleObjectBlob   =   "FRMOUTS.frx":9056
         TabIndex        =   28
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label10 
         Caption         =   "Clave"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Costo"
         Height          =   255
         Left            =   -74400
         TabIndex        =   22
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Pedido"
         Height          =   255
         Left            =   -74400
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Orden"
         Height          =   255
         Left            =   -74400
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   -74400
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remisión"
         Height          =   255
         Left            =   -74400
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Factura"
         Height          =   255
         Left            =   -74400
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin MSDBCtls.DBCombo cboConcepto 
         Bindings        =   "FRMOUTS.frx":975E
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
         _version        =   65536
         _extentx        =   3413
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   16777215
         listfield       =   "DESCRIP"
         boundcolumn     =   "DESCRIP"
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Barras"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOutputs"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public sARTICULO As String

Private Sub cmdExit_Click()
    frmOutputs.Hide
    Unload frmOutputs
End Sub

Private Sub cmdSave_Click()


On Error GoTo Local_Handler_Saveouts

Dim ALMACEN As Database
Dim ENTRADAS As Database

Dim RecALMACEN As Recordset
Dim RecENTRADAS As Recordset

Set ALMACEN = OpenDatabase(NetCashFile)
Set RecALMACEN = ALMACEN.OpenRecordset("ALMACEN")

Set ENTRADAS = OpenDatabase(NetCashFile)
Set RecENTRADAS = ENTRADAS.OpenRecordset("SALIDAS")


    Dim Found As Boolean
    Found = False
    
    RecALMACEN.MoveFirst
    Do While (RecALMACEN.EOF = False)
        If (RecALMACEN.Fields(0).Value = txtClave.Text) Then
                sARTICULO = RecALMACEN.Fields(2).Value
                Found = True
                Exit Do
        End If
    RecALMACEN.MoveNext
    Loop
                
    If Found Then
        RecALMACEN.Edit
        If txtCantidad.Text = "" Then
            txtCantidad.Text = "0"
        End If
        If ((RecALMACEN.Fields(23).Value = 0) Or (RecALMACEN.Fields(23).Value < CInt(txtCantidad.Text))) Then
                msg = MsgBox("No hay existencias disponibles para operar una salida de este artículo, Existencias Actuales: " + CStr(RecALMACEN.Fields(23).Value), vbCritical, PRODUCT_NAME)
                RecALMACEN.Close
                ALMACEN.Close
                Exit Sub
        End If
        RecALMACEN.Fields(23).Value = RecALMACEN.Fields(23).Value - CInt(txtCantidad.Text)
        RecALMACEN.Update
        RecALMACEN.Close
        ALMACEN.Close
    Else
        msg = MsgBox("El artículo no esta dado de alta.", vbCritical, PRODUCT_NAME)
        RecALMACEN.Close
        ALMACEN.Close
        Exit Sub
    End If

    Rem --------- Validates ------------------------------------
    If txtCodigoBarras.Text = "" Then
        txtCodigoBarras.Text = txtClave.Text
    End If
    If txtCantidad.Text = "0" Then
        mg = MsgBox("No se puede guardar esta salida de almacén, debe existir al menos 1 artículo a salir en el campo CANTIDAD.", vbOKOnly, PRODUCT_NAME)
        RecENTRADAS.Close
        ENTRADAS.Close
        Exit Sub
    End If
    If cboConcepto.Text = "" Then
        cboConcepto.Text = " "
    End If
    If txtFecha.Text = "" Then
        txtFecha.Text = Date
    End If
    If txtFactura.Text = "" Then
        txtFactura.Text = " "
    End If
    If txtRemision.Text = "" Then
        txtRemision.Text = " "
    End If
    If txtReferencia.Text = "" Then
        txtReferencia.Text = " "
    End If
    If txtOrden.Text = "" Then
        txtOrden.Text = " "
    End If
    If txtPedido.Text = "" Then
        txtPedido.Text = " "
    End If
    If txtCosto.Text = "" Then
        txtCosto.Text = "0.0"
    End If


    RecENTRADAS.AddNew
        RecENTRADAS.Fields(0).Value = txtClave.Text
        RecENTRADAS.Fields(1).Value = txtCodigoBarras.Text
        RecENTRADAS.Fields(2).Value = CInt(txtCantidad.Text)
        RecENTRADAS.Fields(3).Value = cboConcepto.Text
        RecENTRADAS.Fields(4).Value = txtFecha.Text
        RecENTRADAS.Fields(5).Value = txtFactura.Text
        RecENTRADAS.Fields(6).Value = txtRemision.Text
        RecENTRADAS.Fields(7).Value = txtReferencia.Text
        RecENTRADAS.Fields(8).Value = txtOrden.Text
        RecENTRADAS.Fields(9).Value = txtPedido.Text
        RecENTRADAS.Fields(10).Value = CDbl(txtCosto.Text)
        RecENTRADAS.Fields(11).Value = sARTICULO
    RecENTRADAS.Update
    RecENTRADAS.Close
    ENTRADAS.Close


mg = MsgBox("La salida del artículo fue registrada.", vbOKOnly, PRODUCT_NAME)
txtClave.Text = ""
txtCodigoBarras.Text = ""
txtCantidad.Text = ""
cboConcepto.Text = ""
txtFactura.Text = ""
txtRemision.Text = ""
txtOrden.Text = ""
txtPedido.Text = ""
txtCosto.Text = ""

DBGrid1.Refresh


Exit Sub

Local_Handler_Saveouts:
    x = MsgBox("No se puede completar la operación verifique los campos de entrada.", 64, PRODUCT_NAME)



End Sub

Private Sub cmdSeek_Click()
    CENTRADAS.Recordset.FindFirst ("CLAVE='" + txtClave2.Text + "'")
End Sub


Private Sub Form_Load()
     Data1.DatabaseName = NetCashFile
    CENTRADAS.DatabaseName = NetCashFile
    txtFecha.Text = Date
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    CENTRADAS.Refresh
End Sub

Private Sub txtClave_LostFocus()

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = txtClave.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    txtClave.Text = sTMP


End Sub


Private Sub txtClave2_LostFocus()

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = txtClave2.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    txtClave2.Text = sTMP


End Sub


Private Sub txtCodigoBarras_LostFocus()

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = txtCodigoBarras.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    txtCodigoBarras.Text = sTMP


End Sub


