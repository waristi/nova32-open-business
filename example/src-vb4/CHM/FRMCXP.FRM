VERSION 4.00
Begin VB.Form frmCXP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas por Pagar"
   ClientHeight    =   6165
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5820
   Height          =   6570
   Icon            =   "FRMCXP.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   5940
   Begin VB.CommandButton Command7 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4200
      TabIndex        =   52
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Data PROVEEDORES 
      Caption         =   "PROVEEDORES"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\NETCASH.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "PROVEEDORES"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _version        =   65536
      _extentx        =   9763
      _extenty        =   9551
      _stockprops     =   15
      caption         =   "Cargos"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   4
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Cargos"
      tab(0).controlcount=   5
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label6"
      tab(0).control(1)=   "cbo_Cliente_Cargos"
      tab(0).control(2)=   "Label1"
      tab(0).control(3)=   "txt_Fecha_Cargos"
      tab(0).control(4)=   "Frame2"
      tabcaption(1)   =   "Abonos"
      tab(1).controlcount=   5
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "Frame3"
      tab(1).control(1)=   "txt_Fecha_Abonos"
      tab(1).control(2)=   "Label7"
      tab(1).control(3)=   "cbo_Clientes_Abonos"
      tab(1).control(4)=   "Label10"
      tabcaption(2)   =   "Saldos"
      tab(2).controlcount=   6
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "Label12"
      tab(2).control(1)=   "cbo_Clientes_Saldos"
      tab(2).control(2)=   "Label11"
      tab(2).control(3)=   "Command1"
      tab(2).control(4)=   "Frame4"
      tab(2).control(5)=   "txt_Fecha_Saldos"
      tabcaption(3)   =   "Reportes"
      tab(3).controlcount=   6
      tab(3).controlenabled=   0   'False
      tab(3).control(0)=   "CommonDialog1"
      tab(3).control(1)=   "cbo_Clientes_Reportes"
      tab(3).control(2)=   "Command6"
      tab(3).control(3)=   "Command5"
      tab(3).control(4)=   "Command4"
      tab(3).control(5)=   "Picture1"
      Begin VB.Frame Frame2 
         Caption         =   "Detalle de Transacción"
         Height          =   3255
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   4935
         Begin VB.TextBox txt_Concepto_Cargos 
            Height          =   2295
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   39
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txt_Importe_Cargos 
            Height          =   285
            Left            =   1080
            TabIndex        =   38
            Text            =   "0"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Aplicar Cargo"
            Height          =   375
            Left            =   3480
            TabIndex        =   37
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Importe"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2760
            Width           =   615
         End
      End
      Begin VB.TextBox txt_Fecha_Cargos 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   35
         Text            =   "00/00/0000"
         Top             =   900
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Abono en Cuenta"
         Height          =   3195
         Left            =   -74760
         TabIndex        =   29
         Top             =   1980
         Width           =   4935
         Begin VB.CommandButton Command3 
            Caption         =   "Aplicar Abono"
            Height          =   375
            Left            =   3480
            TabIndex        =   32
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox txt_Importe_Abonos 
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Text            =   "0"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox txt_Concepto_Abonos 
            Height          =   2175
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   30
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label8 
            Caption         =   "Importe"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Concepto"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txt_Fecha_Abonos 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   -73440
         TabIndex        =   28
         Text            =   "00/00/0000"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txt_Fecha_Saldos 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73440
         TabIndex        =   27
         Text            =   "00/00/0000"
         Top             =   900
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cliente"
         Height          =   3615
         Left            =   -74880
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
         Begin VB.TextBox HABER 
            BackColor       =   &H0080FF80&
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Text            =   "0"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox SALDO 
            BackColor       =   &H0080FF80&
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Text            =   "0"
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox tx_Nombre 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox tx_RFC 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox tx_Direccion 
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox tx_Ciudad 
            Height          =   285
            Left            =   1080
            TabIndex        =   12
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox tx_Tel 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox tx_Cel 
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox tx_Mail 
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox DEBE 
            BackColor       =   &H0080FF80&
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Text            =   "0"
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Dirección"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "RFC"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Ciudad"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Tel."
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Movil"
            Height          =   255
            Left            =   2400
            TabIndex        =   21
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "e-Mail"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "DEBE"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "HABER"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "SALDO"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   3240
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3255
         Left            =   -74760
         Picture         =   "FRMCXP.frx":0442
         ScaleHeight     =   3195
         ScaleWidth      =   3075
         TabIndex        =   5
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Historial Crediticio"
         Height          =   375
         Left            =   -72840
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   -71040
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reporteador"
         Height          =   375
         Left            =   -71520
         TabIndex        =   2
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Configurar Hoja"
         Height          =   375
         Left            =   -71520
         TabIndex        =   1
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Clave del Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1260
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo cbo_Cliente_Cargos 
         Bindings        =   "FRMCXP.frx":29B84
         Height          =   315
         Left            =   1560
         TabIndex        =   50
         Top             =   1260
         Width           =   3615
         _version        =   65536
         _extentx        =   6376
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   8454143
         listfield       =   "ID_PROVEEDOR"
         boundcolumn     =   "ID_PROVEEDOR"
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Clave del Cliente"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   1380
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo cbo_Clientes_Abonos 
         Bindings        =   "FRMCXP.frx":29B98
         Height          =   315
         Left            =   -73440
         TabIndex        =   47
         Top             =   1380
         Width           =   3615
         _version        =   65536
         _extentx        =   6376
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   8454143
         listfield       =   "ID_PROVEEDOR"
         boundcolumn     =   "ID_PROVEEDOR"
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   -74760
         TabIndex        =   46
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Clave del Cliente"
         Height          =   255
         Left            =   -74760
         TabIndex        =   45
         Top             =   1260
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo cbo_Clientes_Saldos 
         Bindings        =   "FRMCXP.frx":29BAC
         Height          =   315
         Left            =   -73440
         TabIndex        =   44
         Top             =   1260
         Width           =   3615
         _version        =   65536
         _extentx        =   6376
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   8454143
         listfield       =   "ID_PROVEEDOR"
         boundcolumn     =   "ID_PROVEEDOR"
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   900
         Width           =   615
      End
      Begin MSDBCtls.DBCombo cbo_Clientes_Reportes 
         Bindings        =   "FRMCXP.frx":29BC0
         Height          =   315
         Left            =   -74760
         TabIndex        =   42
         Top             =   840
         Width           =   3615
         _version        =   65536
         _extentx        =   6376
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   8454143
         listfield       =   "ID_PROVEEDOR"
         boundcolumn     =   "ID_PROVEEDOR"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70680
         Top             =   1080
         _version        =   65536
         _extentx        =   847
         _extenty        =   847
         _stockprops     =   0
      End
   End
End
Attribute VB_Name = "frmCXP"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo Local_Handler_SaveCXC

Dim m_DEBE As Double
Dim m_HABER As Double
Dim m_SALDO As Double

Dim DB_CXC As Database
Dim RE_CXC As Recordset
Dim mFOUND As Boolean

'---------- SEEK PROVEEDOR ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("PROVEEDORES")

    mFOUND = False
    RE_CXC.MoveFirst
    Do While (RE_CXC.EOF = False)
        If (Trim(RE_CXC.Fields(0).Value) = Trim(cbo_Clientes_Saldos.Text)) Then
            mFOUND = True
            
            If (RE_CXC.Fields("NOMBRE").Value <> vbNullString) Then
                tx_Nombre.Text = RE_CXC.Fields("NOMBRE").Value
            Else
                tx_Nombre.Text = ""
            End If
            
            
            If (RE_CXC.Fields("RFC").Value <> vbNullString) Then
                tx_RFC.Text = RE_CXC.Fields("RFC").Value
            Else
                tx_RFC.Text = ""
            End If
            
            
            If (RE_CXC.Fields("DIRECCION").Value <> vbNullString) Then
                tx_Direccion.Text = RE_CXC.Fields("DIRECCION").Value
            Else
                tx_Direccion.Text = ""
            End If
            
            If (RE_CXC.Fields("CIUDAD").Value <> vbNullString) Then
                tx_Ciudad.Text = RE_CXC.Fields("CIUDAD").Value
            Else
                tx_Ciudad.Text = ""
            End If
            
            
            If (RE_CXC.Fields("TELEFONO").Value <> vbNullString) Then
                tx_Tel.Text = RE_CXC.Fields("TELEFONO").Value
            Else
                tx_Tel.Text = ""
            End If
            
            
            If (RE_CXC.Fields("MOVIL").Value <> vbNullString) Then
                tx_Cel.Text = RE_CXC.Fields("MOVIL").Value
            Else
                tx_Cel.Text = ""
            End If
            
            
            If (RE_CXC.Fields("EMAIL").Value <> vbNullString) Then
                tx_Mail.Text = RE_CXC.Fields("EMAIL").Value
            Else
                tx_Mail.Text = ""
            End If
                                              
            If (RE_CXC.Fields("DEBE").Value <> vbNull) Then
                m_DEBE = RE_CXC.Fields("DEBE").Value
            Else
                m_DEBE = 0
            End If
                                              
            If (RE_CXC.Fields("HABER").Value <> vbNull) Then
                m_HABER = RE_CXC.Fields("HABER").Value
            Else
                m_HABER = 0
            End If
                                                          
            m_SALDO = m_DEBE - m_HABER
            DEBE.Text = CStr(m_DEBE)
            HABER.Text = CStr(m_HABER)
            SALDO.Text = CStr(m_SALDO)
            Exit Do
        End If
    RE_CXC.MoveNext
    Loop


RE_CXC.Close
DB_CXC.Close

If (mFOUND = False) Then
    x = MsgBox("No se encontro el Proveedor en el Sistema, operación cancelada.", 64, PRODUCT_NAME)
    Exit Sub
End If


Exit Sub

Local_Handler_SaveCXC:
    x = MsgBox("No se puede completar la operación verifique los campos de entrada, o los valores de la cuenta del cliente que intenta abrir.", 64, PRODUCT_NAME)

End Sub

Private Sub Command2_Click()

On Error GoTo Local_Handler_SaveCXC


Dim mCLIENTE As String
Dim DB_CXC As Database
Dim RE_CXC As Recordset
Dim mFOUND As Boolean

'---------- SEEK PROVEEDOR ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("PROVEEDORES")

    mFOUND = False
    RE_CXC.MoveFirst
    Do While (RE_CXC.EOF = False)
        If (Trim(RE_CXC.Fields(0).Value) = Trim(cbo_Cliente_Cargos.Text)) Then
            mFOUND = True
            mCLIENTE = RE_CXC.Fields(1).Value
            Exit Do
        End If
    RE_CXC.MoveNext
    Loop


RE_CXC.Close
DB_CXC.Close

If (mFOUND = False) Then
    x = MsgBox("No se encontro el Proveedor en el Sistema, operación cancelada.", 64, PRODUCT_NAME)
    Exit Sub
End If

'---------- UPDATE CXP ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("CXP")

    RE_CXC.AddNew
        RE_CXC.Fields("FECHA").Value = txt_Fecha_Cargos.Text
        RE_CXC.Fields("HORA").Value = CStr(Time)
        RE_CXC.Fields("ID_PROVEEDOR").Value = cbo_Cliente_Cargos.Text
        RE_CXC.Fields("NOMBRE").Value = mCLIENTE
        RE_CXC.Fields("CONCEPTO").Value = txt_Concepto_Cargos.Text
        RE_CXC.Fields("CARGO").Value = CDbl(txt_Importe_Cargos.Text)
        RE_CXC.Fields("ABONO").Value = 0
    RE_CXC.Update
    RE_CXC.Close
DB_CXC.Close



'---------- UPDATE CLIENTE ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("PROVEEDORES")

    mFOUND = False
    RE_CXC.MoveFirst
    Do While (RE_CXC.EOF = False)
        If (Trim(RE_CXC.Fields(0).Value) = Trim(cbo_Cliente_Cargos.Text)) Then
            RE_CXC.Edit
            RE_CXC.Fields("DEBE").Value = RE_CXC.Fields("DEBE").Value + CDbl(txt_Importe_Cargos.Text)
            RE_CXC.Update
            Exit Do
        End If
    RE_CXC.MoveNext
    Loop


RE_CXC.Close
DB_CXC.Close



x = MsgBox("El Cargo fue aplicado en cuenta de " + cbo_Cliente_Cargos.Text + " a nombre de " + mCLIENTE, 64, PRODUCT_NAME)


cbo_Cliente_Cargos.Text = ""
txt_Concepto_Cargos.Text = ""


Exit Sub

Local_Handler_SaveCXC:
    x = MsgBox("No se puede completar la operación verifique los campos de entrada.", 64, PRODUCT_NAME)

End Sub

Private Sub Command3_Click()

On Error GoTo Local_Handler_SaveCXC


Dim mCLIENTE As String
Dim DB_CXC As Database
Dim RE_CXC As Recordset
Dim mFOUND As Boolean

'---------- SEEK PROVEEDOR ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("PROVEEDORES")

    mFOUND = False
    RE_CXC.MoveFirst
    Do While (RE_CXC.EOF = False)
        If (Trim(RE_CXC.Fields(0).Value) = Trim(cbo_Clientes_Abonos.Text)) Then
            mFOUND = True
            mCLIENTE = RE_CXC.Fields(1).Value
            Exit Do
        End If
    RE_CXC.MoveNext
    Loop


RE_CXC.Close
DB_CXC.Close

If (mFOUND = False) Then
    x = MsgBox("No se encontro el Proveedor en el Sistema, operación cancelada.", 64, PRODUCT_NAME)
    Exit Sub
End If

'---------- UPDATE CXC ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("CXP")

    RE_CXC.AddNew
        RE_CXC.Fields("FECHA").Value = txt_Fecha_Abonos.Text
        RE_CXC.Fields("HORA").Value = CStr(Time)
        RE_CXC.Fields("ID_PROVEEDOR").Value = cbo_Clientes_Abonos.Text
        RE_CXC.Fields("NOMBRE").Value = mCLIENTE
        RE_CXC.Fields("CONCEPTO").Value = txt_Concepto_Abonos.Text
        RE_CXC.Fields("CARGO").Value = 0
        RE_CXC.Fields("ABONO").Value = CDbl(txt_Importe_Abonos.Text)
    RE_CXC.Update
    RE_CXC.Close
DB_CXC.Close



'---------- UPDATE CLIENTE ---------------------------------
Set DB_CXC = OpenDatabase(NetCashFile)
Set RE_CXC = DB_CXC.OpenRecordset("PROVEEDORES")

    mFOUND = False
    RE_CXC.MoveFirst
    Do While (RE_CXC.EOF = False)
        If (Trim(RE_CXC.Fields(0).Value) = Trim(cbo_Clientes_Abonos.Text)) Then
            RE_CXC.Edit
            RE_CXC.Fields("HABER").Value = RE_CXC.Fields("HABER").Value + CDbl(txt_Importe_Abonos.Text)
            RE_CXC.Update
            Exit Do
        End If
    RE_CXC.MoveNext
    Loop


RE_CXC.Close
DB_CXC.Close



x = MsgBox("El Abono fue aplicado en cuenta de " + cbo_Clientes_Abonos.Text + " a nombre de " + mCLIENTE, 64, PRODUCT_NAME)


cbo_Clientes_Abonos.Text = ""
txt_Concepto_Abonos.Text = ""


Exit Sub

Local_Handler_SaveCXC:
    x = MsgBox("No se puede completar la operación verifique los campos de entrada.", 64, PRODUCT_NAME)

End Sub


Private Sub Command4_Click()



Dim TX_DEBE As Double
Dim TX_HABER As Double
Dim TX_SALDO As Double


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Reporte Historial Crediticio"
    CommonDialog1.Filter = "Microsoft Excel (*.xls)|*.xls|Microsoft Word (*.doc)|*.doc|Documento Web (*.htm)|*.htm"
    CommonDialog1.Action = 2
    
    If CommonDialog1.FileName <> "" Then
        
    Dim sEmpresa As String
    Dim sRFC As String
    Dim sDireccion As String
    Dim sCiudad As String
    Dim sEstado As String
    Dim sTelefono As String
    Dim sMail As String
        
    Set objDB = OpenDatabase(NetCashFile)
    Set objRecord = objDB.OpenRecordset("EMPRESA", dbOpenTable, dbReadOnly)
   
    objRecord.MoveFirst
    sEmpresa = objRecord.Fields(0).Value
    sRFC = objRecord.Fields(1).Value
    sDireccion = objRecord.Fields(2).Value
    sCiudad = objRecord.Fields(3).Value
    sEstado = objRecord.Fields(4).Value
    sTelefono = objRecord.Fields(8).Value
    sMail = objRecord.Fields(9).Value
                
            
        'NetCashFile = CommonDialog1.filename
        Open CommonDialog1.FileName For Output As #1
            Print #1, "<html>"
            Print #1, "<body><center>"
            Print #1, "<h1>" + sEmpresa + "</h1>"
            Print #1, "<h3>Reporte de Cuentas por Pagar</h3>"
            Print #1, "<h4>" + CStr(Date) + " RFC: " + sRFC
            Print #1, sDireccion + " " + sCiudad + ", " + sEstado
            Print #1, "Tel. " + sTelefono + " Mail: " + sMail
            Print #1, "</h4></center><br><br>"
                        
            Print #1, "<center><table border='1'>"
            Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
            Print #1, "<td>"
            Print #1, "FECHA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "HORA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CUENTA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PROVEEDOR"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CARGOS"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ABONOS"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            objRecord.Close
            objDB.Close

            
                Set objDB = OpenDatabase(NetCashFile)
                Set objRecord = objDB.OpenRecordset("CXP", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                If bReportAll = False Then
                    If (objRecord.Fields(0).Value >= CDate(DateStart) And objRecord.Fields(0).Value <= CDate(DateEnd)) Then
                       'ok
                       
                       
                       
                        If (Trim(cbo_Clientes_Reportes.Text) = Trim(objRecord.Fields("ID_PROVEEDOR").Value)) Then
                        TX_DEBE = TX_DEBE + objRecord.Fields("CARGO").Value
                        TX_HABER = TX_HABER + objRecord.Fields("ABONO").Value
                                                                     
                        Print #1, "<tr>"
                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(0).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(1).Value
                        Print #1, "</td>"
                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(2).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(3).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(4).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(5).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(6).Value
                        Print #1, "</td>"
                    
                                                                        
                        Print #1, "</tr>"
                        End If
                       
                    Else
                        'not
                    End If
                Else
                
                
                        If (Trim(cbo_Clientes_Reportes.Text) = Trim(objRecord.Fields("ID_PROVEEDOR").Value)) Then
                            TX_DEBE = TX_DEBE + objRecord.Fields("CARGO").Value
                            TX_HABER = TX_HABER + objRecord.Fields("ABONO").Value
                                
                            Print #1, "<tr>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(0).Value
                            Print #1, "</td>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(1).Value
                            Print #1, "</td>"
                                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(2).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(3).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(4).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(5).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(6).Value
                            Print #1, "</td>"
                                                                                
                            Print #1, "</tr>"
                        End If
                
                End If
            
            objRecord.MoveNext
            Loop
            
            
            Rem ------------ TOTALES -------------------------
                    Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " TOTALES "
                    Print #1, "</td>"
                                
                    Print #1, "<td>"
                    Print #1, CStr(TX_DEBE)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_HABER)
                    Print #1, "</td>"
                
                    
                    Print #1, "</tr>"
            
            
            Print #1, "</table>"
            Print #1, "<br><h3>Saldo Neto: $ " + CStr(TX_DEBE - TX_HABER) + " al " + CStr(Date) + "</h3>"
                                   
            Print #1, "</center></body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenExcel(CommonDialog1.FileName)
        
    Else
        
    End If

End Sub

Private Sub AutoOpenExcel(ByVal FileName As String)

On Error GoTo LOCAL_HANDLER_EXCEL_ERROR

    Shell "C:\WINDOWS\EXCEL.EXE " + FileName, vbNormalFocus
    Exit Sub

LOCAL_HANDLER_EXCEL_ERROR:

End Sub

Private Sub Command5_Click()

    Load frmSetups
    frmSetups.Show

End Sub

Private Sub Command6_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.Action = 5

End Sub

Private Sub Command7_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub Form_Load()

    PROVEEDORES.DatabaseName = NetCashFile
    
    txt_Fecha_Cargos.Text = CStr(Date)
    txt_Fecha_Abonos.Text = CStr(Date)
    txt_Fecha_Saldos.Text = CStr(Date)

End Sub


