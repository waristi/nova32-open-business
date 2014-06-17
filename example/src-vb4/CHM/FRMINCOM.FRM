VERSION 4.00
Begin VB.Form frmIncomeCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de ventas"
   ClientHeight    =   6735
   ClientLeft      =   12915
   ClientTop       =   4470
   ClientWidth     =   7215
   Height          =   7140
   Icon            =   "FRMINCOM.frx":0000
   Left            =   12855
   LinkTopic       =   "frmIncomeCancel"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Top             =   4125
   Width           =   7335
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Z:\NETCASH.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "CANING"
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Data QuickView 
      Caption         =   "QuickView"
      Connect         =   "Access"
      DatabaseName    =   "Z:\NETCASH.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "INGRESOS"
      Top             =   7440
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _version        =   65536
      _extentx        =   12303
      _extenty        =   11456
      _stockprops     =   15
      caption         =   "Cancelar"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Cancelar"
      tab(0).controlcount=   14
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Image1"
      tab(0).control(1)=   "Label6"
      tab(0).control(2)=   "Label5"
      tab(0).control(3)=   "Label4"
      tab(0).control(4)=   "Label2"
      tab(0).control(5)=   "DBGrid2"
      tab(0).control(6)=   "Label3"
      tab(0).control(7)=   "txt_Motivo_Cancela"
      tab(0).control(8)=   "txt_Hora_Cancela"
      tab(0).control(9)=   "txt_Fecha_Cancela"
      tab(0).control(10)=   "txt_Autoriza"
      tab(0).control(11)=   "Command1"
      tab(0).control(12)=   "cmd_CancelNow"
      tab(0).control(13)=   "txt_Folio"
      tabcaption(1)   =   "Consultas"
      tab(1).controlcount=   4
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "txt_SeekDate"
      tab(1).control(1)=   "cmd_SeekDate"
      tab(1).control(2)=   "DBGrid1"
      tab(1).control(3)=   "Label1"
      tabcaption(2)   =   "Reporte"
      tab(2).controlcount=   5
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "Picture1"
      tab(2).control(1)=   "Command2"
      tab(2).control(2)=   "Command3"
      tab(2).control(3)=   "Command4"
      tab(2).control(4)=   "CommonDialog1"
      Begin VB.TextBox txt_SeekDate 
         Height          =   285
         Left            =   -73080
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmd_SeekDate 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   -71400
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txt_Folio 
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Text            =   "0"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmd_CancelNow 
         Caption         =   "Cancelar Transacción"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar.."
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_Autoriza 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox txt_Fecha_Cancela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txt_Hora_Cancela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txt_Motivo_Cancela 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   5640
         Width           =   5895
      End
      Begin VB.PictureBox Picture1 
         Height          =   4095
         Left            =   -74280
         Picture         =   "FRMINCOM.frx":0442
         ScaleHeight     =   4035
         ScaleWidth      =   3075
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Generar Reporte"
         Height          =   495
         Left            =   -70320
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reporteador"
         Height          =   495
         Left            =   -70320
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Configurar Hoja"
         Height          =   495
         Left            =   -70320
         TabIndex        =   1
         Top             =   2400
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FRMINCOM.frx":29B84
         Height          =   5415
         Left            =   -74880
         OleObjectBlob   =   "FRMINCOM.frx":29B92
         TabIndex        =   21
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Folio de la transacción:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Folio de la transacción:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "FRMINCOM.frx":2A29A
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "FRMINCOM.frx":2A2AC
         TabIndex        =   18
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Autoriza"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de cancelación"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Hora de Cancelación"
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   5040
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70320
         Top             =   3120
         _version        =   65536
         _extentx        =   847
         _extenty        =   847
         _stockprops     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6360
         Picture         =   "FRMINCOM.frx":2A9B4
         Top             =   480
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmIncomeCancel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim TFECHA As String
Dim TCONCEPTO As String
Dim TNETO As Double
Dim TBRUTO As Double
Dim TDESCUENTO As Double
Dim TIVA As Double
Dim TCOBRO As Double
Dim TCAMBIO As Double
Dim TCAJERO As String
Dim THORA As String
Dim TCAJA As String
Dim TARTICULOS As String
Dim TFOLIO As Integer
Dim TNOMBRE As String
Dim TMOTIVO As String
Dim TAUTORIZA As String
Dim TFECHA_CAN As String
Dim THORA_CAN As String

Private Sub cmd_CancelNow_Click()

On Error GoTo LOCAL_HANDLER_CANCEL_NULL

Dim mFOUND As Boolean

Dim MDB As Database
Dim REC As Recordset


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("INGRESOS")

    REC.MoveFirst
    mFOUND = False

    Do While (REC.EOF <> True)

        If (CInt(txt_Folio.Text) = REC.Fields("FOLIO").Value) Then
            TFECHA = REC.Fields("FECHA").Value
            TCONCEPTO = REC.Fields("CONCEPTO").Value
            TNETO = REC.Fields("NETO").Value
            TBRUTO = REC.Fields("BRUTO").Value
            TDESCUENTO = REC.Fields("DESCUENTO").Value
            TIVA = REC.Fields("IVA").Value
            TCOBRO = REC.Fields("COBRO").Value
            TCAMBIO = REC.Fields("CAMBIO").Value
            TCAJERO = REC.Fields("CAJERO").Value
            THORA = REC.Fields("HORA").Value
            TCAJA = REC.Fields("CAJA").Value
            TARTICULOS = REC.Fields("ARTICULOS").Value
            TFOLIO = REC.Fields("FOLIO").Value
            TNOMBRE = REC.Fields("NOMBRE").Value
            TMOTIVO = txt_Motivo_Cancela.Text
            TAUTORIZA = txt_Autoriza.Text
            TFECHA_CAN = txt_Fecha_Cancela.Text
            THORA_CAN = CStr(Time)
            mFOUND = True
            Exit Do
        End If

    REC.MoveNext
    Loop


REC.Close
MDB.Close


If (mFOUND = False) Then
    x = MsgBox("No existe el folio " + txt_Folio.Text + " a eliminar.", 64, PRODUCT_NAME)
    Exit Sub
End If



x = MsgBox("¿ Desea cancelar la venta con folio " + CStr(TFOLIO) + " de fecha " + CStr(TFECHA) + "  ?", 4, PRODUCT_NAME)

If x = 7 Then
    Exit Sub
End If


'**************************************************************************
' ELIMINAR FISICAMENTE EL REGISTRO DE BASE DE DATOS
'
'**************************************************************************


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("INGRESOS")


    REC.MoveFirst

    Do While (REC.EOF <> True)

        If (CInt(txt_Folio.Text) = REC.Fields("FOLIO").Value) Then
            REC.Delete
            Exit Do
        End If

    REC.MoveNext
    Loop


REC.Close
MDB.Close


'**************************************************************************
' AGREGAR EL REGISTRO DE CANCELACION EN BASE DE DATOS
'
'**************************************************************************


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("CANING")

    REC.AddNew
    REC.Fields("FECHA").Value = TFECHA
    REC.Fields("CONCEPTO").Value = TCONCEPTO
    REC.Fields("NETO").Value = TNETO
    REC.Fields("BRUTO").Value = TBRUTO
    REC.Fields("DESCUENTO").Value = TDESCUENTO
    REC.Fields("IVA").Value = TIVA
    REC.Fields("COBRO").Value = TCOBRO
    REC.Fields("CAMBIO").Value = TCAMBIO
    REC.Fields("CAJERO").Value = TCAJERO
    REC.Fields("HORA").Value = THORA
    REC.Fields("CAJA").Value = TCAJA
    REC.Fields("ARTICULOS").Value = TARTICULOS
    REC.Fields("FOLIO").Value = TFOLIO
    REC.Fields("NOMBRE").Value = TNOMBRE
    REC.Fields("MOTIVO").Value = TMOTIVO
    REC.Fields("AUTORIZA").Value = TAUTORIZA
    REC.Fields("FECHA_CAN").Value = TFECHA_CAN
    REC.Fields("HORA_CAN").Value = THORA_CAN
    REC.Update

REC.Close
MDB.Close


txt_Folio.Text = "0"
txt_Hora_Cancela.Text = CStr(Time)
txt_Motivo_Cancela.Text = ""


x = MsgBox("La venta con Folio " + CStr(TFOLIO) + " de fecha " + TFECHA + " fue cancelada.", 64, PRODUCT_NAME)



LOCAL_HANDLER_CANCEL_NULL:



End Sub

Private Sub AutoOpenExcel(ByVal FileName As String)

On Error GoTo LOCAL_HANDLER_EXCEL_ERROR

    Shell "C:\WINDOWS\EXCEL.EXE " + FileName, vbNormalFocus
    Exit Sub

LOCAL_HANDLER_EXCEL_ERROR:

End Sub

Private Sub cmd_SeekDate_Click()
On Error GoTo LocalSearch

Data1.Recordset.FindFirst ("FOLIO=" + txt_SeekDate.Text)

LocalSearch:
End Sub

Private Sub Command1_Click()

On Error GoTo LOCAL_HANDLER_SEARCH

    QuickView.Recordset.FindFirst ("FOLIO=" + txt_Folio.Text)
    Exit Sub

LOCAL_HANDLER_SEARCH:
    x = MsgBox("Valor de campo de entrada incorrecto, intente nuevamente.", 64, PRODUCT_NAME)

End Sub


Private Sub Command2_Click()


Dim TX_NETO As Double
Dim TX_BRUTO As Double
Dim TX_DESCUENTO As Double
Dim TX_IVA As Double
Dim TX_COBRO As Double
Dim TX_CAMBIO As Double



If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Reporte de Folios Cancelados"
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
            Print #1, "<h3>Reporte de Folios Cancelados</h3>"
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
            Print #1, "TIPO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "NETO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "BRUTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "DESCUENTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "IVA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "COBRO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CAMBIO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CAJERO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "HORA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CAJA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ARTICULOS"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FOLIO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "MOTIVO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "AUTORIZA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FECHA CANCELACION"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "HORA CANCELACION"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            objRecord.Close
            objDB.Close

            
                Set objDB = OpenDatabase(NetCashFile)
                Set objRecord = objDB.OpenRecordset("CANING", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                If bReportAll = False Then
                    If (objRecord.Fields(0).Value >= CDate(DateStart) And objRecord.Fields(0).Value <= CDate(DateEnd)) Then
                       'ok
                       
                        TX_NETO = TX_NETO + objRecord.Fields("NETO").Value
                        TX_BRUTO = TX_BRUTO + objRecord.Fields("BRUTO").Value
                        TX_DESCUENTO = TX_DESCUENTO + objRecord.Fields("DESCUENTO").Value
                        TX_IVA = TX_IVA + objRecord.Fields("IVA").Value
                        TX_COBRO = TX_COBRO + objRecord.Fields("COBRO").Value
                        TX_CAMBIO = TX_CAMBIO + objRecord.Fields("CAMBIO").Value
                                                                     
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
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(7).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(8).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(9).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(10).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(11).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(12).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(13).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(14).Value
                        Print #1, "</td>"
                                                                                                                                                
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(15).Value
                        Print #1, "</td>"
                                                                                                                                                
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(16).Value
                        Print #1, "</td>"
                                                                        
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(17).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "</tr>"
                       
                    Else
                        'not
                    End If
                Else
                
                
                        TX_NETO = TX_NETO + objRecord.Fields("NETO").Value
                        TX_BRUTO = TX_BRUTO + objRecord.Fields("BRUTO").Value
                        TX_DESCUENTO = TX_DESCUENTO + objRecord.Fields("DESCUENTO").Value
                        TX_IVA = TX_IVA + objRecord.Fields("IVA").Value
                        TX_COBRO = TX_COBRO + objRecord.Fields("COBRO").Value
                        TX_CAMBIO = TX_CAMBIO + objRecord.Fields("CAMBIO").Value
                
                
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
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(7).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(8).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(9).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(10).Value
                    Print #1, "</td>"
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(11).Value
                    Print #1, "</td>"
                                                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(12).Value
                    Print #1, "</td>"
                                        
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(13).Value
                    Print #1, "</td>"
                                        
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(14).Value
                    Print #1, "</td>"
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(15).Value
                    Print #1, "</td>"
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(16).Value
                    Print #1, "</td>"
                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(17).Value
                    Print #1, "</td>"
                                        
                                                                                
                    Print #1, "</tr>"
                
                End If
            
            objRecord.MoveNext
            Loop
            
            
            Rem ------------ TOTALES -------------------------
                    Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " TOTALES "
                    Print #1, "</td>"
                                
                    Print #1, "<td>"
                    Print #1, CStr(TX_NETO)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_BRUTO)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_DESCUENTO)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_IVA)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_COBRO)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(TX_CAMBIO)
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
                    Print #1, " "
                    Print #1, "</td>"
                                        
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
                    
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
                    
                    Print #1, "</tr>"
            
            
            Print #1, "</table></center>"
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenExcel(CommonDialog1.FileName)
        
    Else
        
    End If

End Sub

Private Sub Command3_Click()

  Load frmSetups
    frmSetups.Show
    
End Sub

Private Sub Command4_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.Action = 5
    
End Sub

Private Sub Form_Load()

    Data1.DatabaseName = NetCashFile
    QuickView.DatabaseName = NetCashFile
    txt_Autoriza.Text = UserName
    txt_Fecha_Cancela.Text = CStr(Date)
    txt_Hora_Cancela.Text = CStr(Time)

End Sub


