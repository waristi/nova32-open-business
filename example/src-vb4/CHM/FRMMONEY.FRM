VERSION 4.00
Begin VB.Form frmMoneyOuts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos de Caja"
   ClientHeight    =   4905
   ClientLeft      =   1140
   ClientTop       =   1740
   ClientWidth     =   5820
   Height          =   5310
   Icon            =   "FRMMONEY.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Top             =   1395
   Width           =   5940
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\SYSTEM.LEX"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "EGRESOS"
      Top             =   5760
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _version        =   65536
      _extentx        =   9763
      _extenty        =   8070
      _stockprops     =   15
      caption         =   "Registro"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Registro"
      tab(0).controlcount=   16
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label1"
      tab(0).control(1)=   "Label2"
      tab(0).control(2)=   "Label3"
      tab(0).control(3)=   "Label4"
      tab(0).control(4)=   "Label5"
      tab(0).control(5)=   "Label6"
      tab(0).control(6)=   "Label7"
      tab(0).control(7)=   "Image2"
      tab(0).control(8)=   "txt_Fecha"
      tab(0).control(9)=   "txt_Summary"
      tab(0).control(10)=   "txt_Importe"
      tab(0).control(11)=   "txt_IVA"
      tab(0).control(12)=   "txt_Descuento"
      tab(0).control(13)=   "txt_Total"
      tab(0).control(14)=   "txt_Cajero"
      tab(0).control(15)=   "Command1"
      tabcaption(1)   =   "Consultas"
      tab(1).controlcount=   1
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "DBGrid1"
      tabcaption(2)   =   "Reportes"
      tab(2).controlcount=   5
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "CommonDialog1"
      tab(2).control(1)=   "Command4"
      tab(2).control(2)=   "Command3"
      tab(2).control(3)=   "Picture1"
      tab(2).control(4)=   "Command2"
      Begin VB.CommandButton Command2 
         Caption         =   "Generar Reporte"
         Height          =   495
         Left            =   -71400
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Left            =   -74760
         Picture         =   "FRMMONEY.frx":0442
         ScaleHeight     =   3675
         ScaleWidth      =   3075
         TabIndex        =   19
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reporteador"
         Height          =   495
         Left            =   -71400
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Configurar Hoja"
         Height          =   495
         Left            =   -71400
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txt_Cajero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txt_Total 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_Descuento 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_IVA 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_Importe 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_Summary 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txt_Fecha 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4800
         Picture         =   "FRMMONEY.frx":29B84
         Top             =   600
         Width           =   480
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70080
         Top             =   360
         _version        =   65536
         _extentx        =   847
         _extenty        =   847
         _stockprops     =   0
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FRMMONEY.frx":29FC6
         Height          =   3855
         Left            =   -74880
         OleObjectBlob   =   "FRMMONEY.frx":29FD4
         TabIndex        =   16
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label Label7 
         Caption         =   "Cajero"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "IVA"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMoneyOuts"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public MDB As Database
Public REC As Recordset

Private Sub Command1_Click()

On Error GoTo Local_Handler_Save

    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("EGRESOS")

    REC.AddNew
    REC.Fields(0).Value = txt_Fecha.Text
    REC.Fields(1).Value = txt_Summary.Text
    REC.Fields(2).Value = CDbl(txt_Importe.Text)
    REC.Fields(3).Value = CDbl(txt_IVA.Text)
    REC.Fields(4).Value = CDbl(txt_Descuento.Text)
    REC.Fields(5).Value = CDbl(txt_Importe.Text) + CDbl(txt_IVA.Text) - CDbl(txt_Descuento.Text)
    REC.Fields(6).Value = txt_Cajero.Text
    REC.Fields(7).Value = Time
    REC.Fields(8).Value = "01"
    REC.Update

    
    REC.Close
    MDB.Close
    
    x = MsgBox("Egreso registrado en el Sistema.", 64, PRODUCT_NAME)
    
    txt_Summary.Text = ""
    txt_Importe.Text = "0"
    txt_IVA.Text = "0"
    txt_Descuento.Text = "0"
    txt_Total.Text = "0"
    
    DBGrid1.Refresh
    
    Exit Sub
    
Local_Handler_Save:

    z = MsgBox("No se pueden guardar los datos verifique los campos del formulario.", 64, PRODUCT_NAME)

End Sub




Private Sub Command2_Click()

Dim t_Importe As Double
Dim t_IVA As Double
Dim t_Descuento As Double
Dim t_Neto As Double


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Generar Reporte de Egresos"
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
            Print #1, "<h3>Reporte de Egresos</h3>"
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
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "IMPORTE"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "IVA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "DESCUENTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "NETO/TOTAL"
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
            Print #1, "</tr>"
            
            
            Set objRecord = objDB.OpenRecordset("EGRESOS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                       
                If bReportAll = False Then
                
                    If (objRecord.Fields(0).Value >= CDate(DateStart) And objRecord.Fields(0).Value <= CDate(DateEnd)) Then
                       'ok
                       
                        t_Importe = t_Importe + objRecord.Fields(2).Value
                        t_IVA = t_IVA + objRecord.Fields(3).Value
                        t_Descuento = t_Descuento + objRecord.Fields(4).Value
                        t_Neto = t_Neto + objRecord.Fields(5).Value
                       
                       
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
                                                                    
                        Print #1, "</tr>"
                       
                    Else
                        'not
                    End If
                Else
                
                
                    t_Importe = t_Importe + objRecord.Fields(2).Value
                    t_IVA = t_IVA + objRecord.Fields(3).Value
                    t_Descuento = t_Descuento + objRecord.Fields(4).Value
                    t_Neto = t_Neto + objRecord.Fields(5).Value
                
                
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
                                    
                    Print #1, "</tr>"
                
                End If
            
            objRecord.MoveNext
            Loop
            
            
            Rem ---- TOTALES -------
                    Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, " TOTALES "
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(t_Importe)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(t_IVA)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(t_Descuento)
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, CStr(t_Neto)
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
            
            
            Print #1, "</center>"
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

'
'
Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
    txt_Fecha.Text = Date
    txt_Cajero.Text = UserName
End Sub


Private Sub AutoOpenExcel(ByVal FileName As String)

On Error GoTo LOCAL_HANDLER_EXCEL_ERROR

    Shell "C:\WINDOWS\EXCEL.EXE " + FileName, vbNormalFocus
    Exit Sub

LOCAL_HANDLER_EXCEL_ERROR:

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    DBGrid1.Refresh
End Sub

