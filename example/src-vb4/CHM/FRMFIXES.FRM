VERSION 4.00
Begin VB.Form frmFixesActives 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Activos Fijos"
   ClientHeight    =   8535
   ClientLeft      =   3945
   ClientTop       =   1875
   ClientWidth     =   8535
   Height          =   8940
   Icon            =   "FRMFIXES.frx":0000
   Left            =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Top             =   1530
   Width           =   8655
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _version        =   65536
      _extentx        =   14631
      _extenty        =   14631
      _stockprops     =   15
      caption         =   "Activos"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Activos"
      tab(0).controlcount=   26
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "CommonDialog1"
      tab(0).control(1)=   "DBGrid1"
      tab(0).control(2)=   "Label6"
      tab(0).control(3)=   "Label5"
      tab(0).control(4)=   "Label4"
      tab(0).control(5)=   "Label3"
      tab(0).control(6)=   "Label2"
      tab(0).control(7)=   "Label1"
      tab(0).control(8)=   "Image1"
      tab(0).control(9)=   "Text1"
      tab(0).control(10)=   "cmdFirst"
      tab(0).control(11)=   "cmdPrevious"
      tab(0).control(12)=   "cmdNext"
      tab(0).control(13)=   "cmdLast"
      tab(0).control(14)=   "cmdNew"
      tab(0).control(15)=   "cmdSave"
      tab(0).control(16)=   "cmdeleted"
      tab(0).control(17)=   "cmdClosed"
      tab(0).control(18)=   "Data1"
      tab(0).control(19)=   "Command1"
      tab(0).control(20)=   "txt_FISCAL"
      tab(0).control(21)=   "txt_VIDA"
      tab(0).control(22)=   "txt_VENTA"
      tab(0).control(23)=   "txt_ORIGINAL"
      tab(0).control(24)=   "txt_ACTIVO"
      tab(0).control(25)=   "Frame1"
      tabcaption(1)   =   "Detalles"
      tab(1).controlcount=   2
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "Text2"
      tab(1).control(1)=   "Label9"
      tabcaption(2)   =   "Reportes"
      tab(2).controlcount=   4
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "Command4"
      tab(2).control(1)=   "Command3"
      tab(2).control(2)=   "Command2"
      tab(2).control(3)=   "Picture1"
      Begin VB.CommandButton Command4 
         Caption         =   "Configurar Hoja"
         Height          =   375
         Left            =   -70200
         TabIndex        =   33
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reporte de Especificaciones"
         Height          =   375
         Left            =   -70200
         TabIndex        =   32
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reporte de Valores"
         Height          =   375
         Left            =   -70200
         TabIndex        =   31
         Top             =   840
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Height          =   4095
         Left            =   -74520
         Picture         =   "FRMFIXES.frx":0442
         ScaleHeight     =   4035
         ScaleWidth      =   3075
         TabIndex        =   30
         Top             =   840
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Depreciación Anual"
         Height          =   1335
         Left            =   4800
         TabIndex        =   18
         Top             =   1920
         Width           =   3255
         Begin VB.TextBox res_LINEA 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox res_FISCAL 
            Height          =   285
            Left            =   1320
            TabIndex        =   19
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Línea Recta"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Fiscal"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.TextBox txt_ACTIVO 
         DataField       =   "ACTIVO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txt_ORIGINAL 
         DataField       =   "VALOR_ORIGINAL"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txt_VENTA 
         DataField       =   "VALOR_VENTA"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txt_VIDA 
         DataField       =   "VIDA_UTIL"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txt_FISCAL 
         DataField       =   "FISCAL"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\CHM\NETCASH.MDB"
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "ACTIVOS"
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdClosed 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdeleted 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   7680
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   7680
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<<"
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   7680
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   6600
         TabIndex        =   4
         Top             =   7680
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "ID_CLAVE"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         DataField       =   "CARACTERISTICAS"
         DataSource      =   "Data1"
         Height          =   7095
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "FRMFIXES.frx":29B84
         Top             =   960
         Width           =   7815
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6960
         Picture         =   "FRMFIXES.frx":29B90
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Activo Fijo"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Original"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Valor de Venta"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Vida Util"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "% Depreciación Fiscal"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3000
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FRMFIXES.frx":29FD2
         Height          =   4215
         Left            =   240
         OleObjectBlob   =   "FRMFIXES.frx":29FE0
         TabIndex        =   23
         Top             =   3360
         Width           =   7815
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3240
         Top             =   480
         _version        =   65536
         _extentx        =   847
         _extenty        =   847
         _stockprops     =   0
      End
      Begin VB.Label Label9 
         Caption         =   "Especificaciones del Activo Fijo:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmFixesActives"
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
    mg = MsgBox("El Proveedor fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Activo.", vbCritical, PRODUCT_NAME)

End Sub

Private Sub cmdFirst_Click()

On Error GoTo LocalHandlerMoveFirst
    
    Data1.Recordset.MoveFirst
    
    res_LINEA.Text = ""
    res_FISCAL.Text = ""
    
    
LocalHandlerMoveFirst:

End Sub

Private Sub cmdLast_Click()

On Error GoTo LocalHandlerMoveLast

    Data1.Recordset.MoveLast
    res_LINEA.Text = ""
    res_FISCAL.Text = ""
    
    
LocalHandlerMoveLast:

End Sub

Private Sub cmdNew_Click()

On Error GoTo LocalHandlerNew

    Data1.Refresh
    Data1.Recordset.AddNew
    mg = MsgBox("Digite la clave y descripción del nuevo Activo y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
LocalHandlerNew:

End Sub


Private Sub cmdNext_Click()

On Error GoTo LocalHandlerMoveNext
    
    Data1.Recordset.MoveNext
    res_LINEA.Text = ""
    res_FISCAL.Text = ""
    
    Exit Sub
    
LocalHandlerMoveNext:

End Sub

Private Sub cmdPrevious_Click()

On Error GoTo LocalHandlerMovePrev

    Data1.Recordset.MovePrevious
    res_LINEA.Text = ""
    res_FISCAL.Text = ""
    
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
    ms = MsgBox("No se puede guardar el Activo, ya existe.", vbCritical, PRODUCT_NAME)

End Sub


Private Sub Command1_Click()

On Error GoTo COMPUTENEW

res_LINEA.Text = CStr((CDbl(txt_ORIGINAL.Text) - CDbl(txt_VENTA.Text)) / CDbl(txt_VIDA.Text))
res_FISCAL.Text = CStr(CDbl(txt_ORIGINAL.Text) * (CDbl(txt_FISCAL.Text) / 100))


COMPUTENEW:

End Sub


Private Sub Command2_Click()

Dim TX_LINEAL As Double
Dim TX_FISCAL As Double

Dim TX_ORIGINAL As Double
Dim TX_VENTA As Double


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Reporte Activos Fijos"
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
            Print #1, "<h3>Reporte de Activos Fijos</h3>"
            Print #1, "<h4>" + CStr(Date) + " RFC: " + sRFC
            Print #1, sDireccion + " " + sCiudad + ", " + sEstado
            Print #1, "Tel. " + sTelefono + " Mail: " + sMail
            Print #1, "</h4></center><br><br>"
                        
            Print #1, "<center><table border='1'>"
            Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
            Print #1, "<td>"
            Print #1, "CLAVE"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ACTIVO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "VALOR ORIGINAL"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "VALOR DE VENTA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "VIDA UTIL"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PORC. FISCAL"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "DEP.LINEAL"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "DEP.FISCAL"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            objRecord.Close
            objDB.Close

            
                Set objDB = OpenDatabase(NetCashFile)
                Set objRecord = objDB.OpenRecordset("ACTIVOS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                            Print #1, "<tr>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(0).Value
                            Print #1, "</td>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(1).Value
                            Print #1, "</td>"
                                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(2).Value
                            TX_ORIGINAL = TX_ORIGINAL + objRecord.Fields(2).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(3).Value
                            TX_VENTA = TX_VENTA + objRecord.Fields(3).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(4).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(5).Value
                            Print #1, "</td>"
                
                            Print #1, "<td>"
                            TX_LINEAL = TX_LINEAL + (objRecord.Fields("VALOR_ORIGINAL").Value - objRecord.Fields("VALOR_VENTA").Value) / objRecord.Fields("VIDA_UTIL").Value
                            Print #1, (objRecord.Fields("VALOR_ORIGINAL").Value - objRecord.Fields("VALOR_VENTA").Value) / objRecord.Fields("VIDA_UTIL").Value
                            Print #1, "</td>"
                                                                                
                            Print #1, "<td>"
                            TX_FISCAL = TX_FISCAL + (objRecord.Fields("VALOR_ORIGINAL").Value * (objRecord.Fields("FISCAL").Value / 100))
                            Print #1, (objRecord.Fields("VALOR_ORIGINAL").Value * (objRecord.Fields("FISCAL").Value / 100))
                            Print #1, "</td>"
                                                                                
                            Print #1, "</tr>"
                
                
            
            objRecord.MoveNext
            Loop
            
            
            Rem ------------ TOTALES -------------------------
                    Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
                    
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, "TOTALES "
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, TX_ORIGINAL
                    Print #1, "</td>"
            
                    Print #1, "<td>"
                    Print #1, TX_VENTA
                    Print #1, "</td>"
                                
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, " "
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, TX_LINEAL
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, TX_FISCAL
                    Print #1, "</td>"
                    
                    Print #1, "</tr>"
            
            
            Print #1, "</table>"
                                   
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

Private Sub Command3_Click()

Dim TX_LINEAL As Double
Dim TX_FISCAL As Double

Dim TX_ORIGINAL As Double
Dim TX_VENTA As Double


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Reporte Activos Fijos"
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
            Print #1, "<h3>Reporte de Especificaciones de Activos Fijos</h3>"
            Print #1, "<h4>" + CStr(Date) + " RFC: " + sRFC
            Print #1, sDireccion + " " + sCiudad + ", " + sEstado
            Print #1, "Tel. " + sTelefono + " Mail: " + sMail
            Print #1, "</h4></center><br><br>"
                        
            Print #1, "<center><table border='1'>"
            Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
            Print #1, "<td>"
            Print #1, "CLAVE"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ACTIVO"
            Print #1, "</td>"
            Print #1, "</tr></table><br><br>"
            
            
            objRecord.Close
            objDB.Close

            
                Set objDB = OpenDatabase(NetCashFile)
                Set objRecord = objDB.OpenRecordset("ACTIVOS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                            Print #1, "<table border='1'><tr>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(0).Value
                            Print #1, "</td>"
                    
                            Print #1, "<td>"
                            Print #1, objRecord.Fields(1).Value
                            Print #1, "</td>"
                                                                                                                    
                            Print #1, "</tr></table>"
                                
                            Print #1, "<p>" + objRecord.Fields(6).Value + "</p>"
                
            
            objRecord.MoveNext
            Loop
                                                          
            Print #1, "</center></body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenExcel(CommonDialog1.FileName)
        
    Else
        
    End If

End Sub

Private Sub Command4_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.Action = 5

End Sub

Private Sub DBGrid1_Click()

    res_LINEA.Text = ""
    res_FISCAL.Text = ""
    
End Sub


