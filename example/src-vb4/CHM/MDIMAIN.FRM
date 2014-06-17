VERSION 4.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00808080&
   Caption         =   "Nova32 NetCash Small Business 2.0"
   ClientHeight    =   7365
   ClientLeft      =   1305
   ClientTop       =   2400
   ClientWidth     =   13950
   Height          =   8175
   Icon            =   "MDIMAIN.frx":0000
   Left            =   1245
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMAIN.frx":0442
   Top             =   1650
   Width           =   14070
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1440
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin VB.Menu mFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mNewSession 
         Caption         =   "Iniciar sesión..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mCloseSession 
         Caption         =   "Cerrar sesión"
      End
      Begin VB.Menu mFile0 
         Caption         =   "-"
      End
      Begin VB.Menu mDetailsEnterprise 
         Caption         =   "Entidad"
      End
      Begin VB.Menu mFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mPrintSetup 
         Caption         =   "Configurar impresora..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mPrintPreview 
         Caption         =   "Reportes"
      End
      Begin VB.Menu mFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mCopy 
         Caption         =   "Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mCut 
         Caption         =   "Cortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mPaste 
         Caption         =   "Pegar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mStore 
      Caption         =   "&Negocio"
      Begin VB.Menu mStoreProds 
         Caption         =   "Artículos de almacén"
      End
      Begin VB.Menu descadmin 
         Caption         =   "Administrador de descuentos"
      End
      Begin VB.Menu mCatalogs 
         Caption         =   "Catálogos"
         Begin VB.Menu mSubStore 
            Caption         =   "Almacén"
         End
         Begin VB.Menu mSubProv 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu mprclientscats 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mSubFamily 
            Caption         =   "Grupos"
         End
         Begin VB.Menu mSubTypProds 
            Caption         =   "Tipos de Artículos"
         End
         Begin VB.Menu mSubUnits 
            Caption         =   "Unidades"
         End
         Begin VB.Menu mSubCas 
            Caption         =   "Casilleros"
         End
         Begin VB.Menu mSubFloor 
            Caption         =   "Pasillos"
         End
         Begin VB.Menu mSubOperations 
            Caption         =   "Operaciones"
         End
      End
      Begin VB.Menu mStores1 
         Caption         =   "-"
      End
      Begin VB.Menu mInStore 
         Caption         =   "Entradas al almacén"
      End
      Begin VB.Menu mOutStore 
         Caption         =   "Salidas de almacén"
      End
      Begin VB.Menu mStorRefresh 
         Caption         =   "Existencias de almacén"
      End
      Begin VB.Menu mStores2 
         Caption         =   "-"
      End
      Begin VB.Menu mpr_Egresos 
         Caption         =   "Egresos de caja"
      End
      Begin VB.Menu mprFolies 
         Caption         =   "Folio activo"
      End
      Begin VB.Menu mpr_InitialValue 
         Caption         =   "Saldo inicial de caja"
      End
      Begin VB.Menu mpr_CutX 
         Caption         =   "Corte de caja X"
      End
      Begin VB.Menu mprCut_Z 
         Caption         =   "Corte de caja Z"
      End
      Begin VB.Menu mReports 
         Caption         =   "Reportes"
         Begin VB.Menu mSubDocInputs 
            Caption         =   "Entradas al almacén"
         End
         Begin VB.Menu mSubDocOuts 
            Caption         =   "Salidas de almacén"
         End
         Begin VB.Menu mSubDocExist 
            Caption         =   "Existencias de almacén"
         End
      End
   End
   Begin VB.Menu mcash 
      Caption         =   "&Ingresos"
      Begin VB.Menu mSalesPoint 
         Caption         =   "Punto de ventas"
         Shortcut        =   ^T
      End
      Begin VB.Menu mInsumesFixeds 
         Caption         =   "Reporte de venta"
         Shortcut        =   ^G
      End
      Begin VB.Menu lineing1 
         Caption         =   "-"
      End
      Begin VB.Menu mcansales 
         Caption         =   "Cancelación de ventas"
      End
   End
   Begin VB.Menu mprContraloria 
      Caption         =   "&Administración"
      Begin VB.Menu mprCxC 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu mprCxP 
         Caption         =   "Cuentas por Pagar"
      End
      Begin VB.Menu mprActivoFijos 
         Caption         =   "Control de Activos Fijos"
      End
   End
   Begin VB.Menu mprSAT 
      Caption         =   "&SAT"
      Begin VB.Menu mprCFDISAT 
         Caption         =   "Facturación CFDI y Contabilidad"
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "&Procesos"
      Begin VB.Menu mSetups 
         Caption         =   "Configuraciones"
      End
      Begin VB.Menu mUsers 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mDevices 
         Caption         =   "Dispositivos"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "A&yuda"
      Begin VB.Menu mIndex 
         Caption         =   "índice y contenido"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim objDB As Database
Dim objRecord As Recordset

Private Sub AutoOpenExcel(ByVal FileName As String)

On Error GoTo LOCAL_HANDLER_EXCEL_ERROR

    Shell "C:\WINDOWS\EXCEL.EXE " + FileName, vbNormalFocus
    Exit Sub

LOCAL_HANDLER_EXCEL_ERROR:

End Sub







Private Sub descadmin_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmOnSales
    frmOnSales.Show
End Sub

'

Private Sub mAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub


Private Sub mClose_Click()

        UseDataBase = False
        NetCashFile = ""
        Me.Caption = PRODUCT_NAME

End Sub


Private Sub mcansales_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If

    Load frmIncomeCancel
    frmIncomeCancel.Show

End Sub

Private Sub mCloseSession_Click()
    UserLogon = False
    UserProfile = 0
End Sub

Private Sub mDetailsEnterprise_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If

Load frmEnterprise
frmEnterprise.Show


End Sub

Private Sub mDevices_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If

Load frmDispositive
frmDispositive.Show

End Sub

Private Sub MDIForm_Load()

    ERROR_SESSION_MESSAGE = "Debe iniciar Sesión para usar el Sistema."
    bReportAll = True


    Dim Line As String

    Open "C:\CHM\SERVER.INI" For Input As #1
        Input #1, Line
        Input #1, Line
    Close #1

    UseDataBase = True
    NetCashFile = Line
                  
    PRODUCT_NAME = "Nova32 NetCash Small Business 2.0"
    Me.Caption = PRODUCT_NAME

End Sub

Private Sub AutoOpenXLS(ByVal FileName As String)

On Error GoTo LOCAL_HANDLER_EXCEL_ERROR

    Shell "C:\WINDOWS\EXCEL.EXE " + FileName, vbNormalFocus
    Exit Sub

LOCAL_HANDLER_EXCEL_ERROR:

End Sub

Private Sub mExit_Click()
    End
End Sub




Private Sub mIndex_Click()
    
    Shell "EXPLORER.EXE C:\CHM\HELP.PDF", vbNormalFocus
    
End Sub

Private Sub mInStore_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmInputs
    frmInputs.Show
    
End Sub

Private Sub mInsumesFixeds_Click()

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
    CommonDialog1.DialogTitle = "Generar Reporte de Ventas"
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
            Print #1, "<h3>Reporte de Ventas</h3>"
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
            Print #1, "ARTICULOS"
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
            Print #1, "FOLIO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            Set objRecord = objDB.OpenRecordset("INGRESOS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                If bReportAll = False Then
                    If (objRecord.Fields(0).Value >= CDate(DateStart) And objRecord.Fields(0).Value <= CDate(DateEnd)) Then
                       'ok
                       
                        TX_NETO = TX_NETO + objRecord.Fields(2).Value
                        TX_BRUTO = TX_BRUTO + objRecord.Fields(3).Value
                        TX_DESCUENTO = TX_DESCUENTO + objRecord.Fields(4).Value
                        TX_IVA = TX_IVA + objRecord.Fields(5).Value
                        TX_COBRO = TX_COBRO + objRecord.Fields(6).Value
                        TX_CAMBIO = TX_CAMBIO + objRecord.Fields(7).Value
                                                                     
                        Print #1, "<tr>"
                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(0).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(1).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(11).Value
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
                        Print #1, objRecord.Fields(12).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(13).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "</tr>"
                       
                    Else
                        'not
                    End If
                Else
                
                
                    TX_NETO = TX_NETO + objRecord.Fields(2).Value
                    TX_BRUTO = TX_BRUTO + objRecord.Fields(3).Value
                    TX_DESCUENTO = TX_DESCUENTO + objRecord.Fields(4).Value
                    TX_IVA = TX_IVA + objRecord.Fields(5).Value
                    TX_COBRO = TX_COBRO + objRecord.Fields(6).Value
                    TX_CAMBIO = TX_CAMBIO + objRecord.Fields(7).Value
                
                
                    Print #1, "<tr>"
                    
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(0).Value
                    Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(1).Value
                        Print #1, "</td>"
                    
            
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(11).Value
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
                    Print #1, objRecord.Fields(12).Value
                    Print #1, "</td>"
                                                                        
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(13).Value
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
                    Print #1, " "
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
                                        
                    Print #1, "</tr>"
            
            
            Print #1, "</table></center>"
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenXLS(CommonDialog1.FileName)
        
    Else
        
    End If



End Sub




Private Sub mNewEnterprise_Click()
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Crear Empresa"
    CommonDialog1.Filter = "Empresa Etecxa Ventas (*.chm)|*.chm"
    CommonDialog1.Action = 2
    
    If CommonDialog1.FileName <> "" Then
        UseDataBase = True
        NetCashFile = CommonDialog1.FileName
        FileCopy "c:\nova32.dot", NetCashFile
        Me.Caption = "Etecxa Ventas 2.0 Express - " + CommonDialog1.FileName
        Load frmSession
        frmSession.Show
    Else
        UseDataBase = False
    End If

End Sub

Private Sub mNewSession_Click()
    
    If (UseDataBase) Then
        Load frmSession
        frmSession.Show
    Else
        mgs = MsgBox("Debe abrir una Empresa para poder iniciar Sesión", vbCritical, PRODUCT_NAME)
    End If
    
End Sub

Private Sub mOpenEnterprise_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Abrir Empresa"
    CommonDialog1.Filter = "Empresa Etecxa Ventas (*.MDB)|*.MDB"
    CommonDialog1.Action = 1
    
    If CommonDialog1.FileName <> "" Then
        UseDataBase = True
        NetCashFile = CommonDialog1.FileName
        Me.Caption = "Etecxa Ventas 2.0 Express - " + CommonDialog1.FileName
        Load frmSession
        frmSession.Show
    Else
        UseDataBase = False
    End If
    
End Sub

Private Sub mOutStore_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If

    Load frmOutputs
    frmOutputs.Show
    
End Sub

Private Sub mpr_CutX_Click()

TYPE_REPORT_ZX = "Reporte de Corte de Caja"

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If


Rem *********************************************************************
Rem * APERTURA DE BASE DE DATOS                                         *
Rem *                                                                   *
Rem *********************************************************************

    Set objDB = OpenDatabase(NetCashFile)



Rem *********************************************************************
Rem * SALDO INICIAL                                                     *
Rem *                                                                   *
Rem *********************************************************************

Dim TX_SALDO_INICIAL As Double


    Rem ------ LEE SALDO INICIAL -----------------------
    Set objRecord = objDB.OpenRecordset("INICIAL", dbOpenTable, dbReadOnly)
        
        objRecord.MoveFirst
        TX_SALDO_INICIAL = objRecord.Fields(0).Value
        
    objRecord.Close


Rem *********************************************************************
Rem * INGRESO                                                           *
Rem *                                                                   *
Rem *********************************************************************

Dim TX_NETO As Double
Dim TX_BRUTO As Double
Dim TX_DESCUENTO As Double
Dim TX_IVA As Double
Dim TX_COBRO As Double
Dim TX_CAMBIO As Double

    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Generar " + TYPE_REPORT_ZX
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
    
    Rem ------ LEE EMPRESA -----------------------
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
            Print #1, "<h3>" + TYPE_REPORT_ZX + "</h3>"
            Print #1, "<h4>" + CStr(Date) + " RFC: " + sRFC
            Print #1, sDireccion + " " + sCiudad + ", " + sEstado
            Print #1, "Tel. " + sTelefono + " Mail: " + sMail
            Print #1, "</h4></center><br><br>"
            
            
            
            Print #1, "<h3>Saldo Inicial: $" + Format(TX_SALDO_INICIAL, "###,##0.00") + "</h3>"
                        
            Print #1, "<center><table border='1'>"
            Print #1, "<tr style='background-color:#0B0B61;color:#FFFFFF;'>"
            Print #1, "<td>"
            Print #1, "FECHA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ARTICULOS"
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
            Print #1, "FOLIO"
            Print #1, "</td>"
            
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            
            Print #1, "<td>"
            Print #1, "TIPO DE PAGO"
            Print #1, "</td>"
                        
            Print #1, "</tr>"
            
            Set objRecord = objDB.OpenRecordset("INGRESOS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
                                                             
                    If (objRecord.Fields(0).Value = Date) Then
                       'ok
                        Print #1, "<tr>"
                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(0).Value
                        Print #1, "</td>"
                    
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(11).Value
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
                        Print #1, objRecord.Fields(12).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(13).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "<td>"
                        Print #1, objRecord.Fields(1).Value
                        Print #1, "</td>"
                                                                        
                        Print #1, "</tr>"
                       
                        TX_NETO = TX_NETO + objRecord.Fields(2).Value
                        TX_BRUTO = TX_BRUTO + objRecord.Fields(3).Value
                        TX_DESCUENTO = TX_DESCUENTO + objRecord.Fields(4).Value
                        TX_IVA = TX_IVA + objRecord.Fields(5).Value
                        TX_COBRO = TX_COBRO + objRecord.Fields(6).Value
                        TX_CAMBIO = TX_CAMBIO + objRecord.Fields(7).Value
                       
                    Else
                        'not
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
                                        
                    Print #1, "</tr>"
            
            
            Print #1, "</table></center>"
    
    objRecord.Close
    
    Else
        Exit Sub
    End If


    Print #1, "<h3> (+) Ingresos: $" + Format(TX_NETO, "###,##0.00") + "</h3>"


Rem *********************************************************************
Rem * EGRESOS                                                           *
Rem *                                                                   *
Rem *********************************************************************

Dim t_Importe As Double
Dim t_IVA As Double
Dim t_Descuento As Double
Dim t_Neto As Double

                        
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
                                                   
                    If (objRecord.Fields(0).Value = Date) Then
                    'ok
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
                        
                        t_Importe = t_Importe + objRecord.Fields(2).Value
                        t_IVA = t_IVA + objRecord.Fields(3).Value
                        t_Descuento = t_Descuento + objRecord.Fields(4).Value
                        t_Neto = t_Neto + objRecord.Fields(5).Value
                        
                       
                    Else
                        'not
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
            
            
            Print #1, "</table></center>"
            
            
            
            Print #1, "<h3> (-) Egresos: $" + Format(t_Neto, "###,##0.00") + "</h3>"
            
            Dim Result As Double
            
            Result = TX_SALDO_INICIAL + TX_NETO - t_Neto
            
            Print #1, "<h3> (=) Corte de Caja por $" + Format(Result, "###,##0.00") + " realizado el " + CStr(Date) + " a las " + CStr(Time) + " horas. </h3>"
            
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    Call AutoOpenXLS(CommonDialog1.FileName)


End Sub

Private Sub mpr_Egresos_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmMoneyOuts
    frmMoneyOuts.Show
End Sub

Private Sub mpr_InitialValue_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If

    Load frmInitialCash
    frmInitialCash.Show
End Sub

Private Sub mprActivoFijos_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If


Load frmFixesActives
frmFixesActives.Show

End Sub

Private Sub mprCFDISAT_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Shell "C:\CHM\CHBR32 https://rfs.siat.sat.gob.mx/PTSC/RFS/menu/", vbNormalFocus

End Sub

Private Sub mprclientscats_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmClientsCat
    frmClientsCat.Show
End Sub

Private Sub mprCut_Z_Click()

TYPE_REPORT_ZX = "Corte de Caja Z"

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If


x = MsgBox("¿ Esta operación genera el corte de Caja y cierre de día, una vez realizada no podra continuar realizando transacciones con la Caja para la fecha actual, Desea continuar  ?", 4, PRODUCT_NAME)


If x = 7 Then
    Exit Sub
End If


Rem ------- Update  CIERRES --------------------------------------------
Dim MDB As Database
Dim REC As Recordset
Dim Found As Boolean


    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("CIERRES")
        
    Found = False
    REC.MoveFirst
        
    Do While (REC.EOF <> True)
    
        If REC.Fields(0).Value = Date Then
            Found = True
            Exit Do
        End If
        
    REC.MoveNext
    Loop
        
    If Found = False Then
        REC.AddNew
        REC.Fields(0).Value = Date
        REC.Update
    End If

REC.Close
MDB.Close


Call mpr_CutX_Click


Call mpr_InitialValue_Click


x = MsgBox("Corte de Caja Z, realizado, el día contable " + CStr(Date) + " ha sido cerrado.", 64, PRODUCT_NAME)


End Sub

Private Sub mprCxC_Click()


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If


    Load frmCXC
    frmCXC.Show
    
End Sub

Private Sub mprCxP_Click()


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If


    Load frmCXP
    frmCXP.Show

End Sub

Private Sub mprFolies_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If

    Load frmNumber
    frmNumber.Show
End Sub

Private Sub mPrintPreview_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
Load frmSetups
frmSetups.Show

End Sub

Private Sub mPrintSetup_Click()

    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.Action = 5
       
End Sub

Private Sub mSalesPoint_Click()


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    

    Load frmNetCash
    frmNetCash.Show
    
End Sub


Private Sub mSetups_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
Load frmSetups
frmSetups.Show

End Sub

Private Sub mStoreProds_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmCStorageProductsDef
    frmCStorageProductsDef.Show
    
End Sub

Private Sub mStorRefresh_Click()


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmCExist
    frmCExist.Show
    
End Sub

Private Sub mSubCas_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCCas
    frmCCas.Show
    
End Sub

Private Sub mSubDocExist_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Generar Reporte Existencias de Almacén"
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
            Print #1, "<h3>Reporte de Existencias de Almac&eacute;n</h3>"
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
            Print #1, "CODIGO DE BARRAS"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ARTICULO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "MARCA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "MODELO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "SERIE"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "UNIDAD"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "COSTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "MAXIMO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "MINIMO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ALMACEN"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PASILLO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CASILLERO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FACTURA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REMISION"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REFERENCIA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ORDEN"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "TIPO ARTICULO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PEDIDO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FECHA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PRECIO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PROVEEDOR"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CLASIFICACION"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "EXISTENCIAS"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            Set objRecord = objDB.OpenRecordset("ALMACEN", dbOpenTable, dbReadOnly)
            
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
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(18).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(19).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(20).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(21).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(22).Value
                    Print #1, "</td>"
                
                    Print #1, "<td>"
                    Print #1, objRecord.Fields(23).Value
                    Print #1, "</td>"
                    
                            
                    Print #1, "</tr>"

                                
            
            objRecord.MoveNext
            Loop
            
            Print #1, "</table></center>"
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenXLS(CommonDialog1.FileName)
    
    
    Else
        
    End If

End Sub

Private Sub mSubDocInputs_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Generar Reporte Entradas de Almacén"
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
            Print #1, "<h3>Reporte de Entradas de Almac&eacute;n</h3>"
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
            Print #1, "CODIGO DE BARRAS"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CANTIDAD"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FECHA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FACTURA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REMISION"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REFERENCIA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ORDEN"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PEDIDO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "COSTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ARTICULO"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            Set objRecord = objDB.OpenRecordset("ENTRADAS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
            
                If bReportAll = False Then
                    If (objRecord.Fields(4).Value >= CDate(DateStart) And objRecord.Fields(4).Value <= CDate(DateEnd)) Then
                       'ok
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
                                                
                        Print #1, "</tr>"
                       
                    Else
                        'not
                    End If
                Else
                
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
                    
                    Print #1, "</tr>"
                
                End If
            
            objRecord.MoveNext
            Loop
            
            Print #1, "</table></center>"
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    Call AutoOpenXLS(CommonDialog1.FileName)
        
    Else
        
    End If


End Sub

Private Sub mSubDocOuts_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Generar Reporte Salidas de Almacén"
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
            Print #1, "<h3>Reporte de Salidas de Almac&eacute;n</h3>"
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
            Print #1, "CODIGO DE BARRAS"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CANTIDAD"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "CONCEPTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FECHA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "FACTURA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REMISION"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "REFERENCIA"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ORDEN"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "PEDIDO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "COSTO"
            Print #1, "</td>"
            Print #1, "<td>"
            Print #1, "ARTICULO"
            Print #1, "</td>"
            Print #1, "</tr>"
            
            
            Set objRecord = objDB.OpenRecordset("SALIDAS", dbOpenTable, dbReadOnly)
            
            Do While objRecord.EOF = False
                
                If bReportAll = False Then
                    If (objRecord.Fields(4).Value >= CDate(DateStart) And objRecord.Fields(4).Value <= CDate(DateEnd)) Then
                       'ok
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
                        
                        Print #1, "</tr>"
        
                    Else
                        'not
                    End If
                Else
                
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
                    
                    Print #1, "</tr>"

                End If
                                
                
            
            objRecord.MoveNext
            Loop
            
            Print #1, "</table></center>"
            Print #1, "</body>"
            Print #1, "</html>"
            
            
        Close #1
    
    objRecord.Close
    objDB.Close
    
    
    Call AutoOpenXLS(CommonDialog1.FileName)
    
    
    Else
        
    End If

End Sub

Private Sub mSubFamily_Click()

   

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCFamilys
    frmCFamilys.Show
    
End Sub

Private Sub mSubFloor_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCFloor
    frmCFloor.Show
    
End Sub

Private Sub mSubOperations_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmCOperations
    frmCOperations.Show
    
End Sub

Private Sub mSubProv_Click()

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    Load frmCProvider
    frmCProvider.Show
    
End Sub

Private Sub mSubStore_Click()
    
If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCStocks
    frmCStocks.Show
    
End Sub


Private Sub mSubTypProds_Click()
  

If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCTypeArts
    frmCTypeArts.Show


End Sub


Private Sub mSubUnits_Click()


If (UserLogon = False) Then
    mg = MsgBox(ERROR_SESSION_MESSAGE, vbCritical, PRODUCT_NAME)
    Exit Sub
End If
    
    
    Load frmCUnits
    frmCUnits.Show

End Sub


Private Sub mUsers_Click()
 
    If (UseDataBase = False) Then
        mgs = MsgBox("Debe abrir una Empresa para poder acceder al Control de Usuarios", vbCritical, PRODUCT_NAME)
        Exit Sub
    End If
    
bROOT = False

Load frmRoot
frmRoot.Show 1

If bROOT Then
    Load frmCUsers
    frmCUsers.Show 1
End If


End Sub


