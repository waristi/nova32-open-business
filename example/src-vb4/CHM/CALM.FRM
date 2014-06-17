VERSION 4.00
Begin VB.Form frmCStorageProductsDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos de Almacén"
   ClientHeight    =   6540
   ClientLeft      =   1320
   ClientTop       =   1830
   ClientWidth     =   7200
   Height          =   7050
   Icon            =   "CALM.frx":0000
   Left            =   1260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Top             =   1380
   Width           =   7320
   Begin VB.Data CONSALM 
      Caption         =   "CONSALM"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "ALMACEN"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Data OPASILLO 
      Caption         =   "PASILLO"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CPASILLO"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data OTIPOART 
      Caption         =   "TIPOART"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CTIPOART"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data OFAMILIA 
      Caption         =   "FAMILIA"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CFAMILIAS"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data OPROVEED 
      Caption         =   "PROVE"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PROVEEDORES"
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data CASILLERO 
      Caption         =   "CASILLERO"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CCASILLERO"
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data CALMA 
      Caption         =   "CALMA"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CALMACEN"
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data OUNIDAD 
      Caption         =   "UNIDAD"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CUNIDADES"
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data OALMACEN 
      Caption         =   "ALMACEN"
      Connect         =   "Access"
      DatabaseName    =   "C:\NOVA32.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ALMACEN"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   6120
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   6975
      _version        =   65536
      _extentx        =   12303
      _extenty        =   10398
      _stockprops     =   15
      caption         =   "Editar Artículo"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   3
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Editar Artículo"
      tab(0).controlcount=   59
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label21"
      tab(0).control(1)=   "Label20"
      tab(0).control(2)=   "Label19"
      tab(0).control(3)=   "Label18"
      tab(0).control(4)=   "Label17"
      tab(0).control(5)=   "Label16"
      tab(0).control(6)=   "Label15"
      tab(0).control(7)=   "Label14"
      tab(0).control(8)=   "Label13"
      tab(0).control(9)=   "Label12"
      tab(0).control(10)=   "Label11"
      tab(0).control(11)=   "Label10"
      tab(0).control(12)=   "Label9"
      tab(0).control(13)=   "Label8"
      tab(0).control(14)=   "Label7"
      tab(0).control(15)=   "Label6"
      tab(0).control(16)=   "Label5"
      tab(0).control(17)=   "Label4"
      tab(0).control(18)=   "Label3"
      tab(0).control(19)=   "Label1"
      tab(0).control(20)=   "Label2"
      tab(0).control(21)=   "Label22"
      tab(0).control(22)=   "Label23"
      tab(0).control(23)=   "DBCombo1"
      tab(0).control(24)=   "DBCombo2"
      tab(0).control(25)=   "DBCombo3"
      tab(0).control(26)=   "DBCombo4"
      tab(0).control(27)=   "DBCombo5"
      tab(0).control(28)=   "DBCombo6"
      tab(0).control(29)=   "DBCombo7"
      tab(0).control(30)=   "Label47"
      tab(0).control(31)=   "Label49"
      tab(0).control(32)=   "Text16"
      tab(0).control(33)=   "Text15"
      tab(0).control(34)=   "Text14"
      tab(0).control(35)=   "Text13"
      tab(0).control(36)=   "Text12"
      tab(0).control(37)=   "Text11"
      tab(0).control(38)=   "Text10"
      tab(0).control(39)=   "Text9"
      tab(0).control(40)=   "Text8"
      tab(0).control(41)=   "Text7"
      tab(0).control(42)=   "Text6"
      tab(0).control(43)=   "Text5"
      tab(0).control(44)=   "Text4"
      tab(0).control(45)=   "Text3"
      tab(0).control(46)=   "Text1"
      tab(0).control(47)=   "Text2"
      tab(0).control(48)=   "cmdLast"
      tab(0).control(49)=   "cmdNext"
      tab(0).control(50)=   "cmdPrevious"
      tab(0).control(51)=   "cmdFirst"
      tab(0).control(52)=   "cmdeleted"
      tab(0).control(53)=   "cmdSave"
      tab(0).control(54)=   "cmdNew"
      tab(0).control(55)=   "Text40"
      tab(0).control(56)=   "Text42"
      tab(0).control(57)=   "txt_Finder"
      tab(0).control(58)=   "Command5"
      tabcaption(1)   =   "Consultar Artículo"
      tab(1).controlcount=   55
      tab(1).controlenabled=   0   'False
      tab(1).control(0)=   "Text43"
      tab(1).control(1)=   "Text41"
      tab(1).control(2)=   "cmdSeek"
      tab(1).control(3)=   "Text39"
      tab(1).control(4)=   "Text38"
      tab(1).control(5)=   "Text37"
      tab(1).control(6)=   "Text36"
      tab(1).control(7)=   "Text35"
      tab(1).control(8)=   "Text34"
      tab(1).control(9)=   "Text33"
      tab(1).control(10)=   "Text32"
      tab(1).control(11)=   "Text31"
      tab(1).control(12)=   "Text30"
      tab(1).control(13)=   "Text29"
      tab(1).control(14)=   "Text28"
      tab(1).control(15)=   "Command4"
      tab(1).control(16)=   "Command3"
      tab(1).control(17)=   "Command2"
      tab(1).control(18)=   "Command1"
      tab(1).control(19)=   "Text27"
      tab(1).control(20)=   "Text26"
      tab(1).control(21)=   "Text25"
      tab(1).control(22)=   "Text24"
      tab(1).control(23)=   "Text23"
      tab(1).control(24)=   "Text22"
      tab(1).control(25)=   "Text21"
      tab(1).control(26)=   "Text20"
      tab(1).control(27)=   "Text19"
      tab(1).control(28)=   "Text18"
      tab(1).control(29)=   "Text17"
      tab(1).control(30)=   "Label50"
      tab(1).control(31)=   "Label48"
      tab(1).control(32)=   "Label46"
      tab(1).control(33)=   "Label45"
      tab(1).control(34)=   "Label44"
      tab(1).control(35)=   "Label43"
      tab(1).control(36)=   "Label42"
      tab(1).control(37)=   "Label41"
      tab(1).control(38)=   "Label40"
      tab(1).control(39)=   "Label39"
      tab(1).control(40)=   "Label38"
      tab(1).control(41)=   "Label37"
      tab(1).control(42)=   "Label36"
      tab(1).control(43)=   "Label35"
      tab(1).control(44)=   "Label34"
      tab(1).control(45)=   "Label33"
      tab(1).control(46)=   "Label32"
      tab(1).control(47)=   "Label31"
      tab(1).control(48)=   "Label30"
      tab(1).control(49)=   "Label29"
      tab(1).control(50)=   "Label28"
      tab(1).control(51)=   "Label27"
      tab(1).control(52)=   "Label26"
      tab(1).control(53)=   "Label25"
      tab(1).control(54)=   "Label24"
      tabcaption(2)   =   "Mostrar/Editar Todos"
      tab(2).controlcount=   1
      tab(2).controlenabled=   0   'False
      tab(2).control(0)=   "DBGrid1"
      Begin VB.CommandButton Command5 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   2400
         TabIndex        =   116
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt_Finder 
         Height          =   285
         Left            =   240
         TabIndex        =   115
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text43 
         DataField       =   "DESCUENTO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   113
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox Text42 
         BackColor       =   &H0000FFFF&
         DataField       =   "DESCUENTO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   25
         Text            =   "0"
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text41 
         DataField       =   "IVA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   110
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H0000FFFF&
         DataField       =   "IVA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Text            =   "0"
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton cmdSeek 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   -71400
         TabIndex        =   108
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text39 
         DataField       =   "PASILLO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   107
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox Text38 
         DataField       =   "TIPO_ARTICULO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   106
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text37 
         DataField       =   "FAMILIA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   105
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox Text36 
         DataField       =   "PROVEEDOR"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   104
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text35 
         DataField       =   "CASILLERO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   103
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox Text34 
         DataField       =   "ALMACEN"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   102
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Text33 
         DataField       =   "UNIDAD"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   101
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text32 
         DataField       =   "MINIMO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   89
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text31 
         DataField       =   "MAXIMO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   88
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox Text30 
         DataField       =   "PRECIO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   87
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox Text29 
         DataField       =   "COSTO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   86
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text28 
         DataField       =   "FECHA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   85
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">|"
         Height          =   375
         Left            =   -68520
         TabIndex        =   84
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">>"
         Height          =   375
         Left            =   -68880
         TabIndex        =   83
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   375
         Left            =   -69240
         TabIndex        =   82
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "|<"
         Height          =   375
         Left            =   -69600
         TabIndex        =   81
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox Text27 
         DataField       =   "PEDIDO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   68
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         DataField       =   "SERIE"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   67
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         DataField       =   "MODELO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   66
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text24 
         DataField       =   "ORDEN"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   65
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         DataField       =   "REFERENCIA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   64
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         DataField       =   "MARCA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   63
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         DataField       =   "ARTICULO"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   62
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         DataField       =   "REMISION"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   61
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text19 
         DataField       =   "FACTURA"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -70200
         TabIndex        =   60
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         DataField       =   "CODIGO_BARRAS"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   59
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H0000FFFF&
         DataField       =   "CLAVE"
         DataSource      =   "CONSALM"
         Height          =   285
         Left            =   -73440
         TabIndex        =   58
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   4560
         TabIndex        =   57
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5280
         TabIndex        =   24
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton cmdeleted 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6000
         TabIndex        =   56
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   5280
         TabIndex        =   52
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<<"
         Height          =   375
         Left            =   5640
         TabIndex        =   51
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   375
         Left            =   6000
         TabIndex        =   50
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   6360
         TabIndex        =   49
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         DataField       =   "CODIGO_BARRAS"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0000FFFF&
         DataField       =   "CLAVE"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "FACTURA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         DataField       =   "REMISION"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H0000FFFF&
         DataField       =   "ARTICULO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H0000FF00&
         DataField       =   "MARCA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         DataField       =   "REFERENCIA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         DataField       =   "ORDEN"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         DataField       =   "MODELO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         DataField       =   "SERIE"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         DataField       =   "PEDIDO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H0000FFFF&
         DataField       =   "FECHA"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H0000FFFF&
         DataField       =   "COSTO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "0"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H0000FFFF&
         DataField       =   "PRECIO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Text            =   "0"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H0000FF00&
         DataField       =   "MAXIMO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "0"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H0000FF00&
         DataField       =   "MINIMO"
         DataSource      =   "OALMACEN"
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Text            =   "0"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label50 
         Caption         =   "% DESC"
         Height          =   255
         Left            =   -71280
         TabIndex        =   114
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label49 
         Caption         =   "% DESC."
         Height          =   255
         Left            =   3720
         TabIndex        =   112
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label48 
         Caption         =   "% IVA"
         Height          =   255
         Left            =   -71280
         TabIndex        =   111
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label47 
         Caption         =   "% IVA"
         Height          =   255
         Left            =   3720
         TabIndex        =   109
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "Casillero"
         Height          =   255
         Left            =   -74760
         TabIndex        =   100
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "Pasillo"
         Height          =   255
         Left            =   -71280
         TabIndex        =   99
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label44 
         Caption         =   "Almacén"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "Mínimo"
         Height          =   255
         Left            =   -71280
         TabIndex        =   97
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "Máximo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   96
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "Precio"
         Height          =   255
         Left            =   -71280
         TabIndex        =   95
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Costo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   94
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label39 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   -71280
         TabIndex        =   93
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Tipo Unidad"
         Height          =   255
         Left            =   -74760
         TabIndex        =   92
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Grupos"
         Height          =   255
         Left            =   -74760
         TabIndex        =   90
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "Pedido"
         Height          =   255
         Left            =   -71280
         TabIndex        =   80
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Tipo Art."
         Height          =   255
         Left            =   -71280
         TabIndex        =   79
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "Serie"
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   77
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "Orden"
         Height          =   255
         Left            =   -71280
         TabIndex        =   76
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   -71280
         TabIndex        =   75
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Marca"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Remisión"
         Height          =   255
         Left            =   -71280
         TabIndex        =   72
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Factura"
         Height          =   255
         Left            =   -71280
         TabIndex        =   71
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Clave"
         Height          =   255
         Left            =   -74760
         TabIndex        =   70
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Código de Barras"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   840
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CALM.frx":0442
         Height          =   4695
         Left            =   -74880
         OleObjectBlob   =   "CALM.frx":0453
         TabIndex        =   55
         Top             =   480
         Width           =   6735
      End
      Begin MSDBCtls.DBCombo DBCombo7 
         Bindings        =   "CALM.frx":0B5B
         DataField       =   "PASILLO"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   3720
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
      Begin MSDBCtls.DBCombo DBCombo6 
         Bindings        =   "CALM.frx":0B6C
         DataField       =   "TIPO_ARTICULO"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   4800
         TabIndex        =   17
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
      Begin MSDBCtls.DBCombo DBCombo5 
         Bindings        =   "CALM.frx":0B7D
         DataField       =   "FAMILIA"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   5280
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
      Begin MSDBCtls.DBCombo DBCombo4 
         Bindings        =   "CALM.frx":0B8E
         DataField       =   "PROVEEDOR"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   4920
         Width           =   1935
         _version        =   65536
         _extentx        =   3413
         _extenty        =   556
         _stockprops     =   77
         forecolor       =   0
         backcolor       =   16777215
         listfield       =   "NOMBRE"
         boundcolumn     =   "NOMBRE"
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "CALM.frx":0B9F
         DataField       =   "CASILLERO"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   4560
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
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "CALM.frx":0BB1
         DataField       =   "ALMACEN"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   4200
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
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "CALM.frx":0BBF
         DataField       =   "UNIDAD"
         DataSource      =   "OALMACEN"
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   3120
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
      Begin VB.Label Label23 
         Caption         =   "Grupos"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código de Barras"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Factura"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Remisión"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Marca"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Orden"
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Serie"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Art."
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Pedido"
         Height          =   255
         Left            =   3720
         TabIndex        =   37
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Unidad"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Costo"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Precio"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Máximo"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Mínimo"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Almacén"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Pasillo"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Casillero"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4560
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCStorageProductsDef"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Cursor As Integer
Dim Size As Integer

Private Sub cmdClosed_Click()
    frmCStorageProductsDef.Hide
    Unload frmCStorageProductsDef
End Sub

Private Sub cmdeleted_Click()

On Error GoTo LocalHandlerDel

    OALMACEN.Recordset.Delete
    OALMACEN.Refresh
    DBGrid1.Refresh
    mg = MsgBox("El Artículo fue Eliminado.", vbOKOnly, "Nova32 NetCash")
    Text1.Text = ""
    Text2.Text = ""
    Exit Sub

LocalHandlerDel:
    msg = MsgBox("No se puede eliminar el Artículo.", vbCritical, PRODUCT_NAME)
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

Private Sub cmdNew_Click()

On Error GoTo LocalHandlerNew

    OALMACEN.Refresh
    OALMACEN.Recordset.AddNew
    mg = MsgBox("Digite la clave y descripción del nuevo Artículo y despues Guarde los datos.", vbOKOnly, PRODUCT_NAME)
    
LocalHandlerNew:

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

Private Sub cmdSave_Click()

On Error GoTo LocalHandlerSave

    OALMACEN.Recordset.Update
    OALMACEN.Refresh
    DBGrid1.Refresh
    Exit Sub
    
    
LocalHandlerSave:
    ms = MsgBox("No se puede guardar el Artículo, ya existe un registro con esa clave, o un campo es incorrecto.", vbCritical, PRODUCT_NAME)
End Sub


Private Sub cmdSeek_Click()

    CONSALM.Recordset.FindFirst ("CLAVE='" + Text17.Text + "'")
    
End Sub

Private Sub Command1_Click()

On Error GoTo LocalHandlerMoveFirst
    
    CONSALM.Recordset.MoveFirst
    
LocalHandlerMoveFirst:

End Sub

Private Sub Command2_Click()

On Error GoTo LocalHandlerMovePrev

    CONSALM.Recordset.MovePrevious
    Exit Sub
    
LocalHandlerMovePrev:

End Sub

Private Sub Command3_Click()
    
On Error GoTo LocalHandlerMoveNext
    
    CONSALM.Recordset.MoveNext
    Exit Sub
    
LocalHandlerMoveNext:
End Sub

Private Sub Command4_Click()
On Error GoTo LocalHandlerMoveLast

    CONSALM.Recordset.MoveLast
    
LocalHandlerMoveLast:

End Sub

Private Sub Command5_Click()

Dim Found As Boolean
Call cmdFirst_Click

Found = False
    
    Do While (OALMACEN.Recordset.EOF <> True)
    
        If (Trim(Text1.Text) = Trim(txt_Finder.Text)) Then
            Found = True
            Exit Do
        End If
            
        cmdNext_Click
    Loop
    
    
If (Found = False) Then
    x = MsgBox("No se encontro un Formulario con esa clave.", 64, PRODUCT_NAME)
    Call cmdFirst_Click
End If

End Sub

Private Sub Form_Load()
    
    OALMACEN.DatabaseName = NetCashFile
    CALMA.DatabaseName = NetCashFile
    OUNIDAD.DatabaseName = NetCashFile
    CASILLERO.DatabaseName = NetCashFile
    OPROVEED.DatabaseName = NetCashFile
    OFAMILIA.DatabaseName = NetCashFile
    OTIPOART.DatabaseName = NetCashFile
    OPASILLO.DatabaseName = NetCashFile
    CONSALM.DatabaseName = NetCashFile
        
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    CONSALM.Refresh
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


Private Sub Text17_LostFocus()
   
    Dim sTMP As String
    Dim Pos As Integer
    sTMP = Text17.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    Text17.Text = sTMP

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


