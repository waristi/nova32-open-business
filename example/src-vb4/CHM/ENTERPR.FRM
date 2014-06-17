VERSION 4.00
Begin VB.Form frmEnterprise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalles de la Entidad"
   ClientHeight    =   5445
   ClientLeft      =   1545
   ClientTop       =   2280
   ClientWidth     =   6870
   Height          =   5955
   Icon            =   "ENTERPR.frx":0000
   Left            =   1485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Top             =   1830
   Width           =   6990
   Begin VB.Data EMPRESA 
      Caption         =   "EMPRESA"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\SYSTEM.LEX"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "EMPRESA"
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5520
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtDetails 
      DataField       =   "DETALLES"
      DataSource      =   "EMPRESA"
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   23
      Top             =   3960
      Width           =   6615
   End
   Begin VB.TextBox txtemail 
      DataField       =   "EMAIL"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   22
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "TELEFONO"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtSr 
      DataField       =   "REPRESENTANTE"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtCity 
      DataField       =   "CIUDAD"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtServ 
      DataField       =   "GIRO"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtCountry 
      DataField       =   "PAIS"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtEst 
      DataField       =   "ESTADO"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDir 
      DataField       =   "DIRECCION"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   15
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtRFC 
      DataField       =   "RFC"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtEnterprise 
      DataField       =   "EMPRESA"
      DataSource      =   "EMPRESA"
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Top             =   840
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   11
      Top             =   0
      Width           =   6855
      Begin VB.Image Image1 
         Height          =   330
         Left            =   120
         Picture         =   "ENTERPR.frx":0442
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Registro de datos del negocio."
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Detalles"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Representante legal"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Giro"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "País"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Estado"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Ciudad"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "R.F.C."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmEnterprise"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    frmEnterprise.Hide
    Unload frmEnterprise
End Sub


Private Sub cmdOK_Click()
    
If (UserProfile > 0) Then
    x = MsgBox("Acceso denegado,Su cuenta de usuario no dispone de los privilegios suficientes para completar esta operación.", 64, PRODUCT_NAME)
    Exit Sub
End If
    
    
    EMPRESA.UpdateRecord
    EMPRESA.Refresh
    x = MsgBox("Datos actualizados...", mb_iconinformation, PRODUCT_NAME)
End Sub


Private Sub Form_Load()
    EMPRESA.DatabaseName = NetCashFile
End Sub


