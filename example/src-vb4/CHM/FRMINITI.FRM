VERSION 4.00
Begin VB.Form frmInitialCash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldo Inicial de Caja"
   ClientHeight    =   2130
   ClientLeft      =   1140
   ClientTop       =   1740
   ClientWidth     =   4020
   Height          =   2535
   Icon            =   "FRMINITI.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Top             =   1395
   Width           =   4140
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\SYSTEM.LEX"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "INICIAL"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "INICIAL"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "0.0"
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "FRMINITI.frx":0442
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo inicial en Caja para operación del día:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmInitialCash"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Data1.UpdateRecord
    Data1.Refresh
    x = MsgBox("Saldo inicial actualizado...", 64, PRODUCT_NAME)
End Sub

Private Sub Command2_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


