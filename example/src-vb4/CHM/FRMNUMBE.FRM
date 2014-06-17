VERSION 4.00
Begin VB.Form frmNumber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Folios"
   ClientHeight    =   2700
   ClientLeft      =   1140
   ClientTop       =   1740
   ClientWidth     =   5250
   Height          =   3105
   Icon            =   "FRMNUMBE.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Top             =   1395
   Width           =   5370
   Begin VB.TextBox Text2 
      DataField       =   "ID_CAJA"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Text            =   "0"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txt_Serie 
      DataField       =   "ID_SERIE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Text            =   "A"
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      Picture         =   "FRMNUMBE.frx":0442
      ScaleHeight     =   1515
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CHM\SYSTEM.LEX"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "FOLIOS"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_FOLIO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Caja:"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Serie:"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Ingrese la palabra OFF en el campo serie, para desactivarlo."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Folio Activo:"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmNumber"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
    Data1.UpdateRecord
    Data1.Refresh
    x = MsgBox("Nuevo Folio actualizado.", mb_iconinformation, PRODUCT_NAME)

  
End Sub

Private Sub Command2_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub Form_Load()
    Data1.DatabaseName = NetCashFile
End Sub


