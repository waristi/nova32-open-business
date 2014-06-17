VERSION 4.00
Begin VB.Form frmDispositive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dispositivos"
   ClientHeight    =   4215
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5310
   Height          =   4725
   Icon            =   "FRMDISPO.frx":0000
   Left            =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMDISPO.frx":0442
   ScaleHeight     =   4215
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Top             =   1290
   Width           =   5430
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puertos"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   5055
      Begin VB.OptionButton Option1 
         Caption         =   "USB 2.0"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LPT1"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "COM2"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "COM1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "FRMDISPO.frx":0884
      Left            =   1440
      List            =   "FRMDISPO.frx":0891
      TabIndex        =   11
      Text            =   "Xon/Xoff"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "FRMDISPO.frx":08B2
      Left            =   1440
      List            =   "FRMDISPO.frx":08BF
      TabIndex        =   10
      Text            =   "1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FRMDISPO.frx":08CE
      Left            =   1440
      List            =   "FRMDISPO.frx":08E1
      TabIndex        =   9
      Text            =   "Ninguno"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FRMDISPO.frx":090A
      Left            =   1440
      List            =   "FRMDISPO.frx":091D
      TabIndex        =   8
      Text            =   "4"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FRMDISPO.frx":0930
      Left            =   1440
      List            =   "FRMDISPO.frx":094C
      TabIndex        =   7
      Text            =   "9600"
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FRMDISPO.frx":097E
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"FRMDISPO.frx":0DC0
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Control de Flujo"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Bit de Parada"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Bits de Datos"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Baudios"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmDispositive"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    frmDispositive.Hide
    Unload frmDispositive
End Sub

