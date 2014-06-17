VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de..."
   ClientHeight    =   6210
   ClientLeft      =   1905
   ClientTop       =   2415
   ClientWidth     =   8580
   Height          =   6615
   Icon            =   "FRMABOUT.frx":0000
   Left            =   1845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Top             =   2070
   Width           =   8700
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      Picture         =   "FRMABOUT.frx":0442
      ScaleHeight     =   6255
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton Command2 
         Caption         =   "Información del Sistema"
         Height          =   495
         Left            =   6960
         TabIndex        =   9
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   6960
         TabIndex        =   8
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "For Windows NT 16 Bits."
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   7
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"FRMABOUT.frx":46BC8
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   6
         Top             =   5520
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"FRMABOUT.frx":46C94
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   5
         Top             =   4680
         Width           =   5295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Release 201205.1016"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nova32"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   20.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   4440
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Netcash"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   48
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4440
         TabIndex        =   2
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Small Business 2.0"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   21.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4440
         TabIndex        =   1
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   1635
         Left            =   120
         Picture         =   "FRMABOUT.frx":46D21
         Top             =   4680
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Command2_Click()
 
 On Error GoTo Local_Handler_DirectX
 
 
 Shell "DXDIAG", vbNormalFocus
 
 
Local_Handler_DirectX:
 
 
End Sub


Private Sub Form_Load()

    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

End Sub


