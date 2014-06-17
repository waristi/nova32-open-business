VERSION 4.00
Begin VB.Form frmRoot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de Usuarios"
   ClientHeight    =   2535
   ClientLeft      =   5820
   ClientTop       =   4575
   ClientWidth     =   4710
   Height          =   3045
   Icon            =   "FRMROOT.frx":0000
   Left            =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Top             =   4125
   Width           =   4830
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Image Image1 
         Height          =   330
         Left            =   120
         Picture         =   "FRMROOT.frx":0442
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use su password de Administrador para acceder."
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario:        ADMINISTRADOR"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmRoot"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    frmRoot.Hide
    Unload frmRoot
End Sub


Private Sub cmdOK_Click()

 Set objDB = OpenDatabase(NetCashFile)
 Set objRecord = objDB.OpenRecordset("USUARIOS", dbOpenTable, dbReadOnly)

    Dim Found As Boolean
    Found = False
    
    objRecord.MoveFirst
    
    If (objRecord.Fields(3).Value = Text1.Text) Then
                Found = True
    End If
 
                
    If Found Then
        bROOT = True
        objRecord.Close
        objDB.Close
        frmRoot.Hide
        Unload frmRoot
    Else
        bROOT = False
        Label4.Caption = "Clave no valida."
    End If

End Sub


