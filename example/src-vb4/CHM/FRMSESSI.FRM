VERSION 4.00
Begin VB.Form frmSession 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Iniciar Sesion"
   ClientHeight    =   2790
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4350
   Height          =   3195
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   4470
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   4305
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Iniciar sesion e identificarse en Sistema."
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   330
         Left            =   120
         Picture         =   "FRMSESSI.frx":0000
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim objDB As Database
Dim objRecord As Recordset
Private Sub cmdCANCEL_Click()
    
    objRecord.Close
    objDB.Close
    
    frmSession.Hide
    Unload frmSession
    
End Sub

Private Sub cmdOK_Click()
    
    Dim found As Boolean
    found = False
    
    objRecord.MoveFirst
    Do While (objRecord.EOF = False)
        If (objRecord.Fields(2).Value = txtUser.Text) Then
            If (objRecord.Fields(3).Value = txtPassword.Text) Then
                found = True
                Exit Do
            End If
        End If
    objRecord.MoveNext
    Loop
                
    If found Then
        msg = MsgBox("Bienvenido " + objRecord.Fields(2).Value, vbOKOnly, "Acceso:")
        frmSession.Hide
        Unload frmSession
        UserLogon = True
        UserProfile = objRecord.Fields(0).Value
    Else
        txtPassword.Text = ""
        msg = MsgBox("Usuario o Password incorrectos.", vbCritical, "Acceso:")
        UserLogon = False
    End If
End Sub

Private Sub Form_Load()

    Set objDB = OpenDatabase(NetCashFile)
    Set objRecord = objDB.OpenRecordset("USUARIOS", dbOpenTable, dbReadOnly)
    
End Sub


