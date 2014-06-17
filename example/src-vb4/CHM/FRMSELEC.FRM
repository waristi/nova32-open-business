VERSION 4.00
Begin VB.Form frmSelectArt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTAR ARTICULO"
   ClientHeight    =   4770
   ClientLeft      =   1260
   ClientTop       =   1695
   ClientWidth     =   6720
   Height          =   5220
   Icon            =   "FRMSELEC.frx":0000
   Left            =   1200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Top             =   1305
   Width           =   6840
   Begin VB.TextBox txtCLAVEVIEW 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      TabIndex        =   11
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "[F3]  CERRAR"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtEXISTENCIAS 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox txtMARCA 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      TabIndex        =   9
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtPRECIO 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      TabIndex        =   8
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtARTICULO 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label6 
      Caption         =   "CLAVE:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "EXISTENCIAS:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "MARCA:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "PRECIO:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Caption         =   "ARTICULO:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "CLAVE:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelectArt"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClosed_Click()
    frmSelectArt.Hide
    Unload frmSelectArt
End Sub

Private Sub Text1_Click()
    Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case 114 '[F3] CERRAR
        Call cmdClosed_Click
    Case Default

End Select

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    Dim sTMP As String
    Dim Pos As Integer
    sTMP = Text1.Text
    
    Do While (InStr(sTMP, "'") > 0)
       Pos = InStr(sTMP, "'")
       Mid(sTMP, Pos, 1) = "-"
    Loop

    Dim MDB As Database
    Dim REC As Recordset
    
    Dim Found As Boolean
    Found = False

    Set MDB = OpenDatabase(NetCashFile)
    Set REC = MDB.OpenRecordset("ALMACEN", dbOpenTable, dbReadOnly)

    REC.MoveFirst
    Do While (REC.EOF = False)
        If (REC.Fields(0).Value = sTMP) Then
            Found = True
            Exit Do
        End If
    REC.MoveNext
    Loop

    If (Found) Then
        txtCLAVEVIEW.Text = REC.Fields(0).Value
        txtARTICULO.Text = REC.Fields(2).Value
        txtPRECIO.Text = CStr(REC.Fields(20).Value)
        txtMARCA.Text = REC.Fields(3).Value
        txtEXISTENCIAS.Text = CInt(REC.Fields(23).Value)
        Text1.Text = ""
    Else
        M = MsgBox("NO SE ENCONTRO EL ARTICULO.", vbCritical, PRODUCT_NAME)
        Text1.Text = ""
    End If


    REC.Close
    MDB.Close

End If

End Sub


