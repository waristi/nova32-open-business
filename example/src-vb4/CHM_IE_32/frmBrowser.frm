VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SAT"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14955
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   14955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   855
      Left            =   2520
      Picture         =   "frmBrowser.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Avanzar"
      Height          =   855
      Left            =   1680
      Picture         =   "frmBrowser.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Regresar"
      Height          =   855
      Left            =   840
      Picture         =   "frmBrowser.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inicio"
      Height          =   855
      Left            =   0
      Picture         =   "frmBrowser.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   14775
      ExtentX         =   26061
      ExtentY         =   16325
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CMD As String

Private Sub Command1_Click()
    WebBrowser1.Navigate CMD
End Sub

Private Sub Command2_Click()
    
On Error GoTo localHandlerA
    
    WebBrowser1.GoBack
    
localHandlerA:

End Sub

Private Sub Command3_Click()

On Error GoTo localHandlerB
    WebBrowser1.GoForward
    
localHandlerB:
    
End Sub

Private Sub Command4_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    
    CMD = Command
    WebBrowser1.Navigate CMD
    
End Sub
