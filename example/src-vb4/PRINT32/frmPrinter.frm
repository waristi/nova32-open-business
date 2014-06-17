VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   0  'None
   Caption         =   "Impresor"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   Icon            =   "frmPrinter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Imprimiendo Ticket..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


Dim CMD As String

CMD = Command


If Trim(CMD) <> "/autofix" Then
    End
End If


On Error GoTo LocalHandler

    Printer.PrintQuality = -1
    Printer.ScaleMode = 7
    Printer.Font.Name = "Courier"
    Printer.Font.Size = 8
    
    Dim LINE As String
    
    Open "C:\TICKET.TMP" For Input As #1
    
        Do While Not EOF(1)
        
            Input #1, LINE
            Printer.Print LINE
        
        Loop
    
    Close #1
    
    
LocalHandler:
    
    End

End Sub
