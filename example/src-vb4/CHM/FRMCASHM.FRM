VERSION 4.00
Begin VB.Form frmCashMod 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   1365
   ClientTop       =   1695
   ClientWidth     =   10500
   Height          =   9780
   Left            =   1305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Top             =   1305
   Width           =   10620
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmCashMod"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    MsgBox KeyCode
    
    Select Case KeyCode
    
        Case 114
            frmCashMod.Hide
            Unload frmCashMod
        Case Default
    
    End Select

End Sub

