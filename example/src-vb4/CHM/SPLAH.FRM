VERSION 4.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6570
   ClientLeft      =   1650
   ClientTop       =   2130
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Height          =   7080
   Icon            =   "SPLAH.frx":0000
   Left            =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Top             =   1680
   Width           =   8295
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      Picture         =   "SPLAH.frx":0442
      ScaleHeight     =   6615
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ProtectedProduct()

On Error GoTo LOCAL_FAILURE


    FileLen ("C:\WINDOWS\SYSTEM32\AXRUN16.386")


    Open "C:\WINDOWS\SYSTEM32\AXRUN16.SYS" For Input As #1
        Input #1, sACTIVE
    Close #1


    If sACTIVE = "CHM=16" Then
        Rem product activated
        Exit Sub
    End If



    Open "C:\WINDOWS\SYSTEM32\AXRUN16.386" For Input As #1
        Input #1, sEXEC
    Close #1
    
    dEXEC = CInt(sEXEC)

    If dEXEC > 0 Then
        dEXEC = dEXEC - 1
    End If

    sEXEC = CStr(dEXEC)

    Open "C:\WINDOWS\SYSTEM32\AXRUN16.386" For Output As #1
        Print #1, sEXEC
    Close #1


    If dEXEC <= 0 Then
        M = MsgBox("ESTE PRODUCTO NO HA SIDO ACTIVADO, PERIODO DE PRUEBA FINALIZADO, FAVOR DE PONERSE EN CONTACTO CON EL PERSONAL DE NOVA32 PARA ACTIVAR NETCASH", vbCritical + vbOKOnly, PRODUCT_NAME)
        End
    Else
        M = MsgBox("ESTE PRODUCTO NO HA SIDO ACTIVADO, LE RESTAN " + sEXEC + " EJECUCIONES DE 32 OTORGADAS, FAVOR DE PONERSE EN CONTACTO CON EL PERSONAL DE NOVA32 PARA ACTIVAR NOVA32 NETCASH", vbInformation + vbOKOnly, PRODUCT_NAME)
        Exit Sub
    End If


    
    
LOCAL_FAILURE:
    M = MsgBox("PARA PODER USAR ESTA INSTALACION DE NETCASH, DEBE ACTIVARLA EL PERSONAL DE SOPORTE ESPECIALIZADO DE NOVA32, PONGASE EN CONTACTO CON UN DISTRIBUIDOR AUTORIZADO.", vbCritical + vbOKOnly, PRODUCT_NAME)
    End


End Sub


Private Sub Form_Load()
    
Dim sEXEC As String
Dim dEXEC As Integer
Dim sACTIVE As String
Dim PRODUCT_NAME As String


Me.Top = (Screen.Height / 2) - (Me.Height / 2)
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
PRODUCT_NAME = "Nova32 NetCash Small Business 2.0"
    
    
DateStart = Date
DateEnd = Date
    
        
End Sub


Private Sub Timer1_Timer()

    Call ProtectedProduct

    frmSplash.Hide
    Unload frmSplash
 
    Load mdiMain
    mdiMain.Show
 
End Sub


