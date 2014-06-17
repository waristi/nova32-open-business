VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmWeight 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nova32 NetCash"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "frmWeight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtImporte 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   645
      Left            =   3480
      TabIndex        =   7
      Text            =   "0"
      Top             =   4320
      Width           =   4455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2220
      ItemData        =   "frmWeight.frx":030A
      Left            =   240
      List            =   "frmWeight.frx":030C
      TabIndex        =   6
      Top             =   1920
      Width           =   7695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtPeso 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   645
      Left            =   3480
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "AGREGAR A LA VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5880
      TabIndex        =   4
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPesar 
      Appearance      =   0  'Flat
      Caption         =   "PESAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      TabIndex        =   3
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "IMPORTE $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "PRODUCTO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "PESO BASCULA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------ROWS,COLS----------
Dim PRECIOS(1000, 2) As Variant

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal x As Long, _
ByVal y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long) As Long


Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False
    End If

End Function


Private Function ReadPORT_COM1() As String

Dim cBuffer As String

    With MSComm1
        
        If .PortOpen = True Then .PortOpen = False

            .CommPort = 1
            .Settings = "9600,N,8,1"
            .InputLen = 0
            .InputMode = comInputModeText
            .Handshaking = 0
            .PortOpen = True

            cBuffer = ""
            'TORREY keycode is send 'P'
            .Output = Chr$(80)

        Do
            DoEvents
            cBuffer = cBuffer & .Input
        Loop Until InStr(cBuffer, "kg")

        .PortOpen = False
    End With

    ReadPORT_COM1 = cBuffer
    
End Function

Private Sub cmdAceptar_Click()

    Open "C:\CHM\RS232.DAT" For Output As #1
            
            Print #1, List1.Text
            Print #1, txtPeso.Text
            Print #1, txtImporte.Text
                
    Close #1


    End

End Sub

Private Sub cmdPesar_Click()

    Dim strWeight As String
    strWeight = ReadPORT_COM1
    strWeight = Replace(strWeight, "kg", " ")
    txtPeso.Text = CDbl(strWeight)
        
        
    Dim dblVALUE As Double
    Dim strSEEK As String
    Dim dIDX As Integer
    
    strSEEK = List1.Text
            
    For dIDX = 1 To 100
        
        If (Trim(strSEEK) = Trim(PRECIOS(dIDX, 1))) Then
            dblVALUE = PRECIOS(dIDX, 2)
            Exit For
        Else
            dblVALUE = 0
        End If
            
    Next dIDX
        
    txtImporte.Text = (txtPeso.Text * dblVALUE) * 1000
    
End Sub

Private Sub Command1_Click()

    Open "C:\CHM\RS232.DAT" For Output As #1
            
            Print #1, "0"
            Print #1, "0"
            Print #1, "0"
                
    Close #1


    End

End Sub

Private Sub Form_Load()

    If Command <> "/run" Then
        End
    End If


    Rem -------TOPMOST--------------------
    SetTopMostWindow Me.hwnd, True
    Rem ----------------------------------


    Dim sPROD As String
    Dim sVAL As String
    Dim dIDX As Integer
    
    dIDX = 0

    Open "C:\CHM\CATLIST.INI" For Input As #1
    
        Do While Not EOF(1)
        
            dIDX = dIDX + 1
        
            Input #1, sPROD
            Input #1, sVAL
            
            PRECIOS(dIDX, 1) = sPROD
            PRECIOS(dIDX, 2) = sVAL
                    
        Loop
    
    Close #1


    For dIDX = 1 To 100
    
        List1.AddItem PRECIOS(dIDX, 1)

    Next dIDX


End Sub
