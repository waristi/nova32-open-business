VERSION 4.00
Begin VB.Form frmSetups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuraciones del Reporte"
   ClientHeight    =   3105
   ClientLeft      =   1485
   ClientTop       =   2055
   ClientWidth     =   4785
   Height          =   3615
   Icon            =   "FRMSETUP.frx":0000
   Left            =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Top             =   1605
   Width           =   4905
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _version        =   65536
      _extentx        =   8070
      _extenty        =   5106
      _stockprops     =   15
      caption         =   "Reporteador"
      tabsperrow      =   3
      tab             =   0
      taborientation  =   0
      tabs            =   1
      style           =   0
      tabmaxwidth     =   0
      tabheight       =   529
      tabcaption(0)   =   "Reporteador"
      tab(0).controlcount=   7
      tab(0).controlenabled=   -1  'True
      tab(0).control(0)=   "Label1"
      tab(0).control(1)=   "Label2"
      tab(0).control(2)=   "Image1"
      tab(0).control(3)=   "Text1"
      tab(0).control(4)=   "Text2"
      tab(0).control(5)=   "Check1"
      tab(0).control(6)=   "cmdSave"
      Begin VB.CommandButton cmdSave 
         Caption         =   "Aplicar"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Generar Reporte de todo el Ejercicio Contable."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   840
         Picture         =   "FRMSETUP.frx":0442
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSetups"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdClosed_Click()
    frmSetups.Hide
    Unload frmSetups
End Sub


Private Sub cmdSave_Click()

    DateStart = Text1.Text
    DateEnd = Text2.Text
    
    If Check1.Value Then
        bReportAll = True
    Else
        bReportAll = False
    End If
    
    MsgBox "Se Aplicaron los cambios en el Reporteador."
    
End Sub


Private Sub Form_Load()
    Text1.Text = DateStart
    Text2.Text = DateEnd
    
    If bReportAll Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

