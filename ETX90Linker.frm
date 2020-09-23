VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "ETX90 Linker"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAllStop 
      Caption         =   "ALL STOP"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speed"
      Height          =   1815
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   1335
      Begin VB.OptionButton Fast 
         Caption         =   "Fast"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Medium 
         Caption         =   "Medium"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Slow 
         Caption         =   "Slow"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdMvSouth 
      Caption         =   "South"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdMvWest 
      Caption         =   "West"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdMvEast 
      Caption         =   "East"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdMvNorth 
      Caption         =   "North"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu exitme 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAllStop_Click()
MSComm1.PortOpen = True
MSComm1.Output = "#:Qn#"
MSComm1.Output = "#:Qs#"
MSComm1.Output = "#:Qe#"
MSComm1.Output = "#:Qw#"
MSComm1.PortOpen = False
End Sub

Private Sub cmdMvEast_Click()
If MvEast = False Then
    cmdMvEast.Caption = "Moving East"
    MvEast = True
    MSComm1.PortOpen = True
    MSComm1.Output = SpeedSelection
    MSComm1.Output = "#:Me#"
    MSComm1.PortOpen = False
    Exit Sub
End If

If MvEast = True Then
    cmdMvEast.Caption = "East"
    MvEast = False
    MSComm1.PortOpen = True
    MSComm1.Output = "#:Qe#"
    MSComm1.PortOpen = False
End If
End Sub

Private Sub cmdMvNorth_Click()
If MvNorth = False Then
    cmdMvNorth.Caption = "Moving North"
    MvNorth = True
    MSComm1.PortOpen = True
    MSComm1.Output = SpeedSelection
    MSComm1.Output = "#:Mn#"
    MSComm1.PortOpen = False
    Exit Sub
End If

If MvNorth = True Then
    cmdMvNorth.Caption = "North"
    MvNorth = False
    MSComm1.PortOpen = True
    MSComm1.Output = "#:Qn#"
    MSComm1.PortOpen = False
End If
End Sub

Private Sub cmdMvSouth_Click()
If MvSouth = False Then
    cmdMvSouth.Caption = "Moving South"
    MvSouth = True
    MSComm1.PortOpen = True
    MSComm1.Output = SpeedSelection
    MSComm1.Output = "#:Ms#"
    MSComm1.PortOpen = False
    Exit Sub
End If

If MvSouth = True Then
    cmdMvSouth.Caption = "South"
    MvSouth = False
    MSComm1.PortOpen = True
    MSComm1.Output = "#:Qs#"
    MSComm1.PortOpen = False
End If
End Sub

Private Sub cmdMvWest_Click()
If MvWest = False Then
    cmdMvWest.Caption = "Moving West"
    MvWest = True
    MSComm1.PortOpen = True
    MSComm1.Output = SpeedSelection
    MSComm1.Output = "#:Mw#"
    MSComm1.PortOpen = False
    Exit Sub
End If

If MvWest = True Then
    cmdMvWest.Caption = "West"
    MvWest = False
    MSComm1.PortOpen = True
    MSComm1.Output = "#:Qw#"
    MSComm1.PortOpen = False
End If
End Sub



Private Sub Command1_Click()
MSComm1.PortOpen = True
MSComm1.Output = Text6.Text


End Sub

Private Sub Command2_Click()

temp = MSComm1.Input
Text5.Text = temp
MSComm1.PortOpen = False
End Sub

Private Sub exitme_Click()
End
End Sub

Private Sub Fast_Click()
If Fast.Enabled = True Then SpeedSelection = "#:Sw4#"
End Sub

Private Sub Form_Load()
Medium.Value = True
End Sub


Private Sub Medium_Click()
If Medium.Enabled = True Then SpeedSelection = "#:Sw3"

End Sub

Private Sub Slow_Click()
If Slow.Enabled = True Then SpeedSelection = "#:Sw2#"
End Sub
