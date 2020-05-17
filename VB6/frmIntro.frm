VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIntro 
   Caption         =   "LOADING..."
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   4200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   4050
      Left            =   0
      Picture         =   "frmIntro.frx":0000
      Top             =   0
      Width           =   12585
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
    a = 0
End Sub

Private Sub Timer1_Timer()
    a = a + 20
    ProgressBar1.Value = a
    
    If a = 100 Then
        frmBanner.Show
        a = 0
        Timer1.Enabled = False
        Unload Me
    End If
End Sub

