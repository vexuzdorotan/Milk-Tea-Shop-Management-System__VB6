VERSION 5.00
Begin VB.Form frmBanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WELCOME!"
   ClientHeight    =   4995
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H008080FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   0
      Picture         =   "frmBanner.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frmBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public loginTime As String
Public loginTimeEnable As Integer
Public totalIncome As Double

Private Sub cmdProceed_Click()
    loginTime = Format(Time, "h:mm AM/PM")
    loginTimeEnable = 1
    frmBanner.Hide
    frmLogin.Show
End Sub

Private Sub cmdExit_Click()
    Dim ans As String
    
    ans = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
    If ans = vbYes Then
        End
    End If
End Sub
