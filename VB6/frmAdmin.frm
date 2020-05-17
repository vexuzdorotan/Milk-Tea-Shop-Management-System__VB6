VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ADMINISTRATOR"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRetrieve 
      BackColor       =   &H008080FF&
      Caption         =   "Retrieve Receipt"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdSession 
      BackColor       =   &H008080FF&
      Caption         =   "User Login/Logout Session"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton cmdCashiering 
      BackColor       =   &H008080FF&
      Caption         =   "Point of Sales (Cashiering)"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00C0C000&
      Caption         =   "Logout"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CommandButton cmdSales 
      BackColor       =   &H008080FF&
      Caption         =   "View Sales Report"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H008080FF&
      Caption         =   "Edit User"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADMINISTRATOR VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSession_Click()
    frmSession.Show
End Sub

Private Sub Form_Load()
    If isExec = 1 Then
        Label1.Caption = "SUPER ADMIN VIEW"
    Else
        Label1.Caption = "ADMIN VIEW"
    End If
End Sub

Private Sub cmdCashiering_Click()
    frmContent.Show
End Sub

Private Sub cmdRetrieve_Click()
    frmRetrieve.Show
End Sub

Private Sub cmdSales_Click()
    frmSales.Show
End Sub

Private Sub cmdEdit_Click()
    frmEditUser.Show
End Sub

Private Sub cmdLogout_Click()
    Dim ans As String
    
    ans = MsgBox("Are you sure you want to logout?", vbYesNo, "Logout")
    If ans = vbYes Then
        frmLogin.ADO2.Recordset.Fields("TimeOut") = Time
        frmLogin.ADO2.Recordset.update
        
        Unload frmLogin
        frmLogin.Show
        Unload Me
    End If
End Sub
