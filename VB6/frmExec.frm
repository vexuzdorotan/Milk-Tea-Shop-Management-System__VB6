VERSION 5.00
Begin VB.Form frmExec 
   BackColor       =   &H00C0FFFF&
   Caption         =   "EXECUTIVE"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCashiering 
      BackColor       =   &H008080FF&
      Caption         =   "Go to Cashiering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H008080FF&
      Caption         =   "Edit User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdSales 
      BackColor       =   &H008080FF&
      Caption         =   "View Sales Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00C0C000&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "EXECUTIVE VIEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   660
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCashiering_Click()
    frmContent.Show
End Sub

Private Sub cmdEdit_Click()
    frmEditUser.Show
End Sub

Private Sub cmdLogout_Click()
    frmLogin.Show
    Unload Me
End Sub

Private Sub cmdSales_Click()
    frmSales.Show
End Sub

Private Sub Form_Load()
    isAdmin = 0
End Sub
