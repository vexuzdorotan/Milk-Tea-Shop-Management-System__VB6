VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReset 
   BackColor       =   &H00C0FFFF&
   Caption         =   "RESET PASSWORD"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   225
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmReset.frx":0000
      OLEDBString     =   $"frmReset.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblLogin"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2865
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdCheckUser 
      BackColor       =   &H008080FF&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCheckAnswer 
      BackColor       =   &H008080FF&
      Caption         =   "Check"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtAnswer 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox txtNew 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2865
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtConfirm 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2865
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton cmdChangePass 
      BackColor       =   &H008080FF&
      Caption         =   "Change Password"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RESET PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8595
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   10
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   9
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblNew 
      BackColor       =   &H00C0C0FF&
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblConfirm 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheckUser_Click()
    ADO.RecordSource = "SELECT * FROM tblLogin where username='" & txtUser.Text & "'"
    ADO.Refresh
    
    If ADO.Recordset.EOF Then
        MsgBox "User ID not found!", vbExclamation, "Failed"
    Else
        MsgBox "User ID found in the database!", vbInformation, "Success"
        txtAnswer.Enabled = True
        cmdCheckAnswer.Enabled = True
        
        lblQuestion.Caption = ADO.Recordset!Question
        txtUser.Enabled = False
    End If
End Sub

Private Sub cmdCheckAnswer_Click()
    Dim str As String
    str = StrComp(ADO.Recordset.Fields("Answer").Value, txtAnswer.Text, vbTextCompare)
    
    If str = True Then
        MsgBox "Sorry, wrong answer!", vbExclamation, "Failed"
    Else
        MsgBox "The answer is correct!", vbInformation, "Success"
        txtNew.Enabled = True
        txtConfirm.Enabled = True
        cmdChangePass.Enabled = True
        txtAnswer.Enabled = False
    End If
End Sub

Private Sub cmdChangePass_Click()
    If txtNew.Text = txtConfirm.Text Then
        ADO.Recordset.Fields("Password") = txtConfirm.Text
        ADO.Recordset.update
        MsgBox "Password Changed Successfully", vbInformation, "Success"
        Unload frmEditUser
        frmEditUser.Show
        Unload Me
    Else
        MsgBox "Password does not match. Try again!", vbExclamation, "Failed"
        txtNew.Text = ""
        txtConfirm.Text = ""
    End If
End Sub

Private Sub cmdBack_Click()
    If fromLogin = 1 Then
        Unload Me
        frmLogin.Show
    Else
        Unload Me
    End If
End Sub
