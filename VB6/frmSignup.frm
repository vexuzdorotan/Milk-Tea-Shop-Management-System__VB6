VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSignup 
   BackColor       =   &H00C0FFFF&
   Caption         =   "SIGNUP FORM"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSched 
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
      ItemData        =   "frmSignup.frx":0000
      Left            =   2520
      List            =   "frmSignup.frx":000A
      TabIndex        =   19
      Text            =   "Please select your schedule."
      Top             =   3720
      Width           =   5535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSignup.frx":0022
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   5760
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Connect         =   $"frmSignup.frx":0034
      OLEDBString     =   $"frmSignup.frx":00BD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblLogin"
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
   Begin VB.TextBox txtAnswer 
      DataField       =   "Name"
      DataSource      =   "registerado"
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   4935
      Width           =   5535
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H008080FF&
      Caption         =   "Confirm"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C000&
      Caption         =   "Cancel"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   3735
   End
   Begin VB.OptionButton optCashier 
      BackColor       =   &H008080FF&
      Caption         =   "Cashier"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.OptionButton optAdmin 
      BackColor       =   &H008080FF&
      Caption         =   "Admin"
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
      Left            =   2535
      TabIndex        =   1
      Top             =   1335
      Width           =   2520
   End
   Begin VB.ComboBox cboQuestion 
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
      ItemData        =   "frmSignup.frx":0146
      Left            =   2520
      List            =   "frmSignup.frx":0159
      TabIndex        =   6
      Text            =   "Please select your security question."
      Top             =   4320
      Width           =   5535
   End
   Begin VB.TextBox txtConfirm 
      DataField       =   "Name"
      DataSource      =   "registerado"
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
      IMEMode         =   3  'DISABLE
      Left            =   2535
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3135
      Width           =   5535
   End
   Begin VB.TextBox txtPass 
      DataField       =   "Name"
      DataSource      =   "registerado"
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
      IMEMode         =   3  'DISABLE
      Left            =   2535
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2535
      Width           =   5535
   End
   Begin VB.TextBox txtName 
      DataField       =   "FullName"
      DataSource      =   "adoSignup"
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
      Left            =   2535
      TabIndex        =   0
      Top             =   735
      Width           =   5535
   End
   Begin VB.TextBox txtUser 
      DataField       =   "Name"
      DataSource      =   "registerado"
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
      IMEMode         =   3  'DISABLE
      Left            =   2535
      TabIndex        =   3
      Top             =   1935
      Width           =   5535
   End
   Begin VB.Label Label9 
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
      Height          =   420
      Left            =   240
      TabIndex        =   20
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Schedule"
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
      TabIndex        =   18
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Answer"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Height          =   420
      Left            =   240
      TabIndex        =   15
      Top             =   1935
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Account Type"
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
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIGNUP FORM"
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
      TabIndex        =   13
      Top             =   0
      Width           =   8370
   End
   Begin VB.Label Label5 
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
      Height          =   420
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Password"
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
      TabIndex        =   11
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Full Name"
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
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmSignup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim userExisted As String

Private Sub cmdCheck_Click()
    ADO.RecordSource = "Select * from tblLogin where Username='" & txtUser.Text & "'"
    ADO.Refresh
    

End Sub

Private Sub cmdConfirm_Click()
    Dim selectedUser As String
    selectedUser = ""
    
    If optAdmin.Value = True Then
        selectedUser = optAdmin.Caption
    ElseIf optCashier.Value = True Then
        selectedUser = optCashier.Caption
    End If
    
    ADO.RecordSource = "Select * from tblLogin where Username='" & txtUser.Text & "'"
    ADO.Refresh
    
    If ADO.Recordset.EOF Then
        If txtName.Text = "" Or selectedUser = "" Or txtUser.Text = "" Or txtPass.Text = "" Or txtConfirm.Text = "" Or _
            cboQuestion.Text = "Please select your security question." Or txtAnswer.Text = "" Then
                MsgBox "Please fill up all forms!", vbCritical, "Error"
        ElseIf txtPass.Text <> txtConfirm.Text Then
            MsgBox "Password not matched!", vbCritical, "Error"
        Else
            ADO.Recordset.AddNew
            
            If optAdmin.Value = True Then
                ADO.Recordset.Fields("AcctType") = optAdmin.Caption
            Else
                ADO.Recordset.Fields("AcctType") = optCashier.Caption
            End If
            
            ADO.Recordset.Fields("FullName") = txtName.Text
            ADO.Recordset.Fields("Username") = txtUser.Text
            ADO.Recordset.Fields("Password") = txtConfirm.Text
            ADO.Recordset.Fields("Schedule") = cboSched.Text
            ADO.Recordset.Fields("Question") = cboQuestion.Text
            ADO.Recordset.Fields("Answer") = txtAnswer.Text
            
            ADO.Recordset.update
            MsgBox "New user has been successfully added.", vbInformation, "New User Added"
            
            If isAdmin = 1 Then
                frmEditUser.ADO.RecordSource = "SELECT * FROM tblLogin where AcctType='Cashier'"
            End If
            
            ADO.Refresh
            frmEditUser.ADO.Caption = ADO.RecordSource
            frmEditUser.ADO.Refresh
            
            If fromLogin = 0 Then
                Unload frmEditUser
                frmEditUser.Show
                Unload Me
            Else
                Unload Me
            End If
        End If
    Else
        MsgBox "'" & txtUser.Text & "' already existed!"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub optAdmin_Click()
    If optAdmin.Value = True Then
        ADO.RecordSource = "SELECT * FROM tblLogin where AcctType='Admin'"
        ADO.Refresh
        ADO.Caption = ADO.RecordSource
        If ADO.Recordset.RecordCount = 50 Then
            MsgBox "Only 50 admins are allowed!", vbCritical, "Error"
            optAdmin.Value = False
            selectedUser = ""
        End If
    End If
End Sub

Private Sub optCashier_Click()
    If optCashier.Value = True Then
        ADO.RecordSource = "SELECT * FROM tblLogin where AcctType='Cashier'"
        ADO.Refresh
        ADO.Caption = ADO.RecordSource
        If ADO.Recordset.RecordCount = 100 Then
            MsgBox "Only 100 cashiers are allowed!"
        End If
    End If
End Sub
