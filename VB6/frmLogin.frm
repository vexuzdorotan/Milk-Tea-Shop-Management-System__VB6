VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFFF&
   Caption         =   "LOGIN FORM"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLogin.frx":0000
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSAdodcLib.Adodc ADO2 
      Height          =   375
      Left            =   240
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmLogin.frx":0013
      OLEDBString     =   $"frmLogin.frx":009C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblSession"
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
   Begin VB.CommandButton cmdSignup 
      BackColor       =   &H008080FF&
      Caption         =   "Register"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox chkShow 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show Password"
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   240
      Top             =   2610
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   $"frmLogin.frx":0125
      OLEDBString     =   $"frmLogin.frx":01AE
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
   Begin VB.TextBox txtPassword 
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
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1290
      Width           =   3735
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "Rollno"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H008080FF&
      Caption         =   "Login"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton btnReset 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "or"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3435
      TabIndex        =   11
      Top             =   2925
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USER LOGIN"
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
      TabIndex        =   6
      Top             =   0
      Width           =   5640
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1290
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnReset_Click()
    fromLogin = 1
    frmReset.Show
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    Dim a, b, c, x
    
    isExec = 0
    isAdmin = 0
    isCashier = 0
    
    ADO.RecordSource = "SELECT * FROM tblLogin where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'"
    ADO.Refresh
    
    If ADO.Recordset.EOF Then
        MsgBox "Please enter correct username and password!", vbCritical, "Login Failed"
    Else
        x = ADO.Recordset.GetString(, 1)
        a = InStr(1, x, "Super Admin", 1)
        b = InStr(1, x, "Admin", 1)
        c = InStr(1, x, "Cashier", 1)
        
        If a > 0 Then
            ADO.Recordset.MovePrevious
            uNameUser = ADO.Recordset!UserName
            nameUser = ADO.Recordset!FullName
            typeUser = ADO.Recordset!AcctType
            
            ADO2.Recordset.AddNew
            ADO2.Recordset.Fields("AcctType") = typeUser
            ADO2.Recordset.Fields("Name") = nameUser
            ADO2.Recordset.Fields("Username") = txtUsername.Text
            ADO2.Recordset.Fields("Date") = Date
            ADO2.Recordset.Fields("TimeIn") = Time
            ADO2.Recordset.update
            
            MsgBox "Welcome, " & nameUser, vbOKOnly, "Welcome!"
            isExec = 1
            frmAdmin.Show
        ElseIf b > 0 Then
            ADO.Recordset.MovePrevious
            uNameUser = ADO.Recordset!UserName
            nameUser = ADO.Recordset!FullName
            typeUser = ADO.Recordset!AcctType
            
            ADO2.Recordset.AddNew
            ADO2.Recordset.Fields("AcctType") = typeUser
            ADO2.Recordset.Fields("Name") = nameUser
            ADO2.Recordset.Fields("Username") = txtUsername.Text
            ADO2.Recordset.Fields("Date") = Date
            ADO2.Recordset.Fields("TimeIn") = Time
            ADO2.Recordset.update
            
            MsgBox "Welcome, " & nameUser, vbOKOnly, "Welcome!"
            isAdmin = 1
            frmAdmin.Show
        ElseIf c > 0 Then
            isCashier = 1
            ADO.Recordset.MovePrevious
            uNameUser = ADO.Recordset!UserName
            nameUser = ADO.Recordset!FullName
            typeUser = ADO.Recordset!AcctType
            
            ADO2.Recordset.AddNew
            ADO2.Recordset.Fields("AcctType") = typeUser
            ADO2.Recordset.Fields("Name") = nameUser
            ADO2.Recordset.Fields("Username") = txtUsername.Text
            ADO2.Recordset.Fields("Date") = Date
            ADO2.Recordset.Fields("TimeIn") = Time
            ADO2.Recordset.update
            
            MsgBox "Welcome, " & nameUser, vbOKOnly, "Welcome!"
            frmContent.Show
        End If
        frmLogin.Hide
    End If
End Sub

Private Sub chkShow_Click()
    If chkShow.Value = 1 Then
        txtPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "*"
    End If
End Sub

Private Sub cmdExit_Click()
    Dim ans As String
    
    ans = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
    If ans = vbYes Then
        End
    End If
End Sub

Private Sub cmdSignup_Click()
    fromLogin = 1
    frmSignup.Show
End Sub

Private Sub Command1_Click()
    dataReceipt.Sections("Section2").Controls("lblNo").Caption = "ew"
    dataReceipt.Show vbModal
End Sub

Private Sub Form_Load()
    Label3.MousePointer = 99
End Sub
