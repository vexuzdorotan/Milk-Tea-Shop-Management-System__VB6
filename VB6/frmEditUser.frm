VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmEditUser 
   BackColor       =   &H00C0FFFF&
   Caption         =   "EDIT USERS(S)"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPass 
      BackColor       =   &H008080FF&
      Caption         =   "Edit Password"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   495
      Left            =   12960
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Connect         =   $"frmEditUser.frx":0000
      OLEDBString     =   $"frmEditUser.frx":0089
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
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H008080FF&
      Caption         =   "Add New User"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H008080FF&
      Caption         =   "Delete User"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEditUser.frx":0112
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDIT USER(S)"
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
      TabIndex        =   4
      Top             =   0
      Width           =   14955
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPass_Click()
    fromLogin = 0
    frmReset.Show
    frmReset.ADO.RecordSource = "SELECT * FROM tblLogin where username='" & ADO.Recordset!UserName & "'"
    frmReset.ADO.Refresh
    
    frmReset.txtUser.Text = ADO.Recordset!UserName
    frmReset.txtAnswer.Text = ADO.Recordset!Answer
    
    frmReset.txtUser.Enabled = False
    frmReset.txtAnswer.Enabled = False
    frmReset.cmdCheckUser.Enabled = False
    frmReset.cmdCheckAnswer.Enabled = False
    
    frmReset.txtNew.Enabled = True
    frmReset.txtConfirm.Enabled = True
    frmReset.cmdChangePass.Enabled = True
End Sub

Private Sub Form_Load()
    If isAdmin = 1 Then
        ADO.RecordSource = "SELECT * FROM tblLogin where AcctType='Admin' or AcctType='Cashier'"
    End If
    
    'ADO.Refresh
    ADO.Caption = ADO.RecordSource
    ADO.Refresh
    ADO.Recordset.MoveLast
End Sub


Private Sub cmdAdd_Click()
    ADO.RecordSource = "SELECT * FROM tblLogin"
    If ADO.Recordset.RecordCount = 50 Then
        MsgBox "Maximum number of users reached!", vbCritical, "Error"
    Else
        fromLogin = 0
        frmSignup.Show
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim x, y, ans As String
    
    x = ADO.Recordset.GetString(, 1)
    y = InStr(1, x, "Super Admin", 1)
    ADO.Recordset.MovePrevious
    
    If ADO.Recordset!FullName = nameUser Then
        MsgBox "You cannot delete your own account!", vbCritical, "Error"
    Else
        If y > 0 Then
            MsgBox "Super Admin user cannot delete!", vbCritical, "Error Message"
        Else
            ans = MsgBox("Do you want to delete user?", vbYesNo + vbExclamation, "Warning Message")
            If ans = vbYes Then
                ADO.Recordset.Delete
                MsgBox "Record Deleted Successfully", vbInformation, "Delete Record Confirmation"
            Else
                MsgBox "Record Not Deleted", vbInformation, "Record Not Deleted"
            End If
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub
