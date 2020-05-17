VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00C0FFFF&
   Caption         =   "SALES REPORT"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H008080FF&
      Caption         =   "Print Report"
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
      TabIndex        =   16
      Top             =   5640
      Width           =   8775
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearchCashier 
      BackColor       =   &H008080FF&
      Caption         =   "Search"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtCashier 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc ADO2 
      Height          =   330
      Left            =   7560
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"frmSales.frx":0000
      OLEDBString     =   $"frmSales.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblReceipt"
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmSales.frx":0112
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H008080FF&
      Caption         =   "Show All"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H008080FF&
      Caption         =   "Search"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   270
      Left            =   240
      Top             =   5160
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   476
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
      Appearance      =   0
      BackColor       =   8421631
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmSales.frx":0125
      OLEDBString     =   $"frmSales.frx":01AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblSales"
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
   Begin MSDataGridLib.DataGrid gridSales 
      Bindings        =   "frmSales.frx":0237
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton cmdYearly 
      BackColor       =   &H008080FF&
      Caption         =   "Yearly Sales"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdMonthly 
      BackColor       =   &H008080FF&
      Caption         =   "Monthly Sales"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdDaily 
      BackColor       =   &H008080FF&
      Caption         =   "Daily Sales"
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
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   103284739
      CurrentDate     =   43101
      MaxDate         =   43830
      MinDate         =   43101
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   103284739
      CurrentDate     =   43830
      MaxDate         =   43830
      MinDate         =   43101
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6480
      TabIndex        =   17
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cashier Name"
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
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALES REPORT"
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
      TabIndex        =   9
      Top             =   0
      Width           =   9360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "To"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "From"
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
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim date1, date2 As String
Dim i As Integer
Dim tot As Double

Private Sub cmdDaily_Click()
    dataDaily.Sections("Section4").Controls("lblViewed").Caption = "Viewed by: [" & typeUser & _
        "] " & nameUser & " @ " & Date & ", " & Format(Time, "h:mm AM/PM")
        
    If DataEnvironment1.rstblDaily.State = adStateOpen Then
        DataEnvironment1.rstblDaily.Close
    End If
    
    dataDaily.Show
End Sub

Private Sub cmdMonthly_Click()
    dataMonthly.Sections("Section4").Controls("lblViewed").Caption = "Viewed by: [" & typeUser & _
        "] " & nameUser & " @ " & Date & ", " & Format(Time, "h:mm AM/PM")
    
    If DataEnvironment1.rstblMonthly.State = adStateOpen Then
        DataEnvironment1.rstblMonthly.Close
    End If
        
    dataMonthly.Show
End Sub

Private Sub cmdYearly_Click()
    dataYearly.Sections("Section4").Controls("lblViewed").Caption = "Viewed by: [" & typeUser & _
        "] " & nameUser & " @ " & Date & ", " & Format(Time, "h:mm AM/PM")
        
    If DataEnvironment1.rstblYearly.State = adStateOpen Then
        DataEnvironment1.rstblYearly.Close
    End If
    
    dataYearly.Show
End Sub

Private Sub cmdRetrieve_Click()
    frmRetrieve.Show
End Sub

Private Sub cmdPrint_Click()
    Set dataReports.DataSource = ADO
    dataReports.Sections("Section3").Controls("lblTotal").Caption = lblTotal.Caption
    
    dataReports.Sections("Section4").Controls("lblViewed").Caption = "Viewed by: [" & typeUser & _
        "] " & nameUser & " @ " & Date & ", " & Format(Time, "h:mm AM/PM")
        
    dataReports.Show
End Sub

Private Sub cmdSearchCashier_Click()
    ADO.RecordSource = "Select * from tblSales where Cashier='" + txtCashier.Text + "'"
    ADO.Refresh
    
    If ADO.Recordset.EOF Then
        MsgBox "Cashier not found!", vbCritical, "Error"
    Else
        'ADO.Caption = ADO.RecordSource
        tot = 0
        For i = 0 To ADO.Recordset.RecordCount - 1
            tot = tot + CDbl(gridSales.Columns(4).Text)
            ADO.Recordset.MoveNext
        Next i
        
        lblTotal.Caption = "Total: " & Format(tot, "0.00")
    End If
End Sub

Private Sub cmdAll_Click()
    ADO.RecordSource = "Select * from tblSales"
    ADO.Refresh
    
    tot = 0
    For i = 0 To ADO.Recordset.RecordCount - 1
        tot = tot + CDbl(gridSales.Columns(4).Text)
        ADO.Recordset.MoveNext
    Next i
    
    lblTotal.Caption = "Total: " & Format(tot, "0.00")
    
    ADO.Recordset.MoveLast
End Sub

Private Sub cmdSearch_Click()
    date1 = Format(DTPicker1.Value, "mm/dd/yyyy")
    date2 = Format(DTPicker2.Value, "mm/dd/yyyy")
    
    If date2 < date1 Then
        MsgBox "Please select the correct date!", vbCritical, "Warning Message"
    Else
        ADO.RecordSource = "SELECT * FROM tblSales where Date between # " & date1 & " # and # " & date2 & " # "
        ADO.Refresh
        
        If ADO.Recordset.EOF Then
            MsgBox "Record not found!", vbCritical, "Warning Message"
        Else
            ADO.Caption = ADO.RecordSource
            
            tot = 0
            For i = 0 To ADO.Recordset.RecordCount - 1
                tot = tot + CDbl(gridSales.Columns(4).Text)
                ADO.Recordset.MoveNext
            Next i
            
            lblTotal.Caption = "Total: " & Format(tot, "0.00")
            
            ADO.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    If ADO.Recordset.EOF Then
        MsgBox "Record not found!", vbCritical, "Warning Message"
    Else
        For i = 0 To ADO.Recordset.RecordCount - 1
            tot = tot + CDbl(gridSales.Columns(4).Text)
            ADO.Recordset.MoveNext
        Next i
        
        lblTotal.Caption = "Total: " & Format(tot, "0.00")
        
        ADO.Recordset.MoveLast
    End If
End Sub

