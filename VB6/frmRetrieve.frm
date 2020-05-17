VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRetrieve 
   BackColor       =   &H00C0FFFF&
   Caption         =   "RETRIEVE RECEIPT(S)"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H008080FF&
      Caption         =   "Preview Receipt"
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
      Top             =   4440
      Width           =   8295
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1335
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
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRetrieve.frx":0000
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
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
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   240
      Top             =   4080
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
      Connect         =   $"frmRetrieve.frx":0012
      OLEDBString     =   $"frmRetrieve.frx":009B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblReceipt"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   720
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
      Format          =   84213763
      CurrentDate     =   43101
      MaxDate         =   43830
      MinDate         =   43101
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   720
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
      Format          =   84213763
      CurrentDate     =   43830
      MaxDate         =   43830
      MinDate         =   43101
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
      Top             =   720
      Width           =   735
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
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RETRIEVE RECEIPT(S)"
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
      TabIndex        =   1
      Top             =   0
      Width           =   8865
   End
End
Attribute VB_Name = "frmRetrieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim date1, date2 As String

Private Sub cmdPreview_Click()
    If ADO.Recordset.EOF Then
        MsgBox "End of file reached!", vbCritical, "Error!"
    Else
        isNewReceipt = 0
    
        dataReceipt.Sections("Section2").Controls("lblNo").Caption = ADO.Recordset!ReceiptNo
        dataReceipt.Sections("Section2").Controls("lblCashier").Caption = ADO.Recordset!Cashier
        dataReceipt.Sections("Section2").Controls("lblDateTime").Caption = ADO.Recordset!Date & "    " & ADO.Recordset!Time
    
        dataReceipt.Sections("Section2").Controls("lblName").Caption = ADO.Recordset!Name
        dataReceipt.Sections("Section2").Controls("lblAddress").Caption = ADO.Recordset!Address
        dataReceipt.Sections("Section2").Controls("lblContact").Caption = ADO.Recordset!Contact
        dataReceipt.Sections("Section2").Controls("lblPayment").Caption = ADO.Recordset!Payment
        
        dataReceipt.Sections("Section2").Controls("lblItem").Caption = ADO.Recordset!ItemList
        dataReceipt.Sections("Section2").Controls("lblSubtotalList").Caption = ADO.Recordset!AmountList
        
        dataReceipt.Sections("Section2").Controls("lblNoItems").Caption = ADO.Recordset!NoItems
        dataReceipt.Sections("Section2").Controls("lblSubtotal").Caption = Format(ADO.Recordset!subtotal, "0.00")
        dataReceipt.Sections("Section2").Controls("lblVAT").Caption = Format(ADO.Recordset!VAT, "0.00")
        dataReceipt.Sections("Section2").Controls("lblDiscount").Caption = Format(ADO.Recordset!discount, "0.00")
        dataReceipt.Sections("Section2").Controls("lblTotal").Caption = Format(ADO.Recordset!total, "0.00")
        
        dataReceipt.Sections("Section2").Controls("lblPay").Caption = Format(ADO.Recordset!pay, "0.00")
        dataReceipt.Sections("Section2").Controls("lblChange").Caption = Format(ADO.Recordset!Change, "0.00")
    
        dataReceipt.Show
    End If
End Sub

Private Sub cmdSearch_Click()
    date1 = Format(DTPicker1.Value, "mm/dd/yyyy")
    date2 = Format(DTPicker2.Value, "mm/dd/yyyy")
    
    If date2 < date1 Then
        MsgBox "Please select the correct date!", vbCritical, "Warning Message"
    Else
        ADO.RecordSource = "SELECT * FROM tblReceipt where Date between # " & date1 & " # and # " & date2 & " # "
        ADO.Refresh
        
        If ADO.Recordset.EOF Then
            MsgBox "Record not found!", vbCritical, "Warning Message"
        Else
            ADO.Caption = ADO.RecordSource
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
        ADO.Recordset.MoveLast
    End If
End Sub
