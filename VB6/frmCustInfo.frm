VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmCustInfo 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CUSTOMER INFORMATION"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDiscount 
      BackColor       =   &H00C0C0FF&
      Caption         =   "20% Discount"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H008080FF&
      Caption         =   "Enter Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
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
      Left            =   6195
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdReceipt 
      BackColor       =   &H008080FF&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6195
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtContact 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox txtName 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmCustInfo.frx":0000
      Height          =   255
      Left            =   6960
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCustInfo.frx":0013
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
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
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   6120
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
      Connect         =   $"frmCustInfo.frx":0025
      OLEDBString     =   $"frmCustInfo.frx":00AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblReceipt"
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
   Begin MSAdodcLib.Adodc ADO2 
      Height          =   330
      Left            =   6120
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
      Connect         =   $"frmCustInfo.frx":0137
      OLEDBString     =   $"frmCustInfo.frx":01C0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblSales"
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUSTOMER INFO"
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
      TabIndex        =   16
      Top             =   0
      Width           =   7665
   End
   Begin VB.Label lblChange 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Change Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   14
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblPay 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Change Due"
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
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   1560
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Payment"
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
      Height          =   345
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Amount"
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
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Contact No."
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Address:"
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
      TabIndex        =   1
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Name:"
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
      TabIndex        =   0
      Top             =   720
      Width           =   1560
   End
End
Attribute VB_Name = "frmCustInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pay, total, totalTemp, discount As Double
Dim paymentTemp As String
Dim setDate, setTime As String
Dim i As Integer
Dim total2 As Double

Private Sub Form_Load()
    total = frmContent.total
    totalTemp = total
    discount = 0
    pay = 0

    If frmContent.optPayment1.Value = True Or frmContent.optPayment2.Value = True Then
       txtAddress.Enabled = False
       txtContact.Enabled = False
       
       txtAddress.BackColor = &H80000000
       txtContact.BackColor = &H80000000
       
       txtAddress.Text = "N/A"
       txtContact.Text = "N/A"
    End If
    
    Call update
End Sub

Private Sub update()
    lblTotal.Caption = total
    lblPay.Caption = Format(pay, "0.00")
    
    If pay - total > 0 Then
        lblChange.Caption = pay - total
    Else
        lblChange.Caption = 0
    End If
End Sub

Private Sub chkDiscount_Click()
    If chkDiscount.Value = 1 Then
        discount = total * 0.2
    Else
        discount = 0
        total = totalTemp
        pay = 0
    End If
    
    total = total - discount
    Call update
End Sub

Private Sub cmdEnter_Click()
    Dim ans As String
    
    ans = InputBox("Enter amount: ")
    If Val(ans) >= total Then
        pay = Val(ans)
    Else
        MsgBox "Please enter a sufficient amount!", vbCritical, "Error"
        pay = 0
        lblChange.Caption = Format(0, "0.00")
        frmCustInfo.Show
    End If

    Call update
End Sub

Private Sub cmdReceipt_Click()
    Dim ans As String

    If txtName.Text = "" Or txtAddress.Text = "" Or txtContact.Text = "" Or lblPay.Caption = "0.00" Then
        MsgBox "Please fill up all forms!", vbCritical, "Error"
    ElseIf Not IsNumeric(txtContact.Text) And Not txtContact.Text = "N/A" Then
        MsgBox "Please fill up the correct contact no.!", vbCritical, "Error"
    Else
        ans = MsgBox("Proceed?", vbYesNo, "Message")
        If ans = vbYes Then
            isNewReceipt = 1
    
            
            setDate = Format(Date, "mm/dd/yyyy")
            setTime = Format(Time, "h:mm AM/PM")
            
            total2 = lblTotal.Caption
            frmBanner.totalIncome = frmBanner.totalIncome + total2
            
            dataReceipt.Sections("Section2").Controls("lblDateTime").Caption = setDate & "    " & setTime
            dataReceipt.Sections("Section2").Controls("lblName").Caption = frmCustInfo.txtName.Text
            dataReceipt.Sections("Section2").Controls("lblAddress").Caption = frmCustInfo.txtAddress.Text
            dataReceipt.Sections("Section2").Controls("lblContact").Caption = frmCustInfo.txtContact.Text
            dataReceipt.Sections("Section2").Controls("lblPayment").Caption = frmContent.selectPayment
            
            dataReceipt.Sections("Section2").Controls("lblItem").Caption = frmContent.labelItem
            dataReceipt.Sections("Section2").Controls("lblSubtotalList").Caption = frmContent.labelSubtotal
            
            dataReceipt.Sections("Section2").Controls("lblNoItems").Caption = frmContent.qtyTotal
            dataReceipt.Sections("Section2").Controls("lblSubtotal").Caption = Format(total2 * 0.88, "0.00")
            dataReceipt.Sections("Section2").Controls("lblDiscount").Caption = Format(frmCustInfo.discount, "0.00")
            dataReceipt.Sections("Section2").Controls("lblVAT").Caption = Format(total2 * 0.12, "0.00")
            dataReceipt.Sections("Section2").Controls("lblTotal").Caption = Format(total2, "0.00")
            dataReceipt.Sections("Section2").Controls("lblPay").Caption = Format(frmCustInfo.lblPay.Caption, "0.00")
            dataReceipt.Sections("Section2").Controls("lblChange").Caption = Format(frmCustInfo.lblChange.Caption, "0.00")
            
            ADO.Recordset.AddNew
            ADO.Recordset.Fields("Cashier") = nameUser
            ADO.Recordset.Fields("Date") = setDate
            ADO.Recordset.Fields("Time") = setTime
            ADO.Recordset.Fields("Name") = frmCustInfo.txtName.Text
            ADO.Recordset.Fields("Address") = frmCustInfo.txtAddress.Text
            ADO.Recordset.Fields("Contact") = frmCustInfo.txtContact.Text
            ADO.Recordset.Fields("Payment") = frmContent.selectPayment
            
            ADO.Recordset.Fields("ItemList") = frmContent.labelItem
            ADO.Recordset.Fields("AmountList") = frmContent.labelSubtotal
            
            ADO.Recordset.Fields("NoItems") = frmContent.qtyTotal
            ADO.Recordset.Fields("Subtotal") = Format(total2 * 0.88, "0.00")
            ADO.Recordset.Fields("VAT") = Format(total2 * 0.12, "0.00")
            ADO.Recordset.Fields("Discount") = Format(frmCustInfo.discount, "0.00")
            ADO.Recordset.Fields("Total") = Format(total2, "0.00")
            ADO.Recordset.Fields("Pay") = Format(frmCustInfo.lblPay.Caption, "0.00")
            ADO.Recordset.Fields("Change") = Format(frmCustInfo.lblChange.Caption, "0.00")
            ADO.Recordset.update
            
            ADO2.Recordset.AddNew
            ADO2.Recordset.Fields("ReceiptNo") = ADO.Recordset!ReceiptNo
            ADO2.Recordset.Fields("Cashier") = uNameUser
            ADO2.Recordset.Fields("Date") = setDate
            ADO2.Recordset.Fields("Time") = setTime
            ADO2.Recordset.Fields("Total") = Format(total2, "0.00")
            ADO2.Recordset.update
            
            dataReceipt.Sections("Section2").Controls("lblNo").Caption = ADO.Recordset!ReceiptNo
            dataReceipt.Sections("Section2").Controls("lblCashier").Caption = nameUser
            
            frmCustInfo.Hide
            
            Unload frmContent
            frmContent.Show
            dataReceipt.Show
            Unload frmCustInfo
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
    frmContent.Show
End Sub
