VERSION 5.00
Begin VB.Form frmContent 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ORDER FORM"
   ClientHeight    =   11055
   ClientLeft      =   -2925
   ClientTop       =   -1740
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   737
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1358
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdResetAddons 
      BackColor       =   &H008080FF&
      Caption         =   "Reset Add-ons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   9720
      Width           =   1935
   End
   Begin VB.OptionButton optFlavor29 
      BackColor       =   &H008080FF&
      Caption         =   "Taro"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7815
      TabIndex        =   73
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor30 
      BackColor       =   &H008080FF&
      Caption         =   "Wintermelon"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7815
      TabIndex        =   72
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor17 
      BackColor       =   &H008080FF&
      Caption         =   "Oreo"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   71
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor18 
      BackColor       =   &H008080FF&
      Caption         =   "Papaya"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   70
      Top             =   6000
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor22 
      BackColor       =   &H008080FF&
      Caption         =   "Grapes"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   69
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor19 
      BackColor       =   &H008080FF&
      Caption         =   "Pineapple"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   68
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor24 
      BackColor       =   &H008080FF&
      Caption         =   "Mangosteen"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   67
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor16 
      BackColor       =   &H008080FF&
      Caption         =   "Mango"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   66
      Top             =   2640
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor21 
      BackColor       =   &H008080FF&
      Caption         =   "Apple"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   65
      Top             =   2640
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor23 
      BackColor       =   &H008080FF&
      Caption         =   "Lychee"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   64
      Top             =   6000
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor20 
      BackColor       =   &H008080FF&
      Caption         =   "Strawberry"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   63
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor25 
      BackColor       =   &H008080FF&
      Caption         =   "Matcha"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   62
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor26 
      BackColor       =   &H008080FF&
      Caption         =   "Orange"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   61
      Top             =   2640
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor27 
      BackColor       =   &H008080FF&
      Caption         =   "Pomegranate"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   60
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor28 
      BackColor       =   &H008080FF&
      Caption         =   "Raspberry"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   59
      Top             =   6000
      Width           =   1300
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   18360
      Top             =   120
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18720
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearOrder 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
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
      Left            =   18840
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5160
      Width           =   1110
   End
   Begin VB.CommandButton cmdCustInfo 
      BackColor       =   &H008080FF&
      Caption         =   "Confirm Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Service Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      TabIndex        =   45
      Top             =   9480
      Width           =   4575
      Begin VB.OptionButton optPayment3 
         BackColor       =   &H008080FF&
         Caption         =   "Delivery"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optPayment2 
         BackColor       =   &H008080FF&
         Caption         =   "Take Out"
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
         TabIndex        =   47
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optPayment1 
         BackColor       =   &H008080FF&
         Caption         =   "Dine-in"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.OptionButton optFlavor15 
      BackColor       =   &H008080FF&
      Caption         =   "Kiwi"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   44
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor14 
      BackColor       =   &H008080FF&
      Caption         =   "Jackfruit"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   43
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor13 
      BackColor       =   &H008080FF&
      Caption         =   "Coconut"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3465
      TabIndex        =   42
      Top             =   6000
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor12 
      BackColor       =   &H008080FF&
      Caption         =   "Chocolate"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   41
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor11 
      BackColor       =   &H008080FF&
      Caption         =   "Avocado"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   40
      Top             =   2640
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor10 
      BackColor       =   &H008080FF&
      Caption         =   "Vanilla"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor9 
      BackColor       =   &H008080FF&
      Caption         =   "Tamarind"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   38
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor8 
      BackColor       =   &H008080FF&
      Caption         =   "Honeydew"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   37
      Top             =   6000
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor7 
      BackColor       =   &H008080FF&
      Caption         =   "Ginger"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2025
      TabIndex        =   36
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor6 
      BackColor       =   &H008080FF&
      Caption         =   "Durian"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   2640
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor5 
      BackColor       =   &H008080FF&
      Caption         =   "Cucumber"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   9360
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor4 
      BackColor       =   &H008080FF&
      Caption         =   "Coffee Bean"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   7680
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor3 
      BackColor       =   &H008080FF&
      Caption         =   "Calamansi"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   6000
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor2 
      BackColor       =   &H008080FF&
      Caption         =   "Barley"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   4320
      Width           =   1300
   End
   Begin VB.OptionButton optFlavor1 
      BackColor       =   &H008080FF&
      Caption         =   "Almond"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   585
      TabIndex        =   30
      Top             =   2640
      Width           =   1300
   End
   Begin VB.CheckBox chkAdd5 
      BackColor       =   &H008080FF&
      Caption         =   "Red Bean [P10.00]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   29
      Top             =   9360
      Width           =   1905
   End
   Begin VB.CheckBox chkAdd4 
      BackColor       =   &H008080FF&
      Caption         =   "Mini Mochi [P15.00]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   28
      Top             =   7680
      Width           =   1905
   End
   Begin VB.CheckBox chkAdd3 
      BackColor       =   &H008080FF&
      Caption         =   "Ice Cream [P20.00]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   27
      Top             =   6000
      Width           =   1905
   End
   Begin VB.CheckBox chkAdd2 
      BackColor       =   &H008080FF&
      Caption         =   "Grass Jelly [P15.00]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   26
      Top             =   4320
      Width           =   1905
   End
   Begin VB.CheckBox chkAdd1 
      BackColor       =   &H008080FF&
      Caption         =   "Egg Pudding [P25.00]"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   25
      Top             =   2640
      Width           =   1905
   End
   Begin VB.CommandButton cmdDeleteOrder 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
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
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1110
   End
   Begin VB.CommandButton cmdShowOrder 
      BackColor       =   &H008080FF&
      Caption         =   "Show"
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   1110
   End
   Begin VB.CommandButton cmdAddOrder 
      BackColor       =   &H008080FF&
      Caption         =   "Add"
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
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5160
      Width           =   1110
   End
   Begin VB.CommandButton cmdSubQty 
      BackColor       =   &H008080FF&
      Caption         =   "-"
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
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton cmdAddQty 
      BackColor       =   &H008080FF&
      Caption         =   "+"
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
      Left            =   19200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Sugar Content"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      TabIndex        =   13
      Top             =   7560
      Width           =   7215
      Begin VB.OptionButton optSugar5 
         BackColor       =   &H008080FF&
         Caption         =   "0%"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optSugar1 
         BackColor       =   &H008080FF&
         Caption         =   "100%"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optSugar2 
         BackColor       =   &H008080FF&
         Caption         =   "75%"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optSugar3 
         BackColor       =   &H008080FF&
         Caption         =   "50%"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optSugar4 
         BackColor       =   &H008080FF&
         Caption         =   "25%"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Ice Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      TabIndex        =   9
      Top             =   8520
      Width           =   4560
      Begin VB.OptionButton optIce1 
         BackColor       =   &H008080FF&
         Caption         =   "Regular"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optIce2 
         BackColor       =   &H008080FF&
         Caption         =   "Less Ice"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optIce3 
         BackColor       =   &H008080FF&
         Caption         =   "No Ice"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.ListBox lstOrder 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   15240
      TabIndex        =   8
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Cup Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12720
      TabIndex        =   3
      Top             =   6000
      Width           =   5895
      Begin VB.OptionButton optSize4 
         BackColor       =   &H008080FF&
         Caption         =   "Short [P120.00]"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   2445
      End
      Begin VB.OptionButton optSize3 
         BackColor       =   &H008080FF&
         Caption         =   "Tall [P135.00]"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2445
      End
      Begin VB.OptionButton optSize2 
         BackColor       =   &H008080FF&
         Caption         =   "Grande [P150.00]"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   2445
      End
      Begin VB.OptionButton optSize1 
         BackColor       =   &H008080FF&
         Caption         =   "Venti [P165.00]"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2445
      End
   End
   Begin VB.CommandButton cmdType3 
      BackColor       =   &H008080FF&
      Caption         =   "Milk Tea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   2775
   End
   Begin VB.CommandButton cmdType2 
      BackColor       =   &H008080FF&
      Caption         =   "Frappe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   2775
   End
   Begin VB.CommandButton cmdType1 
      BackColor       =   &H008080FF&
      Caption         =   "Hot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      Caption         =   "Qty."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   18960
      TabIndex        =   75
      Top             =   6000
      Width           =   975
   End
   Begin VB.Image Image15 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   7815
      Picture         =   "frmContent.frx":0000
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image14 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   7815
      Picture         =   "frmContent.frx":3183E
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image13 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   600
      Picture         =   "frmContent.frx":654F8
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image12 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   2040
      Picture         =   "frmContent.frx":7866A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   3480
      Picture         =   "frmContent.frx":A23DC
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   4920
      Picture         =   "frmContent.frx":E1A2A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   6360
      Picture         =   "frmContent.frx":11288E
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   600
      Picture         =   "frmContent.frx":139870
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   2040
      Picture         =   "frmContent.frx":1587D0
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   3480
      Picture         =   "frmContent.frx":17AAA9
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   4920
      Picture         =   "frmContent.frx":1AECEF
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   6360
      Picture         =   "frmContent.frx":1E8026
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1305
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   7800
      Picture         =   "frmContent.frx":24A7B9
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   7800
      Picture         =   "frmContent.frx":27DA54
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   7800
      Picture         =   "frmContent.frx":2B5FFF
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   4905
      Picture         =   "frmContent.frx":2E9FFB
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblNameUser 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome, Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   58
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total: "
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
      Left            =   15240
      TabIndex        =   57
      Top             =   4680
      Width           =   4680
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Current Order"
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
      Left            =   12720
      TabIndex        =   56
      Top             =   960
      Width           =   2280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Order Summary"
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
      Left            =   15240
      TabIndex        =   55
      Top             =   960
      Width           =   4665
   End
   Begin VB.Label lblDateTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date and Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   53
      Top             =   480
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Add-ons:"
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
      Left            =   9960
      TabIndex        =   50
      Top             =   960
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Flavors:"
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
      Left            =   600
      TabIndex        =   49
      Top             =   960
      Width           =   8505
   End
   Begin VB.Label lblOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   12720
      TabIndex        =   19
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image imgAdd5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   9960
      Picture         =   "frmContent.frx":316E93
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1905
   End
   Begin VB.Image imgAdd4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   9960
      Picture         =   "frmContent.frx":33D712
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1905
   End
   Begin VB.Image imgAdd3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   9960
      Picture         =   "frmContent.frx":36C739
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Image imgAdd2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   9945
      Picture         =   "frmContent.frx":38720A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Image imgAdd1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   9960
      Picture         =   "frmContent.frx":39D166
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1905
   End
   Begin VB.Image imgFlavor15 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   6345
      Picture         =   "frmContent.frx":3BAA8D
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor14 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   4905
      Picture         =   "frmContent.frx":3EE235
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor13 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   3465
      Picture         =   "frmContent.frx":415663
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor12 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   2025
      Picture         =   "frmContent.frx":447AAD
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor11 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   585
      Picture         =   "frmContent.frx":47EE7F
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Image imgFlavor10 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   6345
      Picture         =   "frmContent.frx":4A7CE5
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image imgFlavor9 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   4905
      Picture         =   "frmContent.frx":4D1413
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image imgFlavor8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   3465
      Picture         =   "frmContent.frx":4F3980
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image imgFlavor7 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   2025
      Picture         =   "frmContent.frx":50F99D
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image imgFlavor6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   585
      Picture         =   "frmContent.frx":5388EB
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Image imgFlavor5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   6345
      Picture         =   "frmContent.frx":5686C9
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Image imgFlavor3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   3465
      Picture         =   "frmContent.frx":58C3A9
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Image imgFlavor2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   2025
      Picture         =   "frmContent.frx":5B06BD
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Image imgFlavor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   585
      Picture         =   "frmContent.frx":5EF35C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1305
   End
End
Attribute VB_Name = "frmContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectType, selectFlavor, selectAddOns, selectSugar, selectSize, selectIce, _
    initType, initAdd, initAdd1, initAdd2, initAdd3, initAdd4, initAdd5, initSize, initIce, _
    add1, add2, add3, add4, add5, orderList(5), orderReceipt(5), orderMsg(5), tempIce As String
Dim qty, orderNo, qtyNo(5) As Integer
Dim amtAddOns, amtSize, subtotal(5) As Double
Public selectPayment, labelItem, labelSubtotal As String
Public qtyTotal As Integer
Public total As Double

Private Sub cmdResetAddons_Click()
    Call resetAddons
End Sub

Private Sub resetAddons()
    chkAdd1.Value = 0
    chkAdd2.Value = 0
    chkAdd3.Value = 0
    chkAdd4.Value = 0
    chkAdd5.Value = 0
    
    initAdd1 = "0"
    initAdd2 = "0"
    initAdd3 = "0"
    initAdd4 = "0"
    initAdd5 = "0"
    
    Call update
End Sub

Private Sub Form_Load()
    lblNameUser.Caption = "Welcome, " & "[" & typeUser & "]" & " " & nameUser
    selectFlavor = ""
    selectAddOns = "" & vbNewLine
    orderNo = 0
    total = 0
    amtSize = 0
    initSize = "Ve"
    initIce = "Re"
    tempIce = ""
    qtyTotal = 0
    
    Call reset
    Call update
End Sub

Private Sub update()
    Call addOns
    
    If qtyNo(orderNo) = 1 Then
        cmdSubQty.Enabled = False
    End If
    
    'If optSize1.Value = False And optSize2.Value = False And optSize3.Value = False And optSize4.Value = False Then
    '    subtotal(orderNo) = 0
    'Else
        subtotal(orderNo) = (amtAddOns + amtSize) * qtyNo(orderNo)
    'End If
    
    initAdd = initAdd1 & initAdd2 & initAdd3 & initAdd4 & initAdd5
    lblTotal.Caption = "Total: " & Format(total, "0.00")
    
    orderReceipt(orderNo) = qtyNo(orderNo) & " " & initType & " " & selectFlavor & " " & initAdd & " " & selectSugar & _
        " " & initSize & " " & initIce
    lblOrder.Caption = "Type: " & selectType & vbNewLine & "Flavor: " & selectFlavor & vbNewLine & "Sugar Content: " & _
        selectSugar & vbNewLine & "Size: " & selectSize & vbNewLine & "Ice Quantity: " & selectIce & vbNewLine & _
        vbNewLine & "Add-ons:" & vbNewLine & selectAddOns & vbNewLine & "Quantity: " & qtyNo(orderNo) & vbNewLine & "Sub-total: " & _
        Format(subtotal(orderNo), "0.00")
End Sub

Private Sub addOns()
    If cmdType1.Value = True Or (add1 = "" And add2 = "" And add3 = "" And add4 = "" And add5 = "") Then
        selectAddOns = "    No Add-ons" & vbNewLine
    Else
        selectAddOns = add1 & add2 & add3 & add4 & add5
    End If
End Sub

Private Sub receipt()
    Dim i As Integer
    i = 0
    labelItem = ""
    labelSubtotal = ""
    
    Do While i < orderNo
        labelItem = labelItem & orderReceipt(i) & vbNewLine
        labelSubtotal = labelSubtotal & Format(subtotal(i), "0.00") & vbNewLine
        i = i + 1
    Loop
    Call update
End Sub

Private Sub reset()
    Call resetFlavors
    Call disableFlavors
        
    Call resetAddons
    
    tempIce = ""
    subtotal(orderNo) = 0
    
    cmdType1.Value = False
    cmdType2.Value = False
    cmdType3.Value = False
    
    optSize1.Value = False
    optSize2.Value = False
    optSize3.Value = False
    optSize4.Value = False
    
    optSugar1.Value = False
    optSugar2.Value = False
    optSugar3.Value = False
    optSugar4.Value = False
    optSugar5.Value = False
    
    optIce1.Value = False
    optIce2.Value = False
    optIce3.Value = False
    
    cmdType1.BackColor = &H8080FF
    cmdType2.BackColor = &H8080FF
    cmdType3.BackColor = &H8080FF
    
    qtyNo(orderNo) = 1
    amtSize = 0
    amtAddOns = 0
    selectType = ""
    selectFlavor = ""
    selectSugar = ""
    selectSize = ""
    selectIce = ""
    
    Call update
End Sub

Private Sub resetFlavors()
    optFlavor1.Value = False
    optFlavor2.Value = False
    optFlavor3.Value = False
    optFlavor4.Value = False
    optFlavor5.Value = False
    optFlavor6.Value = False
    optFlavor7.Value = False
    optFlavor8.Value = False
    optFlavor9.Value = False
    optFlavor10.Value = False
    optFlavor11.Value = False
    optFlavor12.Value = False
    optFlavor13.Value = False
    optFlavor14.Value = False
    optFlavor15.Value = False
    optFlavor16.Value = False
    optFlavor17.Value = False
    optFlavor18.Value = False
    optFlavor19.Value = False
    optFlavor20.Value = False
    optFlavor21.Value = False
    optFlavor22.Value = False
    optFlavor23.Value = False
    optFlavor24.Value = False
    optFlavor25.Value = False
    optFlavor26.Value = False
    optFlavor27.Value = False
    optFlavor28.Value = False
    optFlavor29.Value = False
    optFlavor30.Value = False
End Sub

Private Sub disableFlavors()
    optFlavor1.Enabled = False
    optFlavor2.Enabled = False
    optFlavor3.Enabled = False
    optFlavor4.Enabled = False
    optFlavor5.Enabled = False
    optFlavor6.Enabled = False
    optFlavor7.Enabled = False
    optFlavor8.Enabled = False
    optFlavor9.Enabled = False
    optFlavor10.Enabled = False
    optFlavor11.Enabled = False
    optFlavor12.Enabled = False
    optFlavor13.Enabled = False
    optFlavor14.Enabled = False
    optFlavor15.Enabled = False
    optFlavor16.Enabled = False
    optFlavor17.Enabled = False
    optFlavor18.Enabled = False
    optFlavor19.Enabled = False
    optFlavor20.Enabled = False
    optFlavor21.Enabled = False
    optFlavor22.Enabled = False
    optFlavor23.Enabled = False
    optFlavor24.Enabled = False
    optFlavor25.Enabled = False
    optFlavor26.Enabled = False
    optFlavor27.Enabled = False
    optFlavor28.Enabled = False
    optFlavor29.Enabled = False
    optFlavor30.Enabled = False
End Sub

Private Sub totalAmt()
    Dim i As Integer
    i = 0
    total = 0
    
    Do While i < lstOrder.ListCount
         total = total + subtotal(i)
         i = i + 1
    Loop
End Sub

Private Sub cmdAddOrder_Click()
    If orderNo = 5 Then
        MsgBox "Only 5 kinds of orders per customer!", vbCritical, "Error"
    ElseIf selectType = "" Or selectFlavor = "" Or selectSugar = "" Or selectSize = "" Or selectIce = "" Then
        MsgBox "Please complete all forms!", vbCritical, "Error"
    Else
        lstOrder.AddItem orderReceipt(orderNo) & " @ " & Format(subtotal(orderNo), "0.00")
        orderMsg(orderNo) = lblOrder.Caption
        qtyTotal = qtyTotal + qtyNo(orderNo)
        orderNo = orderNo + 1
        Call totalAmt
        Call reset
        Call update
        'Call totalAmt
        'Call reset
    End If
End Sub

Private Sub cmdShowOrder_Click()
    If lstOrder.ListIndex <> -1 Then
        MsgBox orderMsg(lstOrder.ListIndex), vbOKOnly, "Order Summary"
    Else
        MsgBox "No order is selected!", vbCritical, "Error"
    End If
End Sub

Private Sub cmdDeleteOrder_Click()
    Dim i As Integer
    i = lstOrder.ListIndex
    
    If lstOrder.ListIndex <> -1 Then
        lstOrder.RemoveItem i
        orderNo = orderNo - 1
        qtyTotal = qtyTotal - qtyNo(i)
        orderMsg(i) = orderMsg(i + 1)
        Do While i < lstOrder.ListCount
            orderReceipt(i) = orderReceipt(i + 1)
            subtotal(i) = subtotal(i + 1)
            qtyNo(i) = qtyNo(i + 1)
            i = i + 1
        Loop
    Else
        MsgBox "No order is selected!", vbCritical, "Error"
    End If
    
    Call update
    Call totalAmt
    Call reset
End Sub

Private Sub cmdClearOrder_Click()
    If lstOrder.ListCount = 0 Then
        MsgBox "No order(s) in the list!", vbCritical, "Error"
    Else
        qtyTotal = 0
        If lstOrder.ListCount > 0 Then
            Do While lstOrder.ListCount > 0
                lstOrder.RemoveItem ListIndex
                orderNo = 0
            Loop
        End If
    End If

    Call update
    Call totalAmt
    Call reset
End Sub

Private Sub cmdAddQty_Click()
    qtyNo(orderNo) = qtyNo(orderNo) + 1
    cmdSubQty.Enabled = True
    Call update
End Sub

Private Sub cmdSubQty_Click()
    If qtyNo(orderNo) > 1 Then
        qtyNo(orderNo) = qtyNo(orderNo) - 1
    End If
    Call update
End Sub

Private Sub Timer1_Timer()
    lblDateTime.Caption = Format(Date, "Long Date") & "   " & Time
End Sub

Private Sub cmdCustInfo_Click()
    Dim ans As String
    
    If lstOrder.ListCount = 0 Then
        MsgBox "No order has been made!", vbCritical, "Error"
    ElseIf optPayment1.Value = False And optPayment2.Value = False And optPayment3.Value = False Then
        MsgBox "Please select the service type first!", vbCritical, "Error"
    Else
        ans = MsgBox("Proceed to Payment?", vbYesNo, "Confirm Order")
        If ans = vbYes Then
            Call receipt
            frmContent.Hide
            frmCustInfo.Show
        Else
            frmContent.Show
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Dim ans As String
    
    If isCashier = 1 Then
        ans = MsgBox("Are you sure you want to logout?", vbYesNo, "Logout")
        If ans = vbYes Then
            frmLogin.ADO2.Recordset.Fields("TimeOut") = Time
            frmLogin.ADO2.Recordset.update
            
            Unload frmLogin
            frmLogin.Show
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

'FLAVORS
Private Sub cmdType1_Click()
    Call resetAddons
    
    Call disableFlavors
    optFlavor1.Enabled = True
    optFlavor2.Enabled = True
    optFlavor3.Enabled = True
    optFlavor4.Enabled = True
    optFlavor5.Enabled = True
    optFlavor6.Enabled = True
    optFlavor7.Enabled = True
    optFlavor8.Enabled = True
    optFlavor9.Enabled = True
    optFlavor10.Enabled = True
    
    Call resetFlavors
    selectType = "Hot"
    initType = "Ho"
    selectFlavor = ""
    optIce1.Enabled = False
    optIce2.Enabled = False
    optIce3.Value = True
    chkAdd1.Enabled = False
    chkAdd2.Enabled = False
    chkAdd3.Enabled = False
    chkAdd4.Enabled = False
    chkAdd5.Enabled = False
    cmdType1.BackColor = &HC0C0FF
    cmdType2.BackColor = &H8080FF
    cmdType3.BackColor = &H8080FF
    
    Call update
End Sub

Private Sub cmdType2_Click()
    selectIce = tempIce
    If selectIce = optIce1.Caption Then
        optIce1.Value = True
        optIce2.Value = False
        optIce3.Value = False
    ElseIf selectIce = optIce2.Caption Then
        optIce1.Value = False
        optIce2.Value = True
        optIce3.Value = False
    ElseIf selectIce = optIce3.Caption Then
        optIce1.Value = False
        optIce2.Value = False
        optIce3.Value = True
    Else
        selectIce = ""
        optIce3.Value = False
    End If
    
    Call disableFlavors
    optFlavor11.Enabled = True
    optFlavor12.Enabled = True
    optFlavor13.Enabled = True
    optFlavor14.Enabled = True
    optFlavor15.Enabled = True
    optFlavor16.Enabled = True
    optFlavor17.Enabled = True
    optFlavor18.Enabled = True
    optFlavor19.Enabled = True
    optFlavor20.Enabled = True
    
    Call resetFlavors
    selectType = "Frappe"
    initType = "Fr"
    selectFlavor = ""
    optIce1.Enabled = True
    optIce2.Enabled = True
    optIce3.Enabled = True
    chkAdd1.Enabled = True
    chkAdd2.Enabled = True
    chkAdd3.Enabled = True
    chkAdd4.Enabled = True
    chkAdd5.Enabled = True
    cmdType1.BackColor = &H8080FF
    cmdType2.BackColor = &HC0C0FF
    cmdType3.BackColor = &H8080FF
    
    Call update
End Sub

Private Sub cmdType3_Click()
    selectIce = tempIce
    If selectIce = optIce1.Caption Then
        optIce1.Value = True
        optIce2.Value = False
        optIce3.Value = False
    ElseIf selectIce = optIce2.Caption Then
        optIce1.Value = False
        optIce2.Value = True
        optIce3.Value = False
    ElseIf selectIce = optIce3.Caption Then
        optIce1.Value = False
        optIce2.Value = False
        optIce3.Value = True
    Else
        selectIce = ""
        optIce3.Value = False
    End If
    
    Call disableFlavors
    optFlavor21.Enabled = True
    optFlavor22.Enabled = True
    optFlavor23.Enabled = True
    optFlavor24.Enabled = True
    optFlavor25.Enabled = True
    optFlavor26.Enabled = True
    optFlavor27.Enabled = True
    optFlavor28.Enabled = True
    optFlavor29.Enabled = True
    optFlavor30.Enabled = True
    
    Call resetFlavors
    selectType = "Milk Tea"
    initType = "MT"
    selectFlavor = ""
    optIce1.Enabled = True
    optIce2.Enabled = True
    optIce3.Enabled = True
    chkAdd1.Enabled = True
    chkAdd2.Enabled = True
    chkAdd3.Enabled = True
    chkAdd4.Enabled = True
    chkAdd5.Enabled = True
    
    cmdType1.BackColor = &H8080FF
    cmdType2.BackColor = &H8080FF
    cmdType3.BackColor = &HC0C0FF
    
    Call update
End Sub

Private Sub optFlavor1_Click()
    selectFlavor = optFlavor1.Caption
    Call update
End Sub

Private Sub optFlavor2_Click()
    selectFlavor = optFlavor2.Caption
    Call update
End Sub

Private Sub optFlavor3_Click()
    selectFlavor = optFlavor3.Caption
    Call update
End Sub

Private Sub optFlavor4_Click()
    selectFlavor = optFlavor4.Caption
    Call update
End Sub

Private Sub optFlavor5_Click()
    selectFlavor = optFlavor5.Caption
    Call update
End Sub

Private Sub optFlavor6_Click()
    selectFlavor = optFlavor6.Caption
    Call update
End Sub

Private Sub optFlavor7_Click()
    selectFlavor = optFlavor7.Caption
    Call update
End Sub

Private Sub optFlavor8_Click()
    selectFlavor = optFlavor8.Caption
    Call update
End Sub

Private Sub optFlavor9_Click()
    selectFlavor = optFlavor9.Caption
    Call update
End Sub

Private Sub optFlavor10_Click()
    selectFlavor = optFlavor10.Caption
    Call update
End Sub

Private Sub optFlavor11_Click()
    selectFlavor = optFlavor11.Caption
    Call update
End Sub

Private Sub optFlavor12_Click()
    selectFlavor = optFlavor12.Caption
    Call update
End Sub

Private Sub optFlavor13_Click()
    selectFlavor = optFlavor13.Caption
    Call update
End Sub

Private Sub optFlavor14_Click()
    selectFlavor = optFlavor14.Caption
    Call update
End Sub

Private Sub optFlavor15_Click()
    selectFlavor = optFlavor15.Caption
    Call update
End Sub

Private Sub optFlavor16_Click()
    selectFlavor = optFlavor16.Caption
    Call update
End Sub

Private Sub optFlavor17_Click()
    selectFlavor = optFlavor17.Caption
    Call update
End Sub

Private Sub optFlavor18_Click()
    selectFlavor = optFlavor18.Caption
    Call update
End Sub

Private Sub optFlavor19_Click()
    selectFlavor = optFlavor19.Caption
    Call update
End Sub

Private Sub optFlavor20_Click()
    selectFlavor = optFlavor20.Caption
    Call update
End Sub

Private Sub optFlavor21_Click()
    selectFlavor = optFlavor21.Caption
    Call update
End Sub

Private Sub optFlavor22_Click()
    selectFlavor = optFlavor22.Caption
    Call update
End Sub

Private Sub optFlavor23_Click()
    selectFlavor = optFlavor23.Caption
    Call update
End Sub

Private Sub optFlavor24_Click()
    selectFlavor = optFlavor24.Caption
    Call update
End Sub

Private Sub optFlavor25_Click()
    selectFlavor = optFlavor25.Caption
    Call update
End Sub

Private Sub optFlavor26_Click()
    selectFlavor = optFlavor26.Caption
    Call update
End Sub

Private Sub optFlavor27_Click()
    selectFlavor = optFlavor27.Caption
    Call update
End Sub

Private Sub optFlavor28_Click()
    selectFlavor = optFlavor28.Caption
    Call update
End Sub

Private Sub optFlavor29_Click()
    selectFlavor = optFlavor29.Caption
    Call update
End Sub

Private Sub optFlavor30_Click()
    selectFlavor = optFlavor30.Caption
    Call update
End Sub

'END FLAVORS

'ADD-ONS
Private Sub chkAdd1_Click()
    If chkAdd1.Value = 1 Then
        add1 = "    " & "Egg Pudding" & vbNewLine
        initAdd1 = "1"
        amtAddOns = amtAddOns + 25
    Else
        add1 = ""
        initAdd1 = "0"
        amtAddOns = amtAddOns - 25
    End If

    Call update
End Sub

Private Sub chkAdd2_Click()
    If chkAdd2.Value = 1 Then
        add2 = "    " & "Grass Jelly" & vbNewLine
        initAdd2 = "1"
        amtAddOns = amtAddOns + 15
    Else
        add2 = ""
        initAdd2 = "0"
        amtAddOns = amtAddOns - 15
    End If
    
    Call update
End Sub

Private Sub chkAdd3_Click()
    If chkAdd3.Value = 1 Then
        add3 = "    " & "Ice Cream" & vbNewLine
        initAdd3 = "1"
        amtAddOns = amtAddOns + 20
    Else
        add3 = ""
        initAdd3 = "0"
        amtAddOns = amtAddOns - 20
    End If
    
    Call update
End Sub

Private Sub chkAdd4_Click()
    If chkAdd4.Value = 1 Then
        add4 = "    " & "Mini Mochi" & vbNewLine
        initAdd4 = "1"
        amtAddOns = amtAddOns + 15
    Else
        add4 = ""
        initAdd4 = "0"
        amtAddOns = amtAddOns - 15
    End If
    
    Call update
End Sub

Private Sub chkAdd5_Click()
    If chkAdd5.Value = 1 Then
        add5 = "    " & "Red Bean" & vbNewLine
        initAdd5 = "1"
        amtAddOns = amtAddOns + 10
    Else
        add5 = ""
        initAdd5 = "0"
        amtAddOns = amtAddOns - 10
    End If
    
    Call update
End Sub
'END ADD-ONS

'SUGAR
Private Sub optSugar1_Click()
    selectSugar = optSugar1.Caption
    Call update
End Sub

Private Sub optSugar2_Click()
    selectSugar = optSugar2.Caption
    Call update
End Sub

Private Sub optSugar3_Click()
    selectSugar = optSugar3.Caption
    Call update
End Sub

Private Sub optSugar4_Click()
    selectSugar = optSugar4.Caption
    Call update
End Sub

Private Sub optSugar5_Click()
    selectSugar = optSugar5.Caption
    Call update
End Sub
'END SUGAR

'SIZE
Private Sub optSize1_Click()
    selectSize = "Venti"
    initSize = "Ve"
    amtSize = 165
    Call update
End Sub

Private Sub optSize2_Click()
    selectSize = "Grande"
    initSize = "Gr"
    amtSize = 150
    Call update
End Sub

Private Sub optSize3_Click()
    selectSize = "Tall"
    initSize = "Ta"
    amtSize = 135
    Call update
End Sub

Private Sub optSize4_Click()
    selectSize = "Short"
    initSize = "Sh"
    amtSize = 120
    Call update
End Sub
'END SIZE

'ICE
Private Sub optIce1_Click()
    selectIce = optIce1.Caption
    tempIce = selectIce
    initIce = "Re"
    Call update
End Sub

Private Sub optIce2_Click()
    selectIce = optIce2.Caption
    tempIce = selectIce
    initIce = "Le"
    Call update
End Sub

Private Sub optIce3_Click()
    selectIce = optIce3.Caption
    If cmdType1.Value = False Then
        tempIce = selectIce
    End If
    initIce = "No"
    Call update
End Sub
'END ICE

'PAYMENT
Private Sub optPayment1_Click()
    selectPayment = optPayment1.Caption
End Sub

Private Sub optPayment2_Click()
    selectPayment = optPayment2.Caption
End Sub

Private Sub optPayment3_Click()
    selectPayment = optPayment3.Caption
End Sub
'END PAYMENT
