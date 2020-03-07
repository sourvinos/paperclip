VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form EmployeesTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14355
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   30
      Top             =   5775
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "ƒÁÏÈÔıÒ„ﬂ·"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   5
         Left            =   7350
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   " ÎÂﬂÛÈÏÔ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "¡ÔËﬁÍÂıÛÁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "ƒÈ·„Ò·ˆﬁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "≈˝ÒÂÛÁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   4
         Left            =   5925
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "¡ÍıÒÔ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   9675
      TabIndex        =   15
      Top             =   375
      Width           =   4515
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "EmployeesTransactions.PaymentCategoryID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox txtPaymentCategoryID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtPaymentWayID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox txtEmployeeID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "EmployeesTransactions.EmployeeID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "EmployeesTransactions.ID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "EmployeesTransactions.PaymentWayID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "EmployeesTransactions.TransactionTypeID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox txtTransactionTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1950
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   2296
         Images          =   "EmployeesTransactions.frx":0000
         Version         =   131072
         KeyCount        =   2
         Keys            =   "ˇ"
      End
   End
   Begin UserControls.newDate mskDate 
      Height          =   465
      Left            =   2250
      TabIndex        =   0
      Top             =   1125
      Width           =   1455
      _ExtentX        =   2672
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "01/01/2017"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtEmployeeLastname 
      Height          =   465
      Left            =   2250
      TabIndex        =   2
      Top             =   2175
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   100
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtPaymentCategoryDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   3
      Top             =   2700
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtRemarks 
      Height          =   465
      Left            =   2250
      TabIndex        =   4
      Top             =   3225
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   7275
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":0918
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7275
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2700
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":0EB2
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7725
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":144C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7725
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2700
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":19E6
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newFloat mskAmount 
      Height          =   465
      Left            =   2250
      TabIndex        =   5
      Top             =   3750
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   820
      Alignment       =   1
      ForeColor       =   0
      Text            =   "99.999,99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   7725
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4275
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":1F80
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtPaymentWayDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   6
      Top             =   4275
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7275
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4275
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":251A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtTransactionTypeDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   7
      Top             =   4800
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   6
      Left            =   7275
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4800
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "EmployeesTransactions.frx":2AB4
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newDate mskDateRefersTo 
      Height          =   465
      Left            =   2250
      TabIndex        =   1
      Top             =   1650
      Width           =   1455
      _ExtentX        =   2672
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "01/01/2017"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "¡ˆÔÒ‹"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   450
      TabIndex        =   41
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "‘˝ÔÚ ÛıÌ·ÎÎ·„ﬁÚ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   450
      TabIndex        =   39
      Top             =   4875
      Width           =   1365
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   9
      Left            =   0
      Top             =   9525
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   3975
      Top             =   5250
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   2250
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   2250
      Top             =   6450
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   5175
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   0
      Left            =   1800
      Top             =   2025
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   2850
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   11
      Left            =   450
      TabIndex        =   14
      Top             =   4350
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "–ÔÛ¸"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   450
      TabIndex        =   13
      Top             =   3825
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   " ÈÌﬁÛÂÈÚ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   720
      Left            =   225
      TabIndex        =   12
      Top             =   75
      Width           =   1980
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   " ·ÙÁ„ÔÒﬂ· ·ÏÔÈ‚ﬁÚ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   450
      TabIndex        =   11
      Top             =   2775
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "¡ˆÔÒ‹"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   450
      TabIndex        =   10
      Top             =   3300
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "œÌÔÏ·ÙÂ˛ÌıÏÔ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   450
      TabIndex        =   9
      Top             =   2250
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "«ÏÂÒÔÏÁÌﬂ·"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   450
      TabIndex        =   8
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "EmployeesTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim blnCancel As Boolean
Dim blnPrinterHasBeenSelected As Boolean
Dim lngTrnID As Long
Dim IsError As Boolean

Private Sub AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            blnCancel = True
            ClearFields txtID, txtEmployeeID, txtPaymentCategoryID, txtPaymentWayID, txtTransactionTypeID
            ClearFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
            DisableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
            UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
        End If
        Exit Sub
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Sub

Private Sub DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "EmployeesTransactions", strApplicationName, "ID", txtID.text, True) Then
        blnCancel = True
        ClearFields txtID, txtEmployeeID, txtPaymentCategoryID, txtPaymentWayID, txtTransactionTypeID
        ClearFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
        DisableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    End If
    
End Sub

Private Sub NewRecord()
    
    Dim tmpRecordset As Recordset
    
    blnStatus = True
    blnCancel = False
    
    ClearFields txtID, txtEmployeeID, txtPaymentCategoryID, txtPaymentWayID, txtTransactionTypeID
    ClearFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
    EnableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    
    mskDate.SetFocus
    
End Sub

Private Function PopulateFields(rstRecordset As Recordset)

    With rstRecordset
    
        txtID.text = !ID
        txtEmployeeID.text = !employeeID
        txtPaymentCategoryID.text = !PaymentCategoryID
        txtPaymentWayID.text = !PaymentWayID
        txtTransactionTypeID.text = !TransactionTypeID
        
        mskDate.text = format(!Date, "dd/mm/yyyy")
        mskDateRefersTo.text = format(!DateRefersTo, "dd/mm/yyyy")
        txtEmployeeLastname.text = !Lastname & " " & !Firstname
        mskAmount.text = format(!amount, "#,##0.00")
        txtPaymentCategoryDescription.text = !PaymentCategoryDescription
        txtRemarks.text = !Remarks
        txtPaymentWayDescription.text = !Description
        txtTransactionTypeDescription.text = !TransactionTypeDescription
        
    End With

End Function

Private Function SaveInvoices()

    If MainSaveRecord("CommonDB", "EmployeesTransactions", blnStatus, strApplicationName, "ID", _
        txtID, _
        txtEmployeeID, _
        txtPaymentCategoryID, _
        txtPaymentWayID, _
        txtTransactionTypeID, _
        "1", _
        strCurrentUser) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "EmployeesTransactions", blnStatus, strApplicationName, "ID", _
            txtID.text, _
            mskDate.text, _
            mskDateRefersTo.text, _
            txtEmployeeID.text, _
            txtPaymentCategoryID.text, _
            txtRemarks.text, _
            mskAmount.text, _
            txtPaymentWayID.text, _
            txtTransactionTypeID.text, _
            1, _
            strCurrentUser) <> 0 Then
        ClearFields txtID, txtEmployeeID, txtPaymentCategoryID, txtPaymentWayID, txtTransactionTypeID
        ClearFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
        DisableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
        UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    '«ÏÂÒÔÏÁÌﬂ·
    If Not CheckDate(mskDate.text, strApplicationName) Then
        mskDate.SetFocus
        Exit Function
    End If
    
    '¡ˆÔÒ‹ ÁÏÂÒÔÏÁÌﬂ·
    If Not CheckDate(mskDateRefersTo.text, strApplicationName) Then
        mskDateRefersTo.SetFocus
        Exit Function
    End If
    
    '≈Ò„·Ê¸ÏÂÌÔÚ
    If Len(txtEmployeeID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtEmployeeLastname.SetFocus
        Exit Function
    End If
    
    ' ·ÙÁ„ÔÒﬂ· ·ÏÔÈ‚ﬁÚ
    If Len(txtPaymentCategoryID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPaymentCategoryDescription.SetFocus
        Exit Function
    End If
    
    '‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ
    If Len(txtPaymentWayID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPaymentWayDescription.SetFocus
        Exit Function
    End If
    
    '‘˝ÔÚ ÛıÌ·ÎÎ·„ﬁÚ
    If Len(txtTransactionTypeID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtTransactionTypeDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Sub cmdButton_Click(Index As Integer)

    Dim arrDummy()
    
    Select Case Index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            FindRecords
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Function FindRecords()

    With EmployeesTransactionsIndex
        .Tag = "True"
        .Show 1, Me
    End With

End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            '≈Ò„·Ê¸ÏÂÌÔÚ
            Set tmpRecordset = CheckForMatch("CommonDB", "Employees", "Lastname", "String", txtEmployeeLastname.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "≈˛ÌıÏÔ", "œÌÔÏ·", 0, 40, 40, 1, 0, 0)
                txtEmployeeID.text = tmpTableData.strCode
                txtEmployeeLastname.text = tmpTableData.strFirstField & " " & tmpTableData.strSecondField
            End If
        Case 1
            '≈Ò„·Ê¸ÏÂÌÔÚ
            With Employees
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            ' ·ÙÁ„ÔÒﬂ· ·ÏÔÈ‚ﬁÚ
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentCategories", "Description", "String", txtPaymentCategoryDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "–ÂÒÈ„Ò·ˆﬁ", 0, 40, 1, 0)
                txtPaymentCategoryID.text = tmpTableData.strCode
                txtPaymentCategoryDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            ' ·ÙÁ„ÔÒﬂ· ·ÏÔÈ‚ﬁÚ
            With TablesPaymentCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            '‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentWays", "Description", "String", txtPaymentWayDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "–ÂÒÈ„Ò·ˆﬁ", 0, 40, 1, 0)
                txtPaymentWayID.text = tmpTableData.strCode
                txtPaymentWayDescription.text = tmpTableData.strFirstField
            End If
        Case 5
            ' ·ÙÁ„ÔÒﬂ· ·ÏÔÈ‚ﬁÚ
            With TablesPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case 6
            '‘˝ÔÚ ÛıÌ·ÎÎ·„ﬁÚ
            Set tmpRecordset = CheckForMatch("CommonDB", "TransactionTypes", "Description", "String", txtTransactionTypeDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "–ÂÒÈ„Ò·ˆﬁ", 0, 40, 1, 0)
                txtTransactionTypeID.text = tmpTableData.strCode
                txtTransactionTypeDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Public Function SeekRecord(lngID)

    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    strSQL = "SELECT " _
        & "EmployeesTransactions.ID, EmployeeID, PaymentCategoryID, EmployeesTransactions.Remarks, EmployeesTransactions.PaymentWayID, EmployeesTransactions.TransactionTypeID, Amount, Date, DateRefersTo, " _
        & "Lastname, Firstname, " _
        & "PaymentCategories.Description AS PaymentCategoryDescription, " _
        & "PaymentWays.Description, " _
        & "TransactionTypes.Description AS TransactionTypeDescription " _
        & "FROM ((((EmployeesTransactions " _
        & "INNER JOIN Employees ON EmployeesTransactions.EmployeeID = Employees.ID) " _
        & "INNER JOIN PaymentCategories ON EmployeesTransactions.PaymentCategoryID = PaymentCategories.ID) " _
        & "INNER JOIN PaymentWays ON EmployeesTransactions.PaymentWayID = PaymentWays.ID) " _
        & "INNER JOIN TransactionTypes ON EmployeesTransactions.TransactionTypeID = TransactionTypes.ID) "
        
    'EmployessTransactions.ID
    strThisParameter = "lngID long"
    strThisQuery = "EmployeesTransactions.ID = lngID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngID

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set SeekRecord = rstRecordset
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return

End Function

Public Function DoPostFoundJobs(rstRecordset As Recordset)

    On Error GoTo ErrTrap

    blnStatus = False
    
    PopulateFields rstRecordset
    EnableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
        
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function




Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF7 And cmdButton(3).Enabled, vbKeyF And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False
    ColorizeControls Me, False, False
    blnCancel = True
    ClearFields txtID, txtEmployeeID, txtPaymentCategoryID, txtPaymentWayID, txtTransactionTypeID
    ClearFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
    DisableFields mskDate, mskDateRefersTo, txtEmployeeLastname, txtPaymentCategoryDescription, txtRemarks, mskAmount, txtPaymentWayDescription, txtTransactionTypeDescription
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6)
    UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1

End Sub

Private Sub txtEmployeeLastname_Change()
    
    If txtEmployeeLastname.text = "" Then ClearFields txtEmployeeID
    
End Sub


Private Sub txtEmployeeLastname_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub


Private Sub txtEmployeeLastname_Validate(Cancel As Boolean)

    If txtEmployeeID.text = "" And txtEmployeeLastname.text <> "" Then cmdIndex_Click 0: If txtEmployeeID.text = "" Then Cancel = True
    
End Sub


Private Sub txtPaymentCategoryDescription_Change()

    If txtPaymentCategoryDescription.text = "" Then ClearFields txtPaymentCategoryID
    
End Sub


Private Sub txtPaymentCategoryDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3

End Sub


Private Sub txtPaymentCategoryDescription_Validate(Cancel As Boolean)

    If txtPaymentCategoryID.text = "" And txtPaymentCategoryDescription.text <> "" Then cmdIndex_Click 2: If txtPaymentCategoryID.text = "" Then Cancel = True

End Sub


Private Sub txtPaymentWayDescription_Change()

    If txtPaymentWayDescription.text = "" Then ClearFields txtPaymentWayID

End Sub


Private Sub txtPaymentWayDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    If KeyCode = vbKeyF5 Then cmdIndex_Click 5

End Sub


Private Sub txtPaymentWayDescription_Validate(Cancel As Boolean)

    If txtPaymentWayID.text = "" And txtPaymentWayDescription.text <> "" Then cmdIndex_Click 4: If txtPaymentWayID.text = "" Then Cancel = True
    
End Sub


Private Sub txtTransactionTypeDescription_Change()

    If txtTransactionTypeDescription.text = "" Then ClearFields txtTransactionTypeID

End Sub


Private Sub txtTransactionTypeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6
    
End Sub


Private Sub txtTransactionTypeDescription_Validate(Cancel As Boolean)

    If txtTransactionTypeID.text = "" And txtTransactionTypeDescription.text <> "" Then cmdIndex_Click 6: If txtTransactionTypeID.text = "" Then Cancel = True

End Sub


