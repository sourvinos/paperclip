VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form Employees 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12225
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   1
      Left            =   2625
      TabIndex        =   40
      Top             =   6000
      Width           =   8940
      Begin iGrid300_10Tec.iGrid grdAgreements 
         Height          =   3650
         Left            =   450
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   450
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6429
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483631
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   6
         Left            =   1725
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4350
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   820
         BackColor       =   15133676
         ButtonShape     =   3
         ButtonStyle     =   7
         Caption         =   "쾅一諪蠢塞 憶蜃囚詞"
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
         Height          =   465
         Index           =   7
         Left            =   4500
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4350
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   820
         BackColor       =   15133676
         ButtonShape     =   3
         ButtonStyle     =   7
         Caption         =   "쾡率畯蟄 憶蜃囚詞"
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
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   4
         Left            =   0
         Top             =   3300
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   3
         Left            =   8475
         Top             =   2025
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   2
         Left            =   2550
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   1
         Left            =   0
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmFrame 
      BorderStyle     =   0  'None
      Height          =   5040
      Index           =   0
      Left            =   1875
      TabIndex        =   25
      Top             =   1125
      Width           =   8940
      Begin UserControls.newText txtLastname 
         Height          =   465
         Left            =   2625
         TabIndex        =   0
         Top             =   450
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
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
      Begin UserControls.newText txtShipDescription 
         Height          =   465
         Left            =   2625
         TabIndex        =   5
         Top             =   3075
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   40
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
      Begin UserControls.newText txtSpecialityDescription 
         Height          =   465
         Left            =   2625
         TabIndex        =   4
         Top             =   2550
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   100
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
      Begin UserControls.newText txtFirstname 
         Height          =   465
         Left            =   2625
         TabIndex        =   1
         Top             =   975
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   100
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
      Begin UserControls.newText txtCompanyDescription 
         Height          =   465
         Left            =   2625
         TabIndex        =   2
         Top             =   1500
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   40
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
      Begin UserControls.newText txtPhones 
         Height          =   465
         Left            =   2625
         TabIndex        =   6
         Top             =   3600
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   100
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
         Left            =   7650
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1500
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
         PicNormal       =   "Employees.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   8100
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1500
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
         PicNormal       =   "Employees.frx":059A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtRemarks 
         Height          =   465
         Left            =   2625
         TabIndex        =   7
         Top             =   4125
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   820
         ForeColor       =   4194304
         MaxLength       =   100
         Text            =   "좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋좋"
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
         Index           =   2
         Left            =   7650
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2550
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
         PicNormal       =   "Employees.frx":0B34
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   3
         Left            =   8100
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2550
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
         PicNormal       =   "Employees.frx":10CE
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   4
         Left            =   7650
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3075
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
         PicNormal       =   "Employees.frx":1668
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   5
         Left            =   8100
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3075
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
         PicNormal       =   "Employees.frx":1C02
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newDate mskHireDate 
         Height          =   465
         Left            =   2625
         TabIndex        =   3
         Top             =   2025
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
      Begin VB.Shape Shape2 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   2
         Left            =   8475
         Top             =   1950
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   1
         Left            =   3375
         Top             =   4575
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   0
         Left            =   3450
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "퇴�臧逸"
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
         TabIndex        =   39
         Top             =   525
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "豆誾柝奬"
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
         TabIndex        =   38
         Top             =   3675
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "伎楨�"
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
         TabIndex        =   37
         Top             =   3150
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "헤嚴睛午塞 猝宖揖逋�"
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
         TabIndex        =   36
         Top             =   2100
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "톱昻流晴疊"
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
         TabIndex        =   35
         Top             =   2625
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "權睛�"
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
         Index           =   6
         Left            =   450
         TabIndex        =   34
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "툐誦濬�"
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
         Index           =   7
         Left            =   450
         TabIndex        =   33
         Top             =   1575
         Width           =   1740
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   0
         Left            =   2175
         Top             =   225
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   0
         Left            =   0
         Top             =   2400
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "急畯晴準暢煜"
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
         TabIndex        =   32
         Top             =   4200
         Width           =   1740
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   300
      TabIndex        =   16
      Top             =   6675
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "쾅一諪蠢塞"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   5
         Left            =   7350
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "勘馭猖逸"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "즌炡絲嶪滄"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "쾡率畯蟄"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "퉈遵滄"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   4
         Left            =   5925
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "쥬椿�"
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
         PicOpacity      =   0
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Customer"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   6975
      TabIndex        =   9
      Top             =   6150
      Width           =   4515
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "ShipID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtShipID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
      Begin VB.TextBox txtSpecialityID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "SpecialityID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox txtCompanyID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "CompanyID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "ID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1575
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "Employees.frx":219C
         Version         =   131072
         KeyCount        =   4
         Keys            =   "���"
      End
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   0
      Left            =   450
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "導玎特塞"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388736
      PicOpacity      =   0
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   1
      Left            =   450
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "屠恁蛤牲�"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388736
      PicOpacity      =   0
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   990
      Index           =   1
      Left            =   825
      Top             =   2175
      Width           =   1815
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   990
      Index           =   0
      Left            =   825
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   2
      Left            =   8475
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
      Index           =   1
      Left            =   0
      Top             =   5475
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "퇸信琰靭獐�"
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
      TabIndex        =   8
      Top             =   75
      Width           =   3015
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10800
      Top             =   4575
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3450
      Top             =   7350
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
      Left            =   1500
      Top             =   0
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
      Top             =   2100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   -75
      Top             =   0
      Width           =   840
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnu즌炡絲嶪滄伎不諪瑨晴誼� 
         Caption         =   "즌炡絲嶪滄 足不諪� 彩俉��"
      End
   End
End
Attribute VB_Name = "Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim IsError As Boolean

Private Function AbortProcedure(blnStatus)
    
    If grdAgreements.TextEditText <> "" Then
        grdAgreements.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            btnPanel_Click 0
            ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
            ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
            ClearFields grdAgreements
            DisableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
            DisableFields btnPanel(1)
            DisableFields grdAgreements
            UpdateButtons Me, 7, 1, 0, 0, 1, 0, 1, 0, 0
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function AddGridLine()

    With grdAgreements
            .Enabled = True
            .AddRow
            .CellIcon(.RowCount, "Status") = lstIconList.ItemIndex(2)
            .SetCurCell .RowCount, 2
            .SetFocus
        End With

End Function

Private Function DeleteAgreements()

    Dim lngRow As Long
    
    If IsError Then Exit Function
    
    With grdAgreements
        For lngRow = 1 To .RowCount
             If Not MainDeleteRecord("CommonDB", "EmployeesAgreements", strApplicationName, "ID", .CellValue(lngRow, "ID"), False) Then
                IsError = True
                Exit For
             End If
        Next lngRow
    End With

End Function

Private Function DeleteEmployee()

    If Not MainDeleteRecord("CommonDB", "Employees", strApplicationName, "ID", txtID.text, True) Then
        IsError = True
    End If

End Function

Private Function FindAgreements(lngID)

    'On Error GoTo ErrTrap
    
    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    'Local 靭疊瞬娛毘
    Dim lngIndex As Long
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    Dim tmpRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    '鑒中� SQL
    strSQL = "SELECT ID, EmployeeID, DateFrom, DateTo, Remarks, Amount " _
        & "FROM EmployeesAgreements "

    '퇸信琰靭獐�
    strThisParameter = "lngID Long"
    strThisQuery = "EmployeeID = lngID "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngID
        
    '杜迹贓鱗滄
    strOrder = " ORDER BY DateFrom "
        
    '基艇王剿 疊 揄郁詐紆
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    'SQL
    TempQuery.SQL = strSQL & strOrder
    
    '戡郁詐紆
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    '쥼楨實 剃 recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    '쥼 鴨� 砒� 憶蜃囚毘, 栒損薔
    If rstRecordset.RecordCount = 0 Then Exit Function
    
    If grdAgreements.colCount = 0 Then
        AddColumnsToGrid grdAgreements, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdAgreements"), "04NCNID,04NCDFrom,04NCDTo,30NLNRemarks,04NRFAmount,05NCNStatus,05NCNDeleted", "ID,즌�,툼�,急畯晴準暢煜,器冊,�,�"
    End If
    
    '쬡殮 蜃刷燐� 彩� 足匪藺
    grdAgreements.AddRow , , , , , , , rstRecordset.RecordCount
    
    '伎匪藺
    With grdAgreements
        .Editable = True
        .RowMode = False
    End With
    
    '춧絪殮 剃 足匪藺
    With rstRecordset
        While Not .EOF
            With grdAgreements
                lngRow = lngRow + 1
                .CellValue(lngRow, "ID") = rstRecordset!id
                .CellValue(lngRow, "From") = rstRecordset!dateFrom
                .CellValue(lngRow, "To") = rstRecordset!DateTo
                .CellValue(lngRow, "Remarks") = rstRecordset!Remarks
                .CellValue(lngRow, "Amount") = rstRecordset!amount
            End With
            .MoveNext
        Wend
    End With
    
    '荳邑謂� 孼毗愼虞�
    FindAgreements = True
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    Return

ErrTrap:
    FindAgreements = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function GotoPreviousPanel(formName, intPanelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To formName.btnPanel.Count - 1
    
        If Not formName.btnPanel(intLoop).Enabled Then
            If intLoop - 1 >= 0 Then
                If formName.btnPanel(intLoop - 1).Enabled Then
                    btnPanel_Click intLoop - 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function


Private Function PositionPanels()

    Dim intLoop As Integer
    
    For intLoop = 0 To 1
        frmFrame(intLoop).Visible = False
    Next intLoop
    
    For intLoop = 0 To 1
        btnPanel(intLoop).Enabled = True
        shpBridge(intLoop).Visible = False
        With frmFrame(intLoop)
            .Height = 5040
            .Width = 8940
            .Left = 1875
            .Top = 1125
            .BackColor = &HE0E0E0
        End With
    Next intLoop
    
    btnPanel(0).Enabled = False
    frmFrame(0).Visible = True
    shpBridge(0).Visible = True
    
End Function

Public Function SeekRecord(myID)

    Dim blnEnableDelete As Boolean
    Dim tmpRecordset As Recordset
    Dim rstAgreements As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
    ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    ClearFields grdAgreements
    DisableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    DisableFields grdAgreements
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    
    'SeekRecord = False
    Dim rstRecordset As Recordset
    
    blnEnableDelete = SimpleSeek("EmployeesTransactions", "EmployeeID", myID)
    
    If MainSeekRecord("CommonDB", "Employees", "ID", myID, True, txtID, txtLastname, txtFirstName, txtCompanyID, mskHireDate, txtSpecialityID, txtShipID, txtPhones, txtRemarks) Then
        '툐誦濬�
        Set tmpRecordset = CheckForMatch("CommonDB", "Companies", "ID", "Numeric", txtCompanyID.text)
        txtCompanyID.text = tmpRecordset.Fields(0)
        txtCompanyDescription.text = tmpRecordset.Fields(1)
        '톱昻流晴疊
        Set tmpRecordset = CheckForMatch("CommonDB", "Specialities", "ID", "Numeric", txtSpecialityID.text)
        txtSpecialityID.text = tmpRecordset.Fields(0)
        txtSpecialityDescription.text = tmpRecordset.Fields(1)
        '伎楨�
        If txtShipID.text <> "0" Then
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipID", "Numeric", txtShipID.text)
            txtShipID.text = tmpRecordset.Fields(0)
            txtShipDescription.text = tmpRecordset.Fields(1)
        End If
        EnableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        EnableFields grdAgreements
        EnableFields btnPanel(1)
        UpdateButtons Me, 7, 0, 1, IIf(blnEnableDelete, 1, 0), 0, 1, 0, 1, 1
        blnStatus = False
        'SeekRecord = txtID.text
        Set SeekRecord = rstRecordset
        
        FindAgreements (Val(txtID.text))
        
    End If
    
End Function

Private Function DeleteRecord()
    
    IsError = False
    
    BeginTrans
    
    DeleteEmployee
    DeleteAgreements
    
    If Not IsError Then
        CommitTrans
        btnPanel_Click 0
        ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
        ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
        ClearFields grdAgreements
        DisableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        DisableFields btnPanel(1)
        DisableFields grdAgreements
        UpdateButtons Me, 7, 1, 0, 0, 1, 0, 1, 0, 0
    Else
        Rollback
    End If
    
End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
    ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    ClearFields grdAgreements
    EnableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    EnableFields btnPanel(1)
    EnableFields grdAgreements
    CustomizeGrid grdAgreements
    EditableFields grdAgreements
    UpdateButtons Me, 7, 0, 1, 0, 0, 1, 0, 1, 1
    txtLastname.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If txtShipID.text = "" Then txtShipID.text = "0"
    
    txtID.text = MainSaveRecord("CommonDB", "Employees", blnStatus, strApplicationName, "ID", _
        txtID.text, _
        txtLastname.text, _
        txtFirstName.text, _
        txtCompanyID.text, _
        mskHireDate.text, _
        txtSpecialityID.text, _
        txtShipID.text, _
        txtPhones.text, _
        txtRemarks.text, 1, strCurrentUser)
    
    If txtID.text <> "0" Then SaveAgreements txtID.text
        
    btnPanel_Click 0
    ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
    ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    ClearFields grdAgreements
    DisableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    DisableFields btnPanel(1)
    DisableFields grdAgreements
    UpdateButtons Me, 7, 1, 0, 0, 1, 0, 1, 0, 0
    
End Function

Private Function SaveAgreements(employeeID As Integer)

    Dim lngID As Long
    Dim lngRow As Long
    
    With grdAgreements
        For lngRow = 1 To .RowCount
            'Add Record when Status = Blue and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = 1) And (.CellIcon(lngRow, "Deleted") = -1) Then
                lngID = MainSaveRecord("CommonDB", "EmployeesAgreements", True, strApplicationName, "ID", 0, employeeID, .CellValue(lngRow, "From"), .CellValue(lngRow, "To"), .CellValue(lngRow, "Remarks"), .CellValue(lngRow, "Amount"))
            End If
            'Delete Existing Record when Status = Blank and Deleted = Red
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = 2) Then
                lngID = MainDeleteRecord("CommonDB", "EmployeesAgreements", strApplicationName, "ID", .CellValue(lngRow, "ID"), False)
            End If
            'Update Existing Record when Status = Blank and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = -1) Then
                lngID = MainSaveRecord("CommonDB", "EmployeesAgreements", False, strApplicationName, "ID", .CellValue(lngRow, "ID"), employeeID, .CellValue(lngRow, "From"), .CellValue(lngRow, "To"), .CellValue(lngRow, "Remarks"), .CellValue(lngRow, "Amount"))
            End If
        Next lngRow
    End With
    
    SaveAgreements = employeeID

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    '퇴�臧逸
    If Len(Trim(txtLastname.text)) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 0
        txtLastname.SetFocus
        Exit Function
    End If
    
    '툐誦濬�
    If Len(txtCompanyID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 0
        txtCompanyDescription.SetFocus
        Exit Function
    End If
    
    '헤嚴睛午塞 猝宖揖逋�
    If Not CheckDate(mskHireDate.text, strApplicationName) Then
        btnPanel_Click 0
        mskHireDate.SetFocus
        Exit Function
    End If
    
    '톱昻流晴疊
    If Len(txtSpecialityID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 0
        txtSpecialityDescription.SetFocus
        Exit Function
    End If
    
    '屠恁蛤牲�
    Dim lngRow As Long
    
    With grdAgreements
        For lngRow = 1 To .RowCount
            If Not IsDate(.CellValue(lngRow, "From")) Or Not IsDate(.CellValue(lngRow, "To")) Or .CellValue(lngRow, "Amount") = "" Then
                If MyMsgBox(4, strApplicationName, strAppMessages(1) & lngRow & " 馭奬� 殷寥�", 1) Then
                End If
                btnPanel_Click 1
                .SetCurCell lngRow, 2
                Exit Function
            End If
        Next lngRow
    End With
    
    ValidateFields = True

End Function

Private Sub btnPanel_Click(index As Integer)

    Dim intLoop As Integer
    
    For intLoop = 0 To 1
        btnPanel(intLoop).Enabled = True
        frmFrame(intLoop).Visible = False
        shpBridge(intLoop).Visible = False
    Next intLoop
    
    btnPanel(index).Enabled = False
    frmFrame(index).Visible = True
    shpBridge(index).Visible = True
    
    Select Case index
        '導玎特塞
        Case 0
            If cmdButton(1).Enabled Then
                If txtLastname.Enabled Then
                    txtLastname.SetFocus
                End If
            End If
        '屠恁蛤牲�
        Case 1
            If cmdButton(1).Enabled Then
                If grdAgreements.Enabled And grdAgreements.RowCount > 0 Then
                    With grdAgreements
                        .SetCurCell 1, 2
                        .SetFocus
                        .TabStop = True
                    End With
                End If
            End If
    End Select

End Sub

Private Sub cmdButton_Click(index As Integer)
                                                                                                                                
    Select Case index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            ShowIndex
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
        Case 6
            AddGridLine
        Case 7
            ToggleGridLineToDelete
    End Select

End Sub

Private Function ToggleGridLineToDelete()
    
    With grdAgreements
        If .RowCount > 0 And .CurRow > 0 Then
            If .CellValue(.CurRow, 1) <> "" Then
                .CellIcon(.CurRow, "Deleted") = IIf(.CellIcon(.CurRow, "Deleted") <= 0, lstIconList.ItemIndex(3), lstIconList.ItemIndex(1))
            End If
        End If
    End With
    
End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            '툐誦濬�
            Set tmpRecordset = CheckForMatch("CommonDB", "Companies", "Description", "String", txtCompanyDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "權睛修塞", 0, 40, 1, 0)
                txtCompanyID.text = tmpTableData.strCode
                txtCompanyDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            '툐誦濬�
            With TablesCompanies
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            '톱昻流晴疊
            Set tmpRecordset = CheckForMatch("CommonDB", "Specialities", "Description", "String", txtSpecialityDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "給中蜃囚�", 0, 40, 1, 0)
                txtSpecialityID.text = tmpTableData.strCode
                txtSpecialityDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            '톱昻流晴疊
            With TablesSpecialities
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            '伎楨�
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShipDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "給中蜃囚�", 0, 40, 1, 0)
                txtShipID.text = tmpTableData.strCode
                txtShipDescription.text = tmpTableData.strFirstField
            End If
        Case 5
            '伎楨�
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then Me.Tag = "False"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled And Not btnPanel(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled And Not btnPanel(0).Enabled
            cmdButton_Click 2
        Case vbKeyF7 And cmdButton(3).Enabled, vbKeyF And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyPageUp
            GotoPreviousPanel Me, btnPanel.Count
        Case vbKeyPageDown
            GotoNextPanel Me, btnPanel.Count
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
        
        Case vbKeyInsert And cmdButton(6).Enabled, vbKeyN And CtrlDown And cmdButton(6).Enabled And Not btnPanel(1).Enabled
            cmdButton_Click 6
        Case vbKeyF3 And cmdButton(7).Enabled, vbKeyD And CtrlDown And cmdButton(7).Enabled And Not btnPanel(1).Enabled
            cmdButton_Click 7
            
    End Select

End Function

Private Function GotoNextPanel(formName, ParamArray panels())

    Dim intLoop As Integer
    
    For intLoop = 0 To btnPanel.Count - 1
    
        If Not btnPanel(intLoop).Enabled Then
            If intLoop + 1 <= btnPanel.Count - 1 Then
                If btnPanel(intLoop + 1).Enabled Then
                    btnPanel_Click intLoop + 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function


Private Sub Form_Load()
    
    AddColumnsToGrid grdAgreements, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdAgreements"), "04NCNID,04NCDFrom,04NCDTo,30NLNRemarks,04NRFAmount,05NCNStatus,05NCNDeleted", "ID,즌�,툼�,急畯晴準暢煜,器冊,�,�"
    
    SetUpGrid lstIconList, grdAgreements
    
    PositionPanels
    
    PositionControls Me, False: ColorizeControls Me, , True
    
    ClearFields txtID, txtCompanyID, txtSpecialityID, txtShipID
    ClearFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    ClearFields grdAgreements
    
    DisableFields txtLastname, txtFirstName, txtCompanyDescription, mskHireDate, txtSpecialityDescription, txtShipDescription, txtPhones, txtRemarks
    DisableFields grdAgreements
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    DisableFields btnPanel(0), btnPanel(1)
    
    UpdateButtons Me, 7, 1, 0, 0, 1, 0, 1, 0, 0
    
    ColorizeControls Me, False, False
    
    grdAgreements.RowMode = False
    
    'AddDummyLines grdAgreements, "", "A99/99/9999A", "A99/99/9999A", "좋좋좋좋좋좋좋좋좋좋", "9.999.999,99"

End Sub

Private Sub grdAgreements_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub mnu즌炡絲嶪滄伎不諪瑨晴誼�_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdAgreements", grdAgreements.LayoutCol

End Sub

Private Sub txtCompanyDescription_Change()

    If txtCompanyDescription.text = "" Then ClearFields txtCompanyID

End Sub

Private Sub txtCompanyDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtCompanyDescription_Validate(Cancel As Boolean)
    
    If txtCompanyID.text = "" And txtCompanyDescription.text <> "" Then cmdIndex_Click 6: If txtCompanyID.text = "" Then Cancel = True
    
End Sub

Private Function ShowIndex()

    With EmployeesIndex
        .Tag = "True"
        .lblTitle.Caption = "투遵帖中� 嚴信烈燐薔�"
        .Show 1, Me
    End With

End Function

Private Sub txtShipDescription_Change()

    If txtShipDescription.text = "" Then txtShipID.text = ""

End Sub

Private Sub txtShipDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    If KeyCode = vbKeyF5 Then cmdIndex_Click 5

End Sub

Private Sub txtShipDescription_Validate(Cancel As Boolean)

    If txtShipID.text = "" And txtShipDescription.text <> "" Then cmdIndex_Click 4: If txtShipID.text = "" Then Cancel = True
    
End Sub

Private Sub txtSpecialityDescription_Change()

    If txtSpecialityDescription.text = "" Then ClearFields txtSpecialityID

End Sub

Private Sub txtSpecialityDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3

End Sub

Private Sub txtSpecialityDescription_Validate(Cancel As Boolean)

    If txtSpecialityID.text = "" And txtSpecialityDescription.text <> "" Then cmdIndex_Click 2: If txtSpecialityID.text = "" Then Cancel = True
    
End Sub

