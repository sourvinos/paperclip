VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form RevenuesIndex 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF80FF&
   BorderStyle     =   0  'None
   ClientHeight    =   12075
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   19590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   19590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   20
      Top             =   8025
      Width           =   6090
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Συνέχεια"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388736
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Κλείσιμο"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388736
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Επεξεργασία εγγραφής"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388736
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Νέα αναζήτηση"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388736
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
      Left            =   8625
      TabIndex        =   11
      Top             =   3750
      Width           =   4515
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Destinations.DestinationID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox txtDestinationID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtPaymentWayID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Text            =   "PaymentWays.ID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox txtRevenueSourceID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "RevenueSources.ID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtShipID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Text            =   "Ships.ShipDescription"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Companies.ID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtCompanyID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1950
         _ExtentX        =   953
         _ExtentY        =   953
      End
   End
   Begin VB.Frame frmCriteria 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   0
      Left            =   525
      TabIndex        =   3
      Top             =   2775
      Width           =   8040
      Begin UserControls.newDate mskDateFrom 
         Height          =   465
         Left            =   2175
         TabIndex        =   4
         Top             =   825
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
      Begin UserControls.newDate mskDateTo 
         Height          =   465
         Left            =   3675
         TabIndex        =   5
         Top             =   825
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   0
         Left            =   7200
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1350
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
         PicNormal       =   "RevenuesIndex.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtCompanyDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   30
         Top             =   1350
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
         Index           =   1
         Left            =   7200
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1875
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
         PicNormal       =   "RevenuesIndex.frx":059A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtShipDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   35
         Top             =   1875
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
         Index           =   2
         Left            =   7200
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2400
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
         PicNormal       =   "RevenuesIndex.frx":0B34
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtDestinationDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   38
         Top             =   2400
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
         Index           =   3
         Left            =   7200
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2925
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
         PicNormal       =   "RevenuesIndex.frx":10CE
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtRevenueSourceDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   41
         Top             =   2925
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
         Index           =   4
         Left            =   7200
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3450
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
         PicNormal       =   "RevenuesIndex.frx":1668
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtPaymentWayDescription 
         Height          =   465
         Left            =   2175
         TabIndex        =   44
         Top             =   3450
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
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Τρόπος είσπραξης"
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
         TabIndex        =   45
         Top             =   3525
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προέλευση"
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
         TabIndex        =   42
         Top             =   3000
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προορισμός"
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
         Top             =   2475
         Width           =   1290
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Πλοίο"
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
         TabIndex        =   36
         Top             =   1950
         Width           =   1290
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   315
         Index           =   5
         Left            =   4875
         Top             =   3900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   315
         Index           =   4
         Left            =   2625
         Top             =   525
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Εταιρία"
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
         TabIndex        =   31
         Top             =   1425
         Width           =   1290
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   1
         Left            =   7575
         Top             =   1575
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
         Left            =   1725
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   2
         Left            =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   4200
         Width           =   8040
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         Caption         =   "01/05/2017"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   5100
         TabIndex        =   8
         Top             =   75
         Width           =   2790
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   75
         Width           =   1665
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Εκδοση"
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
         TabIndex        =   6
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   8040
      End
   End
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   8625
      TabIndex        =   0
      Top             =   6375
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "RevenuesIndex.frx":1C02
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "RevenuesIndex.frx":1C1E
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMaster 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Τίτλος"
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
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   75
         Width           =   3765
      End
   End
   Begin iGrid300_10Tec.iGrid grdRevenuesIndex 
      Height          =   6165
      Left            =   450
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1425
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   10874
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερολόγιο εσόδων"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   375
      TabIndex        =   28
      Top             =   150
      Width           =   4470
   End
   Begin VB.Label lblRecordCount 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Βρέθηκαν 99.999 εγγραφές"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   450
      TabIndex        =   27
      Top             =   1050
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   1200
      Top             =   7575
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   2775
      Top             =   8775
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblCriteria 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Κριτήρια αναζήτησης"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   2925
      TabIndex        =   26
      Top             =   1050
      Width           =   12615
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   15525
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   3
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "RevenuesIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields grdRevenuesIndex
        ClearFields lblRecordCount, lblCriteria
        frmCriteria(0).Visible = True
        mskDateFrom.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
    
    If blnStatus Then
        Set RevenuesIndex = Nothing
        Unload Me
    End If

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskDateFrom.text, mskDateTo.text, txtCompanyDescription.text, txtShipDescription.text, txtDestinationDescription.text, txtRevenueSourceDescription.text, txtPaymentWayDescription.text
            EnableGrid grdRevenuesIndex, False
            HighlightRow grdRevenuesIndex, 1, 1, "", True
            UpdateButtons Me, 3, 0, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 3, 1, 0, 0, 1
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                    End If
                End If
            End If
            blnError = False
            blnProcessing = False
            frmCriteria(0).Visible = True
            mskDateFrom.SetFocus
        End If
    End If

End Function

Private Function UpdateCriteriaLabels(dateFrom, DateTo, Company, Ship, Destination, Source, PaymentWay)

    Dim strCriteriaA As String

    strCriteriaA = IIf(dateFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & dateFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(DateTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & DateTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(Company = "", "Εταιρία [ ΟΛΕΣ ] ", "Εταιρία [ " & Company & " ] ")
    strCriteriaA = strCriteriaA & IIf(Ship = "", "Πλοίο [ ΟΛΑ ] ", "Πλοίο [ " & Ship & " ] ")
    strCriteriaA = strCriteriaA & IIf(Destination = "", "Προορισμοί [ ΟΛΟΙ ] ", "Προορισμός [ " & Destination & " ] ")
    strCriteriaA = strCriteriaA & IIf(Source = "", "Προέλευση [ ΟΛΕΣ ] ", "Προέλευση [ " & Source & " ] ")
    strCriteriaA = strCriteriaA & IIf(PaymentWay = "", "Τρόπος είσπραξης [ ΟΛΟΙ ] ", "Τρόπος είσπραξης [ " & PaymentWay & " ] ")
    
    lblCriteria.Caption = strCriteriaA
    
End Function



Private Function RefreshList()
    
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
    
    'Local variables
    Dim lngRow As Long
    Dim strFullInvoice As String
    Dim curTotalRevenue As Currency
    Dim lngTotalPersons As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdRevenuesIndex
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT Revenues.ID AS ID, DateIssue, Companies.Description AS Company, Ships.ShipDescription AS Ship, Destinations.DestinationDescription AS Destination, RevenueSources.Description AS Source, PaymentWays.Description AS PaymentWay, Amount " _
        & "FROM (((((Revenues " _
        & "INNER JOIN Companies ON Revenues.CompanyID = Companies.ID) " _
        & "INNER JOIN Ships ON Revenues.ShipID = Ships.ShipID) " _
        & "INNER JOIN Destinations ON Revenues.DestinationID = Destinations.DestinationID) " _
        & "INNER JOIN RevenueSources ON Revenues.SourceID = RevenueSources.ID) " _
        & "INNER JOIN PaymentWays ON Revenues.PaymentWayID = PaymentWays.ID) "
        
    'Εκδοση Από
    If mskDateFrom.text <> "" Then
        strThisParameter = "datFromDate Date"
        strThisQuery = "DateIssue >= datFromDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskDateFrom.text
    End If
        
    'Εκδοση Εως
    If mskDateTo.text <> "" Then
        strThisParameter = "datToDate Date"
        strThisQuery = "DateIssue <= datToDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskDateTo.text
    End If
    
    'Εταιρία
    If txtCompanyID.text <> "" Then
        strThisParameter = "intCompanyID Integer"
        strThisQuery = "CompanyID = intCompanyID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCompanyID.text)
    End If
    
    'Πλοίο
    If txtShipID.text <> "" Then
        strThisParameter = "intShipID Integer"
        strThisQuery = "Revenues.ShipID = intShipID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtShipID.text)
    End If
    
    'Προορισμός
    If txtDestinationID.text <> "" Then
        strThisParameter = "intDestinationID Integer"
        strThisQuery = "Revenues.DestinationID = intDestinationID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDestinationID.text)
    End If
    
    'Προέλευση
    If txtRevenueSourceID.text <> "" Then
        strThisParameter = "intRevenueSourceID Integer"
        strThisQuery = "SourceID = intRevenueSourceID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtRevenueSourceID.text)
    End If
    
    'Τρόπος είσπραξης
    If txtPaymentWayID.text <> "" Then
        strThisParameter = "intPaymentWayID Integer"
        strThisQuery = "PaymentWayID = intPaymentWayID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPaymentWayID.text)
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY DateIssue"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
    End If
    
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnErrors = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(2).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdRevenuesIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdRevenuesIndex.CellValue(lngRow, "TrnID") = !id
            grdRevenuesIndex.CellValue(lngRow, "DateIssue") = !DateIssue
            grdRevenuesIndex.CellValue(lngRow, "Company") = !Company
            grdRevenuesIndex.CellValue(lngRow, "Ship") = !Ship
            grdRevenuesIndex.CellValue(lngRow, "Destination") = !Destination
            grdRevenuesIndex.CellValue(lngRow, "Source") = !Source
            grdRevenuesIndex.CellValue(lngRow, "PaymentWay") = !PaymentWay
            grdRevenuesIndex.CellValue(lngRow, "Amount") = !amount
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        rstRecordset.Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdRevenuesIndex
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(2).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
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
    blnErrors = True
    ClearFields grdRevenuesIndex, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            EditRecord
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select
    
End Sub

Private Sub cmdButton_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Εταιρία - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Companies", "Description", "String", txtCompanyDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtCompanyID.text = tmpTableData.strCode
                txtCompanyDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'Πλοίο - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShipDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtShipID.text = tmpTableData.strCode
                txtShipDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            'Προορισμός - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            'Προέλευση - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "RevenueSources", "Description", "String", txtRevenueSourceDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtRevenueSourceID.text = tmpTableData.strCode
                txtRevenueSourceDescription.text = tmpTableData.strFirstField
            End If
        Case 4
            'Τρόπος είσπραξης - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentWays", "Description", "String", txtPaymentWayDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPaymentWayID.text = tmpTableData.strCode
                txtPaymentWayDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdRevenuesIndex, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdRevenuesIndex"), _
            "05NCNTrnID,12NCDDateIssue,40NLNCompany,40NLNShip,40NLNDestination,40NLNSource,40NLNPaymentWay,10NRFAmount", _
            "TrnID,Ημερομηνία,Εταιρία,Πλοίο,Προορισμός,Προέλευση,Τρόπος πληρωμής,Ποσό"
        Me.Refresh
        'mskDateFrom.SetFocus
    End If
    
    'AddDummyLines grdRevenuesIndex, "99999", "A99/99/9999A", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "-9.999.999,99"

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
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyE And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False, grdRevenuesIndex
    ColorizeControls Me, False, False
    SetUpGrid lstIconList, grdRevenuesIndex
    
    ClearFields txtCompanyID, txtShipID, txtDestinationID, txtRevenueSourceID, txtPaymentWayID
    ClearFields mskDateFrom, mskDateTo, txtCompanyDescription, txtShipDescription, txtDestinationDescription, txtRevenueSourceDescription, txtPaymentWayDescription
    ClearFields grdRevenuesIndex
    ClearFields lblRecordCount, lblCriteria
    
    EnableFields mskDateFrom, mskDateTo
    UpdateButtons Me, 3, 1, 0, 0, 1

End Sub

Private Sub grdRevenuesIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdRevenuesIndex.Header.Buttons = True

End Sub

Private Sub grdRevenuesIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdRevenuesIndex.Header.Buttons = False

End Sub


Private Sub grdRevenuesIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdRevenuesIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Function EditRecord()
    
    If Not grdRevenuesIndex.Enabled Then Exit Function
        
    Dim rstRecordset As Recordset
    
    Set rstRecordset = SeekRecord(grdRevenuesIndex.CellValue(grdRevenuesIndex.CurRow, "TrnID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    Revenues.DoPostFoundJobs rstRecordset
    Revenues.Tag = "True"
    
    If Revenues.Visible Then
        Set RevenuesIndex = Nothing
        Unload Me
        Revenues.Show
    Else
        Revenues.Show 1, Me
        grdRevenuesIndex.SetFocus
    End If
    
End Function

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
        & "Revenues.ID AS ID, " _
        & "Revenues.DateIssue, " _
        & "Revenues.CompanyID, Companies.Description AS CompanyDescription, " _
        & "Revenues.ShipID, Ships.ShipDescription AS ShipDescription, " _
        & "Revenues.DestinationID, Destinations.DestinationDescription AS DestinationDescription, " _
        & "Revenues.SourceID, RevenueSources.Description AS SourceDesciption, " _
        & "Revenues.PaymentWayID, PaymentWays.Description AS PaymentWayDescription, " _
        & "Amount " _
        & "FROM (((((Revenues " _
        & "INNER JOIN Companies ON Revenues.CompanyID = Companies.ID) " _
        & "INNER JOIN Ships ON Revenues.ShipID = Ships.ShipID) " _
        & "INNER JOIN Destinations ON Revenues.DestinationID = Destinations.DestinationID) " _
        & "INNER JOIN RevenueSources ON Revenues.SourceID = RevenueSources.ID) " _
        & "INNER JOIN PaymentWays ON Revenues.PaymentWayID = PaymentWays.ID) "
        
    'ID
    strThisParameter = "lngID long"
    strThisQuery = "Revenues.ID = lngID"
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


Private Function ValidateFields()

    ValidateFields = False
    
    'Σωστό διάστημα
    If IsDate(mskDateFrom.text) And IsDate(mskDateTo.text) Then
        If CDate(mskDateFrom.text) > CDate(mskDateTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskDateFrom.SetFocus
            Exit Function
        End If
    End If

    ValidateFields = True
    
End Function

Private Sub grdRevenuesIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdRevenuesIndex", grdRevenuesIndex.LayoutCol

End Sub

Private Sub txtCompanyDescription_Change()

    If txtCompanyDescription.text = "" Then ClearFields txtCompanyID

End Sub

Private Sub txtCompanyDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtCompanyDescription_Validate(Cancel As Boolean)

    If txtCompanyID.text = "" And txtCompanyDescription.text <> "" Then cmdIndex_Click 0: If txtCompanyID.text = "" Then Cancel = True

End Sub


Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then ClearFields txtDestinationID
    
End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    
End Sub


Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 2: If txtDestinationID.text = "" Then Cancel = True

End Sub


Private Sub txtPaymentWayDescription_Change()

    If txtPaymentWayDescription.text = "" Then ClearFields txtPaymentWayID

End Sub

Private Sub txtPaymentWayDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    
End Sub


Private Sub txtPaymentWayDescription_Validate(Cancel As Boolean)

    If txtPaymentWayID.text = "" And txtPaymentWayDescription.text <> "" Then cmdIndex_Click 4: If txtPaymentWayID.text = "" Then Cancel = True
    
End Sub

Private Sub txtRevenueSourceDescription_Change()

    If txtRevenueSourceDescription.text = "" Then ClearFields txtRevenueSourceID
    
End Sub


Private Sub txtRevenueSourceDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3
    
End Sub


Private Sub txtRevenueSourceDescription_Validate(Cancel As Boolean)

    If txtRevenueSourceID.text = "" And txtRevenueSourceDescription.text <> "" Then cmdIndex_Click 3: If txtRevenueSourceID.text = "" Then Cancel = True

End Sub

Private Sub txtShipDescription_Change()

    If txtShipDescription.text = "" Then ClearFields txtShipID
    
End Sub


Private Sub txtShipDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    
End Sub


Private Sub txtShipDescription_Validate(Cancel As Boolean)

    If txtShipID.text = "" And txtShipDescription.text <> "" Then cmdIndex_Click 1: If txtShipID.text = "" Then Cancel = True
    
End Sub


