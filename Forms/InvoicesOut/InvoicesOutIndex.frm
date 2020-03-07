VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form InvoicesOutIndex 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   19170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   12900
      TabIndex        =   38
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "InvoicesOutIndex.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "InvoicesOutIndex.frx":001C
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
         TabIndex        =   40
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   75
      TabIndex        =   10
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   30
         Top             =   8850
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
            TabIndex        =   33
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
            Index           =   4
            Left            =   5925
            TabIndex        =   34
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
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3080
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εκτύπωση επιλεγμένων"
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
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   12640511
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία PDF επιλεγμένων"
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
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   5715
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   3000
         Width           =   8040
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   0
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Left            =   3675
            TabIndex        =   1
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
         Begin UserControls.newText txtShipDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   8
            Top             =   3975
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            TabIndex        =   21
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
            PicNormal       =   "InvoicesOutIndex.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtDestinationDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   7
            Top             =   3450
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   40
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            TabIndex        =   22
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
            PicNormal       =   "InvoicesOutIndex.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   6
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
            Index           =   3
            Left            =   7200
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   3975
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
            PicNormal       =   "InvoicesOutIndex.frx":0B6C
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtCodeShortDescriptionA 
            Height          =   465
            Left            =   2175
            TabIndex        =   4
            Top             =   1875
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   3
            Text            =   "ΑΑΑ"
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
            Left            =   2925
            TabIndex        =   49
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
            PicNormal       =   "InvoicesOutIndex.frx":1106
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtInvoiceNo 
            Height          =   465
            Left            =   2175
            TabIndex        =   5
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   10
            Text            =   "AAAAAAAAAA"
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
         Begin UserControls.newDate mskInvoiceDateTripFrom 
            Height          =   465
            Left            =   2175
            TabIndex        =   2
            Top             =   1350
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
         Begin UserControls.newDate mskInvoiceDateTripTo 
            Height          =   465
            Left            =   3675
            TabIndex        =   3
            Top             =   1350
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
         Begin UserControls.newText txtYesOrNoDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   9
            Top             =   4500
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
            MaxLength       =   3
            Text            =   "ΑΑΑ"
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
            Left            =   2925
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   4500
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
            PicNormal       =   "InvoicesOutIndex.frx":16A0
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Με μεταφορά"
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
            TabIndex        =   60
            Top             =   4575
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Αφορά"
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
            TabIndex        =   58
            Top             =   1425
            Width           =   1290
         End
         Begin VB.Label lblCodeDescription 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   3375
            TabIndex        =   56
            Top             =   1875
            Width           =   4200
         End
         Begin VB.Label lblCodeBatch 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ΣΕΙΡΑ Ω"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   4800
            TabIndex        =   55
            Top             =   2100
            Width           =   585
         End
         Begin VB.Label lblHand 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ΜΗΧΑΝΟΓΡΑΦΙΚΟ"
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   3375
            TabIndex        =   54
            Top             =   2100
            Width           =   1350
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Νο παραστατικού"
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
            TabIndex        =   51
            Top             =   2475
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Παραστατικό"
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
            TabIndex        =   50
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2775
            Top             =   4875
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
            Left            =   2775
            Top             =   525
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Συναλλασόμενος"
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
            Top             =   3000
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
            Index           =   4
            Left            =   0
            TabIndex        =   28
            Top             =   5175
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
            Left            =   2625
            TabIndex        =   27
            Top             =   75
            Width           =   5265
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
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   900
            Width           =   1290
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   0
            Left            =   0
            Top             =   1800
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
            Left            =   1725
            Top             =   1950
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
            Left            =   7575
            Top             =   2700
            Visible         =   0   'False
            Width           =   465
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
            TabIndex        =   24
            Top             =   4050
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
            TabIndex        =   23
            Top             =   3525
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
            TabIndex        =   29
            Top             =   0
            Width           =   8040
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   8250
         TabIndex        =   11
         Top             =   5100
         Width           =   4515
         Begin VB.TextBox txtYesOrNoID 
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
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   2325
            Width           =   780
         End
         Begin VB.TextBox Text5 
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
            TabIndex        =   61
            TabStop         =   0   'False
            Text            =   "YesNo"
            Top             =   2325
            Width           =   3540
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
            TabIndex        =   53
            TabStop         =   0   'False
            Text            =   "InvoiceCodeID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceCodeID 
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
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox Text4 
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
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "InvoicePersonID"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtPersonID 
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
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1200
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
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "InvoiceOutShipID"
            Top             =   1950
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "999"
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox txtInvoiceSecondaryRefersTo 
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
            Text            =   "999"
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text8 
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
            Text            =   "InvoiceSecondaryRefersTo"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtDestinationID 
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
            Text            =   "999"
            Top             =   1575
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
            Text            =   "InvoiceOutDestinationID"
            Top             =   1575
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
            Text            =   "InvoiceMasterRefersTo"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceMasterRefersTo 
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
            Text            =   "999"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   2775
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "InvoicesOutIndex.frx":1C3A
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdInvoicesOutIndex 
         Height          =   7290
         Left            =   75
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   12859
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
         Left            =   3975
         TabIndex        =   48
         Top             =   1125
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Επιλεγμένες 0 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3975
         TabIndex        =   47
         Top             =   525
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολα πάνε εδώ"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   3975
         TabIndex        =   46
         Top             =   825
         Width           =   14940
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
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   45
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ημερολόγιο πωλήσεων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   75
         TabIndex        =   37
         Top             =   75
         Width           =   5100
      End
      Begin VB.Shape shpBottomEdge 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Left            =   2550
         Top             =   9900
         Visible         =   0   'False
         Width           =   840
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
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   21975
      Top             =   5550
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
Attribute VB_Name = "InvoicesOutIndex"
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
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdInvoicesOutIndex
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function ColorizeColumns(Color As String, Row As Long, ParamArray Columns() As Variant)

    Dim lngIndex As Long
    
    For lngIndex = 0 To UBound(Columns)
        grdInvoicesOutIndex.CellForeColor(Row, grdInvoicesOutIndex.ColIndex(Columns(lngIndex))) = Color
        'grdInvoicesOutIndex.CellBackColor(Row, grdInvoicesOutIndex.ColIndex(Columns(lngIndex))) = Color
    Next lngIndex

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskInvoiceDateIssueFrom.text, mskInvoiceDateIssueTo.text, txtPersonDescription.text, txtDestinationDescription.text, txtShipDescription.text
            EnableGrid grdInvoicesOutIndex, False
            HighlightRow grdInvoicesOutIndex, 1, 1, "", True
            UpdateButtons Me, 5, 0, 1, 1, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
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
            mskInvoiceDateIssueFrom.SetFocus
        End If
    End If

End Function

Private Function UpdateCriteriaLabels(InvoiceDateIssueFrom, InvoiceDateIssueTo, PersonDescription, DestinationDescription, ShipDescription)

    Dim strCriteriaA As String

    strCriteriaA = IIf(InvoiceDateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & InvoiceDateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(InvoiceDateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & InvoiceDateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(PersonDescription = "", "Συναλλασόμενος [ ΟΛΟΙ ] ", "Συναλλασόμενος [ " & PersonDescription & " ] ")
    strCriteriaA = strCriteriaA & IIf(DestinationDescription = "", "Προορισμός [ ΟΛΟΙ ] ", "Προορισμός [ " & DestinationDescription & " ] ")
    strCriteriaA = strCriteriaA & IIf(ShipDescription = "", "Πλοία [ ΟΛΑ ]", "Πλοίο [ " & ShipDescription & " ]")
    
    lblCriteria.Caption = strCriteriaA
    
End Function


Private Function PrintSelectedInvoices(whatToDo)

    Dim lngRow As Long
    Dim intIndex As Integer
    Dim arrInvoicesTrnID()
    Dim arrCustomerNames()
    
    intIndex = -1
    
    If Not grdInvoicesOutIndex.Enabled Then Exit Function
    
    If Not LinesHaveBeenSelected(grdInvoicesOutIndex) Then
        MyMsgBox 4, strApplicationName, strStandardMessages(6), 1
        Exit Function
    End If
    
    For lngRow = 1 To grdInvoicesOutIndex.RowCount
        If grdInvoicesOutIndex.CellIcon(lngRow, "Selected") >= 1 Then
            intIndex = intIndex + 1
            ReDim Preserve arrInvoicesTrnID(intIndex)
            ReDim Preserve arrCustomerNames(intIndex)
            arrInvoicesTrnID(intIndex) = grdInvoicesOutIndex.CellValue(lngRow, "TrnID")
            arrCustomerNames(intIndex) = grdInvoicesOutIndex.CellValue(lngRow, "CustomerDescription")
        End If
    Next lngRow
    
    InvoicesOut.ProcessSelectedInvoicesForPrinting "", whatToDo, arrInvoicesTrnID, arrCustomerNames 'Called when the array is processed

End Function

Private Function EditRecord()

    If Not grdInvoicesOutIndex.Enabled Then Exit Function
        
    Dim rstRecordset As Recordset
    
    Set rstRecordset = InvoicesOut.SeekRecord(grdInvoicesOutIndex.CellValue(grdInvoicesOutIndex.CurRow, "TrnID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    InvoicesOut.DoPostFoundJobs rstRecordset
    
    If Not InvoicesOut.Visible Then
        InvoicesOut.Show 1, Me
        grdInvoicesOutIndex.SetFocus
    Else
        Unload Me
    End If
    
End Function

Private Function RefreshList()
    
    On Error GoTo ErrTrap

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
    Dim curInvoiceTotalAmount As Currency
    Dim lngInvoiceTotalPersons As Long
    Dim curTotalRevenue As Currency
    
    'Ατομα με μεταφορά
    Dim lngGrandTotalAdultsWithTransfer As Long
    Dim lngGrandTotalKidsWithTransfer As Long
    Dim lngGrandTotalFreeWithTransfer As Long
    Dim lngGrandTotalWithTransfer As Long
    
    Dim lngGrandTotalAdultsWithoutTransfer As Long
    Dim lngGrandTotalKidsWithoutTransfer As Long
    Dim lngGrandTotalFreeWithoutTransfer As Long
    Dim lngGrandTotalWithoutTransfer As Long
    
    Dim lngGrandTotalAdults As Long
    Dim lngGrandTotalKids As Long
    Dim lngGrandTotalFree As Long
    Dim lngGrandTotal As Long
    
    Dim lngWithTransfer As Long
    
    'Ατομα χωρίς μεταφορά
    Dim lngAdultsWithoutTransfer As Long
    Dim lngKidsWithoutTransfer As Long
    Dim lngFreeWithoutTransfer As Long
    Dim lngPersonsWithoutTransfer As Long
    Dim lngWithoutTransfer As Long
    
    'Συνολικά άτομα γραμμής
    Dim lngLineAdults As Long
    Dim lngLineKids As Long
    Dim lngLineFree As Long
    Dim lngLineTotal As Long
    
    'Συνολικά άτομα τελευταία γραμμή
    Dim lngTotalAdults As Long
    Dim lngTotalKids As Long
    Dim lngTotalFree As Long
    Dim lngTotalPersons As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdInvoicesOutIndex
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "InvoiceTrnID, InvoiceDateIssue, InvoiceDateRefersTo, InvoiceNo, InvoiceOutAdultsWithTransfer, InvoiceOutKidsWithTransfer, InvoiceOutFreeWithTransfer, InvoiceOutAdultsWithoutTransfer, InvoiceOutKidsWithoutTransfer, InvoiceOutFreeWithoutTransfer, InvoiceOutAdultsAmountWithTransfer, InvoiceOutKidsAmountWithTransfer, InvoiceOutAdultsAmountWithoutTransfer, InvoiceOutKidsAmountWithoutTransfer, InvoiceOutDirectAmount, " _
        & "Description, " _
        & "DestinationDescription, " _
        & "CodeShortDescriptionB, CodeBatch, CodeDescription, CodeCustomers, " _
        & "ShipDescription " _
        & "FROM (((((Invoices " _
        & "INNER JOIN InvoicesOut ON Invoices.InvoiceTrnID = InvoicesOut.InvoiceOutTrnID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Destinations ON InvoicesOut.InvoiceOutDestinationID = Destinations.DestinationID) " _
        & "INNER JOIN Ships ON InvoicesOut.InvoiceOutShipID = Ships.ShipID) "
        
    'Εγγραφές πωλήσεων
    strThisParameter = "strMasterRefersTo String"
    strThisQuery = "InvoiceMasterRefersTo = strMasterRefersTo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = txtInvoiceMasterRefersTo.text
    
    'Εκδρομές πλοίων ή λεωφορείων
    strThisParameter = "strSecondaryRefersTo String"
    strThisQuery = "InvoiceSecondaryRefersTo = strSecondaryRefersTo"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = txtInvoiceSecondaryRefersTo.text
    
    'Εκδοση Από
    If mskInvoiceDateIssueFrom.text <> "" Then
        strThisParameter = "datFromDate Date"
        strThisQuery = "InvoiceDateIssue >= datFromDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueFrom.text
    End If
        
    'Εκδοση Εως
    If mskInvoiceDateIssueTo.text <> "" Then
        strThisParameter = "datToDate Date"
        strThisQuery = "InvoiceDateIssue <= datToDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueTo.text
    End If
    
    'Εκδρομή Από
    If mskInvoiceDateTripFrom.text <> "" Then
        strThisParameter = "datTripFromDate Date"
        strThisQuery = "InvoiceOutTripDate >= datTripFromDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateTripFrom.text
    End If
    
    'Εκδρομή Εως
    If mskInvoiceDateTripTo.text <> "" Then
        strThisParameter = "datTripToDate Date"
        strThisQuery = "InvoiceOutTripDate <= datTripToDate"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateTripTo.text
    End If
    
    'Τύπος παραστατικού
    If txtInvoiceCodeID.text <> "" Then
        strThisParameter = "lngCodeID Long"
        strThisQuery = "InvoiceCodeID = lngCodeID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtInvoiceCodeID.text)
    End If
    
    'Νο Παραστατικού
    If txtInvoiceNo.text <> "" Then
        strThisParameter = "strInvoiceNo String"
        strThisQuery = "InvoiceNo = strInvoiceNo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = txtInvoiceNo.text
    End If
    
    'Συναλλασόμενος
    If txtPersonID.text <> "" Then
        strThisParameter = "intPersonID Integer"
        strThisQuery = "InvoicePersonID = intPersonID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPersonID.text)
    End If
    
    'Προορισμός
    If txtDestinationID.text <> "" Then
        strThisParameter = "intDestinationID Integer"
        strThisQuery = "InvoiceOutDestinationID = intDestinationID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDestinationID.text)
    End If
    
    'Πλοίο
    If txtShipID.text <> "" Then
        strThisParameter = "intShipID Integer"
        strThisQuery = "InvoiceOutShipID = intShipID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtShipID.text)
    End If
    
    If txtYesOrNoId.text <> "" Then
        'Με Μεταφορά
        If txtYesOrNoId.text = "1" Then
            strThisParameter = "intTransfer Integer"
            strThisQuery = " InvoiceOutAdultsWithTransfer + InvoiceOutKidsWithTransfer + InvoiceOutFreeWithTransfer > intTransfer"
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = 0
        End If
        'Χωρίς Μεταφορά
        If txtYesOrNoId.text = "0" Then
            strThisParameter = "intTransfer Integer"
            strThisQuery = " InvoiceOutAdultsWithOutTransfer + InvoiceOutKidsWithOutTransfer + InvoiceOutFreeWithOutTransfer > intTransfer"
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = 0
        End If
    
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceDateIssue, InvoiceID"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdInvoicesOutIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdInvoicesOutIndex.CellValue(lngRow, "TrnID") = !InvoiceTrnID
            grdInvoicesOutIndex.CellValue(lngRow, "InvoiceDateIssue") = !InvoiceDateIssue
            strFullInvoice = !CodeShortDescriptionB & Space(3 - Len(!CodeShortDescriptionB)) & " "
            strFullInvoice = strFullInvoice & IIf(!CodeBatch <> "", !CodeBatch, "0") & " "
            strFullInvoice = strFullInvoice & Right("00000" & !InvoiceNo, 5)
            curInvoiceTotalAmount = !InvoiceOutAdultsAmountWithTransfer + !InvoiceOutKidsAmountWithTransfer + !InvoiceOutAdultsAmountWithoutTransfer + !InvoiceOutKidsAmountWithoutTransfer + !InvoiceOutDirectAmount
            grdInvoicesOutIndex.CellValue(lngRow, "FullInvoice") = strFullInvoice
            grdInvoicesOutIndex.CellValue(lngRow, "CustomerDescription") = !Description
            grdInvoicesOutIndex.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdInvoicesOutIndex.CellValue(lngRow, "ShipDescription") = !ShipDescription
            grdInvoicesOutIndex.CellValue(lngRow, "InvoiceTotalAmount") = IIf(!CodeCustomers = "+", curInvoiceTotalAmount, -curInvoiceTotalAmount)
            
            'Εμφανίζω Με μεταφορά
            grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutAdultsWithTransfer, -!InvoiceOutAdultsWithTransfer)
            grdInvoicesOutIndex.CellValue(lngRow, "KidsWithTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutKidsWithTransfer, -!InvoiceOutKidsWithTransfer)
            grdInvoicesOutIndex.CellValue(lngRow, "FreeWithTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutFreeWithTransfer, -!InvoiceOutFreeWithTransfer)
            
            'Υπολογίζω σύνολα γραμμής με μεταφορά
            lngWithTransfer = grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "KidsWithTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "FreeWithTransfer")
            
            'Εμφανίζω χωρίς μεταφορά
            grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithoutTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutAdultsWithoutTransfer, -!InvoiceOutAdultsWithoutTransfer)
            grdInvoicesOutIndex.CellValue(lngRow, "KidsWithoutTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutKidsWithoutTransfer, -!InvoiceOutKidsWithoutTransfer)
            grdInvoicesOutIndex.CellValue(lngRow, "FreeWithoutTransfer") = IIf(!CodeCustomers = "+", !InvoiceOutFreeWithoutTransfer, -!InvoiceOutFreeWithoutTransfer)
            
            'Υπολογίζω σύνολα γραμμής χωρίς μεταφορά
            lngWithoutTransfer = grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithoutTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "KidsWithoutTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "FreeWithoutTransfer")
            
            'Υπολογίζω σύνολα γραμμής με μεταφορά και χωρίς μεταφορά
            lngLineAdults = grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithoutTransfer")
            lngLineKids = grdInvoicesOutIndex.CellValue(lngRow, "KidsWithTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "KidsWithoutTransfer")
            lngLineFree = grdInvoicesOutIndex.CellValue(lngRow, "FreeWithTransfer") + grdInvoicesOutIndex.CellValue(lngRow, "FreeWithoutTransfer")
            lngLineTotal = lngLineAdults + lngLineKids + lngLineFree
            
            'Εμφανίζω σύνολα γραμμής με μεταφορά και χωρίς μεταφορά
            grdInvoicesOutIndex.CellValue(lngRow, "LineAdults") = lngLineAdults
            grdInvoicesOutIndex.CellValue(lngRow, "LineKids") = lngLineKids
            grdInvoicesOutIndex.CellValue(lngRow, "LineFree") = lngLineFree
            grdInvoicesOutIndex.CellValue(lngRow, "LineTotal") = lngLineTotal
            
            'Σύνολα γραμμής
            grdInvoicesOutIndex.CellValue(lngRow, "TotalWithTransfer") = lngWithTransfer
            grdInvoicesOutIndex.CellValue(lngRow, "TotalWithoutTransfer") = lngWithoutTransfer
            
            'Χρώματα αριθμών
            InvertColorForNegativeNumbers grdInvoicesOutIndex, lngRow
            
            'Χρώματα στηλών με άτομα
            ColorizeColumns "&H0ec9FF", lngRow, "AdultsWithTransfer", "KidsWithTransfer", "FreeWithTransfer", "TotalWithTransfer"
            ColorizeColumns "&H72d255", lngRow, "AdultsWithoutTransfer", "KidsWithoutTransfer", "FreeWithoutTransfer", "TotalWithoutTransfer"
            ColorizeColumns "&Head999", lngRow, "LineAdults", "LineKids", "LineFree", "LineTotal"
            
            curTotalRevenue = curTotalRevenue + grdInvoicesOutIndex.CellValue(lngRow, "InvoiceTotalAmount")
            
            'Τελευταία γραμμή
            'Με μεταφορά
            lngGrandTotalAdultsWithTransfer = lngGrandTotalAdultsWithTransfer + grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithTransfer")
            lngGrandTotalKidsWithTransfer = lngGrandTotalKidsWithTransfer + grdInvoicesOutIndex.CellValue(lngRow, "KidsWithTransfer")
            lngGrandTotalFreeWithTransfer = lngGrandTotalFreeWithTransfer + grdInvoicesOutIndex.CellValue(lngRow, "FreeWithTransfer")
            lngGrandTotalWithTransfer = lngGrandTotalAdultsWithTransfer + lngGrandTotalKidsWithTransfer + lngGrandTotalFreeWithTransfer
            'Χωρίς μεταφορά
            lngGrandTotalAdultsWithoutTransfer = lngGrandTotalAdultsWithoutTransfer + grdInvoicesOutIndex.CellValue(lngRow, "AdultsWithoutTransfer")
            lngGrandTotalKidsWithoutTransfer = lngGrandTotalKidsWithoutTransfer + grdInvoicesOutIndex.CellValue(lngRow, "KidsWithoutTransfer")
            lngGrandTotalFreeWithoutTransfer = lngGrandTotalFreeWithoutTransfer + grdInvoicesOutIndex.CellValue(lngRow, "FreeWithoutTransfer")
            lngGrandTotalWithoutTransfer = lngGrandTotalAdultsWithoutTransfer + lngGrandTotalKidsWithoutTransfer + lngGrandTotalFreeWithoutTransfer
            'Σύνολα
            lngGrandTotalAdults = lngGrandTotalAdults + lngGrandTotalAdultsWithTransfer + lngGrandTotalAdultsWithoutTransfer
            lngGrandTotalKids = lngGrandTotalKids + lngGrandTotalKidsWithTransfer + lngGrandTotalKidsWithoutTransfer
            lngGrandTotalFree = lngGrandTotalFree + lngGrandTotalFreeWithTransfer + lngGrandTotalFreeWithoutTransfer
            lngGrandTotal = lngGrandTotal + lngGrandTotalAdultsWithTransfer + lngGrandTotalKidsWithTransfer + lngGrandTotalFreeWithTransfer
            
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        rstRecordset.Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdInvoicesOutIndex
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        With grdInvoicesOutIndex
            .AddRow , , , , , , , 2
            .CellValue(grdInvoicesOutIndex.RowCount, "InvoiceTotalAmount") = curTotalRevenue
            'Με μεταφορά
            .CellValue(grdInvoicesOutIndex.RowCount, "AdultsWithTransfer") = lngGrandTotalAdultsWithTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "KidsWithTransfer") = lngGrandTotalKidsWithTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "FreeWithTransfer") = lngGrandTotalFreeWithTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "TotalWithTransfer") = lngGrandTotalWithTransfer
            'Χωρίς μεταφορά
            .CellValue(grdInvoicesOutIndex.RowCount, "AdultsWithoutTransfer") = lngGrandTotalAdultsWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "KidsWithoutTransfer") = lngGrandTotalKidsWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "FreeWithoutTransfer") = lngGrandTotalFreeWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "TotalWithoutTransfer") = lngGrandTotalAdultsWithoutTransfer + lngGrandTotalKidsWithoutTransfer + lngGrandTotalFreeWithoutTransfer
            'Σύνολα
            .CellValue(grdInvoicesOutIndex.RowCount, "LineAdults") = lngGrandTotalAdultsWithTransfer + lngGrandTotalAdultsWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "LineKids") = lngGrandTotalKidsWithTransfer + lngGrandTotalKidsWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "LineFree") = lngGrandTotalFreeWithTransfer + lngGrandTotalFreeWithoutTransfer
            .CellValue(grdInvoicesOutIndex.RowCount, "LineTotal") = .CellValue(grdInvoicesOutIndex.RowCount, "LineAdults") + .CellValue(grdInvoicesOutIndex.RowCount, "LineKids") + .CellValue(grdInvoicesOutIndex.RowCount, "LineFree")
            'Χρώματα
            InvertColorForNegativeNumbers grdInvoicesOutIndex, .RowCount
            ColorizeColumns "&H0ec9FF", grdInvoicesOutIndex.RowCount, "AdultsWithTransfer", "KidsWithTransfer", "FreeWithTransfer", "TotalWithTransfer"
            ColorizeColumns "&H72d255", grdInvoicesOutIndex.RowCount, "AdultsWithoutTransfer", "KidsWithoutTransfer", "FreeWithoutTransfer", "TotalWithoutTransfer"
            ColorizeColumns "&Head999", grdInvoicesOutIndex.RowCount, "LineAdults", "LineKids", "LineFree", "LineTotal"
        End With
    End If
    
    'Τελικές ενέργειες
    cmdButton(4).Caption = "Νέα αναζήτηση"
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
    blnError = True
    ClearFields grdInvoicesOutIndex, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            EditRecord
        Case 2
            PrintSelectedInvoices "Print"
        Case 3
            PrintSelectedInvoices "PDF"
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    'OK
    ValidateFields = False
    
    'Σωστό διάστημα έκδοσης
    If IsDate(mskInvoiceDateIssueFrom.text) And IsDate(mskInvoiceDateIssueTo.text) Then
        If CDate(mskInvoiceDateIssueFrom.text) > CDate(mskInvoiceDateIssueTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom.SetFocus
            Exit Function
        End If
    End If
    
    'Σωστό διάστημα εκδρομής
    If IsDate(mskInvoiceDateTripFrom.text) And IsDate(mskInvoiceDateTripTo.text) Then
        If CDate(mskInvoiceDateTripFrom.text) > CDate(mskInvoiceDateTripTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateTripFrom.SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True

End Function

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub cmdIndex_Click(Index As Integer)
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Παραστατικό - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Codes", "CodeShortDescriptionA, CodeMasterRefersTo", "String, String", txtCodeShortDescriptionA.text, txtInvoiceMasterRefersTo.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 8, 0, 3, 5, 6, 7, 8, 10, 11, "ID", "Συντ. Α'", "Περιγραφή", "Σειρά", "Χειρόγραφο", "Πελάτες", "Τελευταίο Νο", "Ημερομηνία", 0, 6, 40, 6, 10, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 1)
                txtInvoiceCodeID.text = tmpTableData.strCode
                txtCodeShortDescriptionA.text = tmpTableData.strFirstField
                lblCodeDescription.Caption = tmpTableData.strSecondField
                lblCodeBatch.Caption = IIf(txtInvoiceCodeID.text <> "" And tmpTableData.strThirdField <> "", " ΣΕΙΡΑ " & tmpTableData.strThirdField, "")
                lblHand.Caption = IIf(tmpTableData.strFourthField <> "", IIf(tmpTableData.strFourthField = "1", "ΧΕΙΡΟΓΡΑΦΟ", "ΜΗΧΑΝΟΓΡΑΦΙΚΟ"), "")
            End If
        Case 1
            'Πελάτης - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtPersonDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 7, "ID", "Επωνυμία", "Α.Φ.Μ.", 0, 40, 15, 1, 0, 1)
                txtPersonID.text = tmpTableData.strCode
                txtPersonDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            'Προορισμός - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription, ShowInList", "String, Numeric", txtDestinationDescription.text, txtInvoiceSecondaryRefersTo.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            'Πλοίο - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShipDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 6, 0, 1, 3, 4, 5, 6, "ID", "Περιγραφή", "Σημαία", "Αρ. Νηολογίου", "Αρ. Ι.Μ.Ο.", "Διαχειριστής", 0, 40, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0)
                txtShipID.text = tmpTableData.strCode
                txtShipDescription.text = tmpTableData.strFirstField
            End If
        Case 4
            'Transfers
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "Description", "String", txtYesOrNoDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtYesOrNoId.text = tmpTableData.strCode
                txtYesOrNoDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    '"05NCNTrnID,12NCDXInvoiceDateIssue,50NCNFullInvoice,40NLNCustomerDescription,40NLNShipDescription,40NLNDestinationDescription,10NRFInvoiceTotalAmount,10NRIAdultsWithTransfer,10NRIKidsWithTransfer,10NRIFreeWithTransfer,10NRIInvoiceTotalPersons,05NCNSelected"
    
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdInvoicesOutIndex, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdInvoicesOutIndex"), _
            "05NCNTrnID,12NCDXInvoiceDateIssue,50NCNFullInvoice,40NLNCustomerDescription,40NLNShipDescription,40NLNDestinationDescription,10NRFInvoiceTotalAmount,10NRIAdultsWithTransfer,10NRIKidsWithTransfer,10NRIFreeWithTransfer,10NRITotalWithTransfer,10NRIAdultsWithoutTransfer,10NRIKidsWithoutTransfer,10NRIFreeWithoutTransfer,10NRITotalWithoutTransfer,10NRILineAdults,10NRILineKids,10NRILineFree,10NRILineTotal,05NCNSelected", _
            "TrnID,Ημερομηνία έκδοσης,Παραστατικό,Πελάτης,Πλοίο,Προορισμός,Ποσό,Ε,Π,Δ,Σ,Ε,Π,Δ,Σ,Ε,Π,Δ,Σ,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
    End If
            
    'AddDummyLines grdInvoicesOutIndex, "99999", "A99/99/9999A", "AAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "-999999", "-999999"

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
        Case vbKeyP And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select
    
End Function

Private Sub Form_Load()

    PositionControls Me, True, grdInvoicesOutIndex
    ColorizeControls Me, True
    SetUpGrid lstIconList, grdInvoicesOutIndex
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtInvoiceCodeID, txtPersonID, txtDestinationID, txtShipID, lblCodeDescription, lblCodeBatch, lblHand, txtYesOrNoId
    ClearFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, mskInvoiceDateTripFrom, mskInvoiceDateTripTo, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtShipDescription, txtDestinationDescription, txtYesOrNoDescription
    EnableFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, mskInvoiceDateTripFrom, mskInvoiceDateTripTo, txtShipDescription, txtDestinationDescription, txtDestinationDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1

End Sub

Private Sub grdInvoicesOutIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdInvoicesOutIndex.Header.Buttons = True

End Sub

Private Sub grdInvoicesOutIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdInvoicesOutIndex.Header.Buttons = False
    
End Sub

Private Sub grdInvoicesOutIndex_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = ChangeEditButtonStatus(grdInvoicesOutIndex, Me.Tag, lRow, 1)

End Sub

Private Sub grdInvoicesOutIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdInvoicesOutIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdInvoicesOutIndex_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdInvoicesOutIndex.RowCount > 0 Then
        grdInvoicesOutIndex.CellIcon(grdInvoicesOutIndex.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdInvoicesOutIndex, 2, KeyCode, grdInvoicesOutIndex.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdInvoicesOutIndex)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdInvoicesOutIndex, False, "InvoiceTotalAmount", "Ποσό", "decimal", "LineTotal", "Ατομα", "integer")
      End If

End Sub

Private Sub grdInvoicesOutIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdInvoicesOutIndex", grdInvoicesOutIndex.LayoutCol

End Sub

Private Sub txtCodeShortDescriptionA_Change()

    If txtCodeShortDescriptionA.text = "" Then ClearFields txtInvoiceCodeID, lblCodeDescription, lblCodeBatch, lblHand

End Sub

Private Sub txtCodeShortDescriptionA_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub


Private Sub txtCodeShortDescriptionA_Validate(Cancel As Boolean)

    If txtInvoiceCodeID.text = "" And txtCodeShortDescriptionA.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then ClearFields txtDestinationID

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then ClearFields txtPersonID
    
End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    
End Sub


Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtPersonID.text = "" And txtPersonDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtShipDescription_Change()

    If txtShipDescription.text = "" Then ClearFields txtShipID

End Sub

Private Sub txtShipDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtShipDescription_Validate(Cancel As Boolean)

    If txtShipID = "" And txtShipDescription.text <> "" Then cmdIndex_Click 3

End Sub

Private Sub txtDescription_Change()

    If txtYesOrNoDescription.text = "" Then ClearFields txtYesOrNoId
    
End Sub


Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    
End Sub


Private Sub txtDescription_Validate(Cancel As Boolean)
    
    If txtYesOrNoId = "" And txtYesOrNoDescription.text <> "" Then cmdIndex_Click 4
    
End Sub


