VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form EmployeesLedger 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   19170
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   12525
      TabIndex        =   11
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "EmployeesLedger.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "EmployeesLedger.frx":001C
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
         TabIndex        =   13
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2640
         Index           =   0
         Left            =   1725
         TabIndex        =   19
         Top             =   3750
         Width           =   7965
         Begin UserControls.newText txtLastname 
            Height          =   465
            Left            =   2100
            TabIndex        =   1
            Top             =   825
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
            Left            =   7125
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   825
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
            PicNormal       =   "EmployeesLedger.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newDate mskDateFrom 
            Height          =   465
            Left            =   2100
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
         Begin UserControls.newDate mskDateTo 
            Height          =   465
            Left            =   3600
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2700
            Top             =   1800
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
            Left            =   2550
            Top             =   525
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
            Left            =   7500
            Top             =   975
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
            Left            =   1650
            Top             =   675
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
            Left            =   0
            Top             =   975
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
            TabIndex        =   25
            Top             =   2100
            Width           =   7965
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
            Left            =   3000
            TabIndex        =   24
            Top             =   75
            Width           =   4815
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
            TabIndex        =   23
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Ονοματεπώνυμο"
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
            TabIndex        =   22
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Διάστημα"
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
            TabIndex        =   21
            Top             =   1425
            Width           =   1215
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
            TabIndex        =   26
            Top             =   0
            Width           =   7965
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   14
         Top             =   8850
         Width           =   6015
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            TabIndex        =   16
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
         Height          =   1065
         Left            =   7875
         TabIndex        =   6
         Top             =   7650
         Width           =   4515
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
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "EmployeeID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtID 
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
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   450
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   2296
            Images          =   "EmployeesLedger.frx":05D2
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdEmployeesLedger 
         Height          =   7290
         Left            =   75
         TabIndex        =   4
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
         TabIndex        =   10
         Top             =   1125
         Width           =   2565
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
         TabIndex        =   9
         Top             =   525
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
         TabIndex        =   8
         Top             =   825
         Width           =   14940
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
         TabIndex        =   7
         Top             =   1125
         Width           =   14940
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Καρτέλα εργαζόμενου"
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
         TabIndex        =   5
         Top             =   75
         Width           =   4845
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
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "EmployeesLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

'Προοδευτικό υπόλοιπο
Dim curAccBalance As Currency

'Προηγούμενη περίοδος
Dim blnSoFarHasData As Boolean

'Ποσά
Dim curDebitSoFar As Currency
Dim curCreditSoFar As Currency
Dim curBalanceSoFar As Currency


Private Function AddTotalsSoFarForEmployeeToGrid()

    With grdEmployeesLedger
        .AddRow
        .CellValue(.RowCount, "ExpenseDescription") = "ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        '.CellValue(.RowCount, "Debit") = curDebitSoFar
        '.CellValue(.RowCount, "Credit") = curCreditSoFar
        '.CellValue(.RowCount, "Balance") = curAccBalance
        .AddRow
    End With
    
    InvertColorForNegativeNumbers grdEmployeesLedger, grdEmployeesLedger.RowCount - 1

End Function

Private Function CalculateSoFarTotalsForCredit(rstTransactions As Recordset, debitOrCredit As String)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Οφειλές - στήλη πίστωσης
        If debitOrCredit = "Credit" Then
            curCreditSoFar = curCreditSoFar + rstTransactions.Fields("Credit")
        End If
        'Πληρωμή- Στήλη χρέωσης
        If debitOrCredit = "Debit" Then
            curDebitSoFar = curDebitSoFar + rstTransactions.Fields("Debit")
        End If
    End With

End Function

Private Function DoJobs()

    Dim agreements As Recordset
    Dim payments As Recordset
    Dim mergedTable As Recordset
    
    If Not ValidateFields Then Exit Function
        
    Set agreements = GetAgreements(txtID.text, mskDateFrom.text, mskDateTo.text)
    Set payments = GetPayments(txtID.text, mskDateFrom.text, mskDateTo.text)
    
    Set mergedTable = MergeAgreementsAndPayments(mskDateFrom.text, agreements, payments)
    
    frmCriteria(0).Visible = False
    
    'If UBound(mergedArray, 2) = 0 Then
    '    If blnProcessing Then
    '        If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
    '        End If
    '    Else
    '        If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
    '        End If
    '    End If
    '    blnProcessing = False
    '    frmCriteria(0).Visible = True
    '    mskDateFrom.SetFocus
    '    Exit Function
    'End If
    
    RefreshList txtID.text, mskDateFrom.text, mskDateTo.text, mergedTable
    
End Function







Private Function AddPeriodTotalsToGrid(amount)

    With grdEmployeesLedger
        grdEmployeesLedger.AddRow
        .CellValue(.RowCount, "Amount") = amount
    End With
    
    InvertColorForNegativeNumbers grdEmployeesLedger, grdEmployeesLedger.RowCount
    
End Function

Private Function MergeAgreementsAndPayments(dateFrom As Date, agreements As Recordset, payments As Recordset)

    Dim intIndex As Integer
    Dim id As Integer
    Dim z As Integer
    
    Dim balance As Currency
    
    Dim strSQL As String
    Dim rstLedger As Recordset
    
    ReDim arrTransactions(13, 0) As String
    ReDim arrSortedTransactions(13, 0) As String
    
    intIndex = -1
    balance = 0
    
    Set rstLedger = CommonDB.OpenRecordset("EmployeesLedger")
    strSQL = "DELETE * FROM EmployeesLedger"
    CommonDB.Execute (strSQL)
    
    While Not agreements.EOF
       
        GoSub AddAgreementToArray
        
        While Not payments.EOF
        
            If payments.Fields(2) >= agreements.Fields(2) And payments.Fields(2) <= agreements.Fields(3) Then
                GoSub AddPaymentToArray
            End If
            
            payments.MoveNext
        
        Wend
            
        If Not payments.EOF Then
            payments.MoveFirst
        End If
        agreements.MoveNext
        
    Wend
    
    TempQuery.SQL = "SELECT * FROM EmployeesLedger"
    Set rstLedger = TempQuery.OpenRecordset()
    Set MergeAgreementsAndPayments = rstLedger
    
Exit Function
    
AddAgreementToArray:
    balance = balance - agreements.Fields(4)
    id = MainSaveRecord("CommonDB", "EmployeesLedger", True, "", "ID", 0, "Y", "", agreements.Fields(2), agreements.Fields(3), "", "", "", "", "", 0, agreements.Fields(4), balance)
    
    Return
    
AddPaymentToArray:
    balance = balance + payments.Fields(7)
    id = MainSaveRecord("CommonDB", "EmployeesLedger", True, "", "ID", 0, "", "Y", "", "", payments.Fields(2), payments.Fields(3), payments.Fields(4), payments.Fields(5), payments.Fields(6), payments.Fields(7), 0, balance)
    
    Return
    
End Function

Private Function GetPayments(personID As String, fromDate As String, toDate As String)

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
    
    'Recordsets
    Dim rstPayments As Recordset

    'Αρχικές τιμές
    intIndex = 0
    
    'Κυρίως διαδικασία
    strSQL = "SELECT t.ID, Date, DateRefersTo, pc.Description, pw.Description, tp.Description, Remarks, Amount " _
        & "FROM (((EmployeesTransactions t " _
        & "INNER JOIN PaymentCategories pc ON t.PaymentCategoryID = pc.ID) " _
        & "INNER JOIN PaymentWays pw ON t.PaymentWayID = pw.ID) " _
        & "INNER JOIN TransactionTypes tp ON t.TransactionTypeID = tp.ID) "
 
    'Εργαζόμενος
    strThisParameter = "intEmployeeID Integer"
    strThisQuery = "EmployeeID = intEmployeeID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtID.text)
    
    'Από
    strThisParameter = "datFromDate Date"
    strThisQuery = "DateRefersTo >= datFromDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = fromDate
    
    'Εως
    strThisParameter = "datToDate Date"
    strThisQuery = "DateRefersTo <= datToDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = toDate
    
    'Ταξινόμηση
    strOrder = " ORDER BY DateRefersTo"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    'Ανοίγω το recordset
    Set rstPayments = TempQuery.OpenRecordset()
    
    Set GetPayments = rstPayments
    
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
    DisplayErrorMessage True, Err.Description

End Function

Private Function RefreshList(personID As String, fromDate As String, toDate As String, table As Recordset)

    'Loca variables
    Dim lngRow As Long
    Dim blnSoFarHasData As Boolean
    Dim blnPeriodHasData As Boolean

    'Αρχικές τιμές
    lngRow = 1
    blnProcessing = True
    
    'Πλέγμα
    With grdEmployeesLedger
        .Clear
        .Redraw = False
    End With
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, table
    
    'Προσωρινά
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(2).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With table
        If CalculateSoFarTotals(fromDate, table) Then
            blnSoFarHasData = True
            AddTotalsSoFarToGrid
        End If

        'grdEmployeesLedger.AddRow , , , , , , , mergedArray.RecordCount
        'Do Until .EOF
 '           grdEmployeesLedger.CellValue(lngRow, "ID") = !id
 '           grdEmployeesLedger.CellValue(lngRow, "Agreement") = !A
 '           grdEmployeesLedger.CellValue(lngRow, "Transaction") = !T
 '           grdEmployeesLedger.CellValue(lngRow, "From") = !From
 '           grdEmployeesLedger.CellValue(lngRow, "To") = !To
 '           grdEmployeesLedger.CellValue(lngRow, "Date") = !PaymentDate
 '           grdEmployeesLedger.CellValue(lngRow, "PaymentCategory") = !PaymentCategory
 '           grdEmployeesLedger.CellValue(lngRow, "PaymentWay") = !PaymentWay
 '           grdEmployeesLedger.CellValue(lngRow, "TransactionType") = !TransactionType
 '           grdEmployeesLedger.CellValue(lngRow, "Remarks") = !Remarks
 '           grdEmployeesLedger.CellValue(lngRow, "Debit") = !Debit
 '           grdEmployeesLedger.CellValue(lngRow, "Credit") = !Credit
 '           grdEmployeesLedger.CellValue(lngRow, "Balance") = !balance
 '           lngRow = lngRow + 1
 '           DoEvents
 '           If Not blnProcessing Then Exit Do
 '           .MoveNext
 '       Loop
    End With
    
    'Τελικές ενέργειες
    frmProgress.Visible = False
    grdEmployeesLedger.Redraw = True
    grdEmployeesLedger.SetFocus
    'grdEmployeesLedger.SetCurCell 1, 1
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(2).Caption = "Νέα αναζήτηση"
    blnProcessing = False
    
End Function

Private Function AddTotalsSoFarToGrid()

    AddTotalsSoFarForEmployeeToGrid
        
End Function

Private Function CalculateSoFarTotals(fromDate As String, rstTransactions As Recordset)
    
    'Ποσά
    curDebitSoFar = 0
    curCreditSoFar = 0
    curBalanceSoFar = 0

    CalculateSoFarTotals = False
    
    With rstTransactions
        While Not .EOF
            If Not blnProcessing Then Exit Function
            If rstTransactions.Fields("To") < CDate(fromDate) Then
                'Οφειλές - Στήλη πίστωσης
                If rstTransactions.Fields("Credit") <> 0 Then CalculateSoFarTotalsForCredit rstTransactions, "Credit"
                'Πληρωμές - Στήλη χρέωσης
                If rstTransactions.Fields("Debit") <> 0 Then CalculateSoFarTotalsForCredit rstTransactions, "Debit"
                'Εχω εγγραφές!
                CalculateSoFarTotals = True
                'Επόμενη εγγραφή
                rstTransactions.MoveNext
                'Async!
                DoEvents
                'Πρόοδος
                UpdateProgressBar Me
            Else
                curAccBalance = curDebitSoFar - curCreditSoFar
                Exit Function
            End If
        Wend
        'Υπόλοιπο
        curAccBalance = curDebitSoFar - curCreditSoFar
    End With

End Function


Private Function UpdateCriteriaLabels(DateIssueFrom, DateIssueTo, Person)

    Dim strCriteriaA As String

    strCriteriaA = IIf(DateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & DateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(DateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & DateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(Person = "", "Εργαζόμενος [ ΟΛΟΙ ] ", "Εργαζόμενος [ " & Person & " ] ")
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Function EditRecord()

    Dim rstRecordset As Recordset
    
    Set rstRecordset = EmployeesTransactions.SeekRecord(grdEmployeesLedger.CellValue(grdEmployeesLedger.CurRow, "ID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    EmployeesTransactions.DoPostFoundJobs rstRecordset: EmployeesTransactions.Show 1, Me
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            DoJobs
        Case 1
            EditRecord
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Συναλλασόμενος
    If txtLastname.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtLastname.SetFocus
        Exit Function
    End If
    
    'Από
    If mskDateFrom.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDateFrom.SetFocus
        Exit Function
    End If
    
    'Εως
    If mskDateTo.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDateTo.SetFocus
        Exit Function
    End If
    
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

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdEmployeesLedger
        frmCriteria(0).Visible = True
        txtLastname.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
        
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function GetAgreements(personID As String, fromDate As String, toDate As String)

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
    
    'Recordsets
    Dim rstAgreements As Recordset

    'Αρχικές τιμές
    intIndex = 0
    
    'Κυρίως διαδικασία
    strSQL = "SELECT ID, EmployeeID, DateFrom, DateTo, Amount " _
        & "FROM EmployeesAgreements "
 
    'Εργαζόμενος
    strThisParameter = "intEmployeeID Integer"
    strThisQuery = "EmployeeID = intEmployeeID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtID.text)
    
    'Από
    strThisParameter = "datFromDate Date"
    strThisQuery = "DateFrom >= datFromDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = fromDate
    
    'Εως
    strThisParameter = "datToDate Date"
    strThisQuery = "DateTo <= datToDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = toDate
    
    'Ταξινόμηση
    strOrder = " ORDER BY DateTo"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    'Ανοίγω το recordset
    Set rstAgreements = TempQuery.OpenRecordset()
    
    Set GetAgreements = rstAgreements
    
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
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub cmdIndex_Click(Index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Customers - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Employees", "Lastname", "String", txtLastname.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "Επώνυμο", "Ονομα", 0, 40, 40, 1, 0, 0)
                txtID.text = tmpTableData.strCode
                txtLastname.text = tmpTableData.strFirstField & " " & tmpTableData.strSecondField
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdEmployeesLedger, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdEmployeesLedger"), _
            "05ΝCNID,10NCNAgreement,20NCNTransaction,20NCDFrom,20NCDTo,20NCDDate,20NLNPaymentCategory,20NLNPaymentWay,20NLNTransactionType,20NLNRemarks,10NRFDebit,10NRFCredit,10NRFBalance,05NCNSelected", _
            "ID,Α,Τ,Από,Εως,Ημερομηνία,Κατηγορία αμοιβής,Τρόπος πληρωμής,Τύπος συναλλαγής,Παρατηρήσεις,Χρέωση,Πίστωση,Υπόλοιπο,Ε"
        Me.Refresh
    End If
            
    'AddDummyLines grdEmployeesLedger, "12345", "A", "A", "Α99/99/9999Α", "Α99/99/9999Α", "Α99/99/9999Α", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "999999", "999999", "999999", "AAA"
    
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
        Case vbKeyC And CtrlDown And cmdButton(0).Enabled
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

    SetUpGrid lstIconList, grdEmployeesLedger
    PositionControls Me, True, grdEmployeesLedger
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtID
    ClearFields mskDateFrom, mskDateTo, txtLastname
    UpdateButtons Me, 3, 1, 0, 0, 1
    
End Sub

Private Sub grdEmployeesLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdEmployeesLedger_ColHeaderMouseEnter(ByVal lCol As Long)

    grdEmployeesLedger.Header.Buttons = True
    
End Sub


Private Sub grdEmployeesLedger_ColHeaderMouseLeave(ByVal lCol As Long)

    grdEmployeesLedger.Header.Buttons = False

End Sub


Private Sub grdEmployeesLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    
    cmdButton(1).Enabled = ChangeEditButtonStatus(grdEmployeesLedger, Me.Tag, lRow, 1)

End Sub

Private Sub grdEmployeesLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdEmployeesLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal x As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdEmployeesLedger_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdEmployeesLedger.RowCount > 0 Then
        grdEmployeesLedger.CellIcon(grdEmployeesLedger.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdEmployeesLedger, 2, KeyCode, grdEmployeesLedger.CurRow, "ID"))
        lblSelectedGridLines.Caption = CountSelected(grdEmployeesLedger)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdEmployeesLedger, False, "Amount", "Σύνολο", "decimal")
    End If

End Sub

Private Sub grdEmployeesLedger_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdEmployeesLedger", grdEmployeesLedger.LayoutCol

End Sub

Private Sub txtLastname_Change()

    If txtLastname.text = "" Then txtID.text = ""

End Sub

Private Sub txtLastname_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub

Private Sub txtLastname_Validate(Cancel As Boolean)

    If txtID.text = "" And txtLastname.text <> "" Then cmdIndex_Click 0: If txtID.text = "" Then Cancel = True

End Sub

