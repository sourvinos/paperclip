VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form EmployeesBalanceSheet 
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
      TabIndex        =   10
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "EmployeesBalanceSheet.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "EmployeesBalanceSheet.frx":001C
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
         TabIndex        =   12
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
         Height          =   2115
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   3225
         Width           =   5190
         Begin UserControls.newDate mskDateFrom 
            Height          =   465
            Left            =   1800
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
         Begin UserControls.newDate mskDateTo 
            Height          =   465
            Left            =   3300
            TabIndex        =   2
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2700
            Top             =   1275
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
            Left            =   2250
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
            Left            =   4725
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
            Index           =   1
            Left            =   1350
            Top             =   1050
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
            TabIndex        =   22
            Top             =   1575
            Width           =   7665
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
            Left            =   2475
            TabIndex        =   21
            Top             =   75
            Width           =   2565
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
            TabIndex        =   20
            Top             =   75
            Width           =   1665
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
            TabIndex        =   19
            Top             =   900
            Width           =   915
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
            TabIndex        =   23
            Top             =   0
            Width           =   7665
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   13
         Top             =   8850
         Width           =   6015
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
         TabIndex        =   5
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            Images          =   "EmployeesBalanceSheet.frx":0038
            Version         =   131072
            KeyCount        =   2
            Keys            =   "ÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdEmployeesBalanceSheet 
         Height          =   7290
         Left            =   75
         TabIndex        =   3
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1125
         Width           =   14940
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ισοζύγιο εργαζομένων"
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
         TabIndex        =   4
         Top             =   75
         Width           =   4965
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
Attribute VB_Name = "EmployeesBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList(txtID.text, mskDateFrom.text, mskDateTo.text) > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskDateFrom.text, mskDateTo.text
            EnableGrid grdEmployeesBalanceSheet, False
            HighlightRow grdEmployeesBalanceSheet, 1, 1, "", True
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

Private Function AddPeriodTotalsToGrid(amount)

    With grdEmployeesBalanceSheet
        grdEmployeesBalanceSheet.AddRow
        .CellValue(.RowCount, "Payments") = amount
    End With
    
    InvertColorForNegativeNumbers grdEmployeesBalanceSheet, grdEmployeesBalanceSheet.RowCount
    
End Function

Private Function AddTotalsSoFarToGrid()

    'If txtInvoiceMasterRefersTo.text = "1" Then AddTotalsSoFarForExpensesToGrid
    'If txtInvoiceMasterRefersTo.text = "2" Then AddTotalsSoFarForSalesToGrid
        
End Function

Private Function LoadEmployees()



End Function

Private Function UpdateCriteriaLabels(DateIssueFrom, DateIssueTo)

    Dim strCriteriaA As String

    strCriteriaA = IIf(DateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & DateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(DateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & DateIssueTo & " ] ")
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Function EditRecord()

    Dim rstRecordset As Recordset
    
    Set rstRecordset = EmployeesTransactions.SeekRecord(grdEmployeesBalanceSheet.CellValue(grdEmployeesBalanceSheet.CurRow, "ID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    EmployeesTransactions.DoPostFoundJobs rstRecordset: EmployeesTransactions.Show 1, Me
    
End Function

Private Function CreateUnicodeFileForCustomers(strReportTitle, strReportSubTitle1, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim intAdults As Integer
    Dim intKids As Integer
    Dim intFree As Integer
    Dim curAdultsAmount As Currency
    Dim curKidsAmount As Currency
    Dim curDebit As Currency
    Dim curCredit As Currency
    Dim curBalance As Currency

    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    
    GoSub Headers
    
    'Πλέγμα
    With grdEmployeesBalanceSheet
        For lngRow = 1 To grdEmployeesBalanceSheet.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, _
                format(.CellText(lngRow, "Date"), "dd/mm/yy"); _
                Tab(10); .CellText(lngRow, "InvoiceDetails"); _
                Tab(24); Left(.CellText(lngRow, "Destination"), 21); _
                Tab(53 - Len((format(.CellText(lngRow, "Adults"), "#,##0")))); format(.CellText(lngRow, "Adults"), "#,##0"); _
                Tab(61 - Len((format(.CellText(lngRow, "Kids"), "#,##0")))); format(.CellText(lngRow, "Kids"), "#,##0"); _
                Tab(67 - Len((format(.CellText(lngRow, "Free"), "#,##0")))); format(.CellText(lngRow, "Free"), "#,##0"); _
                Tab(81 - Len((format(.CellText(lngRow, "AdultsAmount"), "#,##0.00")))); format(.CellText(lngRow, "AdultsAmount"), "#,##0.00"); _
                Tab(95 - Len((format(.CellText(lngRow, "KidsAmount"), "#,##0.00")))); format(.CellText(lngRow, "KidsAmount"), "#,##0.00"); _
                Tab(109 - Len((format(.CellText(lngRow, "Debit"), "#,##0.00")))); format(.CellText(lngRow, "Debit"), "#,##0.00"); _
                Tab(123 - Len((format(.CellText(lngRow, "Credit"), "#,##0.00")))); format(.CellText(lngRow, "Credit"), "#,##0.00"); _
                Tab(137 - Len((format(.CellText(lngRow, "Balance"), "#,##0.00")))); format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            'Σύνολα
            If .CellText(lngRow, "TrnID") <> "" Then
                intAdults = intAdults + .CellValue(lngRow, "Adults")
                intKids = intKids + .CellValue(lngRow, "Kids")
                intFree = intFree + .CellValue(lngRow, "Free")
                curAdultsAmount = curAdultsAmount + .CellValue(lngRow, "AdultsAmount")
                curKidsAmount = curKidsAmount + .CellValue(lngRow, "KidsAmount")
                curDebit = curDebit + .CellValue(lngRow, "Debit")
                curCredit = curCredit + .CellValue(lngRow, "Credit")
                curBalance = curDebit - curCredit
            End If
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > intReportDetailLines Then
                Print #1, ""
                Print #1, Space(23) & "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                Tab(53 - Len(format(intAdults, "#,##0"))); format(intAdults, "#,##0"); _
                Tab(61 - Len(format(intKids, "#,##0"))); format(intKids, "#,##0"); _
                Tab(67 - Len(format(intFree, "#,##0"))); format(intFree, "#,##0"); _
                Tab(81 - Len(format(curAdultsAmount, "#,##0.00"))); format(curAdultsAmount, "#,##0.00"); _
                Tab(95 - Len(format(curKidsAmount, "#,##0.00"))); format(curKidsAmount, "#,##0.00"); _
                Tab(109 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                Tab(123 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                Tab(137 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                
                GoSub Headers
                
                Print #1, Space(23) & "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(53 - Len(format(intAdults, "#,##0"))); format(intAdults, "#,##0"); _
                    Tab(61 - Len(format(intKids, "#,##0"))); format(intKids, "#,##0"); _
                    Tab(67 - Len(format(intFree, "#,##0"))); format(intFree, "#,##0"); _
                    Tab(81 - Len(format(curAdultsAmount, "#,##0.00"))); format(curAdultsAmount, "#,##0.00"); _
                    Tab(95 - Len(format(curKidsAmount, "#,##0.00"))); format(curKidsAmount, "#,##0.00"); _
                    Tab(109 - Len(format(curDebit, "#,##0.00"))); format(curDebit, "#,##0.00"); _
                    Tab(123 - Len(format(curCredit, "#,##0.00"))); format(curCredit, "#,##0.00"); _
                    Tab(137 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    CreateUnicodeFileForCustomers = True
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 136, intPageNo, strReportTitle, strReportSubTitle1
    PrintColumnHeadings 10, "ΣΤΟΙΧΕΙΟ", 47, "ΕΝΗΛΙ-", 57, "ΠΑΙ-", 64, "ΔΩ-", 73, "ΧΡΕΩΣΕΙΣ", 87, "ΧΡΕΩΣΕΙΣ", 103, "ΣΥΝΟΛΟ"
    PrintColumnHeadings 1, "ΗΜΕΡ/ΝΙΑ", 10, "ΣΕΙΡΑ - Νο", 24, "ΠΡΟΟΡΙΣΜΟΣ", 50, "ΚΕΣ", 58, "ΔΙΑ", 63, "ΡΕΑΝ", 73, "ΕΝΗΛΙΚΩΝ", 88, "ΠΑΙΔΙΩΝ", 102, "ΧΡΕΩΣΗΣ", 116, "ΠΙΣΤΩΣΗ", 129, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 11
      
    Return
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
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

Function AddNumberFormats(sheet As Object, grid As iGrid, format As String, rowOffsetFromTop As Long, ParamArray Columns() As Variant)

    Dim column As Long
    Dim Row As Long
    
    'Excel
    With sheet
        For column = 0 To UBound(Columns)
            Select Case format
                Case "Floats"
                    For Row = 1 To grid.RowCount
                        .Range(Columns(column) & Row + rowOffsetFromTop).NumberFormat = "#,##0.00_);[Red]#,##0.00 "
                    Next Row
                Case "Integers"
                    For Row = 1 To grid.RowCount
                        .Range(Columns(column) & Row + rowOffsetFromTop).NumberFormat = "#,##0_);[Red]#,##0 "
                    Next Row
                Case "Dates"
                    For Row = 1 To grid.RowCount
                        .Range(Columns(column) & Row + rowOffsetFromTop).NumberFormat = "dd-mm-yyyy"
                    Next Row
            End Select
        Next column
    End With

End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
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
        ClearFields grdEmployeesBalanceSheet
        frmCriteria(0).Visible = True
        mskDateFrom.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
        
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function RefreshList(personID As String, fromDate As String, toDate As String)

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
    Dim curTotalPayments As Currency
    
    'Recordsets
    Dim rstTransactions As Recordset

    Dim blnPeriodHasData As Boolean
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    
    frmCriteria(0).Visible = False
    blnPeriodHasData = False
    
    'Πλέγμα
    With grdEmployeesBalanceSheet
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT EmployeesTransactions.EmployeeID, Lastname, Firstname, Sum(EmployeesTransactions.Amount) AS SumOfAmount " _
        & "FROM employeestransactions " _
        & "INNER JOIN Employees ON EmployeesTransactions.EmployeeID = Employees.ID "
        
    'Από
    strThisParameter = "datFromDate Date"
    strThisQuery = "EmployeesTransactions.Date >= datFromDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = fromDate
    
    'Εως
    strThisParameter = "datToDate Date"
    strThisQuery = "EmployeesTransactions.Date <= datToDate"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = toDate
    
    'Ταξινόμηση
    strOrder = " GROUP BY EmployeesTransactions.EmployeeID, Lastname, Firstname"
    
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
    Set rstTransactions = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstTransactions.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstTransactions
    
    'Προσωρινά
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(2).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstTransactions
        If .EOF = False Then
            Do While Not .EOF
                If Not blnProcessing Then Exit Do
                grdEmployeesBalanceSheet.AddRow
                grdEmployeesBalanceSheet.CellValue(grdEmployeesBalanceSheet.RowCount, "ID") = !employeeID
                grdEmployeesBalanceSheet.CellValue(grdEmployeesBalanceSheet.RowCount, "Lastname") = !Lastname
                grdEmployeesBalanceSheet.CellValue(grdEmployeesBalanceSheet.RowCount, "Firstname") = !Firstname
                grdEmployeesBalanceSheet.CellValue(grdEmployeesBalanceSheet.RowCount, "Payments") = !SumOfAmount
                lngRowCount = lngRowCount + 1
                curTotalPayments = curTotalPayments + !SumOfAmount
                UpdateProgressBar Me
                rstTransactions.MoveNext
                DoEvents
            Loop
        End If
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdEmployeesBalanceSheet
        RefreshList = False
    Else
        RefreshList = grdEmployeesBalanceSheet.RowCount
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        grdEmployeesBalanceSheet.Redraw = True
        grdEmployeesBalanceSheet.AddRow
        AddPeriodTotalsToGrid curTotalPayments
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
    blnError = True
    ClearFields grdEmployeesBalanceSheet, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function
Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdEmployeesBalanceSheet, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdEmployeesBalanceSheet"), _
            "05ΝCNID,10NLNLastname,20NLNFirstName,10NRFAgreements,10NRFPayments,10NRFBalance,05NCNSelected", _
            "ID,Επώνυμο,Ονομα,Συμφωνία,Πληρωμές,Υπόλοιπο,Ε"
        Me.Refresh
    End If
            
    'AddDummyLines grdEmployeesBalanceSheet, "12345", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "9999999", "9999999", "9999999"
    
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

    SetUpGrid lstIconList, grdEmployeesBalanceSheet
    PositionControls Me, True, grdEmployeesBalanceSheet
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtID
    ClearFields mskDateFrom, mskDateTo
    UpdateButtons Me, 3, 1, 0, 0, 1
    
End Sub

Private Sub grdEmployeesBalanceSheet_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdEmployeesBalanceSheet_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    
    cmdButton(1).Enabled = ChangeEditButtonStatus(grdEmployeesBalanceSheet, Me.Tag, lRow, 1)

End Sub

Private Sub grdEmployeesBalanceSheet_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdEmployeesBalanceSheet_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal x As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdEmployeesBalanceSheet_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdEmployeesBalanceSheet.RowCount > 0 Then
        grdEmployeesBalanceSheet.CellIcon(grdEmployeesBalanceSheet.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdEmployeesBalanceSheet, 2, KeyCode, grdEmployeesBalanceSheet.CurRow, "ID"))
        lblSelectedGridLines.Caption = CountSelected(grdEmployeesBalanceSheet)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdEmployeesBalanceSheet, False, "Payments", "Πληρωμές", "decimal")
    End If

End Sub

Private Sub grdEmployeesBalanceSheet_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdEmployeesBalanceSheet", grdEmployeesBalanceSheet.LayoutCol

End Sub

