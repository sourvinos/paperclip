VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form EmployeesTransactionsIndex 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
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
      TabIndex        =   14
      Top             =   8025
      Width           =   6090
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
      Left            =   8550
      TabIndex        =   11
      Top             =   6450
      Width           =   4515
      Begin VB.TextBox txtEmployeeID 
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   75
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "EmployeeID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   450
         _ExtentX        =   953
         _ExtentY        =   953
      End
   End
   Begin VB.Frame frmCriteria 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   2640
      Index           =   0
      Left            =   525
      TabIndex        =   3
      Top             =   4875
      Width           =   7965
      Begin UserControls.newDate mskDateFrom 
         Height          =   465
         Left            =   2100
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
         Left            =   3600
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
         Left            =   7125
         TabIndex        =   23
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
         PicNormal       =   "EmployeesTransactionsIndex.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin UserControls.newText txtEmployeeLastname 
         Height          =   465
         Left            =   2100
         TabIndex        =   24
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   315
         Index           =   5
         Left            =   2550
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
         Index           =   3
         Left            =   450
         TabIndex        =   25
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   1
         Left            =   7500
         Top             =   1125
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
         Left            =   1650
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
         Left            =   5025
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
         TabIndex        =   10
         Top             =   0
         Width           =   7965
      End
   End
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   525
      TabIndex        =   0
      Top             =   3675
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
         Picture         =   "EmployeesTransactionsIndex.frx":059A
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "EmployeesTransactionsIndex.frx":05B6
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
   Begin iGrid300_10Tec.iGrid grdEmployeesTransactionsIndex 
      Height          =   6165
      Left            =   450
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1425
      Width           =   12915
      _ExtentX        =   22781
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
      Caption         =   "Εύρεση κινήσεων εργαζομένων"
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
      TabIndex        =   22
      Top             =   150
      Width           =   6885
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
      TabIndex        =   21
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
      Left            =   4200
      TabIndex        =   20
      Top             =   1050
      Width           =   9165
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   13350
      Top             =   3975
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
Attribute VB_Name = "EmployeesTransactionsIndex"
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
        ClearFields grdEmployeesTransactionsIndex, lblRecordCount, lblCriteria
        frmCriteria(0).Visible = True
        mskDateFrom.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskDateFrom.text, mskDateTo.text, txtEmployeeLastname.text
            EnableGrid grdEmployeesTransactionsIndex, False
            HighlightRow grdEmployeesTransactionsIndex, 1, 1, "", True
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

Private Function UpdateCriteriaLabels(InvoiceDateIssueFrom, InvoiceDateIssueTo, PersonDescription)

    Dim strCriteriaA As String

    strCriteriaA = IIf(InvoiceDateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & InvoiceDateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(InvoiceDateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & InvoiceDateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(PersonDescription = "", "Συναλλασόμενος [ ΟΛΟΙ ]", "Συναλλασόμενος [ " & PersonDescription & " ]")
    
    lblCriteria.Caption = strCriteriaA
    
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
    With grdEmployeesTransactionsIndex
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT EmployeesTransactions.ID, Date, Lastname, Firstname, Amount " _
        & "FROM EmployeesTransactions " _
        & "INNER JOIN Employees ON EmployeesTransactions.EmployeeID = Employees.ID "
        
    'Από
    If mskDateFrom.text <> "" Then
        strThisParameter = "datDateFrom Date"
        strThisQuery = "Date >= datDateFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskDateFrom.text
    End If
        
    'Εως
    If mskDateTo.text <> "" Then
        strThisParameter = "datDateTo Date"
        strThisQuery = "Date <= datDateTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskDateTo.text
    End If
    
    'Εργαζόμενος
    If txtEmployeeID.text <> "" Then
        strThisParameter = "intEmployeeID Integer"
        strThisQuery = "EmployeeID = intEmployeeID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtEmployeeID.text)
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY Date, Lastname"
    
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
        grdEmployeesTransactionsIndex.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdEmployeesTransactionsIndex.CellValue(lngRow, "ID") = !ID
            grdEmployeesTransactionsIndex.CellValue(lngRow, "Date") = !Date
            grdEmployeesTransactionsIndex.CellValue(lngRow, "Lastname") = !Lastname
            grdEmployeesTransactionsIndex.CellValue(lngRow, "Firstname") = !Firstname
            grdEmployeesTransactionsIndex.CellValue(lngRow, "Amount") = !Amount
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        rstRecordset.Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdEmployeesTransactionsIndex
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
    ClearFields grdEmployeesTransactionsIndex, frmProgress
    DisplayErrorMessage True, Err.Description

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

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub cmdIndex_Click(Index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
        'Εργαζόμενος
        Set tmpRecordset = CheckForMatch("CommonDB", "Employees", "Lastname", "String", txtEmployeeLastname.text)
        If tmpRecordset.RecordCount > 0 Then
            tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "Επωνυμία", "Α.Φ.Μ.", 0, 40, 15, 1, 0, 1)
            txtEmployeeID.text = tmpTableData.strCode
            txtEmployeeLastname.text = tmpTableData.strFirstField
        End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdEmployeesTransactionsIndex, False, 44, GetSetting(strApplicationName, "Layout Strings", "grdEmployeesTransactionsIndex"), _
            "05NCNID,12NCDDate,40NLNLastname,40NLNFirstname,10NRFAmount", _
            "ID,Ημερομηνία,Επώνυμο,Ονομα,Ποσό"
        Me.Refresh
        mskDateFrom.SetFocus
    End If
    
    'AddDummyLines grdEmployeesTransactionsIndex, "99999", "A99/99/9999A", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "-9.999.999,99"

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

    PositionControls Me, False, grdEmployeesTransactionsIndex
    ColorizeControls Me, False, False
    SetUpGrid lstIconList, grdEmployeesTransactionsIndex
    ClearFields txtEmployeeID, lblRecordCount, lblCriteria
    ClearFields mskDateFrom, mskDateTo, txtEmployeeLastname
    EnableFields mskDateFrom, mskDateTo
    UpdateButtons Me, 3, 1, 0, 0, 1

End Sub

Private Sub grdEmployeesTransactionsIndex_ColHeaderMouseEnter(ByVal lCol As Long)

    grdEmployeesTransactionsIndex.Header.Buttons = True

End Sub

Private Sub grdEmployeesTransactionsIndex_ColHeaderMouseLeave(ByVal lCol As Long)

    grdEmployeesTransactionsIndex.Header.Buttons = False

End Sub


Private Sub grdEmployeesTransactionsIndex_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdEmployeesTransactionsIndex_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Function EditRecord()
    
    If Not grdEmployeesTransactionsIndex.Enabled Then Exit Function
        
    Dim rstRecordset As Recordset
    
    Set rstRecordset = EmployeesTransactions.SeekRecord(grdEmployeesTransactionsIndex.CellValue(grdEmployeesTransactionsIndex.CurRow, "ID"))
                
    If rstRecordset.RecordCount = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(9), 1) Then
        End If
        Exit Function
    End If
    
    EmployeesTransactions.DoPostFoundJobs rstRecordset
    
    Unload Me

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

Private Sub grdEmployeesTransactionsIndex_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdEmployeesTransactionsIndex", grdEmployeesTransactionsIndex.LayoutCol

End Sub

Private Sub txtEmployeeLastname_Change()

    If txtEmployeeLastname.text = "" Then ClearFields txtEmployeeID

End Sub


Private Sub txtEmployeeLastname_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub


Private Sub txtEmployeeLastname_Validate(Cancel As Boolean)

    If txtEmployeeID.text = "" And txtEmployeeLastname.text <> "" Then cmdIndex_Click 0: If txtEmployeeID.text = "" Then Cancel = True

End Sub


