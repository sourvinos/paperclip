VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesBanks 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   15735
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   9
      Top             =   7875
      Width           =   7515
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Δημιουργία"
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
         TabIndex        =   11
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Αποθήκευση"
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Διαγραφή"
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Ακυρο"
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
      Height          =   1065
      Left            =   7875
      TabIndex        =   4
      Top             =   6225
      Width           =   4515
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Banks.ID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   0
         Top             =   450
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   2296
         Images          =   "TablesBanks.frx":0000
         Version         =   131072
         KeyCount        =   2
         Keys            =   "ÿ"
      End
   End
   Begin iGrid300_10Tec.iGrid grdBanks 
      Height          =   6240
      Left            =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   11007
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
   Begin UserControls.newText txtDescription 
      Height          =   465
      Left            =   2400
      TabIndex        =   2
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
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
   Begin UserControls.newText txtAccountsCode 
      Height          =   465
      Left            =   2400
      TabIndex        =   7
      Top             =   1650
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   15
      Text            =   "AAAAAAAAAAAAAAA"
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
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Κωδ. Γεν. Λογιστικής"
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
      Index           =   13
      Left            =   450
      TabIndex        =   8
      Top             =   1725
      Width           =   1515
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   4
      Left            =   7350
      Top             =   900
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   12450
      Top             =   3825
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   7050
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3750
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   0
      Left            =   10275
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
      Top             =   1500
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
      Left            =   1950
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Τράπεζες"
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
      TabIndex        =   3
      Top             =   75
      Width           =   2115
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ονομασία"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1515
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
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   4575
      Top             =   0
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
Attribute VB_Name = "TablesBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow As Long

Dim banks As New ADODB.recordset
Dim bank As New ADODB.recordset

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtID, txtDescription, txtAccountsCode
            DisableFields txtDescription, txtAccountsCode
            grdBanks.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    'If MainDeleteRecord("CommonDB", "Banks", strApplicationName, "BankID", txtID.text, "True") Then
    '    PopulateGrid grdBanks, banks
    '    HighlightRow grdBanks, lngSelectedRow, 2, "", True
    '    ClearFields txtID, txtDescription, txtAccountsCode
    '    DisableFields txtDescription, txtAccountsCode
    '    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    'End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtID, txtDescription, txtAccountsCode
    EnableFields txtDescription, txtAccountsCode
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtDescription.SetFocus

End Function

Private Function PopulateFields(record As ADODB.recordset, ParamArray fields() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(fields)
        If TypeOf fields(intLoop) Is TextBox Or TypeOf fields(intLoop) Is newText Then
            fields(intLoop).text = IIf(Not IsNull(record.fields(intLoop)), record.fields(intLoop), "")
        End If
        If TypeOf fields(intLoop) Is newFloat Then
            fields(intLoop).text = format(record.fields(intLoop), "#,##0.00")
        End If
        If TypeOf fields(intLoop) Is newInteger Then
            fields(intLoop).text = format(record.fields(intLoop), "#,##0")
        End If
        If TypeOf fields(intLoop) Is Label Then
            fields(intLoop).Caption = record.fields(intLoop)
        End If
        If TypeOf fields(intLoop) Is CheckBox Then
            fields(intLoop).Value = IIf(record.fields(intLoop), 1, 0)
        End If
        If TypeOf fields(intLoop) Is OptionButton Then
            fields(intLoop).Value = IIf(record.fields(intLoop), 1, 0)
        End If
        If TypeOf fields(intLoop) Is newDate Then
            fields(intLoop).text = format(record.fields(intLoop), "dd/mm/yyyy")
        End If
    Next intLoop


End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ονομασία
    If Len(txtDescription.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDescription.SetFocus
        Exit Function
    End If
    
    'Κωδ Γεν. Λογιστικής
    'If Len(txtAccountsCode.text) > 0 And Len(txtAccountsCode.text) <> intAccountsCodeLength Then
    '    If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
    '    End If
    '    txtAccountsCode.SetFocus
    '    Exit Function
    'End If
    
    ValidateFields = True

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If SaveRecordToDB("Banks", _
            "Id", "Description", "AccountsCode", "ShowInList", "Username", _
            txtID.text, txtDescription.text, txtAccountsCode.text, 1, "Admin") Then
        Set banks = getDataFromDB("SELECT Id, Description FROM Banks")
        PopulateGrid grdBanks, banks
        HighlightRow grdBanks, lngSelectedRow, 2, txtDescription.text, True
        lngSelectedRow = 0
        ClearFields txtID, txtDescription, txtAccountsCode
        DisableFields txtDescription, txtAccountsCode
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If
    
End Function

Private Function SaveRecordToDB(table As String, ParamArray fields() As Variant)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim strSQL As String
    Dim strFields As String
    Dim fieldCount As Integer
    Dim recordset As ADODB.recordset
    
    Set recordset = New ADODB.recordset
    
    With recordset
        .Source = table
        .ActiveConnection = Connection
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open Options:=adCmdTable
    End With
    
    fieldCount = UBound(fields) + 1
        
    If fields(fieldCount / 2) = "" Then GoSub Add Else GoSub Update
    
    Set recordset = Nothing
    
    SaveRecordToDB = True
    
    Exit Function
    
Add:

    strSQL = "INSERT INTO " + table + " ("
    
    For intLoop = 2 To fieldCount / 2
        strFields = strFields + fields(intLoop - 1) + " = '" + Trim(fields(fieldCount / 2 + intLoop - 1)) + "', "
        Debug.Print strFields
    Next intLoop
    
    strFields = Left(strFields, Len(strFields) - 2)
    
    strSQL = strSQL + strFields + ")"
    
    Connection.Execute strSQL
    
    Return
    
Update:
    
    strSQL = "UPDATE " + table + " SET "
    
    For intLoop = 1 To fieldCount / 2
        strFields = strFields + fields(intLoop - 1) + " = '" + Trim(fields(fieldCount / 2 + intLoop - 1)) + "', "
        Debug.Print strFields
    Next intLoop
    
    strFields = Left(strFields, Len(strFields) - 2)
    
    strSQL = strSQL + strFields + " WHERE Id = " + fields(fieldCount / 2)
    
    Connection.Execute strSQL
    
    Return
    
ErrTrap:
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function SeekRecord(id As Integer)

    Dim strSQL As String
    Dim blnEnableDelete As Boolean
    
    ClearFields txtID, txtDescription, txtAccountsCode
    DisableFields txtDescription, txtAccountsCode
    
    blnEnableDelete = SimpleSeek("PaymentIn", "BankID", grdBanks.CellValue(grdBanks.CurRow, 1))
    If blnEnableDelete Then blnEnableDelete = blnEnableDelete = SimpleSeek("PaymentOut", "BankID", grdBanks.CellValue(grdBanks.CurRow, 1))
    
    strSQL = "SELECT Id, Description, AccountsCode FROM Banks WHERE Id = " & id
    
    Set bank = getDataFromDB(strSQL)
    
    If bank.RecordCount = 1 Then
        PopulateFields bank, txtID, txtDescription, txtAccountsCode
        lngSelectedRow = grdBanks.CurRow
        blnStatus = False
        EnableFields txtDescription, txtAccountsCode
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtDescription.SetFocus
    End If

End Function

Private Sub cmdButton_Click(index As Integer)
                                                                
    Select Case index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdBanks, False, 25, GetSetting(strApplicationName, "Layout Strings", "grdBanks"), "04NCIID,40NLNDescription", "ID,Ονομασία"
        Me.Refresh
        Set banks = getDataFromDB("SELECT Id, Description FROM Banks")
        PopulateGrid grdBanks, banks
    End If
    
    'AddDummyLines grdBanks, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAΑΑΑΑΑΑΑΑΑΑ"

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
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    UpdateColors Me, False
    SetUpGrid lstIconList, grdBanks
    ClearFields txtID, txtDescription, txtAccountsCode
    DisableFields txtDescription, txtAccountsCode
    UpdateButtons Me, 4, 1, 0, 0, 0, 1

End Sub

Private Sub grdBanks_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord grdBanks.CellValue(lRow, 1)

End Sub

Private Sub grdBanks_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdBanks_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord grdBanks.CellValue(grdBanks.CurRow, 1)

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdBanks", grdBanks.LayoutCol

End Sub

