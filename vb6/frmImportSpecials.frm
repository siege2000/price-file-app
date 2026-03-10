VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "C1Sizer.ocx"
Begin VB.Form frmImportSpecials 
   Caption         =   "Import Specials"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   998
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14970
      _cx             =   26405
      _cy             =   13123
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Blank Line"
         Height          =   495
         Left            =   11040
         TabIndex        =   41
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearPOSNote 
         Caption         =   "Clear POS Note"
         Height          =   615
         Left            =   9960
         TabIndex        =   40
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdFindBarcode 
         Caption         =   "Open Stockcards to get barcode or pharmacode for gift"
         Height          =   495
         Left            =   12360
         TabIndex        =   34
         Top             =   6840
         Width           =   2520
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   495
         Left            =   13845
         TabIndex        =   20
         Top             =   6240
         Width           =   1080
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   12645
         TabIndex        =   19
         Top             =   6240
         Width           =   1065
      End
      Begin VB.ComboBox cboFieldDelimiter 
         Height          =   315
         ItemData        =   "frmImportSpecials.frx":0000
         Left            =   12690
         List            =   "frmImportSpecials.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1020
         Width           =   1575
      End
      Begin VB.ComboBox cboRecordDelimiter 
         Height          =   315
         ItemData        =   "frmImportSpecials.frx":002A
         Left            =   12690
         List            =   "frmImportSpecials.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import"
         Height          =   495
         Left            =   11460
         TabIndex        =   16
         Top             =   6240
         Width           =   1050
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Columns"
         Height          =   1755
         Left            =   11265
         TabIndex        =   9
         Top             =   3600
         Width           =   3660
         Begin VB.ComboBox cboPOSNote 
            Height          =   315
            ItemData        =   "frmImportSpecials.frx":0049
            Left            =   1620
            List            =   "frmImportSpecials.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1320
            Width           =   1875
         End
         Begin VB.ComboBox cboBarcode 
            Height          =   315
            ItemData        =   "frmImportSpecials.frx":00B1
            Left            =   1620
            List            =   "frmImportSpecials.frx":00D0
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   960
            Width           =   1875
         End
         Begin VB.ComboBox cboDescription 
            Height          =   315
            ItemData        =   "frmImportSpecials.frx":00EF
            Left            =   1620
            List            =   "frmImportSpecials.frx":010E
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   1875
         End
         Begin VB.ComboBox cboSpecial 
            Height          =   315
            ItemData        =   "frmImportSpecials.frx":012D
            Left            =   1620
            List            =   "frmImportSpecials.frx":014C
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label8 
            Caption         =   "POSNote"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Special Price"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Barcode Pharmacode"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.TextBox txtSpecialName 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   6840
         Width           =   7950
      End
      Begin VB.TextBox txtStartDate 
         Height          =   285
         Left            =   2490
         TabIndex        =   7
         Top             =   6480
         Width           =   1245
      End
      Begin VB.TextBox txtFinishDate 
         Height          =   285
         Left            =   4725
         TabIndex        =   6
         Top             =   6480
         Width           =   1245
      End
      Begin VB.ComboBox cboBrand 
         Height          =   315
         Left            =   12885
         TabIndex        =   5
         Top             =   2880
         Width           =   2025
      End
      Begin VB.CheckBox chkIgnoreFirst 
         Caption         =   "Ignore First Line Of File"
         Height          =   255
         Left            =   11385
         TabIndex        =   4
         Top             =   3240
         Width           =   2595
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   12885
         TabIndex        =   3
         Top             =   2520
         Width           =   1995
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Line"
         Height          =   495
         Left            =   11460
         TabIndex        =   2
         Top             =   5640
         Width           =   1050
      End
      Begin VB.CommandButton cmdFindAndReplace 
         Caption         =   "Bulk Find And Replace"
         Height          =   495
         Left            =   12645
         TabIndex        =   1
         Top             =   5640
         Width           =   2235
      End
      Begin VSFlex7LCtl.VSFlexGrid vsImport 
         Height          =   6435
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11115
         _cx             =   19606
         _cy             =   11351
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportSpecials.frx":016B
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblLines 
         Height          =   375
         Left            =   9120
         TabIndex        =   39
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label lblPBL 
         Caption         =   "Import Barcodes First For PBL From Seperate File"
         Height          =   255
         Left            =   11280
         TabIndex        =   38
         Top             =   5400
         Width           =   3495
      End
      Begin VB.Label lblMultiBuy 
         Caption         =   "Label9"
         Height          =   855
         Left            =   12600
         TabIndex        =   37
         Top             =   0
         Width           =   2295
      End
      Begin VB.Shape shpNewCard 
         BackColor       =   &H00C0C0FF&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   11265
         Top             =   0
         Width           =   465
      End
      Begin VB.Label lblNewCards 
         Caption         =   "New Cards"
         Height          =   195
         Left            =   11820
         TabIndex        =   33
         Top             =   60
         Width           =   825
      End
      Begin VB.Shape shpChange 
         BackColor       =   &H00FFC0C0&
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   11265
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblChange 
         Caption         =   "Changes"
         Height          =   195
         Left            =   11820
         TabIndex        =   32
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblInvalid 
         Caption         =   "Invalid"
         Height          =   195
         Left            =   11820
         TabIndex        =   31
         Top             =   780
         Width           =   480
      End
      Begin VB.Shape shpInvalid 
         BackColor       =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   11265
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblFieldDelimiter 
         Caption         =   "Field Delimiter"
         Height          =   195
         Left            =   11460
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblRecordDelimiter 
         Caption         =   "Record Delimiter"
         Height          =   195
         Left            =   11460
         TabIndex        =   29
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Special Name"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   1545
         TabIndex        =   27
         Top             =   6480
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Finish Date"
         Height          =   255
         Left            =   3855
         TabIndex        =   26
         Top             =   6480
         Width           =   795
      End
      Begin VB.Label lblBrand 
         Caption         =   "Brand"
         Height          =   195
         Left            =   11580
         TabIndex        =   25
         Top             =   2940
         Width           =   1185
      End
      Begin VB.Label lblOddYear 
         Height          =   255
         Left            =   11520
         TabIndex        =   24
         Top             =   1920
         Width           =   3330
      End
      Begin VB.Label lblEvenYear 
         Height          =   255
         Left            =   11460
         TabIndex        =   23
         Top             =   2280
         Width           =   3330
      End
      Begin VB.Label Label7 
         Caption         =   "Password"
         Height          =   195
         Left            =   11580
         TabIndex        =   22
         Top             =   2550
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileAddBrand 
         Caption         =   "Add Brand"
      End
      Begin VB.Menu mnuLoadFile 
         Caption         =   "Load Grid"
      End
      Begin VB.Menu mnuSaveGrid 
         Caption         =   "Save Grid"
      End
      Begin VB.Menu mnuGroupSpecials 
         Caption         =   "Group Specials"
      End
      Begin VB.Menu mnuUNGroupSpecials 
         Caption         =   "Ungroup Specials"
      End
      Begin VB.Menu mmuFileReplace 
         Caption         =   "Bulk Find And Replace"
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Begin VB.Menu mnuSortErrors 
         Caption         =   "Sort On Invalid"
      End
   End
End
Attribute VB_Name = "frmImportSpecials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glbSpecialID As Long
'Const COL_SPECIAL_NEW As Integer = 5
'Const COL_SPECIAL_CHANGE As Integer = 6
'Const COL_SPECIAL_INVALID As Integer = 7
'Const COL_SPECIAL_PRICE As Integer = 3
'Const COL_SPECIAL_NEW As Integer = 4
'Const COL_SPECIAL_CHANGE As Integer = 5
'Const COL_SPECIAL_INVALID As Integer = 6
'Const COL_SPECIAL_PRICE As Integer = 2
'Const COL_SPECIAL_NEW As Integer = 3
'Const COL_SPECIAL_CHANGE As Integer = 4
'Const COL_SPECIAL_INVALID As Integer = 5
'Private Sub cboBrand_Change()
'   Call cboBrand_Click
'End Sub
Dim OldValue As String

Private Sub cboBrand_Click()
    Dim snaSettings As Recordset
    vsImport.Rows = 1
    glbSpecialID = 0
    Me.lblEvenYear.Caption = ""
    Me.lblOddYear.Caption = ""
    If cboBrand.text <> "" Then
        OpenSettingsDB
        If cboBrand.ListIndex >= 0 Then
            Set snaSettings = dbSettings.OpenRecordSet(SQL_Select & "SuppUDSpecialsSettings WHERE BrandID = " & cboBrand.ItemData(cboBrand.ListIndex), dbOpenSnapshot)
            If snaSettings.EOF Then
            
            Else
                cboSpecial.ListIndex = Val(FNulls(snaSettings("Special")))
                
                
                cboDescription.ListIndex = Val(FNulls(snaSettings("Description")))
                
                
                cboBarcode.ListIndex = Val(FNulls(snaSettings("Barcode")))
                
                cboFieldDelimiter.ListIndex = Val(FNulls(snaSettings("FieldDelimiter")))
                cboRecordDelimiter.ListIndex = Val(FNulls(snaSettings("RecordDelimiter")))
                cboPOSNote.ListIndex = Val(FNulls(snaSettings("POSNote")))
                
        '        If CBool(ReadIni(cboBrand.Text, "Import From DB", "FALSE", glbIniLocation$ & "Specials.ini")) Then
        '            Me.chkImportFromDB.Value = vbChecked
        '        Else
        '            Me.chkImportFromDB.Value = vbUnchecked
        '        End If
                
                '16/3/2004 Greg add in option to ignore first line
                If CBool(FNulls(snaSettings("IgnoreFirstLine"))) Then
                    Me.chkIgnoreFirst.value = vbChecked
                Else
                    Me.chkIgnoreFirst.value = vbUnchecked
                End If
            End If
            snaSettings.ClsRS
            Set snaSettings = Nothing
            If cboBrand.ListIndex >= 0 Then
                UpdateCodes (cboBrand.ItemData(cboBrand.ListIndex))
            End If
        End If
        CloseSettingsDB
    End If
End Sub
Private Sub UpdateCodes(ByVal BrandID As Long)
    Dim OddYear As Boolean
    If BrandID > 0 Then
         If Val(Format(Now, "yy")) Mod 2 = 0 Then
            OddYear = False
        Else
            OddYear = True
        End If
        '27/7/2009 Limin added. Should be other way round
        If OddYear Then
            Me.lblOddYear.Caption = "Code:" & "rrGGG" & Format(Now, "yy") & "0r" & BrandID
        Else
            Me.lblOddYear.Caption = "Code:" & BrandID & "rGGG" & Format(Now, "yy") & "0rr"
        End If
'        Me.lblEvenYear.Caption = "Even Year Code:" & BrandID & "rgggyy0rr"
'        Me.lblOddYear.Caption = "Odd Year Code:" & "rrgggyy0r" & BrandID
    End If
End Sub

Private Sub cmdAddItem_Click()
    Me.vsImport.AddItem ""
End Sub

Private Sub cmdClearPOSNote_Click()
    Dim RowCount As Integer
On Error GoTo cmdClearPOSNote_Click_Error
    With Me.vsImport
        For RowCount = 1 To Me.vsImport.Rows - 1
            If .IsSelected(RowCount) Then
                .TextMatrix(RowCount, COL_SPECIAL_POSNOTE) = ""
            End If
        Next RowCount
    End With
  
cmdClearPOSNote_Click_Exit:
   On Error Resume Next
Exit Sub

cmdClearPOSNote_Click_Error:
   If SYSLOG(Err, "cmdClearPOSNote_Click in frmImportSpecials in " & AppVersion()) Then
       Resume cmdClearPOSNote_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim Count As Long
On Error GoTo cmdDelete_Click_Error
    With Me.vsImport
        For Count = .Rows - 1 To 1 Step -1
            If .IsSelected(Count) Then
        'If .Row <= .Rows - 1 Then
                If MsgBox("Delete " & .TextMatrix(Count, 1) & "?", vbYesNo) = vbYes Then
                    .RemoveItem (Count)
                End If
            End If
        Next Count
        'End If
    End With
  
cmdDelete_Click_Exit:
   On Error Resume Next
Exit Sub

cmdDelete_Click_Error:
   If SYSLOG(Err, "cmdDelete_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume cmdDelete_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub cmdFindAndReplace_Click()
'    Dim Find As String
'    Dim Replace As String
'    Dim RowCount As Long
On Error GoTo cmdFindAndReplace_Click_Error

    Call mmuFileReplace_Click
'
'
'    Find = InputBox("Enter String To Find.")
'    Replace = InputBox("Enter String To Replace " & Find & " With.")
'
'    If Find <> "" Then
'        Find = UCase(Find)
'        Replace = UCase(Replace)
'        With vsImport
'            For RowCount = 1 To .Rows - 1
'
'                'If .IsSelected(RowCount) Then
'                    .TextMatrix(RowCount, COL_SPECIAL_DESCRIPTION) = TitleCase(FindandReplace(UCase(.TextMatrix(RowCount, COL_SPECIAL_DESCRIPTION)), Find, Replace))
'                    .TextMatrix(RowCount, COL_SPECIAL_LENGTH) = Len(.TextMatrix(RowCount, COL_SPECIAL_DESCRIPTION))
'                    If Len(.TextMatrix(RowCount, COL_SPECIAL_DESCRIPTION)) > 40 Then
'                        .TextMatrix(RowCount, COL_SPECIAL_INVALID) = "Y"
'                        .Cell(flexcpBackColor, RowCount, 0, .Cols - 1, RowCount) = vbRed
'                    Else
'                        .TextMatrix(RowCount, COL_SPECIAL_INVALID) = "N"
'                        .Cell(flexcpBackColor, RowCount, 0, .Cols - 1, RowCount) = vbWhite
'                    End If
'                'End If
'            Next RowCount
'        End With
'    End If
  
cmdFindAndReplace_Click_Exit:
   On Error Resume Next
Exit Sub

cmdFindAndReplace_Click_Error:
   If SYSLOG(Err, "cmdFindAndReplace_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume cmdFindAndReplace_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub cmdFindBarcode_Click()
    Dim Search As String
    Dim CloseDb As Boolean
    Dim SQL As String
    Dim Parts() As String
    Dim Part As Variant
    Dim SearchSQL As String
    Dim snaDetails As Recordset
    Dim Row As Long
    Dim RowCount As Long
On Error GoTo cmdFindBarcode_Click_Error
   Call Shell(GetLocalLOTSDir & "\Stock.exe /POS_UPdate /nomenu", vbNormalFocus)
'    Search = InputBox("Please enter the tradename to search for.", "Barcode Search")
'    If Search <> "" Then
'        Call OpenSupplierDB(CloseDB)
'        SQL = "SELECT Suppliers.SupplierName, Details.PharmaCode, Details.Description, Details.Barcode" & _
'        " FROM Details INNER JOIN Suppliers ON Details.SupplierID = Suppliers.SupplierID "
'        Parts = Split(Search, " ")
'        For Each Part In Parts
'            If SearchSQL = "" Then
'                SearchSQL = " WHERE Details.Description like('*" & Part & "*')"
'            Else
'                SearchSQL = SearchSQL & " AND Details.Description like('*" & Part & "*')"
'            End If
'        Next Part
'        Set snaDetails = dbSuppliers.OpenRecordSet(SQL & SearchSQL, dbOpenSnapshot)
'        If snaDetails.EOF Then
'            Call MsgBox("No Item Found for search sql " & SearchSQL)
'        Else
'            frmSelect.lstSelect.FormatString = "<Supplier                |<Description                                                   |<Pharmacode|<Barcode               "
'            Do Until snaDetails.EOF
'                Call frmSelect.lstSelect.AddItem(FNulls(snaDetails("SupplierName")) & vbTab & FNulls(snaDetails("Description")) & vbTab & FNulls(snaDetails("PharmaCode")) & vbTab & FNulls(snaDetails("Barcode")))
'                snaDetails.MoveNext
'            Loop
'            frmSelect.Show vbModal
'            If frmSelect.Tag = "1" Then
'                With Me.vsImport
'                    Row = .Row
'                    .TextMatrix(.Row, COL_SPECIAL_GIFTBARCODE) = frmSelect.lstSelect.TextMatrix(frmSelect.lstSelect.Row, 3)
'                    .TextMatrix(.Row, COL_SPECIAL_GIFTPHARMACODE) = frmSelect.lstSelect.TextMatrix(frmSelect.lstSelect.Row, 2)
'                    If Val(.TextMatrix(Row, COL_SPECIAL_GROUPID)) > 0 Then
'                        If MsgBox("Change All Items In the Group To The Same Value?", vbYesNo) = vbYes Then
'                            For RowCount = 1 To .Rows - 1
'                                If Val(.TextMatrix(Row, COL_SPECIAL_GROUPID)) = Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID)) Then
'                                    .TextMatrix(RowCount, COL_SPECIAL_GIFTBARCODE) = .TextMatrix(Row, COL_SPECIAL_GIFTBARCODE)
'                                    .TextMatrix(RowCount, COL_SPECIAL_GIFTPHARMACODE) = .TextMatrix(Row, COL_SPECIAL_GIFTPHARMACODE)
'                                End If
'                            Next RowCount
'                        End If
'
'                    End If
'                End With
'
'            End If
'        End If
'        snaDetails.ClsRS
'        Set snaDetails = Nothing
'        Call CloseSupplierDB(CloseDB)
'    End If
  
cmdFindBarcode_Click_Exit:
   On Error Resume Next
Exit Sub

cmdFindBarcode_Click_Error:
   If SYSLOG(Err, "cmdFindBarcode_Click in frmImportSpecials in " & AppVersion()) Then
       Resume cmdFindBarcode_Click_Exit
   Else
       Resume Next
   End If
End Sub
Private Sub ImportNewFormatPBL(Fields() As String, ByRef FieldColsdetails As GreenCrossNewFormat, ByVal v_LineCount As Integer, ByVal BrandID As Long, ByVal GroupIDS As Dictionary, ByRef v_HighestGroupID As Long)
    Dim Details As GreenCrossNewFormatDetails
    Dim MessageIsQuestion As Boolean
    Dim LivingRewards As Boolean
    Dim Combo As Boolean
    Dim GetLowestFree As Boolean
    Dim GiftBarcode As String
    Dim GiftPharmacode As String
    Dim GroupID As Long
    Dim ThePrices() As String
            Dim TheValue As Variant
    Const NewItem As String = "N"
    Const ChangedItem As String = "N"
    Const InvalidItem As String = "N"
    Const MessageOnly As String = "MESSAGEONLY"
     Details.DealName = Fields(FieldColsdetails.DealNameCol)
    If Details.DealName = "" Then
        Exit Sub
    End If
    If txtSpecialName.text = "" Then
        txtSpecialName.text = Fields(FieldColsdetails.PromotionNameCol)
        glbSpecialID = GetSpecialID(txtSpecialName.text, BrandID, False)
    End If
    If txtStartDate.text = "" Then
        txtStartDate.text = Fields(FieldColsdetails.StartDateCol)
    End If
    If txtFinishDate.text = "" Then
        txtFinishDate.text = Fields(FieldColsdetails.EndDateCol)
    End If
   
    If GroupIDS.Exists(Details.DealName) = False Then
        v_HighestGroupID = v_HighestGroupID + 1
        Call GroupIDS.Add(Details.DealName, v_HighestGroupID)
        
    End If
    GroupID = Val(GroupIDS(Details.DealName))
    Details.DealType = Fields(FieldColsdetails.DealTypeCol)
    Details.DealSubType = Fields(FieldColsdetails.DealSubTypeCol)
    Details.DealDiscountType = Fields(FieldColsdetails.DealDiscountTypeCol)
    Details.DealDiscountValue = Fields(FieldColsdetails.DealDiscountValueCol)
    Details.ProductPrice = Fields(FieldColsdetails.ProductPriceCol)
    Details.TradeName = LEFT(Fields(FieldColsdetails.TradeNameCol), 40)
    
    Select Case UCase(Details.DealType)
    Case "BOGO"
        Details.MultiBuyQty = Fields(FieldColsdetails.MinQtyCol)
        GetLowestFree = True
    Case "POS MESSAGE"
        Details.ProductPrice = MessageOnly
        Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
    Case "PROMO PRICE"
        Details.FreeProduct = Fields(FieldColsdetails.FreeProductCol)
        If (UCase(Details.FreeProduct) = "YES") Then
            Details.DealDiscountType = "%"
            Details.DealDiscountValue = "100"
            Details.ProductPrice = ""
            Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
            MessageIsQuestion = True
        
        End If
    Case "X FOR $Y"
        Details.MultiBuyQty = Fields(FieldColsdetails.MultiBuyQtyCol)
        Details.MultiBuyPrice = Fields(FieldColsdetails.MultiBuyPricecol)
        
        ThePrices = Split(Details.MultiBuyPrice, ",")
        Details.MultiBuyPrice = ""
        For Each TheValue In ThePrices
            If Details.MultiBuyPrice = "" Then
               Details.MultiBuyPrice = ConvertGXHDollarsToDollarsString(TheValue)
            Else
                Details.MultiBuyPrice = Details.MultiBuyPrice & "," & ConvertGXHDollarsToDollarsString(TheValue)
            End If
        Next TheValue
    Case "COMBO"
        Combo = True
        Details.ProductPrice = Fields(FieldColsdetails.DealPriceDollarCol)
    Case "DISCOUNT"
        If UCase(Details.DealSubType) = "TIERED DISCOUNT" Then
            Details.MultiBuyQty = Fields(FieldColsdetails.MultiBuyQtyCol)
            Details.MultiBuyPrice = Fields(FieldColsdetails.MultiBuyPricecol)
            
            ThePrices = Split(Details.MultiBuyPrice, ",")
            Details.MultiBuyPrice = ""
            For Each TheValue In ThePrices
                If Details.MultiBuyPrice = "" Then
                   Details.MultiBuyPrice = TheValue
                Else
                    Details.MultiBuyPrice = Details.MultiBuyPrice & "," & TheValue
                End If
            Next TheValue
        Else
            If Val(Fields(FieldColsdetails.MinQtyCol)) > 1 Then
                MessageIsQuestion = True
                Details.PromotionPOSNote = "Has the customer brought " & Val(Fields(FieldColsdetails.MinQtyCol)) & " or more products from special range?"
            End If
        End If
    Case "GWP"
        Details.ProductPrice = MessageOnly
        Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
    Case Else
        Call MsgBox("Dont know deal type :" & UCase(Details.DealType))
    End Select
    If InStr(UCase(Details.DealSubType), "LIVING REWARDS") > 0 Then
        LivingRewards = True
        Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
    End If
    If InStr(UCase(Details.DealSubType), "GWP") > 0 Then
        Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
    End If
'    Select Case UCase(Details.DealSubType)
'        Case "LIVING REWARDS"
'            LivingRewards = True
'            Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
'        Case "GWP"
'            Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
'        Case "GWP+Living Rewards"
'    End Select
    If UCase(Fields(FieldColsdetails.PromptForPOSNoteCol)) = "YES" Then
        Details.PromotionPOSNote = Fields(FieldColsdetails.PromotionPOSNoteCol)
    End If
    If Val(Fields(FieldColsdetails.MinSpendCol)) > 0 Then
        Details.MultiBuyQty = ">$" & Val(Fields(FieldColsdetails.MinSpendCol))
    End If
    With frmImportSpecials.vsImport
    If Details.DealDiscountType = "%" Then
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & Details.DealDiscountValue & "%" & vbTab & GetLowestFree & vbTab & NewItem & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & GiftPharmacode & vbTab & Combo & vbTab & MessageIsQuestion & vbTab & Details.DealName & vbTab & Details.SecondaryGroupID
    ElseIf Details.DealDiscountType = "$" Then
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & "-" & ConvertGXHDollarsToDollarsString(Details.DealDiscountValue) & vbTab & GetLowestFree & vbTab & NewItem & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & GiftPharmacode & vbTab & Combo & vbTab & MessageIsQuestion & vbTab & Details.DealName & vbTab & Details.SecondaryGroupID
    ElseIf InStr(Details.MultiBuyPrice, ",") Then
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & vbTab & GetLowestFree & vbTab & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & Details.MultiBuyPrice & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & Combo & vbTab & MessageIsQuestion & vbTab & Details.DealName & vbTab & Details.SecondaryGroupID
    ElseIf Details.MultiBuyPrice <> "" Then
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & Details.MultiBuyPrice & vbTab & GetLowestFree & vbTab & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & Combo & vbTab & MessageIsQuestion
    ElseIf Details.ProductPrice = MessageOnly Then
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & vbTab & GetLowestFree & vbTab & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & Combo & vbTab & MessageIsQuestion & vbTab & Details.DealName & vbTab & Details.SecondaryGroupID
    Else
        .AddItem Fields(FieldColsdetails.BarcodeCol) & vbTab & Details.TradeName & vbTab & Len(Details.TradeName) & vbTab & Details.MultiBuyQty & vbTab & ConvertGXHDollarsToDollarsString(Details.ProductPrice) & vbTab & GetLowestFree & vbTab & vbTab & ChangedItem & vbTab & InvalidItem & vbTab & Details.PromotionPOSNote & vbTab & vbTab & LivingRewards & vbTab & GroupID & vbTab & GiftBarcode & vbTab & Combo & vbTab & MessageIsQuestion & vbTab & Details.DealName & vbTab & Details.SecondaryGroupID
        
    End If
    .RowData(.Rows - 1) = .Rows - 1
    End With
    Call UpdateStatusFlag
End Sub
Private Function ConvertGXHDollarsToDollarsString(ByVal v_TheValue As String) As String
    If (InStr(v_TheValue, ".")) > 0 Then
        ConvertGXHDollarsToDollarsString = centsToDollars(dollarsToCents(v_TheValue))
    Else
         ConvertGXHDollarsToDollarsString = centsToDollars(dollarsToCents(Val(v_TheValue) * 100))
    End If
End Function
Private Sub cmdImport_Click()
    Dim hdlImport As Integer
    Dim FileName As String
    Dim FilePath As String
    Dim DefaultDir As String
    Dim RawData As String
    Dim Fields() As String
    Dim Delimiter As String
    Dim LinesRead As Long
    Dim Barcode As String
    Dim Special As String
    Dim Description As String
    Dim Invalid As Boolean
    Dim SpecialPrice As Long
    Dim BrandID As Long
    Dim snaSpecialPrice As Recordset
    Dim Excel As Boolean
    Dim RowsToBeLoaded As Long
    Dim LineCount As Long
    Dim ExitLoop As Boolean
    Dim MaxSheetCount As Integer
    Dim Percent As Boolean
    Dim XML As Boolean
    Dim MyXML As DOMDocument60
    Dim ProductList As MSXML2.IXMLDOMNode
    Dim Product As MSXML2.IXMLDOMNode
    Dim MultiBuy As String
    Dim GetLowest As String
    Dim POSNote As String
    Dim snaDetails As Recordset
    Dim MultiRetail As String
    Dim Flybuys As String
    Dim GroupID As Long
    Dim MaxGroupID As Long
    Dim LastPOSNote As String
    Dim PBLFile As Boolean
    Dim GroupIDS As Dictionary
    Dim GroupRow As Integer
    Dim Progress As ProgressIndicator
    Dim GiftBarcode As String
    Dim GiftBarcodes As Dictionary
    Dim BuyDetails() As String
    Dim PriceDetails() As String
    Dim TheCount As Long
     Dim StartDateCol As Long
     Dim EndDateCol As Long
     Dim TradeNameCol As Long
     Dim POSNoteCol As Long
     Dim POSCouponCol As Long
      Dim BarcodeCol As Long
      Dim PLUCol As Long
      Dim GuidCol As Long
      Dim PriceCol As Long
      Dim DiscountCol As Long
      Dim MultiBuyCol As Long
'      Dim LoyaltyBarcodeCol As Integer
'      Dim FlybuysCol As Integer
      Dim CatalogueText As String
      Dim NextCatalogueText As String
      Dim NewFormatPBL As Boolean
      Dim NewFormatSetting As GreenCrossNewFormat
      
      
   On Error GoTo cmdImport_Click_Error
    
    glbSpecialID = 0
    CatalogueText = Format(Now, "mmmm yyyy") & " Catalogue."
    NextCatalogueText = Format(DateAdd("m", 1, Now), "mmmm yyyy") & " Catalogue."
    'April 2014 Catalogue.
    If cboBrand.ListIndex >= 0 Then
        BrandID = cboBrand.ItemData(cboBrand.ListIndex)
        
        DefaultDir = readini("Defaults", "SpecialsDefaultDir", PathName, GetPOSPricesPath)
        
        FileName = GetOpenFile(DefaultDir, "(CSV (*.csv)|*.csv|text (*.txt)|*.txt|XML (*.xml)|*.xml|Excel (*.xls)|*.xls;*.xlsx|All (*.*)|*.*", FilePath)
        'Change Greg 20/6/06 make sure the cancel works
        If FileName = "" Then
            Exit Sub
        End If
        If Me.vsImport.Rows = 1 Then
            Me.txtSpecialName.text = ""
            Me.txtFinishDate.text = ""
            Me.txtStartDate.text = ""
        End If
        DefaultDir = FileName
        DefaultDir = Replace(DefaultDir, FilePath, "")
        Call writeini("Defaults", "SpecialsDefaultDir", DefaultDir, GetPOSPricesPath)
        OpenSupplierDB
        
        Call SaveSpecialSettings
        UpdateCodes (BrandID)
        OpenSettingsDB
        MousePointer = vbHourglass
        If InStr(UCase(FileName), ".XLS") Then
            Call OpenSpreadsheet(FileName)
            'Change Greg 25/8/06 Make so that multipule worksheets can be loaded
            

            Excel = True
            
            Call frmExcelSheets.SetSheets(GetWorksheetNames, 0)
            MaxSheetCount = WorkSheetCount
            If MaxSheetCount > 1 Then
                frmExcelSheets.Show vbModal
            
            
                For LineCount = 1 To MaxSheetCount
                    If frmExcelSheets.ImportSheet(LineCount, 0) Then
                        Call SetWorksheet(, LineCount)
                        Exit For
                    End If
                Next LineCount
            Else
                Call SetWorksheet(, 1)
            End If
            LineCount = 1
            RowsToBeLoaded = NumberOfRows
            MaxSheetCount = WorkSheetCount
'            proImport.RecordCount = RowsToBeLoaded
        ElseIf InStr(UCase(FileName), ".XML") Then
            hdlImport = FreeFile
            Close #hdlImport
            Open FileName For Input As #hdlImport
            Do Until EOF(hdlImport)
                Line Input #hdlImport, Delimiter
                RawData = RawData & Delimiter & vbCrLf
            Loop
            Close #hdlImport
            hdlImport = FreeFile
            Set MyXML = New DOMDocument60
            Call MyXML.loadXML(RawData)
            Set ProductList = MyXML.documentElement.childNodes.Item(2)
            RowsToBeLoaded = ProductList.childNodes.Length
            LineCount = 1
            XML = True
        Else
            hdlImport = FreeFile
            Open FileName For Input As hdlImport
            Excel = False
            MaxSheetCount = 1
'            proImport.RecordCount = LOF(hdlImport)
        End If
        If Excel Then
            Delimiter$ = "|"
            'Delimiter$ = vbTab
        ElseIf XML Then
            Delimiter = vbTab
        Else
            Select Case UCase(Me.cboFieldDelimiter.text)
                Case "TAB"
                    Delimiter = vbTab
                Case "CR"
                    Delimiter = vbCr
                Case "CRLF"
                    Delimiter = vbCrLf
                Case Else
                    Delimiter = ","
            End Select
        End If
        Set Progress = New ProgressIndicator
        Progress.RecordCount = 10000
        Progress.Show "Importing Special"
        With frmImportSpecials
            If Excel Then
                If .chkIgnoreFirst.value = vbChecked And RowsToBeLoaded <> 0 Then
                     RawData = GetNextRow(1, True)
                     If InStr(UCase(RawData), UCase("Deal type")) > 0 Then
                        NewFormatPBL = True
                        Call GetNewFormatPBLColumns(RawData, Delimiter, NewFormatSetting)
                     ElseIf InStr(UCase(RawData), "SPECIALPROMPTOVERMAX") > 0 Then
                        Call GetPBLColumns(RawData, Delimiter, StartDateCol, EndDateCol, TradeNameCol, POSNoteCol, POSCouponCol, BarcodeCol, PLUCol, GuidCol, PriceCol, DiscountCol, MultiBuyCol)
                        PBLFile = True
                    End If
                     LineCount = 2
                End If
            ElseIf XML Then
                If .chkIgnoreFirst.value = vbChecked Then
                    Set Product = ProductList.childNodes.Item(LineCount - 1)
                    LineCount = 2
                End If
            Else
                If .chkIgnoreFirst.value = vbChecked And EOF(hdlImport) = False Then
                    Line Input #hdlImport, RawData
                    LineCount = 2
                    If InStr(UCase(RawData), "SPECIALPROMPTOVERMAX") > 0 Or (InStr(UCase(RawData), UCase("SpecialPromptForCoupon")) > 0 And InStr(UCase(RawData), UCase("SpecialCouponPrompt")) > 0 And InStr(UCase(RawData), UCase("SpecialDiscount")) > 0) Then
                        Call GetPBLColumns(RawData, Delimiter, StartDateCol, EndDateCol, TradeNameCol, POSNoteCol, POSCouponCol, BarcodeCol, PLUCol, GuidCol, PriceCol, DiscountCol, MultiBuyCol)
                        PBLFile = True
                    End If
                End If
            End If
            '24/3/2004 Greg Add in code to check whether to exit the loop
            If Excel Then
                If LineCount > RowsToBeLoaded Then
                    ExitLoop = True
                Else
                    ExitLoop = False
                End If
            ElseIf XML Then
                If LineCount > RowsToBeLoaded Then
                    ExitLoop = True
                Else
                    ExitLoop = False
                End If
                
            Else
                If EOF(hdlImport) Or Delimiter = "" Then
                    ExitLoop = True
                Else
                    ExitLoop = False
                End If
            End If
            Set GroupIDS = New Dictionary
            Set GiftBarcodes = New Dictionary
            Do Until ExitLoop

               If Excel Then
                    RawData = GetNextRow(LineCount, True)
                    LineCount = LineCount + 1
                    LinesRead = LinesRead + 1
                ElseIf XML Then
                    Set Product = ProductList.childNodes(LineCount - 1)
                    LineCount = LineCount + 1
                    LinesRead = LinesRead + 1
                Else
                    'If Delimiter = "" Or InStr(RawData, Delimiter) = 0 Then
                        Line Input #hdlImport, RawData
                        LinesRead = LinesRead + 1
                    'End If
                End If
                Progress.increment LinesRead / 2
                Barcode = ""
                Description = ""
                Special = ""
                MultiBuy = "1"
                Invalid = True
                Percent = False
                SpecialPrice = -1
                Flybuys = "N"
                GetLowest = "N"
                GiftBarcode = ""

                    If InStr(RawData, Delimiter) > 0 Then
                        
                        Call ConvertDelimitedStringToArray(RawData, Delimiter, Fields)
                        If LinesRead <= 2 Then
                        If PBLFile = False Then
                            If InStr(UCase(RawData), "SPECIALPROMPTOVERMAX") > 0 Then
                                Call GetPBLColumns(RawData, Delimiter, StartDateCol, EndDateCol, TradeNameCol, POSNoteCol, POSCouponCol, BarcodeCol, PLUCol, GuidCol, PriceCol, DiscountCol, MultiBuyCol)
                                PBLFile = True
                            End If
                        End If
                        End If
                        If NewFormatPBL Then
                            Call ImportNewFormatPBL(Fields, NewFormatSetting, LineCount, BrandID, GroupIDS, MaxGroupID)
                            
                            GoTo NextRow
                        ElseIf PBLFile Then
                            If UBound(Fields) < 15 Then
                                 
                                 
                                    
                                Dim Temp As String
                                Temp = RawData
                                If Excel Then
                                    RawData = GetNextRow(LineCount, True)
                                    LineCount = LineCount + 1
                                    'LinesRead = LinesRead + 1
                                Else
                                    Line Input #hdlImport, RawData
                                End If
                                RawData = Temp & RawData
                                Call ConvertDelimitedStringToArray(RawData, Delimiter, Fields)
                                If UBound(Fields) < 28 Then
                                    Temp = RawData
                                    Line Input #hdlImport, RawData
                                    RawData = Temp & RawData
                                    Call ConvertDelimitedStringToArray(RawData, Delimiter, Fields)
                                    If UBound(Fields) < 28 Then
                                        MsgBox ("Test")
                                    End If
                                    
                                End If
                            End If
                        End If
                        Invalid = False
                        'Check to see if it is the promotion name
                        If UCase(Trim(Fields(0))) = "PROMOTION" Then
                            If UBound(Fields) >= 1 Then
                                txtSpecialName.text = Trim(Fields(1))
                                glbSpecialID = GetSpecialID(txtSpecialName.text, BrandID, False)
                            End If
                        'Check to see if it is the Start date
                        ElseIf UCase(Trim(Fields(0))) = "START DATE" Then
                            If UBound(Fields) >= 1 Then
                                txtStartDate.text = FormatDate(Trim(Fields(1)), False)
                            End If
                        'Check to see if it is the finish date
                        ElseIf UCase(Trim(Fields(0))) = "FINISH DATE" Then
                            txtFinishDate.text = FormatDate(Trim(Fields(1)), False)
                        Else
                            'Must be a normal line
                            LinesRead = LinesRead + 1
                            
                            If LinesRead = 1 And (Me.chkIgnoreFirst = vbChecked Or PBLFile) Then
                                Invalid = True
                            Else
                                If PBLFile Then
                                    If IsDate(txtFinishDate.text) = False Then
                                        txtFinishDate.text = Fields(EndDateCol)
                                    
                                    End If
                                    If IsDate(txtStartDate.text) = False Then
                                        txtStartDate.text = Fields(StartDateCol)
                                    
                                    End If
                                    
'                                    Barcode = Fields(LoyaltyBarcodeCol)
'                                    If Barcode = "0" Or InStr(Barcode, "E+") > 0 Then
'                                        Barcode = ""
'                                    End If
'                                    If Barcode = "" Then
                                        Barcode = Fields(BarcodeCol)
                                        If Barcode = "0" Or InStr(Barcode, "E+") > 0 Then
                                            Barcode = ""
                                        End If
                                        If Barcode = "" Then
                                            Barcode = Fields(PLUCol)
                                            If Barcode = "0" Or InStr(Barcode, "E+") > 0 Then
                                                Barcode = ""
                                            End If
                                            If Barcode = "" Then
                                                If GuidCol >= 0 Then
                                                    Dim snaGetBarcode As Recordset
                                                    Set snaGetBarcode = dbSettings.OpenRecordSet("SELECT barcode FROm TBarcodes WHERE tGuid='" & Fields(GuidCol) & "'", dbOpenForwardOnly)
                                                    If snaGetBarcode.EOF Then

                                                    Else
                                                       Barcode = snaGetBarcode("Barcode")
                                                    End If
                                                    snaGetBarcode.ClsRS
                                                    Set snaGetBarcode = Nothing
                                                End If
                                            Else

                                            End If
                                        End If
'                                        If FBool(Fields(FlybuysCol)) Then
'                                            Flybuys = "Y"
'                                        End If
'                                    Else
'                                        Flybuys = "Y"
'                                    End If
                                    Description = Fields(TradeNameCol)
                                     MultiBuy = Fields(MultiBuyCol)
                                     POSNote = Trim(Fields(POSNoteCol))
                                     If POSNote = "" Then
                                        'Change Nigel 26.6.2013 - don't do this as column L (POSCouponCol) contains Toniq only info and Nicki was having to go through and manually remove all this text
                                        POSNote = Fields(POSCouponCol)
                                        'If POSNote <> "" Then
                                        '    MultiBuy = MultiBuy + 1
                                        '    GetLowest = "Y"
                                        'End If
                                     End If
                                     If InStr(UCase(POSNote), "LOYALTY CARD?") > 0 Then
                                        Flybuys = "Y"
                                     End If
                                     'Change Nigel 26.6.2013 - Nicki says to leave this text in
                                     'If Len(POSNote) > 7 And InStr(UCase(POSNote), "CATALOGUE") Then
                                     '  If IsDate(LEFT(POSNote, 7)) Then
                                     '   POSNote = ""
                                     '  End If
                                     'End If
                                     
                                         If InStr(MultiBuy, ",") > 0 Then
                                            Special = Replace(Fields(PriceCol), ",", "")
                                            Temp = Special
                                            Special = ""
                                             BuyDetails = Split(MultiBuy, ",")
                                            For TheCount = 0 To UBound(BuyDetails)
                                                Special = Special & "," & Val(Mid(Temp, (TheCount * 4) + 1, 4)) * Val(BuyDetails(TheCount)) + IIf((TheCount Mod 2) = 1, 1, 0)
                                            Next TheCount
                                            Special = Right(Special, Len(Special) - 1)
                                            
                                         Else
                                             If Fields(DiscountCol) <> "" And Fields(DiscountCol) <> "0" Then
                                                SpecialPrice = (Val(Fields(DiscountCol)) - 1) * 100
                                                Percent = True
                                                Special = SpecialPrice & "%"
                                             ElseIf Fields(PriceCol) <> "" Then
                                                SpecialPrice = Val(Fields(PriceCol))
                                                Special = centsToDollars(SpecialPrice)
                                                Percent = False
                                             Else
                                             
                                             End If
                                            If Val(MultiBuy) > 1 And Percent = False Then
                                                SpecialPrice = SpecialPrice * MultiBuy + IIf((MultiBuy Mod 2) = 0, IIf(SpecialPrice Mod 100 = 0, 0, 1), 0)
                                            End If
                                        End If
                                    
                                Else
                                    If UBound(Fields) >= Me.cboBarcode.ListIndex Then
                                        Barcode = Fields(cboBarcode.ListIndex)
                                        Barcode$ = FindandReplace(Barcode$, """", "")
                                        Barcode$ = FindandReplace(Barcode$, " ", "")
                                        If IsNumeric(Barcode$) And InStr(Barcode$, "E") = 0 Then
                                            
                                        Else
                                            Barcode = ""
                                            Invalid = True
                                        End If
                                    Else
                                        MousePointer = vbDefault
                                        MsgBox "The file isn't in the right format", vbCritical
                                        Exit Sub
                                    End If
                                    If UBound(Fields) >= Me.cboDescription.ListIndex Then
                                        Description = Fields(cboDescription.ListIndex)
                                        Description = FindandReplace(Description, """", "")
                                        If Len(Trim(Description)) > 0 Then
                                            
                                        Else
                                            Description = ""
                                            Invalid = True
                                        End If
                                    Else
                                        MsgBox "The file isn't in the right format", vbCritical
                                        Exit Sub
                                    End If
                                    If UBound(Fields) >= Me.cboSpecial.ListIndex Then
                                        Special = Fields(cboSpecial.ListIndex)
                                        Special = FindandReplace(Special, """", "")
                                        'Check to see if there is a decimal place
                                        If InStr(Special, ".") > 0 Then
                                        
                                        Else
                                            'if there isn't add one in
                                            Special = Special & ".00"
                                        End If
                                        If InStr(Special, "%") Then
                                            Percent = True
                                            SpecialPrice = Val(Special)
                                        Else
                                            SpecialPrice = dollarsToCents(Special)
                                            Percent = False
                                        End If
                                        If SpecialPrice > 0 Then
                                            
                                        Else
                                            SpecialPrice = 0
                                            Invalid = True
                                        End If
                                    Else
                                        MousePointer = vbDefault
                                        MsgBox "The file isn't in the right format", vbCritical
                                        Exit Sub
                                    End If
                                    If cboPOSNote.ListIndex > 0 Then
                                        If UBound(Fields) >= Me.cboPOSNote.ListIndex Then
                                            POSNote = Fields(cboPOSNote.ListIndex)
                                            POSNote$ = FindandReplace(POSNote$, """", "")
                                            
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
'                End If
                
                If Invalid = False Then
                    If Val(MultiBuy) < 1 Then
                        MultiBuy = "1"
                    End If
                    POSNote = FindandReplace(POSNote, CatalogueText, "")
                    POSNote = FindandReplace(POSNote, NextCatalogueText, "")
                    If POSNote <> "" Then
                        Dim TempNote As String
                        TempNote = UCase(POSNote)
                        If InStr(TempNote, "IF THE CUSTOMER IS PURCHASING A QUALIFYING ITEM FOR THIS PROMOTION") Then
                            POSNote = GetValueFromString(POSNote, "", "IF THE CUSTOMER IS PURCHASING A QUALIFYING ITEM FOR THIS PROMOTION")
                        End If
                        If InStr(TempNote, ", PRESS N ON THE") Then
                            POSNote = GetValueFromString(POSNote, "", ", PRESS N ON THE")
                        End If
                        If InStr(TempNote, ", PRESS Y ON THE") Then
                            POSNote = GetValueFromString(POSNote, "", ", PRESS Y ON THE")
                        End If
                        If InStr(TempNote, "PLEASE ENSURE YOU SCAN") Then
                            POSNote = GetValueFromString(POSNote, "", "PLEASE ENSURE YOU SCAN")
                        End If
                        If InStr(TempNote, ", ANSWER Y TO THIS PROMPT ON THE") Then
                            POSNote = GetValueFromString(POSNote, "", ", ANSWER Y TO THIS PROMPT ON THE")
                        End If
                        If Special = "" And InStr(MultiBuy, ">") = 0 Then
                            If InStr(TempNote, "SPEND") > 0 Then
                                If InStr(InStr(TempNote, "SPEND"), TempNote, "OR MORE") > 0 Then
                                    If dollarsToCents(Trim(GetValueFromString(POSNote, "Spend", "Or More"))) > 0 Then
                                        MultiBuy = ">" & Trim(GetValueFromString(POSNote, "Spend", "Or More"))
                                    End If
                                End If
                            End If
                        End If
                        If Special = "" And InStr(MultiBuy, ">") = 0 Then
                            If InStr(TempNote, "SELLING") > 0 Then
                                If InStr(InStr(TempNote, "SELLING"), TempNote, "FROM") > 0 Then
                                    If dollarsToCents(Trim(GetValueFromString(POSNote, "selling", "from"))) > 0 Then
                                        MultiBuy = ">" & Trim(GetValueFromString(POSNote, "selling", "from"))
                                    End If
                                End If
                            End If
                        End If
                        If Special = "" And InStr(MultiBuy, ">") = 0 Then
                            If InStr(TempNote, "SPEND") > 0 Then
                                If InStr(InStr(TempNote, "SPEND"), TempNote, "ON") > 0 Then
                                    If dollarsToCents(Trim(GetValueFromString(POSNote, "spend", "on"))) > 0 Then
                                        MultiBuy = ">" & Trim(GetValueFromString(POSNote, "spend", "on"))
                                    End If
                                End If
                            End If
                        End If
                        If Special = "" And InStr(MultiBuy, ">") = 0 Then
                            If InStr(TempNote, "PURCHASE") > 0 Then
                                If InStr(InStr(TempNote, "PURCHASE"), TempNote, "OR MORE") > 0 Then
                                    If dollarsToCents(Trim(GetValueFromString(POSNote, "PURCHASE", "OR MORE"))) > 0 Then
                                        MultiBuy = ">" & Trim(GetValueFromString(POSNote, "PURCHASE", "OR MORE"))
                                    End If
                                End If
                            End If
                        End If
                        If Special = "" And InStr(MultiBuy, ">") = 0 Then
                            If InStr(TempNote, "PURCHASE") > 0 Then
                                If InStr(InStr(TempNote, "PURCHASE"), TempNote, "OF") > 0 Then
                                    If dollarsToCents(Trim(GetValueFromString(POSNote, "PURCHASE", "OF"))) > 0 Then
                                        MultiBuy = ">" & Trim(GetValueFromString(POSNote, "PURCHASE", "OF"))
                                    End If
                                End If
                            End If
                        End If
                        If GiftBarcode = "" Then
                            If InStr(UCase(POSNote), "SCAN") > 0 Then
                                GiftBarcode = GetBarcode(POSNote, 0)
                                
                            End If
                        End If
                        
                        If POSNote = LastPOSNote Then
                            
                            If GroupID = 0 Then
                                If GroupIDS.Exists(POSNote) Then
                                    GroupID = GroupIDS(POSNote)
                                Else
                                    MaxGroupID = MaxGroupID + 1
                                    GroupID = MaxGroupID
                                    .vsImport.TextMatrix(.vsImport.Rows - 1, COL_SPECIAL_GROUPID) = GroupID
                                    For GroupRow = .vsImport.Rows - 1 To 1 Step -1
                                          If .vsImport.TextMatrix(GroupRow, COL_SPECIAL_POSNOTE) = POSNote Then
                                            .vsImport.TextMatrix(GroupRow, COL_SPECIAL_GROUPID) = GroupID
                                          End If
                                          
                                    Next GroupRow
                                    Call GroupIDS.Add(POSNote, GroupID)
                                    If GiftBarcode <> "" Then
                                        If GiftBarcodes.Exists(GiftBarcode) = False Then
                                            Call GiftBarcodes.Add(GiftBarcode, GroupID)
                                        End If
                                    End If
                                End If
                            Else
                                
                            End If
                        Else
                            If GroupIDS.Exists(POSNote) Then
                                GroupID = GroupIDS(POSNote)
                            Else
                                GroupID = 0
                            End If
                        End If
                        
                    Else
                        GroupID = 0
                    End If
                    LastPOSNote = POSNote
                    
                    With frmImportSpecials.vsImport
                        If Percent Then
                            .AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & SpecialPrice & "%" & vbTab & GetLowest & vbTab & "N" & vbTab & "N" & vbTab & "N" & vbTab & POSNote & vbTab & vbTab & Flybuys & vbTab & GroupID & vbTab & GiftBarcode
                            '.AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & SpecialPrice & "%" & vbTab & "N" & vbTab & "N" & vbTab & "N"
                            '.AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & SpecialPrice & "%" & vbTab & "N" & vbTab & "N" & vbTab & "N"
                        
                        ElseIf InStr(Special, ",") Then
                            .AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & centsToDollars(SpecialPrice) & vbTab & GetLowest & vbTab & "N" & vbTab & "N" & vbTab & "N" & vbTab & POSNote & vbTab & Special & vbTab & Flybuys & vbTab & GroupID & vbTab & GiftBarcode
                        ElseIf SpecialPrice = -1 Then
                            .AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & vbTab & GetLowest & vbTab & "N" & vbTab & "N" & vbTab & "N" & vbTab & POSNote & vbTab & vbTab & Flybuys & vbTab & GroupID & vbTab & GiftBarcode
                        Else
                            .AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & centsToDollars(SpecialPrice) & vbTab & GetLowest & vbTab & "N" & vbTab & "N" & vbTab & "N" & vbTab & POSNote & vbTab & vbTab & Flybuys & vbTab & GroupID & vbTab & GiftBarcode
                            '.AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & MultiBuy & vbTab & centsToDollars(SpecialPrice) & vbTab & "N" & vbTab & "N" & vbTab & "N"
                            '.AddItem Barcode & vbTab & Description & vbTab & Len(Description) & vbTab & centsToDollars(SpecialPrice) & vbTab & "N" & vbTab & "N" & vbTab & "N"
                        End If
                        .RowData(.Rows - 1) = .Rows - 1
                        If Barcode <> "" Then
                            If GiftBarcodes.Exists(Barcode) Then
                                Dim Count As Integer
                                GroupID = GiftBarcodes(Barcode)
                                For GroupRow = .Rows - 1 To 1 Step -1
                                    If .TextMatrix(GroupRow, COL_SPECIAL_GROUPID) = GroupID Then
                                        For Count = 3 To .Cols - 1
                                            .TextMatrix(.Rows - 1, Count) = .TextMatrix(GroupRow, Count)
                                        Next Count
                                        Exit For
                                    End If
                                Next GroupRow
                            End If
                        End If
                        Call UpdateStatusFlag
                    End With
                End If

'                        End If
'                    End If
'                End If
            '24/3/2004 Greg Add in code to check whether to exit the loop
NextRow:
               If Excel Or XML Then
                   If LineCount > RowsToBeLoaded Then
                       ExitLoop = True
                   Else
                       ExitLoop = False
                   End If
               Else
                    If EOF(hdlImport) Then
                        ExitLoop = True
                    End If
'                   If EOF(hdlImport) And Delimiter = "" Then
'                       ExitLoop = True
'                   Else
'                       ExitLoop = False
'                   End If
               End If
            Loop
        End With
    End If
cmdImport_Click_Exit:
   On Error Resume Next
   Unload frmExcelSheets
   CloseSupplierDB
   CloseSettingsDB
   Setlines
   MousePointer = vbDefault
   Exit Sub

cmdImport_Click_Error:
    If SYSLOG(Err, "cmdImport_Click in frmImportSpecials in " & app.ExeName) Then
        Resume cmdImport_Click_Exit
    Else
        Resume Next
    End If
End Sub
Private Sub UpdateStatusFlag()
    With frmImportSpecials.vsImport
     Dim snaSpecialPrice As Recordset
    If glbSpecialID > 0 Then
                            Set snaSpecialPrice = dbSuppliers.OpenRecordSet("SELECT * FROM SpecialItems WHERE Barcode = '" & .TextMatrix(.Rows - 1, COL_SPECIAL_BARCODE) & "' AND SpecialID = " & glbSpecialID, dbOpenSnapshot)
                            If snaSpecialPrice.EOF = True Then
                                .Cell(flexcpBackColor, .Rows - 1, COL_SPECIAL_PRICE) = shpNewCard.FillColor
'                                .Row = .Rows - 1
'                                .Col = COL_SPECIAL_PRICE
'                                .CellBackColor = shpNewCard.FillColor
                                .TextMatrix(.Rows - 1, COL_SPECIAL_NEW) = "Y"
                            Else
                                If Len(Trim(.TextMatrix(.Rows - 1, COL_SPECIAL_DESCRIPTION))) > 50 Then
                                    .TextMatrix(.Rows - 1, COL_SPECIAL_DESCRIPTION) = FNulls(snaSpecialPrice("Description"))
                                End If
                                If .TextMatrix(.Rows - 1, COL_SPECIAL_PRICE) <> FNulls(snaSpecialPrice("SpecialRetail")) Then
                                    .Cell(flexcpBackColor, .Rows - 1, COL_SPECIAL_PRICE) = shpChange.FillColor
'                                    .Row = .Rows - 1
'                                    .Col = COL_SPECIAL_PRICE
'                                    .CellBackColor = shpChange.FillColor
                                    .TextMatrix(.Rows - 1, COL_SPECIAL_CHANGE) = "Y"
                                End If
                                If FNulls(snaSpecialPrice("Description")) <> .TextMatrix(.Rows - 1, COL_SPECIAL_DESCRIPTION) Then
                                    .Cell(flexcpBackColor, .Rows - 1, COL_SPECIAL_DESCRIPTION) = shpChange.FillColor
'                                    .Row = .Rows - 1
'                                    .Col = COL_SPECIAL_DESCRIPTION
'                                    .CellBackColor = shpChange.FillColor
                                    .TextMatrix(.Rows - 1, COL_SPECIAL_CHANGE) = "Y"
                                End If
                                
                            End If
                            snaSpecialPrice.ClsRS
                            Set snaSpecialPrice = Nothing
                        Else
                            .Cell(flexcpBackColor, .Rows - 1, COL_SPECIAL_PRICE) = shpNewCard.FillColor
'                            .Row = .Rows - 1
'                            .Col = COL_SPECIAL_PRICE
'                            .CellBackColor = shpNewCard.FillColor
                            .TextMatrix(.Rows - 1, COL_SPECIAL_NEW) = "Y"
                        End If
                        If Len(Trim(.TextMatrix(.Rows - 1, COL_SPECIAL_DESCRIPTION))) > 50 Then
                            .Cell(flexcpBackColor, .Rows - 1, COL_SPECIAL_DESCRIPTION) = shpInvalid.FillColor
'                            .Row = .Rows - 1
'                            .Col = COL_SPECIAL_DESCRIPTION
'                            .CellBackColor = shpInvalid.FillColor
                            .TextMatrix(.Rows - 1, COL_SPECIAL_INVALID) = "Y"
                        End If
                    End With
End Sub
Private Sub GetNewFormatPBLColumns(ByVal RawData As String, ByVal Delimiter As String, ByRef FieldColsdetails As GreenCrossNewFormat)
'Private Sub GetPBLColumns(ByVal RawData As String, ByVal Delimiter As String, ByRef StartDateCol As Integer, ByRef EndDateCol As Integer, ByRef TradeNameCol As Integer, ByRef POSNoteCol As Integer, ByRef POSCouponCol As Integer, ByRef BarcodeCol As Integer, ByRef PLUCol As Integer, ByRef GuidCol As Integer, ByRef PriceCol As Integer, ByRef DiscountCol As Integer, ByRef MultiBuyCol As Integer, ByRef LoyaltyBarcodeCol As Integer, ByRef FlybuysCol As Integer)
    Dim Details() As String
    Dim TheField As Variant
    Dim FieldCount As Integer
    Dim FieldsFound As Integer
    With FieldColsdetails
    .PromotionNameCol = -1
    .StartDateCol = -1
    .EndDateCol = -1
    
.DealNameCol = -1
 .DealTypeCol = -1
  .DealSubTypeCol = -1
   .TradeNameCol = -1
    .ProductGuidCol = -1
    .GreenCrossCodeCol = -1
    .BarcodeCol = -1
    .PLUCodeCol = -1
    .ProductPriceCol = -1
    .MinQtyCol = -1
    .MaxQtyCol = -1
    .MinSpendCol = -1
    .DealPriceDollarCol = -1
    .DealDiscountTypeCol = -1
    .DealDiscountValueCol = -1
    .MultiBuyQtyCol = -1
    .MultiBuyPricecol = -1
    .PromotionPOSNoteCol = -1
    .PromptForPOSNoteCol = -1
    .FreeProductCol = -1
    
        Details = Split(RawData, Delimiter)
    
    For Each TheField In Details
        
        Select Case UCase(TheField)
       Case "PROMOTION NAME"
            .PromotionNameCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PROMOTION START DATE"
            .StartDateCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PROMOTION END DATE"
            .EndDateCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PRIORITY PROMOTION FLAG"
        Case "STATUS"
        Case "DEAL NAME"
             .DealNameCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DEAL START DATE"
        Case "DEAL END DATE"
        Case "DEAL TYPE"
            .DealTypeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DEAL SUB-TYPE"
            .DealSubTypeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PRODUCT NAME"
            .TradeNameCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PRODUCT GUID"
            .ProductGuidCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "HEAD OFFICE ID", "GXHID"
            .GreenCrossCodeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "BARCODES"
             .BarcodeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "MANUFACTURERS PRODUCT CODE"
        Case "PLU CODE"
            .PLUCodeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PRODUCT PRICE"
        .ProductPriceCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "MIN QTY"
        .MinQtyCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "MAX QTY"
         .MaxQtyCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "MIN SPEND $"
        .MinSpendCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DEAL PRICE $"
        .DealPriceDollarCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DEAL DISCOUNT TYPE"
        .DealDiscountTypeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DEAL DISCOUNT VALUE"
        .DealDiscountValueCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "QUANTITY BREAKS"
        .MultiBuyQtyCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "QUANTITY BREAK PRICES"
        .MultiBuyPricecol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PROMOTION POS NOTE"
         .PromotionPOSNoteCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PROMPT READ PROMO POS NOTE"
          .PromptForPOSNoteCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PROMO RECEIPT NOTE"
        Case "PRODUCT A OR B"
        Case "FREE PRODUCT FLAG"
           .FreeProductCol = FieldCount
            FieldsFound = FieldsFound + 1
        
        End Select
        FieldCount = FieldCount + 1
    Next TheField
    
    Dim TheMessage As String
    If FieldsFound <> 23 Then
        TheMessage = "No Of Fields should be 23 Fields Found " & FieldsFound
        If .PromotionNameCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column PROMOTION NAME Missing "
        End If
        If .StartDateCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column PROMOTION START DATE Missing "
        End If
        If .EndDateCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column PROMOTION END DATE Missing "
        End If
        If .DealNameCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "  Column DEAL NAME Missing "
        End If
        If .DealTypeCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column DEAL TYPE Missing "
        End If
        If .DealSubTypeCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column DEAL SUB-TYPE Missing "
        End If
        If .TradeNameCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column PRODUCT NAME Missing "
        End If
        If .ProductGuidCol = -1 Then
            TheMessage = TheMessage & vbCrLf & "  Column PRODUCT GUID Missing "
        End If
        If .GreenCrossCodeCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "  Column HEAD OFFICE ID Missing "
        End If
        If .BarcodeCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "  Column BARCODES Missing "
        End If
    End If
    If TheMessage <> "" Then
        MsgBox ("Column Names have changed again. Check case statement in GETPBLColumns" & vbCrLf & TheMessage)
    End If
      End With
    
End Sub

Private Sub GetPBLColumns(ByVal RawData As String, ByVal Delimiter As String, ByRef StartDateCol As Long, ByRef EndDateCol As Long, ByRef DescriptionCol As Long, ByRef POSNoteCol As Long, ByRef POSCouponCol As Long, ByRef BarcodeCol As Long, ByRef PLUCol As Long, ByRef GuidCol As Long, ByRef PriceCol As Long, ByRef DiscountCol As Long, ByRef MultiBuyCol As Long)
'Private Sub GetPBLColumns(ByVal RawData As String, ByVal Delimiter As String, ByRef StartDateCol As Integer, ByRef EndDateCol As Integer, ByRef DescriptionCol As Integer, ByRef POSNoteCol As Integer, ByRef POSCouponCol As Integer, ByRef BarcodeCol As Integer, ByRef PLUCol As Integer, ByRef GuidCol As Integer, ByRef PriceCol As Integer, ByRef DiscountCol As Integer, ByRef MultiBuyCol As Integer, ByRef LoyaltyBarcodeCol As Integer, ByRef FlybuysCol As Integer)
    Dim Details() As String
    Dim TheField As Variant
    Dim FieldCount As Integer
    Dim FieldsFound As Integer
    StartDateCol = -1
    EndDateCol = -1
    DescriptionCol = -1
    POSNoteCol = -1
    POSCouponCol = -1
    BarcodeCol = -1
    PLUCol = -1
    GuidCol = -1
    PriceCol = -1
    DiscountCol = -1
    MultiBuyCol = -1
'    LoyaltyBarcodeCol = -1
'    FlybuysCol = -1
    
        Details = Split(RawData, Delimiter)
    
    For Each TheField In Details
        
        Select Case UCase(TheField)
        Case "SPECIALSTART", "SPECIALSTARTDATE"
            StartDateCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALFINISH", "SPECIALFINISHDATE"
            EndDateCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "DESCRIPTION"
            DescriptionCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PRODUCTGUID"
            GuidCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "BARCODES TXT", "TXT BARCODES", "BARCODES TEXT", "BARCODES"
            BarcodeCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "PLUCODE"
            PLUCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALSTOREMAXITEMS"
        Case "SPECIALPRICEBREAK", "SPECIALPRICE"
            PriceCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALDISCOUNTBREAK", "SPECIALDISCOUNT"
            DiscountCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALCOUPONPROMPT"
            POSCouponCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALPOSNOTE"
            POSNoteCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "SPECIALPROMPTOVERMAX"
        Case "SPECIALNODISCOUNT"
        Case "SPECIALNOCLUB"
        Case "SPECIALNORETURN"
'        Case "SPECIALLOYALTY"
'
'            FlybuysCol = FieldCount
'            FieldsFound = FieldsFound + 1
        Case "MANFPRODUCTCODE"
        Case "NONDIMINISHING"
        Case "FREE"
'        Case "LOYALTYBARCODE", "SPECIALLOYALTYBARCODE"
'            LoyaltyBarcodeCol = FieldCount
'            FieldsFound = FieldsFound + 1
        Case "SPECIADOLLARDISCOUNTBREAK"
            
        Case "SPECIALQTYBREAK", "SPECIALQUANTITY"
            MultiBuyCol = FieldCount
            FieldsFound = FieldsFound + 1
        Case "OVERRIDECLUBPOINTS"
        Case "OVERRIDEVALUE"
        Case "SPECIALOVERRIDESALTPRICE"
        Case "SPECIALPROMPTFORPOSNOTEREAD"
        Case "SPECIALRECEIPTNOTE"
        Case "SPECIALPROMPTFORCOUPON"
        Case "SPECIALPOINTSBREAK"
        End Select
        FieldCount = FieldCount + 1
    Next TheField
    Dim TheMessage As String
    If StartDateCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALSTART', 'SPECIALSTARTDATE' Missing"
    End If
    If EndDateCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALFINISH', 'SPECIALFINISHDATE' Missing"
    End If
    If DescriptionCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'DESCRIPTION' Missing"
    End If
    If POSNoteCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALPOSNOTE' Missing"
    End If
    If POSCouponCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALCOUPONPROMPT' Missing"
    End If
    If BarcodeCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'BARCODES TXT', 'TXT BARCODES', 'BARCODES TEXT', 'BARCODES' Missing"
    End If
    
    If PLUCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'PLUCODE' Missing"
    End If
    If PriceCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALPRICEBREAK', 'SPECIALPRICE' Missing"
    End If
    
    If DiscountCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALDISCOUNTBREAK', 'SPECIALDISCOUNT' Missing"
    End If
    
    If MultiBuyCol = -1 Then
        TheMessage = TheMessage & vbCrLf & "'SPECIALQTYBREAK', 'SPECIALQUANTITY' Missing"
    End If
'    If LoyaltyBarcodeCol = -1 Then
'        TheMessage = TheMessage & vbCrLf & "'LOYALTYBARCODE', 'SPECIALLOYALTYBARCODE' Missing"
'    End If
'    If FlybuysCol = -1 Then
'        TheMessage = TheMessage & vbCrLf & "'SPECIALLOYALTY' Missing"
'    End If
    If TheMessage <> "" Then
        MsgBox ("Column Names have changed again. Check case statement in GETPBLColumns" & vbCrLf & TheMessage)
    End If
        
    
End Sub

Private Function GetSpecialID(ByVal SpecialName As String, ByVal BrandID As Long, ByVal CreateNew As Boolean) As Long
    Dim dynSpecial As Recordset
    Dim SQL As String
    Dim SpecialID As Long
    Dim CloseDb As Boolean
    
   On Error GoTo GetSpecialID_Error

    Call OpenSupplierDB(CloseDb)
    If Len(SpecialName) > 50 Then
        SpecialName = LEFT(SpecialName, 50)
    End If
    SQL = "SELECT * FROM Specials "
    SQL$ = SQL & " WHERE SpecialName = " & Chr(34) & SpecialName & Chr(34)
    SQL$ = SQL & " AND BrandID = " & BrandID
    Set dynSpecial = dbSuppliers.OpenRecordSet(SQL$, dbOpenDynaset)
    If dynSpecial.EOF = True Then
        If CreateNew = True Then
            dynSpecial.AddNew
            Call DBValidate(dynSpecial("SpecialName"), SpecialName)
            dynSpecial("BrandID") = BrandID
            SpecialID = dynSpecial("SpecialID")
            dynSpecial.Update
        Else
            SpecialID = 0
        End If
    Else
        SpecialID = dynSpecial("SpecialID")
    End If
    dynSpecial.ClsRS
    Set dynSpecial = Nothing
    Call CloseSupplierDB(CloseDb)

GetSpecialID_Exit:
   On Error Resume Next
   GetSpecialID = SpecialID
   Exit Function

GetSpecialID_Error:
    If SYSLOG(Err, "GetSpecialID in frmImportSpecials in " & app.ExeName) Then
        Resume GetSpecialID_Exit
    Else
        Resume Next
    End If

End Function

Private Sub cmdImportBarcodes_Click()
    Dim Barcode As String
    Dim GUID As String
    Dim Description As String
    Dim hdlFile As Integer
    Dim DefaultDir As String
    Dim FilePath As String
    Dim FileName As String
    Dim Details As String
    Dim TheDetails() As String
    Dim dynDetails As Recordset
    Dim Progress As ProgressIndicator
    Dim Count As Integer
On Error GoTo cmdImportBarcodes_Click_Error
    DefaultDir = readini("Defaults", "SpecialsDefaultDir", PathName, GetPOSPricesPath)
    FileName = GetOpenFile(DefaultDir, "(CSV (*.csv)|*.csv|text (*.txt)|*.txt|All (*.*)|*.*", FilePath)
    If FileExists(FileName) Then
        Set Progress = New ProgressIndicator
        Progress.RecordCount = 10000
        Progress.Show "Importing T Barcodes"
        OpenSettingsDB
        hdlFile = FreeFile
        Close #hdlFile
        Open FileName For Input As hdlFile
        Do Until EOF(hdlFile)
            Count = Count + 1
            Progress.increment Count
            Line Input #hdlFile, Details
            If InStr(Details, vbTab) Then
                TheDetails = Split(Details, vbTab)
            Else
                TheDetails = Split(Details, ",")
            End If
            If IsNumeric(TheDetails(2)) Or InStr(TheDetails(2), ",") Then
                If TheDetails(1) <> "" Then
                    Set dynDetails = dbSettings.OpenRecordSet(SQL_Select & "TBarcodes WHERE TGuid ='" & GUID & "'", dbOpenDynaset)
                    If dynDetails.EOF Then
                        dynDetails.AddNew
                    Else
                        dynDetails.Edit
                    End If
                     dynDetails("TGuid") = TheDetails(1)
                     dynDetails("Barcode") = Replace(TheDetails(2), Chr(34), "")
                    dynDetails.Update
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                End If
            End If
            
        Loop
        Close #hdlFile
        hdlFile = FreeFile
        CloseSettingsDB
        Progress.Hide
        Set Progress = Nothing
    End If
  
cmdImportBarcodes_Click_Exit:
   On Error Resume Next
Exit Sub

cmdImportBarcodes_Click_Error:
   If SYSLOG(Err, "cmdImportBarcodes_Click in frmImportSpecials in " & AppVersion()) Then
       Resume cmdImportBarcodes_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub cmdSave_Click()
    Dim BrandID As Long
    Dim SQL As String
    Dim snaSpecialDetails As Recordset
    Dim Count As Integer
    Dim dynSpecials As Recordset
    'Dim UpdateDetails() As DBUpdateData
    Dim HasError As Boolean
    Dim Progress As ProgressIndicator
On Error GoTo cmdSave_Click_Error
    
    BrandID = cboBrand.ItemData(cboBrand.ListIndex)
    Call Me.vsImport.SaveGrid(GetGridname, flexFileAll)
    Call writeini(str(BrandID), "Last Special Name", Me.txtSpecialName.text, GetINIname)
    Call writeini(str(BrandID), "Last Special SD", Me.txtStartDate.text, GetINIname)
    Call writeini(str(BrandID), "Last Special ED", Me.txtFinishDate.text, GetINIname)
    UpdateCodes (BrandID)
    If txtSpecialName.text = "" Then
        MsgBox "Please enter a special name", vbOKOnly
        Exit Sub
    End If
    If BrandID <= 0 Then
        MsgBox "Please enter a brand name", vbOKOnly
        Exit Sub
    End If
    If IsDate(txtStartDate.text) = False Then
        MsgBox "Please enter a valid start date", vbOKOnly
        Exit Sub
    End If
    If IsDate(txtFinishDate.text) = False Then
        MsgBox "Please enter a valid finish date", vbOKOnly
        Exit Sub
    End If
    If CDate(txtStartDate) > CDate(txtFinishDate) Then
        MsgBox "Start date can't be greater than the finish date", vbOKOnly
        Exit Sub
    End If
    
    glbSpecialID = GetSpecialID(txtSpecialName.text, BrandID, True)
    If glbSpecialID > 0 Then
        OpenSupplierDB
        'Update the start and finish date
        SQL$ = "UPDATE Specials SET StartDate = #" & FormatDate(txtStartDate, False) & "#"
        SQL$ = SQL$ & ",FinishDate = #" & FormatDate(txtFinishDate, False) & "#"
        'Change Greg 22/5/07 Make so that updates the right special id
        SQL$ = SQL$ & " WHERE SpecialID = " & glbSpecialID
        Call dbSuppliers.Execute(SQL$, dbFailOnError)
        'Now delete all existsing items for this special
        If MsgBox("Clear All Previous Info For This Special. Say No Only If You are saving fixes.", vbYesNo) = vbYes Then
            SQL$ = "DELETE * FROM SpecialItems WHERE SpecialID = " & glbSpecialID
            Call dbSuppliers.Execute(SQL$, dbFailOnError)
        End If
        Dim SaveAll As Boolean
        SaveAll = True
'        SQL$ = "DELETE * FROM SpecialItems WHERE SpecialID = " & glbSpecialID
'        Call dbSuppliers.Execute(SQL$, dbFailOnError)
        'now lets add these items in
        With vsImport
            Set Progress = New ProgressIndicator
            Progress.RecordCount = .Rows - 1
            Call Progress.Show("Saving Special")
            For Count% = .Rows - 1 To 0 Step -1
                Call Progress.increment(Progress.RecordCount - Count)
                If SaveAll Or .IsSelected(Count) Then
                If UCase(.TextMatrix(Count%, COL_SPECIAL_INVALID)) = "N" Then
                    HasError = False
                    Set dynSpecials = dbSuppliers.OpenRecordSet(SQL_Select & " SpecialItems WHERE SpecialID = " & glbSpecialID & " AND Barcode = '" & .TextMatrix(Count, COL_SPECIAL_BARCODE) & "'", dbOpenDynaset)
                    If dynSpecials.EOF Then
                        dynSpecials.AddNew
                    Else
                        dynSpecials.Edit
                    End If
                    If IsYesOrTrue(.TextMatrix(Count, COL_SPECIAL_LOWEST_FREE)) Then
                        dynSpecials("lowestFree") = True
                    Else
                        dynSpecials("lowestFree") = False
                    End If
                    '7/11/2016 Limin added: Don't Save Flybuys any more
'                    If UCase(.TextMatrix(Count, COL_SPECIAL_FLYBUYS)) = "Y" Then
'                        dynSpecials("Flybuys") = True
'                    Else
'                        dynSpecials("Flybuys") = False
'                    End If
                    dynSpecials("GroupID") = .TextMatrix(Count, COL_SPECIAL_GROUPID)
                    dynSpecials("SpecialID") = glbSpecialID
                    dynSpecials("Barcode") = .TextMatrix(Count, COL_SPECIAL_BARCODE)
                    
                    dynSpecials("Description") = Trim(.TextMatrix(Count, COL_SPECIAL_DESCRIPTION))
                    If IsYesOrTrue(.TextMatrix(Count, COL_SPECIAL_COMBO)) Then
                        dynSpecials("MultiBuy") = "ALL"
                    ElseIf Val(.TextMatrix(Count, COL_SPECIAL_MULTIBUY)) <= "1" And IsNumeric(.TextMatrix(Count, COL_SPECIAL_MULTIBUY)) And InStr(.TextMatrix(Count, COL_SPECIAL_MULTIBUY), ",") = 0 Then
                        dynSpecials("MultiBuy") = ""
                    Else
                        dynSpecials("MultiBuy") = .TextMatrix(Count, COL_SPECIAL_MULTIBUY)
                    End If
                    
                    dynSpecials("MultirETAIL") = .TextMatrix(Count, COL_SPECIAL_MULTIRETAIL)
'                    Call PrepareUpdateDetails(UpdateDetails, "SpecialID", glbSpecialID, dbLong, True)
'                    Call PrepareUpdateDetails(UpdateDetails, "Barcode", .TextMatrix(Count, COL_SPECIAL_BARCODE), dbText)
'                    Call PrepareUpdateDetails(UpdateDetails, "Description", .TextMatrix(Count, COL_SPECIAL_DESCRIPTION), dbText)
                    'change Greg 3/6/10 Make so can handle multibuys
                    'Call PrepareUpdateDetails(UpdateDetails, "Multibuy", .TextMatrix(Count, COL_SPECIAL_MULTIBUY), dbText)
                    If InStr(.TextMatrix(Count, COL_SPECIAL_PRICE), "%") Then
                        dynSpecials("SpecialRetail") = 0
                        dynSpecials("SpecialpERCENT") = Val(.TextMatrix(Count, COL_SPECIAL_PRICE))
'                        Call PrepareUpdateDetails(UpdateDetails, "SpecialRetail", 0, dbLong)
'                        Call PrepareUpdateDetails(UpdateDetails, "SpecialPercent", Val(.TextMatrix(Count, COL_SPECIAL_PRICE)), dbSingle)
                    ElseIf .TextMatrix(Count, COL_SPECIAL_PRICE) = "" Then
                        dynSpecials("SpecialRetail") = -1
                        dynSpecials("SpecialpERCENT") = 0
                    Else
                        dynSpecials("SpecialRetail") = dollarsToCents(.TextMatrix(Count, COL_SPECIAL_PRICE))
                        dynSpecials("SpecialpERCENT") = 0
'                        Call PrepareUpdateDetails(UpdateDetails, "SpecialRetail", dollarsToCents(.TextMatrix(Count, COL_SPECIAL_PRICE)), dbLong)
'                        Call PrepareUpdateDetails(UpdateDetails, "SpecialPercent", 0, dbSingle)
                    End If
                    Call DBValidate(dynSpecials("SpecialNote"), .TextMatrix(Count, COL_SPECIAL_POSNOTE))
                    Call DBValidate(dynSpecials("GiftBarcode"), .TextMatrix(Count, COL_SPECIAL_GIFTBARCODE))
                    Call DBValidate(dynSpecials("GiftPharmacode"), .TextMatrix(Count, COL_SPECIAL_GIFTPHARMACODE))
                    Call DBValidate(dynSpecials("ComboPurchase"), FBool(.TextMatrix(Count, COL_SPECIAL_COMBO)))
                    Call DBValidate(dynSpecials("MessageIsQuestion"), FBool(.TextMatrix(Count, COL_SPECIAL_MESSAGE_IS_Question)))
                    Call DBValidate(dynSpecials("DealName"), .TextMatrix(Count, COL_DEAL_NAME))
                      Call DBValidate(dynSpecials("SecondaryGroupID"), .TextMatrix(Count, COL_SECONDARY_GROUPID))
'                    If UCase(.TextMatrix(Count, COL_SPECIAL_COMBO)) = "Y" Then
'                        dynSpecials("ComboPurchase") = True
'                    Else
'                        dynSpecials("ComboPurchase") = False
'                    End If
                    'SQL$ = GetDBInsertSQL("SpecialItems", UpdateDetails)
                    On Error GoTo cmdSave_Execute
                    dynSpecials.Update
                    'Call dbSuppliers.Execute(SQL$, dbFailOnError)
                    On Error GoTo cmdSave_Click_Error
                    
                    If HasError = False Then
                        .RemoveItem (Count%)
                    End If
                    dynSpecials.ClsRS
                    Set dynSpecials = Nothing
                Else
                    
                End If
                End If
            Next Count%
        End With
        Progress.Hide
        'New Nigel 15.2.2010 Set the password on the database
        If frmImportSpecials.txtPassword.text <> "" Then
            dbSuppliers.Execute "UPDATE Specials SET Password = '" & EncryptPassword(UCase(Trim(frmImportSpecials.txtPassword.text))) & "' WHERE SpecialID = " & glbSpecialID, dbFailOnError
        End If
        
        CloseSupplierDB
    End If
  
cmdSave_Click_Exit:
   On Error Resume Next
Exit Sub

cmdSave_Click_Error:
   If SYSLOG(Err, "cmdSave_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume cmdSave_Click_Exit
   Else
       Resume Next
   End If
cmdSave_Execute:
    HasError = True
    Call SYSLOG(Err, "cmdSave_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision)
    Resume Next
End Sub
Private Function IsYesOrTrue(ByVal v_Value As String) As Boolean
  IsYesOrTrue = UCase(v_Value) = "Y" Or FBool(v_Value)
End Function
Private Sub Form_Load()
On Error GoTo Form_Load_Error
    Call LoadBrands(cboBrand)
    ClearOldSpecials
    cboDescription.ListIndex = 0
    cboBarcode.ListIndex = 0
    cboSpecial.ListIndex = 0
    vsImport.ColWidth(COL_SPECIAL_NEW) = 0
    vsImport.ColHidden(COL_SPECIAL_NEW) = True
    vsImport.ColWidth(COL_SPECIAL_CHANGE) = 0
    vsImport.ColHidden(COL_SPECIAL_CHANGE) = True
    vsImport.ColWidth(COL_SPECIAL_INVALID) = 0
    vsImport.ColHidden(COL_SPECIAL_INVALID) = True
    Call LoadNumbersIntoListBox(cboDescription)
    Call LoadNumbersIntoListBox(cboBarcode)
    Call LoadNumbersIntoListBox(cboSpecial)
    'New Nigel 26.6.2013
    Call LoadNumbersIntoListBox(cboPOSNote)
    
    Me.lblMultiBuy.Caption = "Multi Buy Fields Are:" & vbCrLf & "ALL -This is a combo Purchase" & vbCrLf & "> Then Value or dollar amount - Purchase x or more" & vbCrLf & "Value Or Dollar Amount- Standard Multi Buy"
  
Form_Load_Exit:
   On Error Resume Next
Exit Sub

Form_Load_Error:
   If SYSLOG(Err, "Form_Load in frmImportSpecials in " & AppVersion()) Then
       Resume Form_Load_Exit
   Else
       Resume Next
   End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub SaveSpecialSettings()
    Dim dynBrand As Recordset
    Dim BrandID As Long
    Dim Count As Integer
    Dim CloseDb As Boolean
    Dim snaSettings As Recordset
   On Error GoTo SaveSpecialSettings_Error

    If cboBrand.ListIndex = -1 Then
        BrandID = 0
    Else
        BrandID = cboBrand.ItemData(cboBrand.ListIndex)
    End If
    If cboBrand.text = "" Then
        Exit Sub
    End If
'    If BrandID = 0 Then
'        Call OpenSupplierDB(CloseDB)
'        Set dynBrand = dbSuppliers.OpenRecordset("SELECT * FROM Brand WHERE BrandName = '" & TitleCase(cboBrand.text) & "'", dbOpenDynaset)
'        If dynBrand.EOF = True Then
'            dynBrand.AddNew
'            BrandID = FNulls(dynBrand("BrandID"))
'            dynBrand("BrandName") = TitleCase(cboBrand.text)
'            dynBrand.Update
'            cboBrand.AddItem TitleCase(cboBrand.text)
'            cboBrand.ItemData(cboBrand.ListCount - 1) = BrandID
''            cboBrand.Sorted = True
'            cboBrand.Refresh
'        Else
'            BrandID = FNulls(dynBrand("BrandID"))
'        End If
'        For Count = 0 To cboBrand.ListCount - 1
'            If cboBrand.ItemData(Count%) = BrandID Then
'                cboBrand.ListIndex = Count%
'                Exit For
'            End If
'        Next Count%
'        Call CloseSupplierDB(CloseDB)
'    End If
    If cboBrand.text <> "" Then
        OpenSettingsDB
        Set snaSettings = dbSettings.OpenRecordSet(SQL_Select & "SuppUDSpecialsSettings WHERE BrandID = " & cboBrand.ItemData(cboBrand.ListIndex), dbOpenDynaset)
        If snaSettings.EOF Then
            snaSettings.AddNew
        Else
            snaSettings.Edit
        End If
        snaSettings("BrandID") = BrandID
         snaSettings("Special") = cboSpecial.ListIndex
         snaSettings("Description") = cboDescription.ListIndex
         snaSettings("Barcode") = cboBarcode.ListIndex
         snaSettings("FieldDelimiter") = cboFieldDelimiter.ListIndex
         snaSettings("RecordDelimiter") = cboRecordDelimiter.ListIndex
        '16/3/2004 Greg add in option to ignore first line
        If Me.chkIgnoreFirst.value = vbChecked Then
            snaSettings("IgnoreFirstLine") = "TRUE"
        Else
            snaSettings("IgnoreFirstLine") = "FALSE"
        End If
        snaSettings("POSNote") = cboPOSNote.ListIndex
        snaSettings.Update
        snaSettings.ClsRS
        Set snaSettings = Nothing
    End If

SaveSpecialSettings_Exit:
   On Error Resume Next
   Exit Sub

SaveSpecialSettings_Error:
    If SYSLOG(Err, "SaveSpecialSettings in frmImportSpecials in " & app.ExeName) Then
        Resume SaveSpecialSettings_Exit
    Else
        Resume Next
    End If
End Sub
Public Sub LoadBrands(ByVal ListBox As ComboBox)
    
    On Error GoTo ErrTrap
    
    Dim snaBrand As Recordset
    
    Call OpenSupplierDB
    
    Set snaBrand = dbSuppliers.OpenRecordSet("SELECT * FROM Brand Order by BrandName", dbOpenSnapshot)
    With ListBox
        .Clear
        Do Until snaBrand.EOF
            .AddItem snaBrand("BrandName")
            .ItemData(.ListCount - 1) = Val(FNulls(snaBrand("BrandID")))
            snaBrand.MoveNext
        Loop
    End With
    snaBrand.ClsRS
    Set snaBrand = Nothing
    Call CloseSupplierDB

Exit Sub
ErrTrap:
    Call SYSLOG(Err, "LoadBrands in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    End
End Sub

Private Sub mmuFileReplace_Click()
    If Me.MousePointer <> vbHourglass Then
        Call frmFindAndReplace.Show
    End If
End Sub

Private Sub mnuFileAddBrand_Click()
    Dim BrandName As String
On Error GoTo mnuFileAddBrand_Click_Error
    BrandName = InputBox("Please enter the new brand name?")
    If BrandName <> "" Then
    Call OpenSupplierDB
    Call dbSuppliers.Execute("INSERT INTO Brand([BrandName]) VALUES(" & Chr(34) & BrandName & Chr(34) & ")", dbFailOnError)
    Call CloseSupplierDB
   Call LoadBrands(cboBrand)
    End If
  
mnuFileAddBrand_Click_Exit:
   On Error Resume Next
Exit Sub

mnuFileAddBrand_Click_Error:
   If SYSLOG(Err, "mnuFileAddBrand_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume mnuFileAddBrand_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub mnuGroupSpecials_Click()
    Dim GroupID As Long
    Dim RowGroupID As Long
    Dim RowCount As Long
On Error GoTo mnuGroupSpecials_Click_Error
    With vsImport
        For RowCount = 1 To .Rows - 1
            If Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID)) > RowGroupID Then
               RowGroupID = Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID))
            End If
            If .IsSelected(RowCount) Then
                If Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID)) > 0 Then
                    GroupID = Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID))
                    Exit For
                End If
            End If
        Next RowCount
        If GroupID = 0 Then
            GroupID = RowGroupID + 1
        End If
        For RowCount = 1 To .Rows - 1
            'Change Nigel 30.10.2013
            If .IsSelected(RowCount) = True Then
                .TextMatrix(RowCount, COL_SPECIAL_GROUPID) = GroupID
            End If
        Next RowCount
    End With
  
mnuGroupSpecials_Click_Exit:
   On Error Resume Next
Exit Sub

mnuGroupSpecials_Click_Error:
   If SYSLOG(Err, "mnuGroupSpecials_Click in frmImportSpecials in " & AppVersion()) Then
       Resume mnuGroupSpecials_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub mnuLoadFile_Click()
    Dim BrandID As Long
On Error GoTo mnuLoadFile_Click_Error
    BrandID = GetValueFromListBox(cboBrand)
    If BrandID = 0 Then
        MsgBox ("Please select a brand to import")
    Else
       Call Me.vsImport.LoadGrid(GetGridname, flexFileAll)
    Me.txtSpecialName.text = readini(BrandID, "Last Special Name", "", GetINIname)
       Me.txtStartDate.text = readini(BrandID, "Last Special SD", "", GetINIname)
       Me.txtFinishDate.text = readini(BrandID, "Last Special ED", "", GetINIname)
       Setlines
    End If
mnuLoadFile_Click_Exit:
   On Error Resume Next
Exit Sub

mnuLoadFile_Click_Error:
   If SYSLOG(Err, "mnuLoadFile_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume mnuLoadFile_Click_Exit
   Else
       Resume Next
   End If
    
End Sub
Private Sub Setlines()
    Me.lblLines.Caption = "Items:" & Me.vsImport.Rows - 1
End Sub
Private Function GetGridname() As String
    If DirExists(SpecialsFilePath) = False Then
        Call Make_The_Dir(SpecialsFilePath)
    End If
    GetGridname = SpecialsFilePath & GetValueFromListBox(cboBrand) & ".asc"
End Function
Private Function GetINIname() As String
    GetINIname = SpecialsFilePath & "SpecialSave.ini"
End Function
Private Sub mnuSaveGrid_Click()
    On Error Resume Next
    Dim BrandID As Long
    BrandID = cboBrand.ItemData(cboBrand.ListIndex)
    Call Me.vsImport.SaveGrid(GetGridname, flexFileAll)
    Call writeini(str(BrandID), "Last Special Name", Me.txtSpecialName.text, GetINIname)
    Call writeini(str(BrandID), "Last Special SD", Me.txtStartDate.text, GetINIname)
    Call writeini(str(BrandID), "Last Special ED", Me.txtFinishDate.text, GetINIname)
End Sub

Private Sub mnuSortErrors_Click()
On Error GoTo mnuSortErrors_Click_Error
     vsImport.Col = COL_SPECIAL_INVALID
    vsImport.Sort = flexSortStringDescending
  
mnuSortErrors_Click_Exit:
   On Error Resume Next
Exit Sub

mnuSortErrors_Click_Error:
   If SYSLOG(Err, "mnuSortErrors_Click in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume mnuSortErrors_Click_Exit
   Else
       Resume Next
   End If
End Sub



Private Sub mnuUNGroupSpecials_Click()
    Dim RowCount As Long
On Error GoTo mnuUNGroupSpecials_Click_Error
    With vsImport
        For RowCount = 1 To .Rows - 1
            If .IsSelected(RowCount) Then
                .TextMatrix(RowCount, COL_SPECIAL_GROUPID) = 0
            End If
        Next RowCount
    End With
  
mnuUNGroupSpecials_Click_Exit:
   On Error Resume Next
Exit Sub

mnuUNGroupSpecials_Click_Error:
   If SYSLOG(Err, "mnuUNGroupSpecials_Click in frmImportSpecials in " & AppVersion()) Then
       Resume mnuUNGroupSpecials_Click_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub txtFinishDate_GotFocus()
    Call txtBox_GotFocus(txtFinishDate)
End Sub


Private Sub txtFinishDate_LostFocus()
    Call txtBox_LostFocus(txtFinishDate, , True)
End Sub

Private Sub txtSpecialName_GotFocus()
    Call txtBox_GotFocus(txtSpecialName)
End Sub


Private Sub txtSpecialName_LostFocus()
    Call txtBox_LostFocus(txtSpecialName)
End Sub

Private Sub txtStartDate_GotFocus()
    Call txtBox_GotFocus(txtStartDate)
End Sub

Private Sub txtStartDate_LostFocus()
    Call txtBox_LostFocus(txtStartDate, , True)
End Sub

Private Sub vsImport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo vsImport_AfterEdit_Error
    
    
    With vsImport
    
        'New Nigel 24.3.2010
        .Editable = flexEDNone

        If Col = COL_SPECIAL_DESCRIPTION Then
        
            .TextMatrix(Row, COL_SPECIAL_LENGTH) = Len(.TextMatrix(Row, Col))
            If Col = COL_SPECIAL_LOWEST_FREE Then
                If dollarsToCents(.TextMatrix(Row, COL_SPECIAL_PRICE)) = 0 Then
                    .TextMatrix(Row, COL_SPECIAL_PRICE) = ""
                End If
            End If
            If Len(.TextMatrix(Row, Col)) > 40 Then
                .TextMatrix(Row, COL_SPECIAL_INVALID) = "Y"
                 .Cell(flexcpBackColor, Row, 0, Row, vsImport.Cols - 1) = vbRed
    '            .Row = Row
    '            .Col = Col
    '            .CellBackColor = shpInvalid.FillColor
            Else
                .TextMatrix(Row, COL_SPECIAL_INVALID) = "N"
                 .Cell(flexcpBackColor, Row, 0, Row, vsImport.Cols - 1) = vbWhite
    '            .Row = Row
    '            .Col = Col
    '            .CellBackColor = vbWhite
            End If
        
        
        End If
        If Me.MousePointer = vbHourglass Then
            vsImport.Editable = flexEDKbdMouse
            Exit Sub
        End If
        If Col <> COL_SPECIAL_GROUPID Then
            If Val(.TextMatrix(Row, COL_SPECIAL_GROUPID)) > 0 Then
                If OldValue <> .TextMatrix(Row, Col) Then
                    If MsgBox("Change All Items In the Group To The Same Value?", vbYesNo) = vbYes Then
                        Dim RowCount As Long
                        For RowCount = 1 To .Rows - 1
                            If Val(.TextMatrix(Row, COL_SPECIAL_GROUPID)) = Val(.TextMatrix(RowCount, COL_SPECIAL_GROUPID)) Then
                                .TextMatrix(RowCount, Col) = .TextMatrix(Row, Col)
                                If Col = COL_SPECIAL_LOWEST_FREE Then
                                    If dollarsToCents(.TextMatrix(RowCount, COL_SPECIAL_PRICE)) = 0 Then
                                        .TextMatrix(RowCount, COL_SPECIAL_PRICE) = ""
                                    End If
                                End If
                            End If
                        Next RowCount
                    End If
                End If
            End If
        End If
        'New Nigel 30.10.2013
        
        
    End With
vsImport_AfterEdit_Exit:
   On Error Resume Next
   vsImport.Editable = flexEDKbdMouse
Exit Sub

vsImport_AfterEdit_Error:
   If SYSLOG(Err, "vsImport_AfterEdit in frmImportSpecials in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume vsImport_AfterEdit_Exit
   Else
       Resume Next
   End If
End Sub

Private Sub vsImport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
On Error GoTo vsImport_BeforeEdit_Error
    
    'New Nigel 30.10.2013
    If Me.MousePointer = vbHourglass Then
        Exit Sub
    End If
    If vsImport.MouseRow = 0 Then
        Exit Sub
    End If
    
    OldValue = vsImport.TextMatrix(Row, Col)
    'New Nigel 30.10.2013
'    Me.MousePointer = vbHourglass
'    vsImport.WordWrap = True
'    vsImport.EditCell
'    Me.MousePointer = vbNormal
  
vsImport_BeforeEdit_Exit:
   On Error Resume Next
Exit Sub

vsImport_BeforeEdit_Error:
   If SYSLOG(Err, "vsImport_BeforeEdit in frmImportSpecials in " & AppVersion()) Then
       Resume vsImport_BeforeEdit_Exit
   Else
       Resume Next
   End If
End Sub

'Private Sub vsImport_EnterCell()
    'Change Nigel 30.10.2013
    'vsImport.EditCell
    'vsImport.EditSelStart = 0
'End Sub
