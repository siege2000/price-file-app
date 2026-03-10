Attribute VB_Name = "basUpdateSupplier"
Option Explicit
Global Const COL_IMPORT_SUPPLIER = 0
Global Const COL_IMPORT_CODE = 1
Global Const COL_IMPORT_PHARMACODE = 2
Global Const COL_IMPORT_OLD_DESCRIPTION = 3
Global Const COL_IMPORT_DESCRIPTION = 4
Global Const COL_IMPORT_LENGTH = 5
Global Const COL_IMPORT_OLDCOST = 6
Global Const COL_IMPORT_COST = 7
Global Const COL_IMPORT_RETAIL = 8
Global Const COL_IMPORT_BARCODE = 9
Global Const COL_IMPORT_OUTERS = 10
Global Const COL_IMPORT_PRODUCTGROUP = 11
Global Const COL_IMPORT_GST = 12
Global Const COL_IMPORT_COSTINCLGST = 13
Global Const COL_IMPORT_NEW = 14
Global Const COL_IMPORT_CHANGE = 15
Global Const COL_IMPORT_INVALID = 16
Global Const COL_IMPORT_SUPPLIERID = 17
Global Const COL_IMPORT_LAYOUT = 18
Global Const COL_IMPORT_CHECK = 19
Global Const COL_IMPORT_API_CODE = 20
Global Const COL_IMPORT_SIGMA_CODE = 21
Global Const COL_IMPORT_FAULDING_CODE = 22
'Global Const COL_IMPORT_CODE = 0
'Global Const COL_IMPORT_PHARMACODE = 1
'Global Const COL_IMPORT_DESCRIPTION = 2
'Global Const COL_IMPORT_OLDCOST = 3
'Global Const COL_IMPORT_COST = 4
'Global Const COL_IMPORT_RETAIL = 5
'Global Const COL_IMPORT_BARCODE = 6
'Global Const COL_IMPORT_OUTERS = 7
'Global Const COL_IMPORT_PRODUCTGROUP = 8
'Global Const COL_IMPORT_GST = 9
'Global Const COL_IMPORT_COSTINCLGST = 10
'Global Const COL_IMPORT_NEW = 11
'Global Const COL_IMPORT_CHANGE = 12
'Global Const COL_IMPORT_INVALID = 13
'Global Const COL_IMPORT_SUPPLIERID = 14
'Global Const COL_IMPORT_LAYOUT = 15
'Global Const COL_IMPORT_GST = 8
'Global Const COL_IMPORT_COSTINCLGST = 9
'Global Const COL_IMPORT_NEW = 10
'Global Const COL_IMPORT_CHANGE = 11
'Global Const COL_IMPORT_INVALID = 12
'Global Const COL_IMPORT_SUPPLIERID = 13
'Global Const COL_IMPORT_LAYOUT = 14
Global BasePath As String
Global SupplierFilePath As String
Global SpecialsFilePath As String
Global ProcessedPath As String
Global FailedPath As String
'Global glbIniLocation As String
Global Const MAX_WORKSHEETS = 30
Global Const MAX_REPLACE = 50
Global glbSaveAfterImport As Boolean
Global dbSettings As Database
Global glbSettingsDBOpen As Boolean

Global Const INDEX_PLU = 0
Global Const INDEX_DESCRIPTION = 1
Global Const INDEX_ALTDESCRIPTION = 2
Global Const INDEX_BARCODE = 3
Global Const INDEX_APNBARCODE = 4
Global Const INDEX_MANUFACTURER = 5
Global Const INDEX_RETAIL1 = 6
Global Const INDEX_UNITOFRETAIL1 = 7
Global Const INDEX_RETAIL2 = 8
Global Const INDEX_UNITOFRETAIL2 = 9
Global Const INDEX_CHANGEINDICATOR = 10
Global Const INDEX_MARKUP = 11
Global Const INDEX_NORMALRETAIL = 12
Global Const INDEX_WHOLESALETAX = 13
Global Const INDEX_GSTFLAG = 14
Global Const INDEX_SCHEDULE = 15
Global Const INDEX_GSTRETAIL = 16
Global Const INDEX_PRODUCTGROUPID = 17
Global Const INDEX_SUBPRODUCTGROUPID = 18
Global Const INDEX_PACKSIZE = 19
Global Const INDEX_UNITCOST = 20
Global Const INDEX_UNITRETAIL = 21
Global Const INDEX_UNITCOST2 = 22
Global Const INDEX_UNITCOST3 = 23
Global Const INDEX_RETAIL3 = 24
Global Const INDEX_UNITOFRETAIL3 = 25
Global Const FAULDING_SUPPLIERID = 2 '9/1/2007 Limin added
Global Const SIGMA_SUPPLIERID = 3
Global Const API_SUPPLIERID = 4

Global Const COL_SPECIAL_BARCODE As Integer = 0
Global Const COL_SPECIAL_DESCRIPTION As Integer = 1
Global Const COL_SPECIAL_LENGTH As Integer = 2
Global Const COL_SPECIAL_MULTIBUY As Integer = 3
Global Const COL_SPECIAL_PRICE As Integer = 4
Global Const COL_SPECIAL_LOWEST_FREE As Integer = 5
Global Const COL_SPECIAL_NEW As Integer = 6
Global Const COL_SPECIAL_CHANGE As Integer = 7
Global Const COL_SPECIAL_INVALID As Integer = 8
Global Const COL_SPECIAL_POSNOTE As Integer = 9
Global Const COL_SPECIAL_MULTIRETAIL As Integer = 10
Global Const COL_SPECIAL_FLYBUYS As Integer = 11
Global Const COL_SPECIAL_GROUPID As Integer = 12
Global Const COL_SPECIAL_GIFTBARCODE As Integer = 13
Global Const COL_SPECIAL_GIFTPHARMACODE As Integer = 14
Global Const COL_SPECIAL_COMBO As Integer = 15
Global Const COL_SPECIAL_MESSAGE_IS_Question As Integer = 16
Global Const COL_DEAL_NAME As Integer = 17
Global Const COL_SECONDARY_GROUPID As Integer = 18
Public Type GreenCrossNewFormat
     PromotionNameCol As Integer
     
      StartDateCol As Integer
      EndDateCol As Integer
      DealNameCol As Integer
      DealTypeCol As Integer
      DealSubTypeCol As Integer
      TradeNameCol As Integer
      ProductGuidCol As Integer
      GreenCrossCodeCol As Integer
      BarcodeCol As Integer
      PLUCodeCol As Integer
      MinQtyCol As Integer
      MaxQtyCol As Integer
      MinSpendCol As Integer
      DealPriceDollarCol As Integer
      DealDiscountTypeCol As Integer
       DealDiscountValueCol As Integer
       MultiBuyQtyCol As Integer
      MultiBuyPricecol As Integer
      PromotionPOSNoteCol As Integer
      PromptForPOSNoteCol As Integer
      FreeProductCol As Integer
      ProductPriceCol As Integer
End Type
Public Type GreenCrossNewFormatDetails
     PromotionName As String
     
      StartDate As String
      EndDate As String
      DealName As String
      DealType As String
      DealSubType As String
      TradeName As String
      ProductGuid As String
      GreenCrossCode As String
      Barcode As String
      PLUCode As String
      MinQty As String
      MaxQty As String
      MinSpend As String
      DealPriceDollar As String
      DealDiscountType As String
       DealDiscountValue As String
       MultiBuyQty As String
      MultiBuyPrice As String
      PromotionPOSNote As String
      PromptForPOSNote As String
      FreeProduct As String
      ProductPrice As String
      SecondaryGroupID As String
End Type
Sub Main()
On Error GoTo Errh
    Dim Names() As String
'    glbIniLocation$ = Trim(ReadIni("POS Prices", "DBPath", "", App.Path & "\POSPrices.ini")) '"C:\Update\Pos Prices\Suppliers.mdb"
'    If InStr(UCase(glbIniLocation$), "POS PRICES") > 0 Then
'        glbIniLocation$ = LEFT(glbIniLocation$, InStr(UCase(glbIniLocation$), "POS PRICES") - 1)
'    Else
'        glbIniLocation$ = App.Path & "\"
'    End If
    If InStr(UCase(Command), "OZBRO") Then
        glbDBType = OZBro
        BasePath = readini("POS Prices", "DBPath", "", GetPOSPricesPath)
       
        If LEFT(BasePath, 2) = "\\" Then
            SupplierFilePath = Replace(UCase(BasePath), "POS PRICES\SUPPLIERS.MDB", "SupplierFiles\")
            SpecialsFilePath = Replace(UCase(BasePath), "POS PRICES\SUPPLIERS.MDB", "Specials\")
            BasePath = Replace(UCase(BasePath), "UPDATEOZ\POS PRICES\SUPPLIERS.MDB", "PhcyPro\OZPOSUpdates")
            
        Else
            BasePath = "C:\PhcyPro\OZPOSUpdates"
             SupplierFilePath = "C:\UpdateOZ\SupplierFiles\"
        End If
    Else
        glbDBType = BigBro
        BasePath = readini("POS Prices", "DBPath", "", GetPOSPricesPath)
        If LEFT(BasePath, 2) = "\\" Then
             SupplierFilePath = Replace(UCase(BasePath), "POS PRICES\SUPPLIERS.MDB", "SupplierFiles\")
             SpecialsFilePath = Replace(UCase(BasePath), "POS PRICES\SUPPLIERS.MDB", "Specials\")
            BasePath = Replace(UCase(BasePath), "POS PRICES\SUPPLIERS.MDB", "NZPOSUpdates")
            
        Else
            BasePath = "C:\Lots\NZPOSUpdates"
            SupplierFilePath = "C:\UpdateNZ\SupplierFiles\"
            SupplierFilePath = "C:\UpdateNZ\Specials\"
        End If
    End If
    
'    If LEFT(App.Path, 2) = "\\" Then
'        BasePath = Replace(UCase(App.Path), "UPDATENZ", "LOTS\NZPOSUpdates")
'    Else
'        BasePath = "C:\LOTS\NZPOSUpdates"
'    End If
    'Call SYSLOG(Nothing, BasePath)
    ProcessedPath = BasePath & "\AutoProcessed"
    FailedPath = BasePath & "\Failed"
    
    If InStr(UCase(Command), "/DRUG_UPDATE") > 0 And InStr(UCase(Command), "/POS_UPDATE") > 0 Then
        'Change Greg 26/6/06 Add in importing of suppliers
        frmOptions.Show vbModal
        
        If frmOptions.Tag = "OK" Then
            
            If frmOptions.optUpdate(0).value Then
                glbSaveAfterImport = False
                frmImport.Show
            ElseIf frmOptions.optUpdate(1).value Then
                frmUpdateSupplier.Show
            ElseIf frmOptions.optUpdate(2).value Then
                frmImportSpecials.Show
            ElseIf frmOptions.optUpdate(3).value Then
                ImportPharmacodeProducts
            ElseIf frmOptions.optUpdate(4).value Then
            ElseIf frmOptions.optUpdate(5).value Then
                CheckSupplierEmails
            ElseIf frmOptions.optUpdate(6).value Then
                'EmailSites
            ElseIf frmOptions.optUpdate(7).value Then
            UpdateStockGuid
            End If
        End If
        Unload frmOptions
'        If MsgBox("YES = Import File" & vbCrLf & "NO = Update Items", vbInformation + vbYesNo) = vbYes Then
'            frmImport.Show
'        Else
'            frmUpdateSupplier.Show
'        End If
    Else
        MsgBox "This program can only run with the command Line parameter '/DRUG_UPDATE' so please add this." & vbCrLf & "Current Command Line Is:" & Command, vbCritical, "Command Line '/DRUG_UPDATE' required"
        End
    End If
    Exit Sub
Errh:
    MsgBox Err.Description
End Sub
Public Function GetBarcode(ByVal POSNote As String, ByVal StartFrom As Long) As String
    Dim StartPoint As Long
    Dim EndPoint As Long
    Dim Result As String
    Dim Count As Integer
On Error GoTo GetBarcode_Error
    If StartFrom = 0 Then
        StartFrom = 1
    End If
    For Count = 0 To 9
        If InStr(StartFrom, POSNote, Count) > 0 Then
            If InStr(StartFrom, POSNote, Count) < StartPoint Or StartPoint = 0 Then
                StartPoint = InStr(StartFrom, POSNote, Count)
            End If
        End If
    Next Count
    If StartPoint > 0 Then
        Result = ""
        Count = 0
        For Count = StartPoint To Len(POSNote)
            If IsNumeric(Mid(POSNote, Count, 1)) Then
                Result = Result & Mid(POSNote, Count, 1)
            Else
                Exit For
            End If
            
        Next Count
        EndPoint = Count
        If Len(Result) > 6 Then
            GetBarcode = Result
        Else
            GetBarcode = GetBarcode(POSNote, EndPoint + 1)
        End If
    End If
  
GetBarcode_Exit:
   On Error Resume Next
Exit Function

GetBarcode_Error:
   If SYSLOG(Err, "GetBarcode in basUpdateSupplier in " & AppVersion()) Then
       Resume GetBarcode_Exit
   Else
       Resume Next
   End If
End Function
'Public Sub EmailSites()
'    Dim SQL As String
'    Dim DateSQL As String
'    Dim StartDate As String
'    Dim text As String
'    Dim snaSuppliers As Recordset
'    Dim ItemFound As Boolean
'    StartDate = readini("Email", "LastDate", Now - 2, App.Path & "\POSPrices.ini")
'    DateSQL = " between #" & FormatDate(StartDate, False) & "# AND #" & FormatDate(Now, False) & "#"
'    SQL = "SELECT Suppliers.SupplierName FROM Suppliers INNER JOIN Details ON Suppliers.SupplierID = Details.SupplierID" & _
'    " WHERE (Suppliers.SuppCreated " & DateSQL & ") OR (Details.LastChange " & DateSQL & ") OR (Details.DateCreated " & DateSQL & ") OR (Suppliers.LastUpdated " & DateSQL & ")" & _
'    " GROUP BY Suppliers.SupplierName"
'    Call OpenSupplierDB
'
'    Set snaSuppliers = dbSuppliers.OpenRecordSet(SQL$, dbOpenSnapshot)
'    If snaSuppliers.EOF Then
'
'    Else
'        ItemFound = True
'        Do Until snaSuppliers.EOF
'            frmSupplier.vsSupplier.AddItem FNulls(snaSuppliers("SupplierName"))
'            'text = text & vbCrLf & FNulls(snaSuppliers("SupplierName"))
'            snaSuppliers.MoveNext
'        Loop
'    End If
'    snaSuppliers.ClsRS
'    Set snaSuppliers = Nothing
'    Call CloseSupplierDB
'    If ItemFound Then
'    'If text <> "" Then
'        text$ = "The following suppliers have been updated." & vbCrLf & _
'        "From the LOTS Main Menu select Tools - LOTS Updates - Download Supplier Prices." & vbCrLf & _
'        "This will download the latest file and can be done while you are working on other machines." & vbCrLf & _
'        "If you have any problems or any other suppliers you want HealthSoft to obtain the files for" & vbCrLf & _
'        "Please email priceupdates@HealthSoft.co.nz or ask your rep to email for details."
'        OpenDB
'        'Change Nigel 25.2.2010
'        Call frmPreview.LOTSStartDocument("LOTS Supplier Update" & vbNewLine & text & vbNewLine, "", pprA4, orPortrait, "Reports", True)
'        frmPreview.Caption = GetLocalLOTSDir & "\LOTS Supplier Update.pdf"
'        Call frmPreview.LOTSPrintGrid("", frmSupplier.vsSupplier.hWnd)
'        Call frmPreview.LOTSFinishDocument
'        'Call ExportFlex(frmSupplier.vsSupplier, GetLocalLOTSDir, "Export.html", , , , , , , text$)
'        Unload frmSupplier
'        Set snaSuppliers = dbBigBro.OpenRecordSet("SELECT CustomerName,ShopEmail FROM Customer WHERE ShopEmail <> '' and not  isnull(ShopEmail) ", dbOpenSnapshot)
'        If snaSuppliers.EOF Then
'
'        Else
'            Do Until snaSuppliers.EOF
'                'Change Nigel 25.2.2010
'                '29/3/2010 Limin removed "\" in "\\LOTS Supplier Update.pdf"
'                Call SendEmailUsingSMTP("", "Supplier Update for " & FormatDate(Now, False), "PriceUpdates@HealthSoft.co.nz", FNulls(snaSuppliers("ShopEmail")), True, True, False, False, GetLocalLOTSDir & "\LOTS Supplier Update.pdf")
'                'Call SendEmailUsingSMTP("", "Supplier Update for " & FormatDate(Now, False), "PriceUpdates@HealthSoft.co.nz", FNulls(snaSuppliers("ShopEmail")), True, False, True, False, GetLocalLOTSDir & "\Export.html")
'                snaSuppliers.MoveNext
'            Loop
'        End If
'        snaSuppliers.ClsRS
'        Set snaSuppliers = Nothing
'        CloseDB
'    End If
'    MsgBox "LOTS Supplier Update PDF has been emailed..."
'
'    Call writeini("Email", "LastDate", Now, App.Path & "\POSPrices.ini")
'
'End Sub
Public Sub ImportPharmacodeProducts()
    Dim PLU As String
    Dim TradeName As String
    Dim hdlPLUFile As Integer
    Dim FilePath As String
    Dim FileName As String
    Dim data As String
    Dim dynSupplier As Recordset
    Dim SupplierID As Long
    Dim LineCount As Long
    Dim DirCount As Integer
    Dim FileCount As Integer
    Dim Database As String
    Dim DateCreated As Date
    Dim DateModified As Date
    Dim BarcodesFound As Long

On Error GoTo ImportPharmacodeProducts_Error
    Call OpenSupplierDB
    'Get the supplier we are saving this info under
    Set dynSupplier = dbSuppliers.OpenRecordSet("SELECT * FROM Suppliers WHERE SupplierName = 'OTC Barcodes'", dbOpenDynaset)
    If dynSupplier.EOF Then
        dynSupplier.AddNew
        dynSupplier("SupplierName") = "OTC Barcodes"
        dynSupplier("SupplierCode") = "OTC"
        dynSupplier("LastUpdated") = Now
        SupplierID = dynSupplier("SupplierID")
        dynSupplier.Update
    Else
        dynSupplier.Edit
        dynSupplier("LastUpdated") = Now
        dynSupplier.Update
        SupplierID = dynSupplier("SupplierID")
    End If
    dynSupplier.ClsRS
    Set dynSupplier = Nothing
    If MsgBox("Do you want to import a new pharmacode file.?", vbYesNo) = vbYes Then
        FilePath = GetOpenFile(PathName & "\Guild\", "*.txt", FileName)
        If FileExists(FilePath) Then
            
            'Now load all the Pharmacodes from the allbrands file
            hdlPLUFile = FreeFile
            Close #hdlPLUFile
            Open FilePath For Input As #hdlPLUFile
            Do Until EOF(hdlPLUFile)
                Line Input #hdlPLUFile, data
    '            LineCount = LineCount + 1
                If LineCount > 1 Then
                    PLU = LEFT(data, InStr(data, vbTab) - 1)
                    'Get rid of the plu
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    'Get rid of the generic code
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    TradeName = LEFT(data, InStr(data, vbTab) - 1)
                    'get rid of the first part of the tradename
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    'Get rid of the manufacturer
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    'Get rid of the ean code
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    TradeName = TradeName & " " & LEFT(data, InStr(data, vbTab) - 1)
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    TradeName = TradeName & " " & LEFT(data, InStr(data, vbTab) - 1)
                    data = Right(data, Len(data) - InStr(data, vbTab))
                    TradeName = TradeName & " " & LEFT(data, InStr(data, vbTab) - 1)
                    'Save any new pharmacodes to the supplier database
                    Set dynSupplier = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE PharmaCode = '" & PLU & "' AND  SupplierID = " & SupplierID, dbOpenDynaset)
                    If dynSupplier.EOF Then
                        dynSupplier.AddNew
                        dynSupplier("SupplierID") = SupplierID
                        dynSupplier("PharmaCode") = PLU
                        Call DBValidate(dynSupplier("Description"), UCase(TradeName))
                        dynSupplier("LastUpdated") = Now
                        dynSupplier.Update
                    End If
                    dynSupplier.ClsRS
                    Set dynSupplier = Nothing
                End If
            Loop
        End If
    End If
    
    
    Set dynSupplier = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID, dbOpenDynaset)
    If dynSupplier.EOF Then
    
    Else
        
        If FileExists(GetLocalLOTSDir & "\LOTSDB\Lots.mdb") Then
            If FileExists(GetLocalLOTSDir & "\LOTSDB\LotsA.mdb") Then
                Kill GetLocalLOTSDir & "\LOTSDB\LotsA.mdb"
            End If
            Name GetLocalLOTSDir & "\LOTSDB\Lots.mdb" As GetLocalLOTSDir & "\LOTSDB\LotsA.mdb"
        End If
        
        
        frmFiles.Hide
        frmFiles.BackDirs.Path = "\\hs-Server\Lots Databases\"
        frmFiles.BackFiles.Pattern = "*.zip"
 
        For DirCount = 0 To frmFiles.BackDirs.ListCount - 1
            frmFiles.BackFiles.Path = frmFiles.BackDirs.LIst(DirCount)
            For FileCount = 0 To frmFiles.BackFiles.ListCount - 1
                Call GetAllFileDates(frmFiles.BackFiles.Path & "\" & frmFiles.BackFiles.LIst(FileCount), DateCreated, DateModified)
                If DateCreated > DateAdd("m", -3, Now) Or DateModified > DateAdd("m", -3, Now) Then
                    If RestoreFileFromServer(frmFiles.BackFiles.LIst(FileCount), frmFiles.BackFiles.Path) Then
                         BarcodesFound = BarcodesFound + UpdateBarcodes(dynSupplier)
                        Exit For
                    End If
                End If
            Next FileCount
        Next DirCount
        FileCopy GetLocalLOTSDir & "\LOTSDB\LotsA.mdb", GetLocalLOTSDir & "\LOTSDB\Lots.mdb"
    End If
    dynSupplier.ClsRS
    Set dynSupplier = Nothing
ImportPharmacodeProducts_Exit:
   On Error Resume Next
   End
Exit Sub

ImportPharmacodeProducts_Error:
   If SYSLOG(Err, "ImportPharmacodeProducts in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume ImportPharmacodeProducts_Exit
   Else
       Resume Next
   End If
End Sub
Public Function RestoreFileFromServer(ByVal FileName As String, ByVal FilePath As String)
On Error GoTo RestoreFileFromServer_Error
    Dim ErrorMessage As String
    With frmLOTSZip.zipBackup
        .ExtractDirectory = GetLocalLOTSDir & "\LOTSDB"
        .ZipFileName = FilePath & "\" & FileName
        .UsePaths = False
        .FilesToProcess = "*.mdb"
        frmLOTSZip.Tag = "RESTORE"
        frmLOTSZip.cmdCancel.Visible = True
    
        DoEvents
        frmLOTSZip.Show 1
        DoEvents
        ErrorMessage = translateZipError(Val(frmLOTSZip.Tag))
        If ErrorMessage = "" Then
            RestoreFileFromServer = True
        Else
            
        End If
    End With
RestoreFileFromServer_Exit:
   On Error Resume Next
   Unload frmLOTSZip
Exit Function

RestoreFileFromServer_Error:
   If SYSLOG(Err, "RestoreFileFromServer in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume RestoreFileFromServer_Exit
   Else
       Resume Next
   End If
End Function
Public Function UpdateBarcodes(ByRef Pharmacodes As Recordset) As Long
    Dim snaBarcodes As Recordset
    Dim Progress As ProgressIndicator
    Dim UpdateBarcode As Boolean
    Dim Barcode As String
On Error GoTo UpdateBarcodes_Error
    Set dbLots = Workspaces(0).OpenDatabase(GetLocalLOTSDir & "\LOTSDB\Lots.MDB")
    With Pharmacodes
        .MoveLast
        .MoveFirst
        Set Progress = New ProgressIndicator
        Progress.RecordCount = Pharmacodes.RecordCount
        Progress.Show "Downloading DB"
        Do Until .EOF
            Progress.increment Pharmacodes.AbsolutePosition
            Set snaBarcodes = dbLots.OpenRecordSet("SELECT Barcode FROM Barcode INNER JOIN Stock ON Barcode.StockID = Stock.StockID WHERE PLU = '" & FNulls(Pharmacodes("Pharmacode")) & "' ORDER BY BarcodeID", dbOpenSnapshot)
            If snaBarcodes.EOF Then
            
            Else
                snaBarcodes.MoveLast
                If Val(FNulls(snaBarcodes("Barcode"))) <> Val(FNulls(Pharmacodes("Barcode"))) Then
                    UpdateBarcode = False
                     Barcode = FNulls(Pharmacodes("Barcode"))
                     'Remove any leading zeros
                     Do Until Val(LEFT(Barcode, 1)) <> 0
                        Barcode = Right(Barcode, Len(Barcode) - 1)
                     Loop
                     If Val(FNulls(Pharmacodes("Barcode"))) = "" Then
                        UpdateBarcode = True
                     End If
                     If Barcode = 13 Then
                        UpdateBarcode = True
                     End If
                     If UpdateBarcode = True Then
                        .Edit
                        Call SYSLOG(Err, "PDE: " & FNulls(Pharmacodes("Pharmacode")) & " Old Barcode: " & FNulls(Pharmacodes("Barcode")) & " NewBarcode: " & Barcode)
                        Pharmacodes("Barcode") = Barcode
                        .Update
                    End If
                End If
            End If
            snaBarcodes.ClsRS
            Set snaBarcodes = Nothing
            .MoveNext
        Loop
    End With
    dbLots.ClsDB
    Set dbLots = Nothing
    Set Progress = Nothing
UpdateBarcodes_Exit:
   On Error Resume Next
Exit Function

UpdateBarcodes_Error:
   If SYSLOG(Err, "UpdateBarcodes in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume UpdateBarcodes_Exit
   Else
       Resume Next
   End If
End Function
Sub LoadSuppliers(ByVal cboSupplier As ComboBox)
    
On Error GoTo ErrTrap
    Dim CloseDB As Boolean
    Dim snaSuppliers As Recordset
    
    Call OpenSupplierDB(CloseDB)
    cboSupplier.Clear
    Set snaSuppliers = dbSuppliers.OpenRecordSet("SELECT SupplierName,SupplierID,SettingsGuid FROM Suppliers Order by SupplierName", dbOpenSnapshot)
    Do Until snaSuppliers.EOF
        cboSupplier.AddItem FNulls(snaSuppliers("SupplierName"))
        cboSupplier.ItemData(cboSupplier.ListCount - 1) = Val(FNulls(snaSuppliers("SupplierID")))
        If FNulls(snaSuppliers("SettingsGuid")) = "" Then
            Call dbSuppliers.Execute("UPdate Suppliers SET SettingsGuid = '" & GetNewGuid & "' WHERE SupplierID = " & Val(FNulls(snaSuppliers("SupplierID"))))
            
        End If
        snaSuppliers.MoveNext
    Loop
    snaSuppliers.ClsRS
    Set snaSuppliers = Nothing
    Call CloseSupplierDB(CloseDB)
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "LoadSuppliers in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    End
End Sub

Sub GetSupplierCode()

On Error GoTo ErrTrap

    Dim snaDetails As Recordset
    Dim SQL As String
    
    frmUpdateSupplier.txtCost.text = ""
    
    Call OpenSupplierDB
    
    SQL$ = "SELECT * FROM Details"
    SQL$ = SQL$ & " WHERE SupplierID = " & frmUpdateSupplier.cboSupplier.ItemData(frmUpdateSupplier.cboSupplier.ListIndex)
    SQL$ = SQL$ & " AND Code = '" & frmUpdateSupplier.txtCode.text & "'"
    
    Set snaDetails = dbSuppliers.OpenRecordSet(SQL$, dbOpenSnapshot)
    
    If snaDetails.RecordCount Then
        frmUpdateSupplier.txtCost.text = centsToDollars(FNulls(snaDetails("Cost")))
        frmUpdateSupplier.txtBarcode.text = Val(FNulls(snaDetails("Barcode")))
        frmUpdateSupplier.txtDescription.text = Trim(FNulls(snaDetails("Description")))
        frmUpdateSupplier.txtOuters.text = Val(FNulls(snaDetails("Outers")))
        frmUpdateSupplier.txtRetail.text = centsToDollars(FNulls(snaDetails("Retail")))
        frmUpdateSupplier.txtCost.SetFocus
    Else
        Beep
        If MsgBox("'" & frmUpdateSupplier.txtCode.text & "' not found" & vbCrLf & vbCrLf & "Add New", vbInformation + vbYesNo + vbDefaultButton2, "Add New Item") = vbYes Then
            SQL$ = "INSERT INTO Details (SupplierID,Code) VALUES "
            SQL$ = SQL$ & "(" & frmUpdateSupplier.cboSupplier.ItemData(frmUpdateSupplier.cboSupplier.ListIndex)
            SQL$ = SQL$ & ",'" & frmUpdateSupplier.txtCode.text & "')"
            dbSuppliers.Execute SQL$, dbFailOnError
            frmUpdateSupplier.txtCost.text = ""
            frmUpdateSupplier.txtCost.SetFocus
        Else
            frmUpdateSupplier.txtCost.text = ""
            frmUpdateSupplier.txtCode.SelStart = 0
            frmUpdateSupplier.txtCode.SelLength = Len(frmUpdateSupplier.txtCode.text)
        End If
    End If
    
    snaDetails.ClsRS
    Set snaDetails = Nothing
    Call CloseSupplierDB
        
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "GetSupplierCode in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    End
End Sub

Sub SaveSupplierCost()
    
On Error GoTo ErrTrap
    
    Dim dynDetails As Recordset
    Dim SQL As String
    
    If IsNumeric(frmUpdateSupplier.txtCost.text) And dollarsToCents(frmUpdateSupplier.txtCost.text) >= 0 And dollarsToCents(frmUpdateSupplier.txtCost.text) < 99999 Then ' $999.99
        
        frmUpdateSupplier.txtCost.text = centsToDollars(dollarsToCents(frmUpdateSupplier.txtCost.text))
        
        Call OpenSupplierDB
        
        SQL$ = "SELECT * FROM Details"
        SQL$ = SQL$ & " WHERE SupplierID = " & frmUpdateSupplier.cboSupplier.ItemData(frmUpdateSupplier.cboSupplier.ListIndex)
        SQL$ = SQL$ & " AND Code = '" & frmUpdateSupplier.txtCode.text & "'"
        
        Set dynDetails = dbSuppliers.OpenRecordSet(SQL$, dbOpenDynaset)
        If dynDetails.RecordCount Then
            dynDetails.Edit
            dynDetails("Cost") = dollarsToCents(frmUpdateSupplier.txtCost.text)
            dynDetails("Retail") = dollarsToCents(frmUpdateSupplier.txtRetail.text)
            dynDetails("Barcode") = Trim(frmUpdateSupplier.txtBarcode.text)
            dynDetails("Outers") = Val(frmUpdateSupplier.txtOuters.text)
            dynDetails("Description") = Trim(frmUpdateSupplier.txtDescription.text)
            dynDetails("LastUpdated") = Now
            dynDetails.Update
        End If
        
        dbSuppliers.Execute "UPDATE Suppliers SET LastUpdated = Now WHERE SupplierID = " & frmUpdateSupplier.cboSupplier.ItemData(frmUpdateSupplier.cboSupplier.ListIndex), dbFailOnError
        
        Call CloseSupplierDB
        
        frmUpdateSupplier.txtRetail.text = ""
        frmUpdateSupplier.txtBarcode.text = ""
        frmUpdateSupplier.txtOuters.text = ""
        frmUpdateSupplier.txtDescription.text = ""
        
        frmUpdateSupplier.txtCode.SetFocus
    Else
        frmUpdateSupplier.txtCost.text = ""
        frmUpdateSupplier.txtCost.SetFocus
    End If
    
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "SaveSupplierCost in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    End
End Sub
Public Function HasNumericChar(ByVal strText As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : HasNumericChar
' DateTime  : 11/11/2010
' Author    : Limin
' Purpose   : Check any more than 4 numeric chars in strText '0: 48,  9: 57
'---------------------------------------------------------------------------------------
On Error GoTo HasNumericChar_Error
  Dim n As Long
  Dim Count As Long
  Debug.Print strText
  For n = 1 To Len(strText)
    If Asc(Mid(strText, n, 1)) < 58 And Asc(Mid(strText, n, 1)) > 47 Then
      Count = Count + 1
      If Count > 4 Then Exit For
    End If
  Next
  If Count > 4 Then HasNumericChar = True
HasNumericChar_Exit:
   On Error Resume Next
Exit Function

HasNumericChar_Error:
   If SYSLOG(Err, "HasNumericChar in CommonUtilities in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
     Resume HasNumericChar_Exit
   Else
     Resume Next
   End If

End Function
Public Sub ConvertStringToArray(ByVal InputString As String, ByRef OutputString() As String)
    'New Nigel 29.10.2009 - copied straight from DrugID code
    On Error GoTo ErrTrap
    Dim PLU As String
    Dim Description As String
    Dim ALTDESCRIPTION As String
    Dim Barcode As String
    Dim APNBARCODE As String
    Dim Manufacturer As String
    Dim Retail1 As Long
    Dim UNITOFRETAIL1 As Long
    Dim Retail2 As Long
    Dim UNITOFRETAIL2 As Long
    Dim CHANGEINDICATOR As String
    Dim Markup As Long
    Dim NORMALRETAIL As Long
    Dim WHOLESALETAX As Long
    Dim GSTFlag As String
    Dim Schedule As Long
    Dim GSTRETAIL As Long
    Dim ProductGroupID As Long
    Dim SUBPRODUCTGROUPID As Long
    Dim PackSize As Long
    Dim Retail3 As Long
    Dim UNITOFRETAIL3  As Long
    

    'Change Nigel 28.5.2009
    ReDim OutputString(INDEX_UNITOFRETAIL3)
    'ReDim OutputString(INDEX_UNITRETAIL)
    If Len(InputString) > 200 Then
        PLU = Mid(InputString, 1, 6)
        Description = Mid(InputString, 7, 30)
        ALTDESCRIPTION = Mid(InputString, 37, 30)
        Barcode = Mid(InputString, 67, 14)
        APNBARCODE = Mid(InputString, 83, 14)
        Manufacturer = Mid(InputString, 99, 3)
        'Retail1 = Mid(InputString, 124, 8)
        'UNITOFRETAIL1 = Mid(InputString, 132, 4)
        Retail2 = Mid(InputString, 136, 8)
        UNITOFRETAIL2 = Mid(InputString, 144, 4)
        '5/4/2004 Limin User cost3 and retail3
        Retail3 = Mid(InputString, 148, 8)
        UNITOFRETAIL3 = Mid(InputString, 156, 4)
        'use  retail3
        'Change Nigel 4.6.2009 - not sure we defauly Retail1 to Retail3?
        'Retail1 = Retail3
        'UNITOFRETAIL1 = UNITOFRETAIL3
        Retail1 = Mid(InputString, 124, 8)
        UNITOFRETAIL1 = Mid(InputString, 132, 4)

        CHANGEINDICATOR = Mid(InputString, 194, 1)
        Markup = Mid(InputString, 195, 5)
        NORMALRETAIL = Mid(InputString, 200, 6)
        WHOLESALETAX = Mid(InputString, 207, 5)
        GSTFlag = Mid(InputString, 212, 1)
        Schedule = Val(Mid(InputString, 213, 1))
        GSTRETAIL = Mid(InputString, 214, 6)
        ProductGroupID = Mid(InputString, 184, 2)
        SUBPRODUCTGROUPID = Mid(InputString, 186, 2)
        'If IsNumeric(Mid(InputString, 102, 8)) Then
        '    PackSize = Mid(InputString, 102, 8)
        'Else
        '    PackSize = 0
        'End If
    
        OutputString(INDEX_PLU) = PLU
        OutputString(INDEX_DESCRIPTION) = Description
        If PLU = "393728" Then
            Debug.Print
        End If
        
        
        'OutputString(INDEX_ALTDESCRIPTION) = FilterDescription(Description)
        
        OutputString(INDEX_BARCODE) = Barcode
        OutputString(INDEX_APNBARCODE) = APNBARCODE
        OutputString(INDEX_MANUFACTURER) = Manufacturer
        OutputString(INDEX_RETAIL1) = centsToDollars(Retail1)
        OutputString(INDEX_UNITOFRETAIL1) = UNITOFRETAIL1
        OutputString(INDEX_RETAIL2) = centsToDollars(Retail2)
        OutputString(INDEX_UNITOFRETAIL2) = UNITOFRETAIL2
        OutputString(INDEX_CHANGEINDICATOR) = CHANGEINDICATOR
        OutputString(INDEX_MARKUP) = Markup / 100
        OutputString(INDEX_NORMALRETAIL) = centsToDollars(NORMALRETAIL)
        OutputString(INDEX_WHOLESALETAX) = WHOLESALETAX / 100
        OutputString(INDEX_GSTFLAG) = GSTFlag
        OutputString(INDEX_SCHEDULE) = Schedule
        OutputString(INDEX_GSTRETAIL) = centsToDollars(GSTRETAIL)
        OutputString(INDEX_PRODUCTGROUPID) = ProductGroupID
        OutputString(INDEX_SUBPRODUCTGROUPID) = SUBPRODUCTGROUPID
        'New Nigel 4.6.2009
        OutputString(INDEX_RETAIL3) = centsToDollars(Retail3)
        OutputString(INDEX_UNITOFRETAIL3) = UNITOFRETAIL3
        
        Dim IsPackSize As Boolean
                    
        IsPackSize = False
        If GSTFlag = "F" Then
            If Abs(CLng(Retail1 + Retail1 * Markup / 10000) - NORMALRETAIL) > 4 Then
                IsPackSize = True
            End If
        ElseIf GSTFlag = "N" Then
            If Abs(CLng(Retail1 + Retail1 * Markup / 10000) - NORMALRETAIL) > 4 Then
                IsPackSize = True
            End If
        ElseIf GSTFlag = "Y" Then
            If Abs(CLng((Retail1 + Retail1 * Markup / 10000) * (1 + 0.1)) - NORMALRETAIL) > 4 Then
                IsPackSize = True
            End If
        
        End If
        
        Dim UnitCost As Long
        'New Nigel 28.5.2009
        Dim UnitCost2 As Long
        Dim UnitCost3 As Long
        
        If IsPackSize Then
            If IsNumeric(Mid(InputString, 110, 4)) Then
                PackSize = Mid(InputString, 110, 4)
                If PackSize > 0 Then
                    UnitCost = Retail1 / PackSize
                    'New Nigel 28.5.2009
                    UnitCost2 = Retail2 / PackSize
                    UnitCost3 = Retail3 / PackSize
                Else
                    UnitCost = Retail1
                    'New Nigel 28.5.2009
                    UnitCost2 = Retail2
                    UnitCost3 = Retail3
                End If
            ElseIf IsNumeric(Mid(InputString, 114, 4)) Then
                PackSize = Mid(InputString, 114, 4)
                If PackSize > 0 Then
                    UnitCost = Retail1 / PackSize
                    'New Nigel 28.5.2009
                    UnitCost2 = Retail2 / PackSize
                    UnitCost3 = Retail3 / PackSize
                End If
            Else
                PackSize = 0
            End If

            OutputString(INDEX_PACKSIZE) = PackSize
            OutputString(INDEX_UNITCOST) = centsToDollars(UnitCost)
            'New Nigel 28.5.2009
            OutputString(INDEX_UNITCOST2) = centsToDollars(UnitCost2)
            OutputString(INDEX_UNITCOST3) = centsToDollars(UnitCost3)
        Else
            OutputString(INDEX_PACKSIZE) = "1"
            OutputString(INDEX_UNITCOST) = OutputString(INDEX_RETAIL1)
            'New Nigel 28.5.2009
            OutputString(INDEX_UNITCOST2) = OutputString(INDEX_RETAIL2)
            OutputString(INDEX_UNITCOST3) = OutputString(INDEX_RETAIL3)
        End If
        
        If OutputString(INDEX_PACKSIZE) = "0" Then
            OutputString(INDEX_PACKSIZE) = "1"
        End If
        
        If NORMALRETAIL > 0 Then
            OutputString(INDEX_UNITRETAIL) = centsToDollars(NORMALRETAIL)
        Else
            If GSTFlag = "F" Then
                NORMALRETAIL = CLng(Retail1 + Retail1 * Markup / 10000)
            ElseIf GSTFlag = "N" Then
                NORMALRETAIL = CLng(Retail1 + Retail1 * Markup / 10000)
            ElseIf GSTFlag = "Y" Then
                NORMALRETAIL = CLng((Retail1 + Retail1 * Markup / 10000) * (1 + 0.1))
            End If
            OutputString(INDEX_UNITRETAIL) = centsToDollars(NORMALRETAIL)
        End If
        'Dim LineToAdd As String
        'Dim nTemp As Long
        
        'LineToAdd = OutputString(0)
        'For nTemp = 1 To INDEX_UNITRETAIL
        '   LineToAdd = LineToAdd & vbTab & OutputString(nTemp)
        'Next
       '
       ' With frmUpdateMenu.vsPrices
       '     .AddItem LineToAdd
       ' End With
    End If
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "ConvertStringToArray in " & app.ExeName)
    Resume Next
    
End Sub
Public Function GetGSTFlags(ByVal SuppID As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : GetGSTFlags
' DateTime  : 19/03/2010
' Author    : LiminOZ
' Purpose   : check what the GST Flags are can be YNF or "TAF"
'"F" stands for Free To End User.
'"A" is without GST.
'"T" is with GST.
 
'---------------------------------------------------------------------------------------
'
On Error GoTo GetGSTFlags_Error
  Dim CloseDB As Boolean
  Dim tblSupp As Recordset
  If SuppID = 0 Then Exit Function
If dbSuppliers Is Nothing Then
  Call OpenSupplierDB
  CloseDB = True
End If

  Set tblSupp = dbSuppliers.OpenRecordSet(SQL_Select & "Suppliers WHERE SupplierID = " & SuppID, dbOpenSnapshot)
  If tblSupp.EOF Then
  Else
    
    tblSupp.MoveFirst
    GetGSTFlags = Trim(FNulls(tblSupp("GSTFlags")))
  
  End If
  tblSupp.ClsRS
  Set tblSupp = Nothing
If CloseDB Then
  Call CloseSupplierDB
End If


GetGSTFlags_Exit:
   On Error Resume Next
Exit Function

GetGSTFlags_Error:
   If SYSLOG(Err, "GetGSTFlags in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
     Resume GetGSTFlags_Exit
   Else
     Resume Next
   End If
End Function

Public Function getSuppGST(ByVal GSTFlags As String, ByVal theGST As String) As String
    If GSTFlags = "" Or theGST = "" Then Exit Function
    Select Case InStr(GSTFlags, theGST$)
      Case 1
        getSuppGST = "Y"
      Case 2
        getSuppGST = "N"
      Case 3
        getSuppGST = "F"
    End Select
End Function
Public Function ImportFile(ByVal SettingsSupplierID As Long, ByVal FileName As String, Optional ByVal AutoImport As Boolean = False, Optional ByVal SupplierName As String) As Boolean
'Sub ImportFile(ByVal FileName As String, ByVal SupplierID As Long)
On Error GoTo ErrTrap

    Dim data() As String
    Dim RawData As String
    Dim hdlImport As Long
    Dim Failed As Boolean
    
    Dim Code As String
    Dim Cost As Long
    Dim Outers As Long
    Dim Barcode As String
    Dim Description As String
    Dim Retail As Long
    Dim Pharmacode As String
    Dim Prefix As String
    Dim Suffix As String
    Dim lngTemp As Long
'    Dim InvalidDescriptions As Long
'    Dim InvalidBarcodesFound As Long
    Dim proImport As New ProgressIndicator
    
    Dim FieldDelimiter As String
    Dim RecordDelimiter As String
    Dim rawdata2 As String
    Dim rawdataPos As Long
    
    Dim LastDescription2 As String
    Dim LastCost As Long
    Dim LastRetail As Long
'    Dim GST As Long
'    Dim CostInclGST As Long
    
    Dim Excel As Boolean
    Dim LineCount As Long
    Dim RowsToBeLoaded As Long
    Dim ExitLoop As Boolean
    Dim SheetCount As Long
    Dim MaxSheetCount As Integer
    Dim Header() As String
    Dim ColumnWidth As Integer
    Dim Column As Integer
    Dim ColumnsAcross As Integer
    Dim StartPOS As Integer
    Dim SkipSheet As Boolean
    Dim SheetLayout As Integer
    Dim SupplierID As Long
    Dim FindReplaceCount As Integer
    Dim Find As String
    Dim Replace As String
    Dim Temp As String
    Dim ProductGroup As String
    Dim GST As String
    Dim GSTValue As String
    Dim CostInclGST As Long
    Dim Supplier As String
    Dim ExistingItemFound As Boolean
    Dim SuppHeader As String '15/10/2010 Limin added
    Dim APICode As String
     Dim SigmaCode As String
     Dim FauldingCode As String
              Dim Item() As String
     Dim GSTFlags As String
    'New Nigel 13.2.2013
    'Dim NumberImported As Long
    'Dim NumberSkipped As Long
    Dim TimesRetried As Long
    Dim AddPossiblePharmacodes As Boolean
    Dim PharmacodeQuestionAsked As Boolean
    'Greg 24/3/2004 Check to see if it is an excel sheet
    Call SYSLOG(Nothing, "Opening SpreadSheet")
    If InStr(UCase(FileName), ".XLS") Then
        Call OpenSpreadsheet(FileName)
        'Change Greg 25/8/06 Make so that multipule worksheets can be loaded
        MaxSheetCount = WorkSheetCount
'        Call SetWorksheet(, 1)
'        LineCount = 1
        Excel = True
'        If MaxSheetCount > 1 Then
            Call frmExcelSheets.SetSheets(GetWorksheetNames, SettingsSupplierID)
'        End If
    Else
        'Change Greg 25/8/06 Make so that multipule worksheets can be loaded
        MaxSheetCount = 1
        hdlImport = FreeFile
        Open FileName For Input As hdlImport
        Excel = False
    End If
    ExitLoop = True
    'Change ends
    If AutoImport = False Then
        If frmImport.vsImport.Rows > 1 Then
            If MsgBox("Clear out all existing data?" & vbCrLf & "Yes if Reimporting the file or No if an additional file.", vbYesNo) = vbYes Then
                frmImport.vsImport.Rows = 1
                frmImport.vsImport.Redraw = False
                frmImport.InvalidBarcodesFound = 0
                frmImport.InvalidDescriptions = 0
            End If
        Else
            frmImport.vsImport.Rows = 1
            frmImport.vsImport.Redraw = False
            frmImport.InvalidBarcodesFound = 0
            frmImport.InvalidDescriptions = 0
        End If
        If MaxSheetCount > 1 Then
            frmExcelSheets.Show vbModal
        End If
    End If
    
    proImport.Show "Importing Price Update..."
    GSTFlags = GetGSTFlags(SettingsSupplierID)
    For SheetCount = 1 To MaxSheetCount
        SkipSheet = False
        SheetLayout = 0
        'Greg 24/3/2004 Get number of rows
        If Excel Then
'            If MaxSheetCount > 1 Then
                SkipSheet = Not frmExcelSheets.ImportSheet(SheetCount, SheetLayout)
'            End If
            If Not SkipSheet Then
                'Change Greg 25/8/06 Make so that multipule worksheets can be loaded
                Call SetWorksheet(, SheetCount)
                LineCount = 1
                RowsToBeLoaded = NumberOfRows
                proImport.RecordCount = RowsToBeLoaded
            End If
        Else
            proImport.RecordCount = LOF(hdlImport)
        End If
        If Not SkipSheet Then
            If frmImport.chkHasMultiColumns(SheetLayout).value = vbChecked Then
                ColumnsAcross = frmImport.txtNoOfColumns(SheetLayout).text
                ColumnWidth = frmImport.txtColumnsAcross(SheetLayout).text
            Else
                ColumnsAcross = 1
                ColumnWidth = 1
            End If
            ReDim Header(ColumnsAcross) As String
            With frmImport.cboSaveSupplier(SheetLayout)
                SupplierID = .ItemData(.ListIndex)
                Supplier = .text
                If frmImport.chkClear.value = vbChecked Then
                    frmImport.SuppliersProcessed = frmImport.SuppliersProcessed & "," & SupplierID
                End If
            End With
            With frmImport
                'Greg 24/3/2004 Data from excel is tab delimited
                If Excel Then
                    FieldDelimiter$ = vbTab
                Else
                    Select Case UCase(.cboFieldDelimiter.text)
                    Case "COMMA": FieldDelimiter$ = ","
                    Case "TAB": FieldDelimiter$ = vbTab
                    Case "CR": FieldDelimiter$ = vbCr
                    Case "CRLF": FieldDelimiter$ = vbCrLf
                    Case Else: FieldDelimiter$ = vbTab
                    End Select
                End If
            
                Select Case UCase(.cboRecordDelimiter.text)
                Case "LF"
                    RecordDelimiter = vbLf
                Case "CR", "CRLF"
                    RecordDelimiter = ""
                End Select
                rawdataPos& = 1
                '16/3/2004 Greg add in code to ignore first line
                If Excel Then
                    If .chkIgnoreFirst(SheetLayout).value = vbChecked And RowsToBeLoaded <> 0 Then
                         RawData = GetNextRow(1)
                         LineCount = 2
                    End If
                Else
                    If .chkIgnoreFirst(SheetLayout).value = vbChecked And EOF(hdlImport) = False Then
                        Line Input #hdlImport, RawData
                    End If
                End If
                '24/3/2004 Greg Add in code to check whether to exit the loop
                If Excel Then
                    If LineCount > RowsToBeLoaded Then
                        ExitLoop = True
                    Else
                        ExitLoop = False
                    End If
                Else
                    If EOF(hdlImport) And ((InStr(rawdataPos& + 1, RawData$, vbLf) = 0) Or RecordDelimiter = "") Then
                        ExitLoop = True
                    Else
                        ExitLoop = False
                    End If
                End If
                Do Until ExitLoop
                'Do Until EOF(hdlImport) And (InStr(rawdataPos& + 1, RawData$, vbLf) = 0 Or RecordDelimiter = "")
                    If Excel Then
TryAgain:
                        RawData = GetNextRow(LineCount)
                        If RawData = "" And glbClipboardError = True Then
                            RawData = GetNextRow(LineCount)
                            If RawData = "" And glbClipboardError = True Then
                                RawData = GetNextRow(LineCount)
                            End If
                        End If
                        'New Nigel 13.2.2013 - remove leading vbtab which stuffs up the line import (I've seen lines in Where On Earth file where the first char is a vbtab instead of their code)
                        If LEFT(RawData, 1) = vbTab And SettingsSupplierID <> 496 Then  'this is garbage - I don't know who where on earth are or I would just use their id
                            RawData = Mid(RawData, 2)
                        End If
                        LineCount = LineCount + 1
                    Else
                        If RecordDelimiter = "" Or InStr(rawdataPos + 1, RawData, RecordDelimiter) = 0 Then
                            Line Input #hdlImport, RawData
                        End If
                    End If
''                    'New Nigel 13.2.2013 - this seems to fix issue with some Excel rows not importing (I don't know why it fixes it, but it does)
''                    DoEvents
                    '( InStr(rawdata, FieldDelimiter$) > 0 And
                    If Not (InStr(UCase(RawData), "COST") > 0 And InStr(UCase(RawData), "DESCRIPTION") > 0) _
                     And Not (InStr(UCase(RawData), "PRODUCT") > 0 And InStr(UCase(RawData), "DESCRIPTION") > 0) Then
                        'I have no idea what this code is trying to do
                        'i think it is looking for a split line and getting that data
                        If InStr(RawData, vbLf) > 0 And Excel = False Then
                            If InStr(rawdataPos& + 1, RawData$, vbLf) - rawdataPos& - 1 > 0 Then
                                rawdata2$ = Mid(RawData$, rawdataPos& + 1, InStr(rawdataPos& + 1, RawData$, vbLf) - rawdataPos& - 1)
                                rawdataPos& = InStr(rawdataPos& + 1, RawData$, vbLf)
                            Else
                                rawdata2$ = RawData$
                            End If
                            
                            proImport.increment rawdataPos&
                        Else
                            rawdata2$ = RawData$
                            RecordDelimiter = ""
                            rawdataPos = 0
                            
                            If Excel Then
                                proImport.increment LineCount
                            Else
                                proImport.increment Loc(hdlImport) * 128
                            End If
                        End If
                    
                        'If UCase(frmImport.cboSupplier.Text) = "ALMAY" Then
                        '    Call ConvertStringToArray(rawdata2$, data$)
                        'Else
                        'Greg 24/3/2004 Check that something got read in
                        'changed so that loads data with no numeric values as this could be a header
                        If Len(rawdata2) > 0 Then
                        'If Len(rawdata2) > 0 and hasnumericdata(rawdata2) Then
                            If (UCase(SupplierName) = "GUARDIAN" Or UCase(SupplierName) = "AMCAL" Or FieldDelimiter$ = "FIXEDLENGTH") And InStr(rawdata2, FieldDelimiter) = 0 And glbDBType = OZBro Then
                                If LEFT$(rawdata2, 1) = " " Or LEFT$(rawdata2, 1) = Chr$(0) Or Len(rawdata2) < 10 Or InStr(1, RawData, "FULLAMCAL") Or InStr(1, RawData, "FULLGUARDIAN") Then
                                
                                Else
                           
                                    Call ConvertStringToArray(rawdata2, Item())
                                    If Item(0) <> "" Then
                                        Description$ = UCase(Trim(Item(INDEX_DESCRIPTION)))
                                        Pharmacode$ = Item(INDEX_PLU)
                                        Code$ = Item(INDEX_PLU)
                                        Cost& = dollarsToCents(Item(INDEX_UNITCOST))
                                        Retail& = dollarsToCents(Item(INDEX_RETAIL1))
                                        
                                        GST$ = Item(INDEX_GSTFLAG)
                                        
                                        If GSTFlags <> "" Then
                                          GST$ = getSuppGST(GSTFlags, GST)
                                        Else
                                          If GST$ = "Y" Then
                                              'This flag is already the same in Sigma file as in Prices.mdb
                                              GST$ = "Y"
                                          ElseIf GST$ = "N" Then
                                              GST$ = "F"
                                          ElseIf GST$ = "F" Then
                                              If frmImport.chkOwnTable(SheetLayout).value = vbChecked Then
                                                'don't change
                                              Else
                                                GST$ = "N"
                                              End If
                                          Else
                                              'Assume GST is chargable
                                              GST$ = "Y"
                                          End If
                                        End If
                                        Barcode = Trim(Item(INDEX_BARCODE))
                                        ProductGroup = Val(Item(INDEX_PRODUCTGROUPID))
                                        Outers = 1
                                        With frmImport.vsImport
                                            ImportFile = True
                                            .AddItem ""
                      
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SUPPLIER) = Supplier$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_CODE) = Code$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_PHARMACODE) = Pharmacode
                                            .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = FindandReplace(FindandReplace(Description$, vbCr, ""), vbLf, "")
                                            .TextMatrix(.Rows - 1, COL_IMPORT_COST) = centsToDollars(Cost&)
                                            If frmImport.chkRetail(0).value = vbChecked Then
                                                .TextMatrix(.Rows - 1, COL_IMPORT_RETAIL) = centsToDollars(Retail&)
                                            Else
                                                .TextMatrix(.Rows - 1, COL_IMPORT_RETAIL) = centsToDollars(0)
                                            End If
                                            .TextMatrix(.Rows - 1, COL_IMPORT_BARCODE) = Barcode$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_OUTERS) = Outers
                                            .TextMatrix(.Rows - 1, COL_IMPORT_PRODUCTGROUP) = ProductGroup
                                            .TextMatrix(.Rows - 1, COL_IMPORT_GST) = GST
                                            .TextMatrix(.Rows - 1, COL_IMPORT_COSTINCLGST) = centsToDollars(CostInclGST)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_LENGTH) = Len(Description)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_CHECK) = "N"
                                            If Len(TitleCase(Description$)) > 40 Then
                                                .Row = .Rows - 1
                                                .Col = COL_IMPORT_DESCRIPTION
                                                .CellBackColor = frmImport.shpInvalid.FillColor
                                                frmImport.InvalidDescriptions = frmImport.InvalidDescriptions + 1
                                                .TextMatrix(.Rows - 1, COL_IMPORT_INVALID) = "Y"
                                            End If
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SUPPLIERID) = SupplierID
                                            .TextMatrix(.Rows - 1, COL_IMPORT_LAYOUT) = SheetLayout
                                            'chaneg Greg 7/1/11 make so can save the api and faulding codes
                                            .TextMatrix(.Rows - 1, COL_IMPORT_API_CODE) = APICode
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SIGMA_CODE) = SigmaCode
                                            .TextMatrix(.Rows - 1, COL_IMPORT_FAULDING_CODE) = FauldingCode
                                        End With
                                    End If
                                End If
                            Else
                                Call ConvertDelimitedStringToArray(rawdata2$, FieldDelimiter$, data$)
                                'End If
                                For Column = 0 To ColumnsAcross - 1
                                    StartPOS = Column * ColumnWidth
                                    Code$ = ""
                                    Cost& = 0
                                    Retail& = 0
                                    Barcode$ = ""
                                    Description$ = ""
                                    Outers& = 1
                                    Pharmacode$ = ""
                '                    GST = 0
                                    Suffix = ""
                                    Prefix = ""
                                    ProductGroup = ""
                                    SuppHeader = "" '15/10/2010 Limin added
                                    APICode = ""
                                    SigmaCode = ""
                                    FauldingCode = ""
                                    GST = ""
                                    CostInclGST = 0
                                    'Item Code...
                                    If UBound(data$) >= (StartPOS + .cboCode(SheetLayout).ListIndex) And .chkCode(SheetLayout).value = vbChecked Then
                                        If Trim(data$(StartPOS + .cboCode(SheetLayout).ListIndex)) <> "" Then
                                            Code$ = Trim(data$(StartPOS + .cboCode(SheetLayout).ListIndex))
                                            'Change Greg 8/1/07 Do the find and replace after we have tried removeing from the description
            '                                Code$ = FindandReplace(Code$, """", "")
            '                                Code$ = FindandReplace(Code$, " ", "")
            '                                Code$ = Trim(Code$)
                                        
                                        End If
                                    End If
                                
                                    'Item PharmaCode...
                                    If UBound(data$) >= (StartPOS + .cboPharmacode(SheetLayout).ListIndex) And .chkPharmacode(SheetLayout).value = vbChecked Then
                                        Pharmacode$ = CleanValue(data$(StartPOS + .cboPharmacode(SheetLayout).ListIndex), True)
    '                                    Pharmacode$ = Trim(data$(StartPOS + .cboPharmacode(SheetLayout).ListIndex))
    '                                    Pharmacode$ = FindandReplace(Pharmacode$, """", "")
    '                                    'Change Greg 3/3/07 add removal of spaces from the number
    '                                    Pharmacode$ = FindandReplace(Pharmacode$, " ", "")
    '                                    Pharmacode$ = FindandReplace(Pharmacode$, Chr(160), "")
                                        If IsPharmacodeValid(Pharmacode) Then
                                            
                                        Else
                                            If Code = "" Then
                                                Code = Pharmacode
                                            End If
                                            Pharmacode$ = ""
                                        End If
    '                                    If IsPharmacodeValid(data$(StartPOS + .cboPharmacode(SheetLayout).ListIndex)) Then
    '
    '                                        Pharmacode$ = FindandReplace(Pharmacode$, """", "")
    '                                        'Change Greg 3/3/07 add removal of spaces from the number
    '                                        Pharmacode$ = FindandReplace(Pharmacode$, " ", "")
    '                                    End If
                                    End If
                                    'chaneg Greg 15/3/10 If pharmacode use as pharmacode not code
                                    'change gReg 14/6/10 Make so pharmacode is in both pharmacode and
                                    'code list
                                    If Code <> "" Then
                                        If Pharmacode = "" Then
                                            If glbDBType = OZBro Then
                                            
                                            Else
                                            'Change Nigel 22.5.2013 - if the Pharmacode is a number, keep it as both the code and Pharmacode, as people like PBL now send through non-valid Pharmacode codes, which can be used for ordering
                                                If IsPharmacodeValid(Code) Or IsNumeric(Code) > 0 Then
                                                    If PharmacodeQuestionAsked = False Then
                                                        If MsgBox("This file contains possible pharmacodes. Do you want to import these as pharmacodes?", vbYesNo + vbDefaultButton2) = vbYes Then
                                                            AddPossiblePharmacodes = True
                                                        End If
                                                        PharmacodeQuestionAsked = True
                                                    End If
                                                'If IsPharmacodeValid(Code) Then
                                                    If AddPossiblePharmacodes Then
                                                        Pharmacode = Code
                                                    End If
        '                                            Code = ""
                                                End If
                                            End If
                                        End If
                                    End If
                                    'New Nigel 22.5.2013
                                    If IsPharmacodeValid(Pharmacode) = True And Code = "" Then
                                        Code = Pharmacode
                                    End If
                                    
                                    'Description...
                                    If UBound(data$) >= (StartPOS + .cboDescription(SheetLayout).ListIndex) And .chkDescription(SheetLayout).value = vbChecked Then
    '                                    If Trim(data$(StartPOS + .cboDescription(SheetLayout).ListIndex)) <> "" Then
                                            Description$ = CleanValue(data$(StartPOS + .cboDescription(SheetLayout).ListIndex), False)
    '                                        Description$ = Trim(data$(StartPOS + .cboDescription(SheetLayout).ListIndex))
    '                                        Description$ = FindandReplace(Description$, """", "")
    '                                    End If
                                    End If
                                    If glbDBType = OZBro Then
                                        If UBound(data$) >= (StartPOS + .cboAPI(SheetLayout).ListIndex) And .chkAPI(SheetLayout).value = vbChecked Then
                                            APICode$ = CleanValue(data$(StartPOS + .cboAPI(SheetLayout).ListIndex), True)
                                        End If
                                        If UBound(data$) >= (StartPOS + .cboSigma(SheetLayout).ListIndex) And .chkSigma(SheetLayout).value = vbChecked Then
                                            SigmaCode$ = CleanValue(data$(StartPOS + .cboSigma(SheetLayout).ListIndex), True)
                                        End If
                                        If UBound(data$) >= (StartPOS + .cboFaulding(SheetLayout).ListIndex) And .chkFaulding(SheetLayout).value = vbChecked Then
                                            FauldingCode$ = CleanValue(data$(StartPOS + .cboFaulding(SheetLayout).ListIndex), True)
                                        End If
                                        If UBound(data$) >= .cboGST(SheetLayout).ListIndex And .chkGST(SheetLayout).value = vbChecked Then
                                          If GSTFlags <> "" And Trim(data$(.cboGST(SheetLayout).ListIndex)) <> "" Then
                                              GST$ = getSuppGST(GSTFlags, Trim(data$(.cboGST(SheetLayout).ListIndex)))
                                          Else
                                              If IsNumeric(data$(.cboGST(SheetLayout).ListIndex)) Then
                                                  GST = (data$(.cboGST(SheetLayout).ListIndex))
                                                  GST = FindandReplace(CStr(GST), """", "")
                                              'Change Greg Add in GST handling for flags
                                              ElseIf UCase(Trim(data$(.cboGST(SheetLayout).ListIndex))) = "N" Or UCase(Trim(data$(.cboGST(SheetLayout).ListIndex))) = "NO" Then
                                                  GST = "N"
                                                  GSTValue = "N"
                                              ElseIf UCase(Trim(data$(.cboGST(SheetLayout).ListIndex))) = "Y" Or UCase(Trim(data$(.cboGST(SheetLayout).ListIndex))) = "YES" Then
                                                  GST = "Y"
                                                  GSTValue = "Y"
                                              ElseIf Trim(data$(.cboGST(SheetLayout).ListIndex)) = "" Or Trim(data$(.cboGST(SheetLayout).ListIndex)) = " " Then
                                                  If GSTValue = "Y" Then
                                                      GST = "N"
                                                  ElseIf GSTValue = "N" Then
                                                      GST = "Y"
                                                  End If
                                                '10/3/2010 Limin Nadglen
                                              ElseIf UCase(Trim(data$(.cboGST(SheetLayout).ListIndex))) = "F" Then
                                                  GST = "F"
                                              End If
                                            End If
                                        End If
                                        Temp = 0
                                        'Cost incl GST
                                        If UBound(data$) >= .cboCostInclGST(SheetLayout).ListIndex And .chkCostInclGST(SheetLayout).value = vbChecked Then
                                            '10/3/2010 Limin added "."
                                            If InStr(data$(.cboCostInclGST(SheetLayout).ListIndex), "$") > 0 Or InStr(data$(.cboCostInclGST(SheetLayout).ListIndex), ".") Then
                                                lngTemp = dollarsToCents(data$(.cboCostInclGST(SheetLayout).ListIndex))
                                            '12/3/2010 Limin converted to dollars
                                            ElseIf InStr(data$(.cboCostInclGST(SheetLayout).ListIndex), ".") = 0 Then
                                              Retail& = dollarsToCents("$" & data$(.cboCostInclGST(SheetLayout).ListIndex))
                                            Else
                                                lngTemp& = Val(data$(.cboCostInclGST(SheetLayout).ListIndex))
                                            End If
                                            If IsNumeric(lngTemp&) Then
                                                'Change Greg 18/3/2004 If divide by outer and we have outer value
                                                'then calculate the new cost
                                                If .chkDivideCost(SheetLayout) = vbChecked And Outers& > 1 Then
                                                    CostInclGST& = Val(lngTemp&) / Outers&
                                                Else
                                                    CostInclGST& = Val(lngTemp&)
                                                End If
                                                
                                                CostInclGST& = FindandReplace(CStr(CostInclGST&), """", "")
                                                '3/5/2005 Limin Need to take GST off from CostInclGST&
                                                If GST = "Y" And frmImport.chkOwnTable(SheetLayout).value = vbUnchecked Then
                                                  Cost& = CostInclGST& * 1 / (1 + 10 / 100)
                                                End If
                                                'Change ends
                                            End If
                                        End If
                                    Else
                                    
                                    End If
                                    'Change Greg 31/10/07 Add in product group
                                    If UBound(data$) >= (StartPOS + .cboProductGroup(SheetLayout).ListIndex) And .chkProductGroup(SheetLayout).value = vbChecked Then
                                        ProductGroup$ = CleanValue(data$(StartPOS + .cboProductGroup(SheetLayout).ListIndex), False)
    '                                    If Trim(data$(StartPOS + .cboProductGroup(SheetLayout).ListIndex)) <> "" Then
    '                                        ProductGroup$ = Trim(data$(StartPOS + .cboProductGroup(SheetLayout).ListIndex))
    '                                        ProductGroup$ = FindandReplace(ProductGroup$, """", "")
    '                                    End If
                                    End If
                                    
                                    'Item Cost...
                                    If UBound(data$) >= (StartPOS + .cboCost(SheetLayout).ListIndex) And .chkCost(SheetLayout).value = vbChecked Then
                                        data$(StartPOS + .cboCost(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboCost(SheetLayout).ListIndex), "*", "")
                                        'Change Greg 4/8/08 Remove any nzd notation
                                        data$(StartPOS + .cboCost(SheetLayout).ListIndex) = FindandReplace(UCase(data$(StartPOS + .cboCost(SheetLayout).ListIndex)), "NZD", "")
                                        If IsNumeric(data$(StartPOS + .cboCost(SheetLayout).ListIndex)) Then
                                            'Change Greg 16/3/2004 Check if decimal point and if not add one in
                                            If InStr(data$(StartPOS + .cboCost(SheetLayout).ListIndex), ".") Then
                                                'do nothing
                                            Else
                                                data$(StartPOS + .cboCost(SheetLayout).ListIndex) = data$(StartPOS + .cboCost(SheetLayout).ListIndex) & ".00"
                                            End If
                                            'Change ends
                                            Cost& = dollarsToCents(data$(StartPOS + .cboCost(SheetLayout).ListIndex))
                                        End If
                                    End If
                                
                                    'Item Retail...
                                    If UBound(data$) >= (StartPOS + .cboRetail(SheetLayout).ListIndex) And .chkRetail(SheetLayout).value = vbChecked Then
                                        'Change Greg 4/8/08 Remove any nzd notation
                                        data$(StartPOS + .cboRetail(SheetLayout).ListIndex) = FindandReplace(UCase(data$(StartPOS + .cboRetail(SheetLayout).ListIndex)), "NZD", "")
                                        If IsNumeric(data$(StartPOS + .cboRetail(SheetLayout).ListIndex)) Then
                                            If InStr(data$(StartPOS + .cboRetail(SheetLayout).ListIndex), ".") > 0 Or InStr(data$(StartPOS + .cboRetail(SheetLayout).ListIndex), "$") > 0 Then
                                                Retail& = dollarsToCents(data$(StartPOS + .cboRetail(SheetLayout).ListIndex))
                                            Else
                                                Retail& = dollarsToCents(data$(StartPOS + .cboRetail(SheetLayout).ListIndex) & ".00")
                                            End If
                                        End If
                                    End If
                                    'Change greg 17/8/06 add in change to the way suffix and prefix works
                                    If UBound(data$) >= StartPOS + .cboDescription2(SheetLayout).ListIndex And .chkDescription2(SheetLayout).value = vbChecked Then
                                        'Change Greg 21/8/06 Description can't be blank
                                        If Trim(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex)) <> "" And Trim(Description) <> "" Then
                '                        If Trim(Data$(StartPOS + .cboDescription2(SheetLayout).listindex)) <> "" Then
                                            'Change Greg 5/6/08 Make so that for naturo pharm we can handle two suffixes
                                            If .chkSuffixAsPrefix(SheetLayout).value = vbChecked Then
                                                If UCase(Right(Trim(Description), Len(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex)))) <> UCase(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex)) Then
                                                    Description = Trim(Description & " " & data$(StartPOS + .cboDescription2(SheetLayout).ListIndex))
                                                End If
                                            Else
                                                If UCase(LEFT(Trim(Description), Len(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex)))) <> UCase(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex)) Then
                                                    Description = Trim(data$(StartPOS + .cboDescription2(SheetLayout).ListIndex) & " " & Description)
                                                End If
                                            End If
                                        End If
                                    End If
                                    If UBound(data$) >= StartPOS + .cboSuffix(SheetLayout).ListIndex And .chkSuffix(SheetLayout).value = vbChecked Then
                                        'Change Greg 21/8/06 Description can't be blank
                                        If Trim(data$(StartPOS + .cboSuffix(SheetLayout).ListIndex)) <> "" And Trim(Description) <> "" Then
                '                        If Trim(Data$(StartPOS + .cboSuffix(SheetLayout).listindex)) <> "" Then
                                            If UCase(LEFT(Trim(Description), Len(data$(StartPOS + .cboSuffix(SheetLayout).ListIndex)))) <> UCase(data$(StartPOS + .cboSuffix(SheetLayout).ListIndex)) Then
                                                Description = Trim(Description & " " & data$(StartPOS + .cboSuffix(SheetLayout).ListIndex))
                                            End If
                                        End If
                                    End If
                                    If Header(Column) <> "" And Description <> "" And .chkHeader(SheetLayout).value = vbChecked Then
    '                                If Header <> "" And Description <> "" And .chkAddDescription(SheetLayout).value = vbChecked Then
                                        'Change Greg 11/4/07 Add in header where possible
                                        If UCase(LEFT(Trim(Description$), Len(Header(Column)))) <> UCase(Header(Column)) Then
                                            'Change Greg 24/4/07 add the header to the end of the description if ticked
                                            If .chkHeaderToEnd(SheetLayout).value = vbChecked Then
                                                Description = Description & " " & Header(Column)
                                            Else
                                                Description = Header(Column) & " " & Description
                                            End If
                                        End If
                                    End If
                                    If Description <> "" And .txtDescHeader(SheetLayout).text <> "" And .chkAddDescription(SheetLayout).value = vbChecked Then
                                        If UCase(LEFT(Trim(Description$), Len(.txtDescHeader(SheetLayout).text))) <> UCase(.txtDescHeader(SheetLayout).text) Then
                                            Description = UCase(.txtDescHeader(SheetLayout).text) & " " & Description
                                            SuppHeader = TitleCase(.txtDescHeader(SheetLayout).text) & " " '15/10/2010 Limin added
                                        End If
                                    End If
                                    'Change Greg 7/11/06 make so supplier code can be removed automatically from list
                                    If Code <> "" And Description <> "" And .chkRemoveCode(SheetLayout).value = vbChecked Then
                                        'Change Greg 18/12/06 Remove any extra spaces from the description
                                        Description = Trim(FindandReplace(Description, Code, ""))
            '                            Description = FindandReplace(Description, Code, "")
                                    End If
                                    'Change Greg now take the code and remove unwanted info
                                    If Code <> "" Then
                                        'Do the find and replace after we have tried removeing from the description
                                        Code$ = FindandReplace(Code$, """", "")
                                        Code$ = FindandReplace(Code$, " ", "")
                                        Code$ = Trim(Code$)
                                    End If
                                    'Now try removing the code again just in case we missed it
                                    If Code <> "" And Description <> "" And .chkRemoveCode(SheetLayout).value = vbChecked Then
                                        'Change Greg 18/12/06 Remove any extra spaces from the description
                                        Description = Trim(FindandReplace(Description, Code, ""))
            '                            Description = FindandReplace(Description, Code, "")
                                    End If
                                    'Change Greg 24/4/07 Add some letters from the code to the description
                                    If Code <> "" And Description <> "" And .chkAddCodeToDescription(SheetLayout).value = vbChecked Then
                                        Temp = ""
                                        If IsNumeric(.txtLetters(SheetLayout).text) Then
                                            Temp = LEFT(Code, .txtLetters(SheetLayout).text)
                                        Else
                                            Do Until IsNumeric(Mid(Code, Len(Code) - Len(Temp), 1))
                                                Temp = Mid(Code, Len(Code) - Len(Temp), 1) & Temp
                                                If Len(Code) = Len(Temp) Then
                                                    Exit Do
                                                End If
                                            Loop
                                        End If
                                        Description = Trim(Description & " " & Temp)
                                    End If
                                    'Description 2...
                '                    If UBound(Data$) >= .cboDescription2(SheetLayout).listindex And .chkDescription2(SheetLayout).value = vbChecked Then
                '                        If Trim(LastDescription2) <> "" And Trim(Description$) <> "" Then
                '                            If frmImport.chkDescription2(SheetLayout).value = vbChecked Then
                '                                If frmImport.cboDescription2(SheetLayout).listindex >= 0 Then
                '                                    If InStr(UCase(Description$), UCase(LastDescription2)) = 0 Then
                '                                        Description$ = LastDescription2 & " " & Description$
                '                                        If Cost& = 0 And Retail& = 0 Then
                '                                            Cost& = LastCost&
                '                                            Retail& = LastRetail&
                '                                        End If
                '                                    End If
                '                                End If
                '                            End If
                '                        End If
                '                    End If
                                
                                    'Barcode...
                                    If UBound(data$) >= StartPOS + .cboBarcode(SheetLayout).ListIndex And .chkBarcode(SheetLayout).value = vbChecked Then
                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(CleanValue(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), True), "-", "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), " ", "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), """", "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), vbTab, "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), vbLf, "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), vbCr, "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), "-", "")
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = Trim(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex))
'                                        data$(StartPOS + .cboBarcode(SheetLayout).ListIndex) = FindandReplace(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), Chr(160), "")
                                        If IsNumeric(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex)) Then
                                            Barcode$ = data$(StartPOS + .cboBarcode(SheetLayout).ListIndex)
'                                            Barcode$ = FindandReplace(Barcode$, """", "")
'                                            Barcode$ = FindandReplace(Barcode$, " ", "")
'
'                                            Barcode$ = Trim(Barcode$)
    '                                    ElseIf InStr(data$(StartPOS + .cboBarcode(SheetLayout).ListIndex), "E") > 0 Then
    '                                        'Change Greg 8/1/07 Add in count for invalid barcodes to form so we can update
    '                                        frmImport.InvalidBarcodesFound = frmImport.InvalidBarcodesFound + 1
    '        '                                InvalidBarcodesFound = InvalidBarcodesFound + 1
                                        End If
                                        If InStr(Barcode, "E") > 0 Then
                                            frmImport.InvalidBarcodesFound = frmImport.InvalidBarcodesFound + 1
                                            Barcode = ""
                                        End If
                                    End If
                                    
                                    'Outers...
                                    If UBound(data$) >= StartPOS + .cboOuters(SheetLayout).ListIndex And .chkOuters(SheetLayout).value = vbChecked Then
                                        If IsNumeric(data$(StartPOS + .cboOuters(SheetLayout).ListIndex)) Then
                                            Outers& = Val(data$(StartPOS + .cboOuters(SheetLayout).ListIndex))
                                            'Change Greg 23/8/2005 Changed as value so no strings included
                '                            Outers& = FindandReplace(Outers&, """", "")
                                        End If
                                        If Outers& < 1 Then
                                            Outers& = 1
                                        End If
                                    End If
                                
                                
                                    'GST
                '                    If UBound(data$) >= .cboGST(SheetLayout).listindex And .chkGST(SheetLayout).value = vbChecked Then
                '                        If IsNumeric(data$(.cboGST(SheetLayout).listindex)) Then
                '                            GST& = Val(data$(.cboGST(SheetLayout).listindex))
                '                            GST& = FindandReplace(GST&, """", "")
                '                        'Change Greg Add in GST handling for flags
                '                        ElseIf UCase(Trim(data$(.cboGST(SheetLayout).listindex))) = "N" Or UCase(Trim(data$(.cboGST(SheetLayout).listindex))) = "NO" Then
                '                            GST& = 0
                '                        ElseIf UCase(Trim(data$(.cboGST(SheetLayout).listindex))) = "Y" Or UCase(Trim(data$(.cboGST(SheetLayout).listindex))) = "YES" Then
                '                            GST& = 10
                '                        End If
                '                    End If
                                
                                
                                    'Cost incl GST
                                    'Greg Change 30/3/2004 Not required in NZ
                '                    If UBound(data$) >= .cboCostInclGST(SheetLayout).listindex And .chkCostInclGST(SheetLayout).value = vbChecked Then
                '                        If IsNumeric(data$(.cboCostInclGST(SheetLayout).listindex)) Then
                '                            'Change Greg 18/3/2004 If divide by outer and we have outer value
                '                            'then calculate the new cost
                '                            If .chkDivideCost = vbChecked And Outers& > 1 Then
                '                                CostInclGST& = Val(data$(.cboCostInclGST(SheetLayout).listindex)) / Outers&
                '                            Else
                '                                CostInclGST& = Val(data$(.cboCostInclGST(SheetLayout).listindex))
                '                            End If
                '                            CostInclGST& = FindandReplace(CostInclGST&, """", "")
                '                            'Change ends
                '                        End If
                '                    End If
                                    'If PharmaCode = "2073358" Then
                                    '    Debug.Print
                                    'End If
                                    
                                    'Now change in items in the find replace list
                                    Description = UCase(Description)
                                    With frmImport.vsFindReplace
                                        For FindReplaceCount = 1 To .Rows - 1
                                            Find = UCase(Trim(.TextMatrix(FindReplaceCount, 0)))
                                            Replace = UCase(Trim(.TextMatrix(FindReplaceCount, 1)))
                                            If Find <> "" Then
                                                Description = FindandReplace(Description, Find, Replace)
                                            End If
                                        Next FindReplaceCount
                                    End With
                                    'Change Greg 18/12/06 Allow file to be saved if no code
                                    If ((Val(Cost) > 0 Or (frmImport.chkImportZeroCost.value = vbChecked And (Trim(Code$) <> "" Or Trim(Pharmacode$) <> ""))) And Description$ <> "") Then
                                        'NumberImported = NumberImported + 1
                                        TimesRetried = 0
                                        
            '                        If (Val(Cost) > 0 Or Description$ <> "") Then
            '                        If (Trim(Code$) <> "" Or Trim(Pharmacode$) <> "") And (Val(Cost) > 0 Or Description$ <> "") Then
                                        With frmImport.vsImport
                                            ImportFile = True
                                            .AddItem ""
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SUPPLIER) = Supplier$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_CODE) = Code$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_PHARMACODE) = Pharmacode
                                            '8/7/2004 Limin prefix
                                            .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = TitleCase(Description$)
                                            '15/10/2010 Limin added to conver uppercase of Header
                                            If SuppHeader <> "" Then
                                                If LEFT(TitleCase(Description$), Len(SuppHeader)) = TitleCase(SuppHeader) Then
                                                    .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = UCase(SuppHeader) & Right(TitleCase(Description$), Len(Description$) - Len(SuppHeader))
                                                End If
                                            End If
                                            
                                            
                                            '.TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = Trim(frmImport.txtDescHeader.Text) & " " & TitleCase(Description$)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_COST) = centsToDollars(Cost&)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_RETAIL) = centsToDollars(Retail&)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_BARCODE) = Barcode$
                                            .TextMatrix(.Rows - 1, COL_IMPORT_OUTERS) = Outers
                                            .TextMatrix(.Rows - 1, COL_IMPORT_PRODUCTGROUP) = ProductGroup
                                            'Change Greg 23/3/2004 Add in GST and Cost INcl GST
                                            'Greg Change 30/3/2004 Not required in NZ
                                            .TextMatrix(.Rows - 1, COL_IMPORT_GST) = GST
                                            .TextMatrix(.Rows - 1, COL_IMPORT_COSTINCLGST) = centsToDollars(CostInclGST)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_LENGTH) = Len(Description)
                                            .TextMatrix(.Rows - 1, COL_IMPORT_CHECK) = "N"
                                            If Len(TitleCase(Description$)) > 40 Then
                                                .Row = .Rows - 1
                                                .Col = COL_IMPORT_DESCRIPTION
                                                .CellBackColor = frmImport.shpInvalid.FillColor
                                                'Change Greg 8/1/07 Add in count for invalid descriptions to form so we can update
                                                frmImport.InvalidDescriptions = frmImport.InvalidDescriptions + 1
            '                                    InvalidDescriptions = InvalidDescriptions + 1
                                                'Change Greg 26/6/06 Add in option to sort invalid details
                                                .TextMatrix(.Rows - 1, COL_IMPORT_INVALID) = "Y"
                                            End If
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SUPPLIERID) = SupplierID
                                            .TextMatrix(.Rows - 1, COL_IMPORT_LAYOUT) = SheetLayout
                                            .TextMatrix(.Rows - 1, COL_IMPORT_API_CODE) = APICode
                                            .TextMatrix(.Rows - 1, COL_IMPORT_SIGMA_CODE) = SigmaCode
                                            .TextMatrix(.Rows - 1, COL_IMPORT_FAULDING_CODE) = FauldingCode
                                        End With
                                    'Change Greg 17/8/06 This no longer works as requested so change
                '                    ElseIf frmImport.chkDescription2(SheetLayout).value = vbChecked And Trim(Code$) = "" Then
                '                        LastDescription2$ = ""
                '                        LastCost& = 0
                '                        LastRetail& = 0
                '                        If frmImport.cboDescription2(SheetLayout).listindex >= 0 Then
                '                            LastDescription2$ = Data$(frmImport.cboDescription2(SheetLayout).listindex)
                '
                '                            If frmImport.chkCost(SheetLayout).value = vbChecked Then
                '                                LastCost& = dollarsToCents(Data$(.cboCost(SheetLayout).listindex))
                '                            End If
                '
                '                            If frmImport.chkRetail(SheetLayout).value = vbChecked Then
                '                                LastRetail& = dollarsToCents(Data$(.cboRetail(SheetLayout).listindex))
                '                            End If
                '                        End If
                                    Else
                                        'NumberSkipped = NumberSkipped + 1
                                        If UBound(data$) >= StartPOS + .cboHeader(SheetLayout).ListIndex And .chkHeader(SheetLayout).value = vbChecked Then
                                            If Trim(data(StartPOS + .cboHeader(SheetLayout).ListIndex)) <> "" Then
                                                'Change Greg 17/7/08 Make so if none in header code clears ot out
                                                If Trim(UCase(data$(StartPOS + .cboHeader(SheetLayout).ListIndex))) = "NONE" Then
                                                    Header(Column) = ""
                                                Else
                                                    Header(Column) = data$(StartPOS + .cboHeader(SheetLayout).ListIndex)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Column
                            End If
                        Else
                            'MsgBox "LineSkipped. LineCount: " & LineCount & vbNewLine & "RawData: " & RawData & vbNewLine & "Rawdata2: " & rawdata2
                            TimesRetried = TimesRetried + 1
                            If TimesRetried < 3 Then
                                GoTo TryAgain
                            End If
                            'NumberSkipped = NumberSkipped + 1
                        End If
                        
                    End If
                    '24/3/2004 Greg Add in code to check whether to exit the loop
                    If Excel Then
                        If LineCount > RowsToBeLoaded Then
                            ExitLoop = True
                        Else
                            ExitLoop = False
                        End If
                    Else
                        If EOF(hdlImport) And ((InStr(rawdataPos& + 1, RawData$, vbLf) = 0) Or RecordDelimiter = "") Then
                            ExitLoop = True
                        Else
                            ExitLoop = False
                        End If
                    End If
                Loop
            
            End With
        End If
    Next SheetCount
    
    'MsgBox "NumberImported: " & NumberImported & vbNewLine & "NumberSkipped: " & NumberSkipped & vbNewLine & "LineCount: " & LineCount
    
    Unload frmExcelSheets
    proImport.Hide
    Set proImport = Nothing
    If Excel Then
        CloseSpreadsheet
    Else
        Close hdlImport
    End If
    frmImport.vsImport.Redraw = True
    'Change Greg 8/1/07 save invalid details in main form so that we can update message
    frmImport.lblInfo.Caption = "Total Items: " & frmImport.vsImport.Rows - 1 & vbCrLf
    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Barcodes: " & frmImport.InvalidBarcodesFound & vbCrLf
    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Descriptions: " & frmImport.InvalidDescriptions
'    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Barcodes: " & InvalidBarcodesFound & vbCrLf
'    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Descriptions: " & InvalidDescriptions
    
Exit Function
ErrTrap:
    Call SYSLOG(Err, "ImportFile in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    Resume Next
End Function
Sub UpdateFlex()
'Sub UpdateFlex(ByVal SupplierID As Long)

On Error GoTo ErrTrap

    Dim proImport As ProgressIndicator
    Dim snaDetails As Recordset
    Dim n As Long
    Dim SupplierID As Long
    Dim Layout As Long
    Dim ExistingItemFound As Boolean
    Call OpenSupplierDB
    
'    Set snaDetails = dbSuppliers.OpenRecordset("SELECT * FROM Details WHERE SupplierID = " & SupplierID&, dbOpenSnapshot)
'    If snaDetails.RecordCount Then
    
        Set proImport = New ProgressIndicator
        proImport.Show "Comparing Import..."
        proImport.RecordCount = frmImport.vsImport.Rows - 1
    
        For n& = 1 To frmImport.vsImport.Rows - 1
        
            proImport.increment n&
            SupplierID = Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SUPPLIERID))
            Layout = Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_LAYOUT))
            If SupplierID = 0 Then
                 frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SUPPLIERID) = frmImport.cboSupplier.ItemData(frmImport.cboSupplier.ListIndex)
            End If
            If SupplierID > 0 Then
                If IsPharmacodeValid(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) Then
                    Set snaDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Pharmacode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE) & "'", dbOpenSnapshot)
                    If snaDetails.EOF Then
                        snaDetails.ClsRS
                        Set snaDetails = Nothing
                        Set snaDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'", dbOpenSnapshot)
                    End If
'                    snaDetails.FindFirst "Pharmacode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE) & "'"
                Else
                    Set snaDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'", dbOpenSnapshot)
'                    snaDetails.FindFirst "Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'"
                End If
                
                'Change Greg 16/3/2004 If not found and we have loaded barcodes check again using barcode
                If snaDetails.EOF Then
'                If snaDetails.NoMatch Then
                    If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) <> "" Then
                        snaDetails.ClsRS
                        Set snaDetails = Nothing
    '            If snaDetails.NoMatch And frmImport.chkBarcode(SheetLayout).Value = vbChecked Then
                        Set snaDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'", dbOpenSnapshot)
'                        snaDetails.FindFirst "Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'"
                    End If
                End If
                'Change ends
                If snaDetails.EOF Then
'                If snaDetails.NoMatch Then
                    frmImport.vsImport.Col = COL_IMPORT_CODE
                    frmImport.vsImport.Row = n&
                    frmImport.vsImport.CellBackColor = frmImport.shpNewCard.FillColor
                    'Change Greg 26/6/06 Add in option to sort new cards
                    frmImport.vsImport.TextMatrix(n, COL_IMPORT_NEW) = "Y"
                Else
                    ExistingItemFound = True
                    frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLDCOST) = centsToDollars(FNulls(snaDetails("Cost")))
                    frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION) = Trim(FNulls(snaDetails("Description")))
                    'Description...
                    If frmImport.chkDescription(Layout).value = vbUnchecked Then
                        frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) = Trim(FNulls(snaDetails("Description")))
                    ElseIf Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)) <> Trim(FNulls(snaDetails("Description"))) Then
                        frmImport.vsImport.Col = COL_IMPORT_DESCRIPTION
                        frmImport.vsImport.Row = n&
                        'If the description is incorrect and we have a valid description then use it
                        If frmImport.vsImport.TextMatrix(n, COL_IMPORT_INVALID) = "Y" Then
'                            frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) = Trim(FNulls(snaDetails("Description")))
                            frmImport.vsImport.TextMatrix(n, COL_IMPORT_INVALID) = "N"
                            frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHECK) = "Y"
                            frmImport.vsImport.Row = n&
                            frmImport.vsImport.CellForeColor = vbBlack
                            frmImport.vsImport.CellBackColor = frmImport.shpCheck.FillColor
                            frmImport.InvalidDescriptions = frmImport.InvalidDescriptions - 1
                            
                            
                        Else
                            frmImport.vsImport.CellBackColor = frmImport.shpChange.FillColor
                            'CHange Greg 26/6/06 Add in code to sort by changes
                            frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHANGE) = "Y"
                            frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHECK) = "Y"
                            frmImport.vsImport.Row = n&
                            frmImport.vsImport.CellForeColor = vbBlack
                            'frmImport.vsImport.CellBackColor = frmImport.shpCheck.FillColor
                        End If
                    End If
                    
                    'Cost...
                    If frmImport.chkCost(Layout).value = vbUnchecked Then
                        frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST) = centsToDollars(FNulls(snaDetails("Cost")))
                    ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST) <> centsToDollars(FNulls(snaDetails("Cost"))) Then
                        frmImport.vsImport.Col = COL_IMPORT_COST
                        frmImport.vsImport.Row = n&
                        frmImport.vsImport.CellBackColor = frmImport.shpChange.FillColor
                        'CHange Greg 26/6/06 Add in code to sort by changes
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHANGE) = "Y"
                        
                    End If
                    
                    'Retail...
                    If frmImport.chkRetail(Layout).value = vbUnchecked Then
                        frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL) = centsToDollars(FNulls(snaDetails("Retail")))
                    ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL) <> centsToDollars(FNulls(snaDetails("Retail"))) Then
                        frmImport.vsImport.Col = COL_IMPORT_RETAIL
                        frmImport.vsImport.Row = n&
                        frmImport.vsImport.CellBackColor = frmImport.shpChange.FillColor
                        'CHange Greg 26/6/06 Add in code to sort by changes
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHANGE) = "Y"
                    End If
                    
                    'Barcode...
                    If frmImport.chkBarcode(Layout).value = vbUnchecked Then
                        frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) = FNulls(snaDetails("Barcode"))
                    ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) <> FNulls(snaDetails("Barcode")) Then
                        frmImport.vsImport.Col = COL_IMPORT_BARCODE
                        frmImport.vsImport.Row = n&
                        frmImport.vsImport.CellBackColor = frmImport.shpChange.FillColor
                        'CHange Greg 26/6/06 Add in code to sort by changes
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHANGE) = "Y"
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHECK) = "Y"
                    End If
                    
                    'Outers...
                    If frmImport.chkOuters(Layout).value = vbUnchecked Then
                        frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS) = Trim(FNulls(snaDetails("Outers")))
                    ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS) <> Trim(FNulls(snaDetails("Outers"))) Then
                        frmImport.vsImport.Col = COL_IMPORT_OUTERS
                        frmImport.vsImport.Row = n&
                        frmImport.vsImport.CellBackColor = frmImport.shpChange.FillColor
                        'CHange Greg 26/6/06 Add in code to sort by changes
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHANGE) = "Y"
                        frmImport.vsImport.TextMatrix(n, COL_IMPORT_CHECK) = "Y"
                    End If
                    
                        
                        
                        
                End If
                snaDetails.ClsRS
                Set snaDetails = Nothing
                
            Else
                MsgBox "Somthing is wrong here please inform the programmers."
            End If
        Next
        If ExistingItemFound = False Then
            '14/10/2010 Limin changed from dbLots dbSupplier.
            Set snaDetails = dbSuppliers.OpenRecordSet("SELECT SupplierID FROM Details WHERE SupplierID = " & frmImport.cboSupplier.ItemData(frmImport.cboSupplier.ListIndex), dbOpenSnapshot)
            If snaDetails.EOF Then
            Else
                Call MsgBox("No existing products have been found and this is not a new supplier. Please double check the file and supplier.")
            End If
            snaDetails.ClsRS
            Set snaDetails = Nothing
            
        End If
        frmImport.SetInfoCaption
'    End If
    
    Set proImport = Nothing
    
'    snaDetails.ClsRS
'    Set snaDetails = Nothing
    
    Call CloseSupplierDB
    'Change Greg 8/1/07 Make so invalid descriptions count updates
    frmImport.lblInfo.Caption = "Total Items: " & frmImport.vsImport.Rows - 1 & vbCrLf
    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Barcodes: " & frmImport.InvalidBarcodesFound & vbCrLf
    frmImport.lblInfo.Caption = frmImport.lblInfo.Caption & "Invalid Descriptions: " & frmImport.InvalidDescriptions
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "ImportFile in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    Resume Next
End Sub
Public Function GetPSLNetID(ByVal Sheet As String) As String
'        If .cboPSLNetSupplierID.ListIndex > 0 Then
'                If .PSLIDS.Count > .cboPSLNetSupplierID.ItemData(.cboPSLNetSupplierID.ListIndex) Then
'                    temp = .PSLIDS("R" & .cboPSLNetSupplierID.ItemData(.cboPSLNetSupplierID.ListIndex))
'                    temp = Right(temp, Len(temp) - InStr(temp, ":"))
'                    tblSupplier("PSLNetID") = temp
'    '                tblSupplier("PSLNetID") = .PSLIDS("R" & .cboPSLNetSupplierID.ItemData(.cboPSLNetSupplierID.ListIndex))
'                Else
'                    tblSupplier("PSLNetID") = ""
'                End If
'            Else
'                tblSupplier("PSLNetID") = ""
'            End If
End Function
Function CreateSupplierInPOSMasterFromSupplierFile(ByVal v_SupplierDBID As Long) As Long
    Dim SupplierID As Long
    Dim snaDetails As Recordset
    
    Set snaDetails = dbSuppliers.OpenRecordSet("SELECT Suppliername,RemoveSupplier,Password,SettingsGuid FROM Suppliers WHERE SupplierID = " & v_SupplierDBID, dbOpenSnapshot)
    If snaDetails.EOF Then
    
    Else
        Dim snaMasterFile As Recordset
        Set snaMasterFile = dbLots.OpenRecordSet(SQL_Select & "Supplier WHERE SupplierGuid ='" & FNulls(snaDetails("SettingsGuid")) & "'", dbOpenDynaset)
        If snaMasterFile.EOF Then
            Set snaMasterFile = dbLots.OpenRecordSet(SQL_Select & "Supplier WHERE Suppliername =" & WrapTextInSingleQuotes(FNulls(snaDetails("Suppliername"))), dbOpenDynaset)
        End If
        If snaMasterFile.EOF Then
            snaMasterFile.AddNew
            snaMasterFile("SupplierGuid") = FNulls(snaDetails("SettingsGuid"))
            
        Else
            snaMasterFile.Edit
        End If
        Call DBValidate(snaMasterFile("SupplierName"), FNulls(snaDetails("Suppliername")))
        Call DBValidate(snaMasterFile("BuyingGroup"), FNulls(snaDetails("RemoveSupplier")))
        snaMasterFile("Restricted") = Len(FNulls(snaDetails("Password"))) > 0
        If FBool(FNulls(snaMasterFile("BuyingGroup"))) Then
            snaMasterFile("PriceList") = True
        End If
        CreateSupplierInPOSMasterFromSupplierFile = Val(FNulls(snaMasterFile("SupplierID")))
        snaMasterFile.Update
        
        snaMasterFile.ClsRS
        Set snaMasterFile = Nothing
        Set snaMasterFile = dbLots.OpenRecordSet(SQL_Select & "SuppUD WHERE SettingsGuid ='" & FNulls(snaDetails("SettingsGuid")) & "'", dbOpenDynaset)
        If snaMasterFile.EOF Then
           Set snaMasterFile = dbLots.OpenRecordSet(SQL_Select & "SuppUD WHERE SupplierUpdateID =" & v_SupplierDBID, dbOpenDynaset)
        End If
        If snaMasterFile.EOF Then
            snaMasterFile.AddNew
            
            
        Else
            snaMasterFile.Edit
        End If
        Call DBValidate(snaMasterFile("SupplierUpdateID"), v_SupplierDBID)
        Call DBValidate(snaMasterFile("SupplierID"), CreateSupplierInPOSMasterFromSupplierFile)
        Call DBValidate(snaMasterFile("OrderSuppID"), CreateSupplierInPOSMasterFromSupplierFile)
        Call DBValidate(snaMasterFile("SupplierName"), FNulls(snaDetails("Suppliername")))
        Call DBValidate(snaMasterFile("SettingsGuid"), FNulls(snaDetails("SettingsGuid")))
        If UsePOSUpdateSettings Then
            Call DBValidate(snaMasterFile("AutoProcess"), True)
            Call DBValidate(snaMasterFile("SupplierCost"), True)
        End If
        
        snaMasterFile.Update
        snaMasterFile.ClsRS
        Set snaMasterFile = Nothing
        
    End If
     snaDetails.ClsRS
     Set snaDetails = Nothing
End Function
Function GetLOTSSupplierID(ByVal v_SupplierDBID As Long) As Long
   GetLOTSSupplierID = CreateSupplierInPOSMasterFromSupplierFile(v_SupplierDBID)
'    Dim snaDetails As Recordset
'    Set snaDetails = dbSuppliers.OpenRecordSet("SELECT SettingsGuid FROM Suppliers WHERE SupplierID = " & v_SupplierDBID, dbOpenSnapshot)
'    If snaDetails.EOF Then
'
'    Else
'        Dim TheGuid As String
'        TheGuid = FNulls(snaDetails("SettingsGuid"))
'        GetLOTSSupplierID = Val(GetAnyField("Supplier", "SupplierID", WrapText(TheGuid), "SupplierGuid"))
'    End If
'    snaDetails.ClsRS
'    Set snaDetails = Nothing
End Function
Sub SaveDetails(ByVal DefaultSupplierID As Long, ByVal AutoUpdate As Boolean, ByVal DateUpdated As String, ByVal SaveCheck As Boolean)

On Error GoTo ErrTrap

    Dim dynDetails As Recordset
'    Dim DateUpdated As String
    Dim n As Long
    Dim proSave As New ProgressIndicator
    Dim FoundMatch As Boolean
    Dim SupplierID As Long
    Dim CostChange As Boolean
    Dim MyRows As Long
    Dim CloseDB As Boolean
    Dim LOTSSupplierID As Long
    Dim SettingsGuid As String
    Dim PartcodeGuid As String
    Dim StockcardGuid As String
    
    Dim NewItems As Long
    Dim UpdatedItems As Long
    Dim ItemsSkipped As Long
    Dim FoundBy As String
    Dim v_TheGUids As Dictionary
    Set v_TheGUids = New Dictionary
    Call OpenSupplierDB(CloseDB)
   If DataBaseOpen("LOTS") = False Then
        Exit Sub
   End If
    If frmImport.chkOwnTable(0).value = vbChecked Then
      Call SaveOwnSupplierTable(DefaultSupplierID, AutoUpdate, DateUpdated, SaveCheck)
      Exit Sub
    End If
'    Set dynDetails = dbSuppliers.OpenRecordset("SELECT * FROM Details WHERE SupplierID = " & SupplierID&, dbOpenDynaset)
    
    'Change Greg 7/11/06
    '31/10/2007 Limin fixed problems with date on 31
    
    'DateUpdated = Day(Now) & " " & frmImport.cboDate
    proSave.Show "Saving..."
    MyRows = frmImport.vsImport.Rows - 1
    proSave.RecordCount = MyRows
'    If frmImport.chkClear.Value = vbChecked Then
'        Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID, dbFailOnError)
'    End If
    'New Nigel 12.2.2013
'    If frmImport.chkClearExistingRecords.value = vbChecked Then
'        Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID, dbFailOnError)
'    End If
'
    Dim SaveTime As Date
    SaveTime = Now
    For n& = MyRows To 1 Step -1
        proSave.increment MyRows - n&
        SupplierID = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SUPPLIERID)
        If SupplierID = 0 Then
            SupplierID = DefaultSupplierID
        End If
        If v_TheGUids.Exists("S" & SupplierID) Then
            LOTSSupplierID = v_TheGUids("S" & SupplierID)
        Else
           
            
            
            LOTSSupplierID = CreateSupplierInPOSMasterFromSupplierFile(SupplierID)
            
            Call v_TheGUids.Add("S" & SupplierID, LOTSSupplierID)
            If frmImport.chkClearExistingRecords.value = vbChecked Then
                Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & SupplierID, dbFailOnError)
                'mark all stockcards as deleted in master pos file
                'then as we add in partcode will unmark them
                Call ExecuteSafe("UPDATE Partcode SET pcDeleted = true where SupplierID = " & LOTSSupplierID)
            End If
        End If
        If n& Mod 500 = 0 Then
            proSave.Show "Saving..."
            DoEvents
        End If
        FoundBy = ""
        
        If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CHECK) = "Y" And SaveCheck Then
            If Len(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)) > 40 And Len(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION)) < 40 Then
                frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION)
            End If
        End If
        
        'If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) = "9417679077133" Then
        '    Debug.Print
        'End If

        If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_INVALID) = "Y" And frmImport.chkTrimAndSave.value = vbUnchecked Then
            ItemsSkipped = ItemsSkipped + 1
        ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CHECK) = "Y" And SaveCheck = False Then
            ItemsSkipped = ItemsSkipped + 1
        Else
            'Change Greg 7/9/2004 Search by all codes if they are avaliable
            FoundMatch = False
            
            
            
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) <> "" Then
                Set dynDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'", dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                    FoundBy = "Barcode:" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)
                    
                End If
            End If
            
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) <> "" And FoundMatch = False Then
                Set dynDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND PharmaCode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE) & "'", dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "PharmaCode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                    FoundBy = "Pharmacode:" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)
                End If
            End If
            
            If Len(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE))) > 1 And FoundMatch = False Then
                Set dynDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'", dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                    FoundBy = "Partcode:" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)
                End If
            End If
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)) <> "" And FoundMatch = False Then
                'Change Greg 15/9/2004 If no other match tried then use name
                If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)) = "" And Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) = "" And Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) = "" Then
                    Set dynDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Description Like(" & Chr(34) & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) & Chr(34) & ")", dbOpenDynaset)
                    If dynDetails.EOF Then
                        dynDetails.ClsRS
                        Set dynDetails = Nothing
                    Else
    '                dynDetails.FindFirst "Description Like('" & Replace(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION), "'", "''") & "')"
    '                If dynDetails.NoMatch = False Then
                        FoundMatch = True
                        FoundBy = "Description:" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)
                    End If
                End If
            End If
            If dynDetails Is Nothing Then
                Set dynDetails = dbSuppliers.OpenRecordSet("SELECT * FROM Details WHERE SupplierID = " & SupplierID& & " AND Description = ''", dbOpenDynaset)
            End If
            If FoundMatch = False Then
            'If dynDetails.NoMatch Then
                NewItems = NewItems + 1
                dynDetails.AddNew
                dynDetails("SupplierID") = SupplierID&
                dynDetails("DateCreated") = frmImport.cboDate.text
            Else
                dynDetails.Edit
                UpdatedItems = UpdatedItems + 1
            End If
            CostChange = False
            Call DBValidate(dynDetails("Code"), Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)))
            dynDetails("Pharmacode") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)
            If dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST)) <> FNulls(dynDetails("Cost")) Then
                CostChange = True
            End If
            dynDetails("Cost") = dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST))
            
            If frmImport.vsImport.ColWidth(COL_IMPORT_RETAIL) > 0 Then
                If Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL)) > 0 Or frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL) <> "" Then
                    dynDetails("Retail") = dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL))
                End If
            End If
            If IsNumeric(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE))) Then
                If InStr(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)), ",") Then
                    frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) = GetValueFromString(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE), "", ",")
                End If
                Call DBValidate(dynDetails("Barcode"), Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)))
            End If
            'Change Greg 31/10/7 Add in product group
            dynDetails("ProductGroup") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PRODUCTGROUP)
            If glbDBType = OZBro Then
               If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_GST) <> "" Then
                    'Change Greg 23/3/2004 Allow for numbers in gst
                    If UCase(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_GST)) = "N" Then
                        dynDetails("FreeGST") = True
                    Else '"Y"
                        dynDetails("FreeGST") = False
                    End If
                Else
                    dynDetails("FreeGST") = False
                End If
    
                
                If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COSTINCLGST) <> "" Then
                    dynDetails("CostInclGST") = dollarsToCents((frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COSTINCLGST)))
                Else
                    dynDetails("CostInclGST") = 0
                End If
                 'Change Greg 7/1/11 make so records the api ,sigma and faulding codes
                dynDetails("API") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_API_CODE)
                dynDetails("Sigma") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SIGMA_CODE)
                dynDetails("Mayne") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_FAULDING_CODE)
            End If
            If Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS)) > 0 Then
                dynDetails("Outers") = Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS))
            End If
           frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) = FindandReplace(FindandReplace(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION), vbCr, ""), vbLf, "")
            If Len(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))) > 40 Then
                'MsgBox "Invalid Description " & Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))
                If Len(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION))) > 40 Then
                    dynDetails("Description") = LEFT(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)), 40)
                ElseIf Len(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION))) = 0 Then
                    dynDetails("Description") = LEFT(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)), 40)
                Else
                    dynDetails("Description") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION))
                End If
            Else
                dynDetails("Description") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))
            End If
            If glbDBType = OZBro Then
                Call UpdateMasterPOSDb(FNulls(dynDetails("Pharmacode")), FNulls(dynDetails("Code")), LOTSSupplierID, FNulls(dynDetails("Description")), FNulls(dynDetails("Cost")), FNulls(dynDetails("Retail")), FNulls(dynDetails("Outers")), FNulls(dynDetails("API")), FNulls(dynDetails("Sigma")), FNulls(dynDetails("Mayne")), FNulls(dynDetails("Barcode")), FNulls(dynDetails("FreeGST")), PartcodeGuid, StockcardGuid)
            Else
                Call UpdateMasterPOSDb(FNulls(dynDetails("Pharmacode")), FNulls(dynDetails("Code")), LOTSSupplierID, FNulls(dynDetails("Description")), FNulls(dynDetails("Cost")), FNulls(dynDetails("Retail")), FNulls(dynDetails("Outers")), "", "", "", FNulls(dynDetails("Barcode")), False, PartcodeGuid, StockcardGuid)
            End If
            dynDetails("PartcodeGuid") = PartcodeGuid
            dynDetails("StockcardGuid") = StockcardGuid
            '31/10/2007 Limin
            dynDetails("LastUpdated") = FormatDate(frmImport.cboDate.text, False)
            If CostChange Then
                dynDetails("LastChange") = FormatDate(frmImport.cboDate.text, False)
                End If
            dynDetails.Update
            dbSuppliers.Execute "UPDATE Suppliers SET LastUpdated = Now WHERE SupplierID = " & SupplierID, dbFailOnError
            dbLots.Execute "UPDATE SuppUD Set LastUpdated = #" & FormatDate(frmImport.cboDate.text, False) & "# where supplierid = " & LOTSSupplierID
            dbLots.Execute "UPDATE Supplier Set supDateModified = #" & FormatDate(frmImport.cboDate.text, False) & "# where supplierid = " & LOTSSupplierID
            frmImport.vsImport.RemoveItem (n&)
        End If
    Next
    UpdateLastUpdatedInDatabase
'    If frmImport.chkClear.Value = vbChecked Then
        If AutoUpdate = False Then
            If glbSaveAfterImport Then
                If frmImport.SuppliersProcessed <> "" Then
                    frmImport.SuppliersProcessed = Right(frmImport.SuppliersProcessed, Len(frmImport.SuppliersProcessed) - 1)
                    Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID  in (" & frmImport.SuppliersProcessed & ") AND LastUpdated <> #" & frmImport.cboDate.text & "#", dbFailOnError)
                    frmImport.SuppliersProcessed = ""
                Else
    '                Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID & " AND LastUpdated <> #" & DateUpdated & "#", dbFailOnError)
                End If
            Else
                If frmImport.chkClear.value = vbChecked Then
                    Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID & " AND LastUpdated <> #" & FormatDate(frmImport.cboDate.text, False) & "#", dbFailOnError)
                End If
            End If
        Else
'            If frmImport.chkClear.Value = vbChecked Then
'                frmImport.SuppliersProcessed = frmImport.SuppliersProcessed & "," & DefaultSupplierID
'            End If
        End If
'    End If
    proSave.Hide
    Set proSave = Nothing
    If Not dynDetails Is Nothing Then
        dynDetails.ClsRS
        Set dynDetails = Nothing
    End If
    If frmImport.chkClearExistingRecords.value = vbChecked And UpdatedItems > 0 Then
        Call MsgBox("Supplier Saved: New Items:" & NewItems & " Items Changed:" & UpdatedItems & " Items Skipped:" & ItemsSkipped & " since this is set to clear items that means duplicate codes in the import file.")
    Else
        Call MsgBox("Supplier Saved: New Items:" & NewItems & " Items Changed:" & UpdatedItems & " Items Skipped:" & ItemsSkipped)
    End If
    frmImport.chkClearExistingRecords.value = vbUnchecked
    Call UPdatePassword(DefaultSupplierID)
    'Change Greg 23/8/2005 Set the password on the database
'    If frmImport.txtPassword.text <> "" Then
'        dbSuppliers.Execute "UPDATE Suppliers SET Password = '" & EncryptPassword(UCase(Trim(frmImport.txtPassword.text))) & "' WHERE SupplierID = " & SupplierID, dbFailOnError
'    End If
    Call CloseSupplierDB(CloseDB)

    
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "SaveDetails in " & app.ExeName)
    MsgBox Err.Description, vbExclamation, Err.NUmber
    Resume Next
End Sub
Public Sub UPdatePassword(ByVal SupplierID As Long)
     If frmImport.txtPassword.text <> "" Then
        dbSuppliers.Execute "UPDATE Suppliers SET Password = '" & EncryptPassword(UCase(Trim(frmImport.txtPassword.text))) & "' WHERE SupplierID = " & SupplierID, dbFailOnError
    End If
End Sub
Public Sub SaveOwnSupplierTable(ByVal DefaultSupplierID As Long, ByVal AutoUpdate As Boolean, ByVal DateUpdated As String, ByVal SaveCheck As Boolean)
'---------------------------------------------------------------------------------------
' Procedure : SaveOwnSupplierTable
' DateTime  : 10/03/2010 15:51
' Author    : LiminOZ
' Purpose   : This works for
'TBLNAME: Nadglen
'Pharmacode: API
'---------------------------------------------------------------------------------------
Dim tblName As String
    Dim dynDetails As Recordset
'    Dim DateUpdated As String
    Dim n As Long
    Dim proSave As New ProgressIndicator
    Dim FoundMatch As Boolean
    Dim SupplierID As Long
    Dim CostChange As Boolean
    Dim MyRows As Long
    Dim CloseDB As Boolean



On Error GoTo SaveOwnSupplierTable_Error
  tblName = "Nadglen"
    Call OpenSupplierDB(CloseDB)
    Call dbSuppliers.Execute("Delete FROM " & tblName, dbFailOnError)
    proSave.Show "Saving..."
    MyRows = frmImport.vsImport.Rows - 1
    proSave.RecordCount = MyRows
    
    For n& = MyRows To 1 Step -1
        proSave.increment MyRows - n&
        SupplierID = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SUPPLIERID)
        If SupplierID = 0 Then
            SupplierID = DefaultSupplierID
        End If
        If n& Mod 500 = 0 Then
            proSave.Show "Saving..."
            DoEvents
        End If
        If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CHECK) = "Y" And SaveCheck Then
            If Len(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)) > 40 And Len(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION)) < 40 Then
                frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OLD_DESCRIPTION)
            End If
        End If
        If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_INVALID) = "Y" Then
            
        ElseIf frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CHECK) = "Y" And SaveCheck = False Then
        Else
            'Change Greg 7/9/2004 Search by all codes if they are avaliable
            FoundMatch = False
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) <> "" Then
                Set dynDetails = dbSuppliers.OpenRecordSet(SQL_Select & tblName & " WHERE Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'", dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "Barcode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                End If
            End If
            
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) <> "" And FoundMatch = False Then
                Set dynDetails = dbSuppliers.OpenRecordSet(SQL_Select & tblName & " WHERE API = " & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE), dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "PharmaCode = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                End If
            End If
            
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)) <> "" And FoundMatch = False Then
                Set dynDetails = dbSuppliers.OpenRecordSet(SQL_Select & tblName & " WHERE Code = " & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE), dbOpenDynaset)
                If dynDetails.EOF Then
                    dynDetails.ClsRS
                    Set dynDetails = Nothing
                Else
    '            dynDetails.FindFirst "Code = '" & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE) & "'"
    '            If dynDetails.NoMatch = False Then
                    FoundMatch = True
                End If
            End If
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)) <> "" And FoundMatch = False Then
                'Change Greg 15/9/2004 If no other match tried then use name
                If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)) = "" And Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE)) = "" And Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) = "" Then
                    Set dynDetails = dbSuppliers.OpenRecordSet(SQL_Select & tblName & " WHERE Description Like(" & Chr(34) & frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION) & Chr(34) & ")", dbOpenDynaset)
                    If dynDetails.EOF Then
                        dynDetails.ClsRS
                        Set dynDetails = Nothing
                    Else
    '                dynDetails.FindFirst "Description Like('" & Replace(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION), "'", "''") & "')"
    '                If dynDetails.NoMatch = False Then
                        FoundMatch = True
                    End If
                End If
            End If
            If dynDetails Is Nothing Then
                Set dynDetails = dbSuppliers.OpenRecordSet(SQL_Select & tblName & " WHERE Description = ''", dbOpenDynaset)
            End If
            If FoundMatch = False Then
            'If dynDetails.NoMatch Then
                dynDetails.AddNew
                 
            Else
                dynDetails.Edit
            End If
            CostChange = False
            dynDetails("Code") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE))
            '2/3/2011 Limin API is the code for Nadglen
            If IsNumeric(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE))) Then
              dynDetails("API") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_CODE)
            End If
            'If IsNumeric(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)) Then
            '  dynDetails("API") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PHARMACODE)
            'End If
            If dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST)) <> FNulls(dynDetails("Cost")) Then
                CostChange = True
            End If
            dynDetails("Cost") = dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COST))
            
            If frmImport.vsImport.ColWidth(COL_IMPORT_RETAIL) > 0 Then
                If Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL)) > 0 Or frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL) <> "" Then
            '        dynDetails("Retail") = dollarsToCents(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_RETAIL))
                End If
            End If
            If IsNumeric(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE))) Then
                dynDetails("Barcode") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_BARCODE))
            End If
            
            '2/3/2011 Limin will update Faulding and Sigma
            If IsNumeric(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_FAULDING_CODE))) Then
              dynDetails("Mayne") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_FAULDING_CODE))
            End If
            If IsNumeric(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SIGMA_CODE))) Then
              dynDetails("Sigma") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_SIGMA_CODE))
            End If
            'Change Greg 31/10/7 Add in product group
            'dynDetails("ProductGroup") = frmImport.vsImport.TextMatrix(n&, COL_IMPORT_PRODUCTGROUP)
             

            
            'If frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COSTINCLGST) <> "" Then
            '    dynDetails("CostInclGST") = dollarsToCents((frmImport.vsImport.TextMatrix(n&, COL_IMPORT_COSTINCLGST)))
            'Else
            '    dynDetails("CostInclGST") = 0
            'End If
            
            'If Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS)) > 0 Then
            '    dynDetails("Outers") = Val(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_OUTERS))
            'End If
            If Len(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))) > 40 Then
                'MsgBox "Invalid Description " & Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))
                dynDetails("Description") = LEFT(Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION)), 40)
            Else
                dynDetails("Description") = Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_DESCRIPTION))
            End If
            If Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_GST)) <> "" Then
              Call DBValidate(dynDetails("TaxCode"), Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_GST)))
              If InStr("YN", Trim(frmImport.vsImport.TextMatrix(n&, COL_IMPORT_GST))) Then
                Call DBValidate(dynDetails("GST"), 10)
              Else
                Call DBValidate(dynDetails("GST"), 0)
              End If
            End If
            '31/10/2007 Limin
           'dynDetails("LastUpdated") = frmImport.cboDate.Text
           ' If CostChange Then
           '     dynDetails("LastChange") = frmImport.cboDate.Text
           ' End If
            dynDetails.Update
            
            frmImport.vsImport.RemoveItem (n&)
        End If
    Next
    dbSuppliers.Execute "UPDATE Suppliers SET LastUpdated = Now WHERE SupplierID = " & SupplierID, dbFailOnError
'    If frmImport.chkClear.Value = vbChecked Then
        If AutoUpdate = False Then
            If glbSaveAfterImport Then
                If frmImport.SuppliersProcessed <> "" Then
                    frmImport.SuppliersProcessed = Right(frmImport.SuppliersProcessed, Len(frmImport.SuppliersProcessed) - 1)
                    Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID  in (" & frmImport.SuppliersProcessed & ") AND LastUpdated <> #" & frmImport.cboDate.text & "#", dbFailOnError)
                    frmImport.SuppliersProcessed = ""
                Else
    '                Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID & " AND LastUpdated <> #" & DateUpdated & "#", dbFailOnError)
                End If
            Else
                If frmImport.chkClear.value = vbChecked Then
                    Call dbSuppliers.Execute("DELETE * FROM DETAILS Where SupplierID = " & DefaultSupplierID & " AND LastUpdated <> #" & frmImport.cboDate.text & "#", dbFailOnError)
                End If
            End If
        Else
'            If frmImport.chkClear.Value = vbChecked Then
'                frmImport.SuppliersProcessed = frmImport.SuppliersProcessed & "," & DefaultSupplierID
'            End If
        End If
'    End If
    proSave.Hide
    Set proSave = Nothing
    If Not dynDetails Is Nothing Then
        dynDetails.ClsRS
        Set dynDetails = Nothing
    End If
    
    'Change Greg 23/8/2005 Set the password on the database
    If frmImport.txtPassword.text <> "" Then
        dbSuppliers.Execute "UPDATE Suppliers SET Password = '" & EncryptPassword(UCase(Trim(frmImport.txtPassword.text))) & "' WHERE SupplierID = " & SupplierID, dbFailOnError
    End If
    Call CloseSupplierDB(CloseDB)

SaveOwnSupplierTable_Exit:
   On Error Resume Next
Exit Sub

SaveOwnSupplierTable_Error:
   If SYSLOG(Err, "SaveOwnSupplierTable in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
     Resume SaveOwnSupplierTable_Exit
   Else
     Resume Next
   End If


End Sub
'Change so that the file now looks to the combo box settings
Public Sub ImportDBFile(ByVal FileName As String, ByVal SupplierName As String)
On Error GoTo ErrTrap
    Dim dbPOSSupplier As Database

    Dim Failed As Boolean

    Dim Code As String
    Dim Cost As Single
    Dim Outers As Long
    Dim Barcode As String
    Dim Description As String
    Dim Retail As Single
    Dim Pharmacode As String

    Dim InvalidDescriptions As Long
    Dim InvalidBarcodesFound As Long
    Dim proImport As New ProgressIndicator

    Dim rawdata2 As String
    Dim rawdataPos As Long

    Dim LastDescription2 As String
    Dim LastCost As Long
    Dim LastRetail As Long
    Dim SupplierID As Long
    Dim snapPOSPrice As Recordset
'    Dim GSTFree As Boolean
'    Dim CostInclGST As Long

    'Open POS Supplier MDB
    '.vsImport.FormatString = "<Code         |<Description                                                         |>Cost           |>Retail       |<Barcode                     "
    '.vsImport.Cols = 5
    With frmImport
        SupplierID = .cboSupplier.ItemData(.cboSupplier.ListIndex)
        If .chkPharmacode(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_PHARMACODE) = 0
        Else
           .vsImport.ColWidth(COL_IMPORT_PHARMACODE) = 1000
        End If
        If .chkCode(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_CODE) = 0
        End If
        If .chkOuters(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_OUTERS) = 0
        End If
        .vsImport.ColWidth(COL_IMPORT_OLDCOST) = 0
        If .chkBarcode(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_BARCODE) = 0
        End If
        If .chkCostInclGST(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_COSTINCLGST) = 0
        End If
        If .chkGST(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_GST) = 0
        End If

        If .chkRetail(0).value = vbUnchecked Then
            .vsImport.ColWidth(COL_IMPORT_RETAIL) = 0

        End If
        .vsImport.Refresh
        Set dbPOSSupplier = Workspaces(0).OpenDatabase(FileName)
        Set snapPOSPrice = dbPOSSupplier.OpenRecordSet("SELECT * FROM Price", dbOpenSnapshot)

        If snapPOSPrice.RecordCount > 0 Then
            snapPOSPrice.MoveLast
            snapPOSPrice.MoveFirst
        Else
            GoTo ExitImportDBFile
        End If
        .vsImport.Rows = 1
        .vsImport.Redraw = False

        proImport.Show "Importing Price Update..."
        proImport.RecordCount = snapPOSPrice.RecordCount


         While Not snapPOSPrice.EOF
            proImport.increment snapPOSPrice.AbsolutePosition

            If .chkCode(0).value = vbChecked Then
                Code$ = Trim(FNulls(snapPOSPrice(Val(.cboCode(0).text) - 1)))
                '28/7/2004 Limin add Header
                If .chkAddCode(0).value = vbChecked Then
                    If Trim(FNulls(snapPOSPrice("Code"))) <> "" Then
                        Code$ = .txtDescHeader(0).text & " " & Code$
    '                    Code$ = .txtDescHeader.Text & " " & Trim(FNulls(snapPOSPrice("Code")))
                    Else
                        Code$ = ""
                    End If

    '            Else
    '                Code$ = Trim(FNulls(snapPOSPrice(Val(.cboCode.Text) - 1)))
    '                Code$ = Trim(FNulls(snapPOSPrice("Code")))
                End If
                'Code$ = Trim(FNulls(snapPOSPrice("Code")))
                'If Code = "9873" Then
                '    Debug.Print
                'End If


            End If
            If .chkPharmacode(0).value = vbChecked Then
                Pharmacode$ = Trim(FNulls(snapPOSPrice(Val(.cboPharmacode(0).text) - 1)))
                'PharmaCode$ = Trim(FNulls(snapPOSPrice("PharmaCode")))
            End If

            If .chkCost(0).value = vbChecked Then
                If InStr(Trim(FNulls(snapPOSPrice(Val(.cboCost(0).text) - 1))), ".") = 0 And InStr(Trim(FNulls(snapPOSPrice(Val(.cboCost(0).text) - 1))), "$") = 0 Then
                    'Must be dollars
                    Cost = dollarsToCents(CStr(100 * Val((FNulls(snapPOSPrice(Val(.cboCost(0).text) - 1))))))
                Else
                    Cost = dollarsToCents(Trim(FNulls(snapPOSPrice(Val(.cboCost(0).text) - 1))))
                End If
            End If
            If .chkRetail(0).value = vbChecked Then
                If InStr(Trim(FNulls(snapPOSPrice(Val(.cboRetail(0).text) - 1))), ".") = 0 And InStr(Trim(FNulls(snapPOSPrice(Val(.cboRetail(0).text) - 1))), "$") = 0 Then
                    'must be dollars
                    Retail = dollarsToCents(CStr(100 * Trim(FNulls(snapPOSPrice(Val(.cboRetail(0).text) - 1)))))
                Else
                    Retail = dollarsToCents(Trim(FNulls(snapPOSPrice(Val(.cboRetail(0).text) - 1))))
                End If
            End If
'            If InStr(Trim(FNulls(snapPOSPrice("Cost"))), "$") > 0 Then
'                Cost = dollarsToCents(Trim(FNulls(snapPOSPrice("Cost"))))
'            Else
'                Cost = Val(FNulls(snapPOSPrice("Cost")))
'            End If
            'Greg Change 30/3/2004 Not required in NZ
    '        If .chkGST.Value = vbChecked Then
    '            If IsNumeric(FNulls(snapPOSPrice("GST"))) Then
    '                If Val(FNulls(snapPOSPrice("GST"))) > 0 Then
    '                    GSTFree = False
    '                Else
    '                    GSTFree = True
    '                End If
    '            ElseIf UCase(Trim(FNulls(snapPOSPrice("GST")))) = "N" Or UCase(Trim(FNulls(snapPOSPrice("GST")))) = "NO" Then
    '                GSTFree = True
    '            ElseIf UCase(Trim(FNulls(snapPOSPrice("GST")))) = "Y" Or UCase(Trim(FNulls(snapPOSPrice("GST")))) = "YES" Then
    '                GSTFree = False
    '            End If
    '        End If


'            If InStr(Trim(FNulls(snapPOSPrice("Retail"))), "$") > 0 Or InStr(Trim(FNulls(snapPOSPrice("Retail"))), ".") > 0 Then
'                Retail = dollarsToCents(Trim(FNulls(snapPOSPrice("Retail"))))
'            Else
'                Retail = Val(FNulls(snapPOSPrice("Retail")))
'            End If

            '14/11/2003 Limin Add Cost incl GST
            'Greg Change 30/3/2004 Not required in NZ
    '        If .chkCostInclGST.Value = vbChecked Then
    '            If InStr(Trim(FNulls(snapPOSPrice("CostInclGST"))), "$") > 0 Or InStr(Trim(FNulls(snapPOSPrice("CostInclGST"))), ".") > 0 Then
    '                CostInclGST = dollarsToCents(Trim(FNulls(snapPOSPrice("CostInclGST"))))
    '            Else
    '                CostInclGST = Val(FNulls(snapPOSPrice("Retail")))
    '            End If
    '        End If

            If .chkBarcode(0).value = vbChecked Then
                Barcode$ = Trim(FNulls(snapPOSPrice(Val(.cboBarcode(0).text) - 1)))
                'Barcode$ = Trim(FNulls(snapPOSPrice("Barcode")))
            End If
            If .chkOuters(0).value = vbChecked Then
                Outers& = Val((FNulls(snapPOSPrice(Val(.cboOuters(0).text) - 1))))
                'Outers& = Val((FNulls(snapPOSPrice("OuterSize"))))
            End If
            'Change Greg 25/8/06 Change the way the code handles the descriptions
            Description$ = Trim(FNulls(snapPOSPrice("Description")))
            If .chkSuffix(0).value = vbChecked Then
                If Trim(FNulls(snapPOSPrice(Val(.cboSuffix(0).text) - 1))) <> "" Then
                    Description$ = Description & " " & Trim(FNulls(snapPOSPrice(Val(.cboSuffix(0).text) - 1)))
                End If
            End If
            If .chkDescription2(0).value = vbChecked Then
                If Trim(FNulls(snapPOSPrice(Val(.cboDescription2(0).text) - 1))) <> "" Then
                    Description$ = Trim(FNulls(snapPOSPrice(Val(.cboDescription2(0).text) - 1))) & " " & Description
                End If
            End If
            '28/7/2004 Limin Add prefix
'            If .chkAddDescription.Value = vbChecked Then
'                If Trim(FNulls(snapPOSPrice(Val(.cboDescription2.text) - 1))) <> "" Then
'                'If Trim(FNulls(snapPOSPrice("Description2"))) <> "" Then
'                    Description$ = Trim(FNulls(snapPOSPrice(Val(.cboDescription2.text) - 1))) & " " & Trim(FNulls(snapPOSPrice(Val(.cboDescription.text) - 1)))
'                    'Description$ = Trim(FNulls(snapPOSPrice("Description2"))) & " " & Trim(FNulls(snapPOSPrice("Description")))
'                Else
'                    Description$ = Trim(FNulls(snapPOSPrice(Val(.cboDescription.text) - 1)))
'                End If
'                'Change Greg 17/8/06 Add in saving of suffix to details
'                If Trim(FNulls(snapPOSPrice(Val(.cboSuffix.text) - 1))) <> "" Then
'                    Description$ = Trim(Description & " " & FNulls(snapPOSPrice(Val(.cboSuffix.text) - 1)))
'                End If
'            Else
'                Description$ = Trim(FNulls(snapPOSPrice("Description")))
'            End If

            'Outers& = 0
            'PharmaCode$ = ""
                'Item Code...
            If (Trim(Code$) <> "" Or Trim(Pharmacode$) <> "") And (Val(Cost) > 0 Or Description$ <> "") Or UCase(SupplierName) Like "STARLET*" Then
                With .vsImport
                    '.vsImport.FormatString = "<Code         |<Description                                                         |>Cost           |>Retail       |<Barcode                     "
                    '.AddItem Code$ & vbTab & Trim(.txtDescHeader.Text) & " " & TitleCase(Description$) & vbTab & Cost & vbTab & Retail & vbTab & Barcode
                    .AddItem ""
                    If .chkCode.value = vbChecked Then
                        .TextMatrix(.Rows - 1, COL_IMPORT_CODE) = Code$
                    End If

                    If .chkPharmacode.value = vbChecked Then
                        .TextMatrix(.Rows - 1, COL_IMPORT_PHARMACODE) = Pharmacode$
                    End If
                    If .chkOuters.value = vbChecked Then
                        .TextMatrix(.Rows - 1, COL_IMPORT_OUTERS) = Outers&
                    End If

                    'Change 23/3/2004 Use value for GST
                    'Greg Change 30/3/2004 Not required in NZ
    '                If .chkGST.Value = vbChecked Then
    '                    If GSTFree Then
    '                        .TextMatrix(.Rows - 1, COL_IMPORT_GST) = 0
    '                        '.TextMatrix(.Rows - 1, COL_IMPORT_GST) = "Free"
    '                    Else
    '                        .TextMatrix(.Rows - 1, COL_IMPORT_GST) = 10
    '                        '.TextMatrix(.Rows - 1, COL_IMPORT_GST) = "NOT Free"
    '                    End If
    '                End If
                    'Greg Change 30/3/2004 Not required in NZ
    '                If .chkCostInclGST.Value = vbChecked Then
    '                    .TextMatrix(.Rows - 1, COL_IMPORT_COSTINCLGST) = centsToDollars(CostInclGST&)
    '                End If


                    '.TextMatrix(.Rows - 1, COL_IMPORT_PHARMACODE) = PharmaCode
                    '8/7/2004 Limin
                    If Len(TitleCase(Description$)) > 40 Then
                        .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = TitleCase(Description$)
                    Else
                        .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = TitleCase(Description$)
                    End If
                    .TextMatrix(.Rows - 1, COL_IMPORT_COST) = centsToDollars(Cost)
                    .TextMatrix(.Rows - 1, COL_IMPORT_RETAIL) = centsToDollars(Retail)
                    .TextMatrix(.Rows - 1, COL_IMPORT_BARCODE) = Barcode$
                    '.TextMatrix(.Rows - 1, COL_IMPORT_OUTERS) = Outers
                    
                    'If Len(Trim(.txtDescHeader.Text) & " " & TitleCase(Description$)) > 40 Then
                    If Len(.TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION)) > 40 Then
                        .Row = .Rows - 1
                        .Col = COL_IMPORT_DESCRIPTION
                        .CellBackColor = .shpInvalid.FillColor
                        InvalidDescriptions = InvalidDescriptions + 1
                        .TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION) = LEFT(.TextMatrix(.Rows - 1, COL_IMPORT_DESCRIPTION), 40)
                    End If
                    .TextMatrix(.Rows - 1, COL_IMPORT_SUPPLIERID) = SupplierID
                    .TextMatrix(.Rows - 1, COL_IMPORT_LAYOUT) = 0
                End With
            End If
            snapPOSPrice.MoveNext
        Wend
        proImport.Hide
        Set proImport = Nothing
        .vsImport.Redraw = True

        .lblInfo.Caption = "Total Items: " & .vsImport.Rows - 1 & vbCrLf
        .lblInfo.Caption = .lblInfo.Caption & "Invalid Barcodes: " & InvalidBarcodesFound & vbCrLf
        .lblInfo.Caption = .lblInfo.Caption & "Invalid Descriptions: " & InvalidDescriptions
    End With
ExitImportDBFile:

    snapPOSPrice.ClsRS
    Set snapPOSPrice = Nothing
    dbPOSSupplier.ClsDB
    Set dbPOSSupplier = Nothing

Exit Sub
ErrTrap:
    Call SYSLOG(Err, "ImportFile in " & app.ExeName)
    'MsgBox Err.Description, vbExclamation, Err.Number
    Resume Next
End Sub

'Public Function ConvertStringToArray(ByVal RawData As String, data() As String)
'    On Error GoTo ErrTrap
'    If UCase(.cboSupplier.Text) = "ALMAY" Then
'        With
'            'Item Code...
'            If UBound(data$) >= .cboCode.ListIndex And .chkCode.Value = vbChecked Then
'                If Trim(data$(.cboCode.ListIndex)) <> "" Then
'                    Code$ = Left(RawData$, 17)
'                    'Code$ = Trim(data$(.cboCode.ListIndex))
'                    Code$ = FindandReplace(Code$, """", "")
'                    Code$ = Trim(Code$)
'                End If
'            End If
'
'            'Item PharmaCode...
'            'If UBound(data$) >= .cboPharmacode.ListIndex And .chkPharmacode.Value = vbChecked Then
'            '    If IsPharmacodeValid(data$(.cboPharmacode.ListIndex)) Then
'            '        PharmaCode$ = Trim(data$(.cboPharmacode.ListIndex))
'            '        PharmaCode$ = FindandReplace(PharmaCode$, """", "")
'            '    End If
'            'End If
'
'            'Description...
'            If UBound(data$) >= .cboDescription.ListIndex And .chkDescription.Value = vbChecked Then
'                If Trim(data$(.cboDescription.ListIndex)) <> "" Then
'                    Description$ = Trim(data$(.cboDescription.ListIndex))
'                    Description$ = FindandReplace(Description$, """", "")
'                End If
'            End If
'
'            'Item Cost...
'            If UBound(data$) >= .cboCost.ListIndex And .chkCost.Value = vbChecked Then
'                If IsNumeric(data$(.cboCost.ListIndex)) Then
'                    Cost& = dollarsToCents(data$(.cboCost.ListIndex))
'                End If
'            End If
'
'            'Item Retail...
'            If UBound(data$) >= .cboRetail.ListIndex And .chkRetail.Value = vbChecked Then
'                If IsNumeric(data$(.cboRetail.ListIndex)) Then
'                    Retail& = dollarsToCents(data$(.cboRetail.ListIndex))
'                End If
'            End If
'
'            'Description 2...
'            If UBound(data$) >= .cboDescription2.ListIndex And .chkDescription2.Value = vbChecked Then
'                If Trim(LastDescription2) <> "" And Trim(Description$) <> "" Then
'                    If .chkDescription2.Value = vbChecked Then
'                        If .cboDescription2.ListIndex >= 0 Then
'                            If InStr(UCase(Description$), UCase(LastDescription2)) = 0 Then
'                                Description$ = LastDescription2 & " " & Description$
'                                If Cost& = 0 And Retail& = 0 Then
'                                    Cost& = LastCost&
'                                    Retail& = LastRetail&
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'
'            'Barcode...
'            If UBound(data$) >= .cboBarcode.ListIndex And .chkBarcode.Value = vbChecked Then
'                data$(.cboBarcode.ListIndex) = FindandReplace(data$(.cboBarcode.ListIndex), " ", "")
'                data$(.cboBarcode.ListIndex) = FindandReplace(data$(.cboBarcode.ListIndex), """", "")
'                If IsNumeric(data$(.cboBarcode.ListIndex)) And InStr(data$(.cboBarcode.ListIndex), "E") = 0 Then
'                    Barcode$ = data$(.cboBarcode.ListIndex)
'                    Barcode$ = FindandReplace(Barcode$, """", "")
'                ElseIf InStr(data$(.cboBarcode.ListIndex), "E") > 0 Then
'                    InvalidBarcodesFound = InvalidBarcodesFound + 1
'                End If
'            End If
'
'            'Outers...
'            If UBound(data$) >= .cboOuters.ListIndex And .chkOuters.Value = vbChecked Then
'                If IsNumeric(data$(.cboOuters.ListIndex)) Then
'                    Outers& = Val(data$(.cboOuters.ListIndex))
'                    Outers& = FindandReplace(Outers&, """", "")
'                End If
'            End If
'
'
'
'
'            'GST
'            If UBound(data$) >= .cboGST.ListIndex And .chkGST.Value = vbChecked Then
'                If IsNumeric(data$(.cboGST.ListIndex)) Then
'                    GST& = Val(data$(.cboGST.ListIndex))
'                    GST& = FindandReplace(GST&, """", "")
'                End If
'            End If
'
'
'            'Cost incl GST
'            If UBound(data$) >= .cboCostInclGST.ListIndex And .chkCostInclGST.Value = vbChecked Then
'                If IsNumeric(data$(.cboCostInclGST.ListIndex)) Then
'                    CostInclGST& = Val(data$(.cboCostInclGST.ListIndex))
'                    CostInclGST& = FindandReplace(CostInclGST&, """", "")
'                End If
'            End If
'        End With
'
'
'
'    End If
'
'
'
'
'Exit Function
'ErrTrap:
'    If Syslog(Err, "ConvertStringToArray in " & App.EXEName) Then
'        Exit Function
'    Else
'        Resume Next
'    End If
'
'
'End Function
Public Sub LoadNumbersIntoListBox(ByRef cboListBox As ComboBox)
    Dim Count As Integer
On Error GoTo LoadNumbersIntoListBox_Error
    cboListBox.Clear
    For Count% = 1 To 20
        cboListBox.AddItem Count%
    Next Count%

LoadNumbersIntoListBox_Exit:
   On Error Resume Next
Exit Sub

LoadNumbersIntoListBox_Error:
   If SYSLOG(Err, "LoadNumbersIntoListBox in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume LoadNumbersIntoListBox_Exit
   Else
       Resume Next
   End If
End Sub
Public Sub ClearOldSpecials()
On Error GoTo ErrorHandler
    Call OpenSupplierDB
    Call dbSuppliers.Execute("DELETE SpecialItems.* FROM Specials INNER JOIN SpecialItems ON Specials.SpecialID = SpecialItems.SpecialID WHERE FinishDate < dateadd('m',-2,now)", dbFailOnError)
    Call dbSuppliers.Execute("DELETE Specials.* FROM Specials WHERE FinishDate < dateadd('m',-2,now)", dbFailOnError)
        
    Call CloseSupplierDB

Procedure_Exit:
   On Error Resume Next
Exit Sub
ErrorHandler:
   If SYSLOG(Err, "ClearOldSpecials in basUpdateSupplier in " & app.ExeName & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume Procedure_Exit
   Else
       Resume Next
   End If
    
End Sub
Public Function CleanValue(ByVal value As String, ByVal RemoveSpaces As Boolean) As String
On Error GoTo CleanValue_Error
    value = Trim(value)
    value = FindandReplace(value, vbCr, "")
    value = FindandReplace(value, vbLf, "")
    value = FindandReplace(value, vbTab, "")
    value = FindandReplace(value, """", "")
    value = FindandReplace(value, "'", "")

    If RemoveSpaces Then
        value = FindandReplace(value, " ", "")
        value = FindandReplace(value, Chr(160), "")
    End If
  
CleanValue_Exit:
   On Error Resume Next
   CleanValue = value
Exit Function

CleanValue_Error:
   If SYSLOG(Err, "CleanValue in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume CleanValue_Exit
   Else
       Resume Next
   End If
End Function
Public Sub CheckSupplierEmails()
    Dim EmailAddresses As Collection
    Dim CurrentMessage As Long
    Dim SupplierEmail As Boolean
    Dim AttachmentNo As Integer
    Dim SupplierID As Long
    Dim SupplierName As String
    Dim FileName As String
    Dim n As Long
    Dim ProcessedSuppliers As String
    Dim FailedIMport As String
    Dim Settings As SuppSettings
    Dim DeleteEmail As Boolean
    Dim dFlag As Integer
On Error GoTo CheckSupplierEmails_Error
    OpenSupplierDB
     OpenSettingsDB
    Set EmailAddresses = LoadEmailAddresses()
    If DirExists(ProcessedPath) = False Then
        Make_The_Dir (ProcessedPath)
    End If
    With frmImport.MAPIMessages1
        'Open up a MAPI session:
        frmImport.MAPISession1.NewSession = True
        frmImport.MAPISession1.UserName = "Outlook"
        frmImport.MAPISession1.SignOn
        If frmImport.MAPISession1.SessionID > 0 Then
            'Point the MAPI messages control to the open MAPI session:
            .SessionID = frmImport.MAPISession1.SessionID
            .FetchUnreadOnly = True
            'Get Messages from Inbox
            .Fetch
            'Loop through messages in Inbox
            CurrentMessage = 0
            Call SYSLOG(Nothing, "AddressesLoaded:" & EmailAddresses.Count)
            Do Until CurrentMessage >= .MsgCount
               SupplierEmail = False
               DeleteEmail = False
               SupplierID = 0
                .MsgIndex = CurrentMessage
                Call SYSLOG(Nothing, "Email Address:" & .MsgOrigAddress & " Message:" & .MsgSubject)
                For Each Settings In EmailAddresses
                    
                    If InStr(Trim(UCase(.MsgOrigAddress)), UCase(Settings.EmailAddress)) Then
                        SupplierID = Settings.SupplierID
                        SupplierName = Settings.SupplierName
                        Exit For
                    Else
                        'Call SYSLOG(Nothing, "Addresseschecked:" & Settings.EmailAddress)
                    End If
                Next Settings
                If SupplierID <> 0 Then
                    If .AttachmentPathName <> "" Then
                    End If
                    n& = 0
                    Do While n& < .AttachmentCount
                        .AttachmentIndex = n&
                        FileName = UCase(.AttachmentPathName)
                        dFlag = 1
                        If InStr(FileName, "CSV") <> 0 Or InStr(FileName, "XLS") <> 0 Or InStr(FileName, "TXT") <> 0 Then
                            DeleteEmail = False
                            dFlag = 2
                            Call SetValueInListBox(frmImport.cboSupplier, SupplierID)
                            Call frmImport.LoadSupplierSettings(SupplierID)
                            dFlag = 3
                            SupplierName = frmImport.cboSupplier.text
                            If frmImport.cboSupplier.ItemData(frmImport.cboSupplier.ListIndex) > 0 Then
                            dFlag = 4
                                If ImportFile(SupplierID, FileName, True) Then
                                    dFlag = 5
                                    SupplierEmail = True
                                    
                                    If FileExists(FileName) Then
                                       Call FileCopy(FileName, ProcessedPath & "\" & SupplierName & n & ".xls")
                                       If FileExists(ProcessedPath & "\" & SupplierName & n & ".xls") Then
                                            DeleteEmail = True
                                       End If
                                    End If
                                Else
                                    If FileExists(FileName) Then
                                       Call FileCopy(FileName, FailedPath & "\" & SupplierName & n & ".xls")
                                    End If
                                    FailedIMport = FailedIMport & vbCrLf & SupplierName
                                End If
                            End If
                        End If
                        n = n + 1
                    Loop
                    
                End If

                If SupplierEmail Then
                    dFlag = 6
                    ProcessedSuppliers = ProcessedSuppliers & "," & SupplierName
                    glbSaveAfterImport = True
                    Call SaveDetails(SupplierID, True, frmImport.cboDate.text, False)
                    If DeleteEmail Then
                        .Delete mapMessageDelete
                    Else
                        CurrentMessage = CurrentMessage + 1
                    End If
                Else
                    CurrentMessage = CurrentMessage + 1
                End If
            Loop
        
            frmImport.Refresh
            
            'Close MAPI mail session:
            frmImport.MAPISession1.SignOff
        End If
    End With
    CloseSupplierDB
    CloseSettingsDB
    MsgBox ("Imported " & ProcessedSuppliers & vbCrLf & " and " & FailedIMport & " failed to import.")
    frmImport.Show vbModal
    
CheckSupplierEmails_Exit:
   On Error Resume Next
Exit Sub

CheckSupplierEmails_Error:
   If SYSLOG(Err, "CheckSupplierEmails in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision & " DFlag:" & dFlag) Then
       Resume CheckSupplierEmails_Exit
   Else
       Resume Next
   End If
End Sub
Public Function LoadEmailAddresses() As Collection
    Dim NewCollection As Collection
    Dim snaSuppliers As Recordset
    Dim Email As String
    Dim Settings As SuppSettings
On Error GoTo LoadEmailAddresses_Error
   
    Set NewCollection = New Collection
    Set snaSuppliers = dbSettings.OpenRecordSet("SELECT SupplierID,Email FROM Settings WHERE Email <> ''", dbOpenSnapshot)
    If snaSuppliers.EOF Then
    
    Else
        Do Until snaSuppliers.EOF
            Email = FNulls(snaSuppliers("Email"))
            If Email <> "" Then
                Set Settings = New SuppSettings
                Settings.SupplierID = FNulls(snaSuppliers("SupplierID"))
                Settings.SupplierName = ""
                Settings.EmailAddress = Email
                Call NewCollection.Add(Settings)
            End If
            snaSuppliers.MoveNext
        Loop
    End If
    snaSuppliers.ClsRS
    Set snaSuppliers = Nothing
    Set LoadEmailAddresses = NewCollection
    
LoadEmailAddresses_Exit:
   On Error Resume Next
Exit Function

LoadEmailAddresses_Error:
   If SYSLOG(Err, "LoadEmailAddresses in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume LoadEmailAddresses_Exit
   Else
       Resume Next
   End If
End Function
Public Sub UpdateStockGuid()
    Dim dynDetails As Recordset
    Dim v_TheGUids As Dictionary
On Error GoTo UpdateStockGuid_Error
    Set v_TheGUids = New Dictionary
    Call OpenSupplierDB
    Set dynDetails = dbSuppliers.OpenRecordSet("SELECT Details.* FROM DETAILS INNER JOIN SUppliers on Details.SupplierID = Suppliers.SupplierID WHERE PartcodeGuid is null or StockcardGuid is null", dbOpenDynaset)
    If dynDetails.EOF Then
    
    Else
        Do Until dynDetails.EOF
            
            Dim LOTSSupplierID As Long
            If v_TheGUids.Exists("S" & FNulls(dynDetails("SupplierID"))) Then
                LOTSSupplierID = v_TheGUids("S" & FNulls(dynDetails("SupplierID")))
            Else
                LOTSSupplierID = GetLOTSSupplierID(FNulls(dynDetails("SupplierID")))
                Call v_TheGUids.Add("S" & FNulls(dynDetails("SupplierID")), LOTSSupplierID)
            End If
            If LOTSSupplierID > 0 Then
                Dim StockID As Long
                 If glbDBType = OZBro Then
                    StockID = GetStockIDFromMasterFile(FNulls(dynDetails("Pharmacode")), FNulls(dynDetails("Code")), LOTSSupplierID, FNulls(dynDetails("Description")), FNulls(dynDetails("API")), FNulls(dynDetails("Sigma")), FNulls(dynDetails("Mayne")), FNulls(dynDetails("Barcode")))
                Else
                    StockID = GetStockIDFromMasterFile(FNulls(dynDetails("Pharmacode")), FNulls(dynDetails("Code")), LOTSSupplierID, FNulls(dynDetails("Description")), "", "", "", FNulls(dynDetails("Barcode")))
                End If
                If StockID > 0 Then
                Dim snaDetails As Recordset
                Set snaDetails = dbLots.OpenRecordSet("SELECT StockGuid,PartcodeGuid FROM Stock INNER JOIN Partcode ON Stock.StockID = Partcode.StockID WHERE Stock.StockID = " & StockID & " AND SupplierID = " & LOTSSupplierID, dbOpenSnapshot)
                If snaDetails.EOF Then
                
                Else
                    dynDetails.Edit
                    dynDetails("StockcardGuid") = snaDetails("StockGuid")
                    dynDetails("PartcodeGuid") = snaDetails("PartcodeGuid")
                    dynDetails.Update
                End If
                snaDetails.ClsRS
                Set snaDetails = Nothing
                End If
            End If
            dynDetails.MoveNext
        Loop
    End If
  
UpdateStockGuid_Exit:
   On Error Resume Next
Exit Sub

UpdateStockGuid_Error:
   If SYSLOG(Err, "UpdateStockGuid in basUpdateSupplier in " & AppVersion()) Then
       Resume UpdateStockGuid_Exit
   Else
       Resume Next
   End If
        
End Sub
Public Sub UpdateMasterPOSDb(ByVal v_PLU As String, ByVal v_Partcode As String, ByVal v_SupplierID As Long, ByVal v_TradeName As String, ByVal v_Cost As Long, ByVal v_Retail As Long, ByVal v_OuterSize As Single, ByVal v_APICode As String, ByVal v_SigmaCode As String, ByVal v_FauldingCode As String, ByVal v_Barcode As String, ByVal v_FreeGST As Boolean, ByRef PartcodeGuid As String, ByRef StockcardGuid As String)
    Dim StockID As Long
    Dim dynDetails As Recordset
    Dim HasChange As Boolean
On Error GoTo UpdateMasterPOSDb_Error
    StockID = GetStockIDFromMasterFile(v_PLU, v_Partcode, v_SupplierID, v_TradeName, v_APICode, v_SigmaCode, v_FauldingCode, v_Barcode)
    If StockID = 0 Then
        Set dynDetails = dbLots.OpenRecordSet(SQL_Select & "Stock WHERE StockID =0", dbOpenDynaset)
        dynDetails.AddNew
        dynDetails("Tradename") = v_TradeName
        dynDetails("PLU") = v_PLU
        dynDetails("StockType") = CardType.pos
        dynDetails("DateCreated") = Now
        dynDetails("LastUpdated") = Now
        dynDetails("Reorder") = True
        dynDetails("POSLookup") = True
        dynDetails("PackSize") = 1
        dynDetails("StockGuid") = GetNewGuid
        If glbDBType = OZBro Then
            dynDetails("FreeToEndUser") = v_FreeGST
        Else
            dynDetails("StockGST") = getGST
        End If
        
        StockcardGuid = FNulls(dynDetails("StockGuid"))
        StockID = Val(FNulls(dynDetails("StockID")))
        dynDetails.Update
        dynDetails.ClsRS
        Set dynDetails = Nothing
    Else
        StockcardGuid = IndexedField("Stock", "StockGuid", StockID, "StockID")
        If glbDBType = OZBro Then
            Call ExecuteSafe("UPDATE Stock SET reorder = true,skDeleted=false,FreeToEndUser = " & v_FreeGST & ", LastUpdated = getdate() WHERE StockID = " & StockID & " and (reorder = true or skdeleted =true or FreeToEndUser <> " & v_FreeGST & ")")
        Else
            Call ExecuteSafe("UPDATE Stock SET reorder = true,skDeleted=false,StockGST=" & getGST & ", LastUpdated = getdate()  WHERE StockID = " & StockID & " and (reorder = true or skdeleted =true or StockGST<>" & getGST & ")")
        End If
    End If
    Set dynDetails = dbLots.OpenRecordSet(SQL_Select & "Partcode WHERE StockID =" & StockID & " AND SupplierID = " & v_SupplierID, dbOpenDynaset)
    If dynDetails.EOF Then
        dynDetails.AddNew
        dynDetails("PartcodeGuid") = GetNewGuid
        dynDetails("PartcodeDateModified") = Now
        HasChange = True
    Else
        dynDetails.Edit
        If FNulls(dynDetails("Partcode")) <> v_Partcode Or FNulls(dynDetails("pcOuterSize")) <> v_OuterSize Or FNulls(dynDetails("pcRRP")) <> v_Retail Or FNulls(dynDetails("SupplierListCost")) <> v_Cost Then
            dynDetails("PartcodeDateModified") = Now
            HasChange = True
        End If
    End If
    PartcodeGuid = FNulls(dynDetails("PartcodeGuid"))
    dynDetails("Partcode") = v_Partcode
    dynDetails("pcDeleted") = False
    dynDetails("pcOuterSize") = v_OuterSize
    dynDetails("pcRRP") = v_Retail
    dynDetails("SupplierListCost") = v_Cost
    dynDetails("StockID") = StockID
    dynDetails("SupplierID") = v_SupplierID
    dynDetails.Update
    dynDetails.ClsRS
    Set dynDetails = Nothing
    If v_Barcode <> "" Then
        Set dynDetails = dbLots.OpenRecordSet(SQL_Select & "Barcode WHERE barcode = " & WrapText(v_Barcode), dbOpenDynaset)
        If dynDetails.EOF Then
            dynDetails.AddNew
            dynDetails("BarcodeGuid") = GetNewGuid
            dynDetails("BarcodeDateModified") = Now
            HasChange = True
        Else
            dynDetails.Edit
            If FNulls(dynDetails("StockID")) <> StockID Then
                HasChange = True
                dynDetails("BarcodeDateModified") = Now
            End If
        End If
        dynDetails("Barcode") = v_Barcode
        dynDetails("StockID") = StockID
        dynDetails.Update
        dynDetails.ClsRS
        Set dynDetails = Nothing
    End If
    If HasChange Then
        Call ExecuteSafe("UPDATE Stock SET LastUpdated = getdate() WHERE StockID = " & StockID)
    End If
UpdateMasterPOSDb_Exit:
   On Error Resume Next
Exit Sub

UpdateMasterPOSDb_Error:
   If SYSLOG(Err, "UpdateMasterPOSDb in basUpdateSupplier in " & AppVersion()) Then
       Resume UpdateMasterPOSDb_Exit
   Else
       Resume Next
   End If
End Sub
Public Sub UpdateLastUpdatedInDatabase()
' Call ExecuteSafe("update supplier set supdatemodified = M.TheDate from( select max(PartcodeDateModified) As TheDate,Partcode.SupplierID from supplier inner join partcode on supplier.SupplierID = partcode.SupplierID group by partcode.supplierid)M inner join supplier on M.SupplierID = supplier.SupplierID ")
' Call ExecuteSafe("update Stock set LastUpdated = M.TheDate from( select max(PartcodeDateModified) As TheDate,Partcode.StockID from  partcode  group by partcode.StockID)M inner join Stock on M.StockID = Stock.StockID ")
End Sub
Public Function GetIDFromPartcodeTable(ByVal SupplierID As Long, ByVal Partcode As String) As Long
On Error GoTo GetIDFromPartcodeTable_Error
    If Partcode <> "" Then
        Dim snaDetails As Recordset
        Set snaDetails = dbLots.OpenRecordSet("SELECT StockID FROM Partcode WHERE SupplierID =" & SupplierID & " AND PArtcode =" & WrapText(Partcode), dbOpenSnapshot)
        If snaDetails.EOF Then
        
        Else
            GetIDFromPartcodeTable = Val(FNulls(snaDetails("StockID")))
        End If
        snaDetails.ClsRS
        Set snaDetails = Nothing
    End If
  
GetIDFromPartcodeTable_Exit:
   On Error Resume Next
Exit Function

GetIDFromPartcodeTable_Error:
   If SYSLOG(Err, "GetIDFromPartcodeTable in basUpdateSupplier in " & AppVersion()) Then
       Resume GetIDFromPartcodeTable_Exit
   Else
       Resume Next
   End If
End Function
Public Function GetStockIDFromMasterFile(ByVal v_PLU As String, ByVal v_Partcode As String, ByVal v_SupplierID As Long, ByVal v_TradeName As String, ByVal v_APICode As String, ByVal v_SigmaCode As String, ByVal v_FauldingCode As String, ByVal v_Barcode As String) As Long
    Dim StockID As Long
    
On Error GoTo GetStockIDFromMasterFile_Error
    If glbDBType = OZBro Then
        StockID = GetIDFromPartcodeTable(API_SUPPLIERID, v_APICode)
        If StockID = 0 Then
            StockID = GetIDFromPartcodeTable(FAULDING_SUPPLIERID, v_FauldingCode)
        End If
        If StockID = 0 Then
            StockID = GetIDFromPartcodeTable(SIGMA_SUPPLIERID, v_SigmaCode)
        End If
    Else
        If v_PLU <> "" Then
            StockID = Val(GetAnyField("Stock", "StockID", WrapText(v_PLU), "PLU"))
            
        End If
    End If
    If StockID = 0 Then
        If v_Barcode <> "" Then
            StockID = Val(GetAnyField("barcode", "StockID", WrapText(v_Barcode), "Barcode"))
            
        End If
    End If
    If StockID = 0 Then
        If v_TradeName <> "" Then
            v_TradeName = Replace(v_TradeName, Chr(34), "")
            StockID = Val(GetAnyField("Stock", "StockID", WrapText(v_TradeName), "Tradename"))
        End If
    End If
    If StockID = 0 Then
        If v_Partcode <> "" Then
            StockID = Val(GetIDFromPartcodeTable(v_SupplierID, v_Partcode))
        End If
    End If
    
  
GetStockIDFromMasterFile_Exit:
   On Error Resume Next
   GetStockIDFromMasterFile = StockID
Exit Function

GetStockIDFromMasterFile_Error:
   If SYSLOG(Err, "GetStockIDFromMasterFile in basUpdateSupplier in " & AppVersion()) Then
       Resume GetStockIDFromMasterFile_Exit
   Else
       Resume Next
   End If
End Function
Public Function GetSupplierSettingsDatabaseName() As String
   GetSupplierSettingsDatabaseName = "NZRxOneSupplierSettings"
End Function
Sub OpenSettingsDB(Optional ByRef CloseDB As Boolean)
On Error GoTo ErrTrap

    Dim ws As Workspace
    Dim DBLocation As String
    Dim TempMasterPath As String 'Greg 31/8/2004
    If glbSettingsDBOpen = True Then
        CloseDB = False
    Else
        Dim strSvr As String
        Dim strCon As String
        Dim strPing As String
        Dim strInst As String
        Dim strConStr As String
        Call LoadSQLSettingsFromINI("POS SQL Server", GetDBINIFileName, strSvr, strCon, strPing, strInst, "")

        Call UpdateConnectionString(strCon, True)
        strConStr = MakeConString(strCon, strSvr, strInst, GetSupplierSettingsDatabaseName)
    
        Set dbSettings = OpenDatabase(strConStr, "timeout=3")
        'Change Greg 31/7/06 Add in db open code
        CloseDB = True
        glbSettingsDBOpen = True
    End If
'    Set ws = Workspaces(0)
'
'    If InStr(UCase(Command), "/DRUG_UPDATE") > 0 Then
'        'Change greg 6/9/06 use the dblocation first to try and find the database
'
'        DBLocation = readini("POS Prices", "SettingsDBPath", "", GetPOSPricesPath)
'
'        If FileExists(DBLocation) And Trim(DBLocation) <> "" Then
'                'DB Found
'
'        Else
'            If FileExists("C:\UpdateNZ\Pos Prices\Settings.mdb") Then
'                DBLocation = "C:\UpdateNZ\Pos Prices\Settings.mdb"
'
'            Else
''            DBLocation = ReadIni("POS Prices", "DBPath", "", App.Path & "\POSPrices.ini")
''            If FileExists(DBLocation) And Trim(DBLocation) <> "" Then
''                'DB Found
''            Else
'                MsgBox "POS Settings database not found", vbCritical
'                End
'            End If
'        End If
'    Else
'        'Change Greg 31/8/2004 ALLOW access across the network
'        '15/2/2006 Limin should change to another way
'        'Change Greg 16/6/06 Add in new master lots code
'        TempMasterPath$ = GetMasterLotsDir
''        If IsThisTheMaster Then
''            TempMasterPath$ = GetLocalLOTSDir
''        Else
''
''            TempMasterPath$ = GetMasterPathToRoot & "\LOTS"
''        End If
'        If glbHOPriceUpdate Then
'            'Use DB Name...
'            If FileExists(TempMasterPath$ & "\Update\HOSettings.mdb") Then
'                DBLocation = TempMasterPath$ & "\Update\HOSettings.mdb"
''            If FileExists(GetLocalLOTSDir & "\Update\HOSettingss.mdb") Then
''                DBLocation = GetLocalLOTSDir & "\Update\HOSettingss.mdb"
'            Else
'                Call SYSLOG(Err, "Database '\Update\HOSettingss.mdb' not found in OpenSettingsDB in " & app.Path)
'                Exit Sub
'            End If
'        Else
'            If FileExists(TempMasterPath$ & "\Update\Settings.mdb") Then
'                'Use Default DB...
'                DBLocation = TempMasterPath$ & "\Update\Settings.mdb"
''            If FileExists(GetLocalLOTSDir & "\Update\Settingss.mdb") Then
''                'Use Default DB...
''                DBLocation = GetLocalLOTSDir & "\Update\Settingss.mdb"
'            Else
'                Call SYSLOG(Err, "Database not found in OpenSettingsDB in " & app.Path)
'                Exit Sub
'            End If
'        End If
'    End If
'    If glbSettingsDBOpen = True Then
'        CloseDB = False
'    Else
'        Set dbSettings = ws.OpenDatabase(DBLocation)
'        'Change Greg 31/7/06 Add in db open code
'        CloseDB = True
'        glbSettingsDBOpen = True
'    End If
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "OpenSettingsDB in " & app.ExeName)
End Sub

Sub CloseSettingsDB(Optional ByVal CloseDB As Boolean = True)
    'Change Greg 31/7/06 Add in db open code
On Error GoTo CloseSettingsDB_Error
    If CloseDB Then
        If dbSettings Is Nothing = False Then
            dbSettings.ClsDB
            Set dbSettings = Nothing
        End If
        glbSettingsDBOpen = False
    End If
  
CloseSettingsDB_Exit:
   On Error Resume Next
Exit Sub

CloseSettingsDB_Error:
   If SYSLOG(Err, "CloseSettingsDB in basUpdateSupplier in " & app.ExeName & " " & Right(app.Major, 2) & " " & app.Minor & " " & app.Revision) Then
       Resume CloseSettingsDB_Exit
   Else
       Resume Next
   End If
End Sub
Sub ConvertDelimitedStringToArray(DelimitedString As String, Delimiter As String, OutputArray() As String)
    On Error GoTo ErrTrap
    ' breaks up a delimited string based at every incident of 'Delimiter' and returns the results
    ' in OutputArray.
    If InStr(DelimitedString, Delimiter & Chr(34)) > 0 Then
        Dim TempArray() As String
        Dim TheWord As Variant
        Dim InsideQuote As Boolean
        Dim FirstEntry As Boolean
        TempArray = Split(DelimitedString$, Delimiter)
        ReDim OutputArray(0 To 0)
        FirstEntry = True
        For Each TheWord In TempArray
            
            If TheWord = "" And InsideQuote = False Then
                ReDim Preserve OutputArray(0 To UBound(OutputArray) + 1)
                
            Else
                If FirstEntry Then
                    OutputArray(UBound(OutputArray)) = TheWord
                Else
                    If InsideQuote Then
                        If Right(TheWord, 1) = Chr(34) Then
                            InsideQuote = False
                        End If
                        OutputArray(UBound(OutputArray)) = OutputArray(UBound(OutputArray)) & Delimiter & TheWord
                    Else
                        ReDim Preserve OutputArray(0 To UBound(OutputArray) + 1)
                        OutputArray(UBound(OutputArray)) = TheWord
                        If LEFT(TheWord, 1) = Chr(34) And Right(TheWord, 1) <> Chr(34) Then
                            InsideQuote = True
                        End If
                    End If
                End If
                
            End If
            OutputArray(UBound(OutputArray)) = Replace(OutputArray(UBound(OutputArray)), Chr(34), "")
            FirstEntry = False
        Next TheWord
    Else
        OutputArray = Split(DelimitedString$, Delimiter)
    End If
Exit Sub
ErrTrap:
    Call SYSLOG(Err, "ConvertDelimitedStringToArray")
    Resume Next
End Sub
Public Function PathName() As String
    
    
    
    PathName = app.Path
    If IsProgrammersMachine Then
    
    End If
End Function
