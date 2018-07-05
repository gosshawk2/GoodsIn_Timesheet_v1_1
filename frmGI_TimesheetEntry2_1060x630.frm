VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGI_TimesheetEntry2_1060x630 
   Caption         =   "GOODSIN TIMESHEET RP"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19905
   OleObjectBlob   =   "frmGI_TimesheetEntry2_1060x630.frx":0000
End
Attribute VB_Name = "frmGI_TimesheetEntry2_1060x630"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '64bit
''Goods In Timesheet Main Program - written by Daniel Goss 2018 version RP7.32

Private Sub btnAddData_Click()
    'Loop through controls and grab the TAG property value. This sets the column number for the textbox contents to be inserted:
    
    
End Sub

Function GetNextAvailablerows(TartgetWB As Workbook, WorksheetName As String, StartRow As Long, CheckCol As Long) As Long
Dim IDX As Long
Dim LastRow As Long
Dim sht As Worksheet
Dim BlankRow As Long

Set sht = TartgetWB.Sheets(WorksheetName)
GetNextAvailablerows = StartRow
IDX = StartRow
LastRow = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
'BlankRow = sht.Cells.Find(" ", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

'Me.txtActualCages.Top = 0.9
GetNextAvailablerows = LastRow + 1

End Function

Private Sub btnAddOperative_Click()
    'Dynamically Add Timer Controls:
    'Add one to the index for each type of control:
    'txtBox_Index = 5 (1+4)
    'combo_Index = 3 (1+2)
    'cmd_Index = 3 (1+2)
    Dim strDeliveryRef As String
    Dim strDeliveryDate As String
    Dim ASN As String
    Dim SearchCriteria As String
    Dim ControlDBTable As String
    Dim JustTAG As Boolean
    Dim FieldsTableLoadedOK As Boolean
    Dim ControlFieldsTable As String
    Dim allLookupFields As Variant
    Dim RowIDX As Long
    Dim Fieldnames As String
    Dim SortFields As String
    Dim Reversed As Boolean
    Dim TotalRows As Long
    
    ControlFieldsTable = "tblFieldsAndTAGS"
    ControlDBTable = "tblOperatives"
    
    strDeliveryDate = Me.txtDeliveryDate.Text
    strDeliveryRef = Me.txtDeliveryRef.Text
    ASN = Me.txtASNNum.Text
    
    Reversed = False
    SearchCriteria = "TableName = " & Chr(34) & ControlDBTable & Chr(34)
    SortFields = "TABLENAME,TAGID"
    JustTAG = True
    FieldsTableLoadedOK = LoadAccessDBTable(ControlFieldsTable, AccessDBpath, JustTAG, SearchCriteria, SortFields, Reversed, _
            False, "", "", Nothing, allLookupFields)
            
    RowIDX = 11
    Do While RowIDX <= 14
        If Len(Fieldnames) = 0 Then
            Fieldnames = allLookupFields(2, RowIDX)
        Else
            Fieldnames = Fieldnames & "," & allLookupFields(2, RowIDX)
        End If
        RowIDX = RowIDX + 1
    Loop
    TotalRows = OperativeCount - 1
    Call AddNewOperatives(OperativeCount, TextTAGID, strDeliveryDate, strDeliveryRef, ASN, 400, Fieldnames, TotalRows)
    
    Me.txtTotalOperatives.Text = CStr(OperativeCount)
    
    'MsgBox ("OpID: " & CStr(OperativeCount) & " ,Text TAGID:" & CStr(TextTAGID) & " ,Time TAGID:" & CStr(TimeTAGID) & " ,btn TAGID:" & CStr(btnTAGID))
    'MsgBox ("Command Button Next Index:" & CStr(cmd_Index) & " textbox:" & CStr(txtBox_Index) & " combo:" & CStr(combo_Index))
    
    
    
End Sub

Private Sub CommandButton2_Click()

'Unload UserForm1

End Sub

Private Sub btnArrivedOnTime_Click()
    'BUTTON: Arrived On Time ?
    Me.txtArrivedONTime.Font.Name = "Tahoma"
    If UCase(Me.txtArrivedONTime.Text) = "YES" Then
        Me.txtArrivedONTime.Text = "NO"
        Me.txtArrivedONTime.Tag = "0"
        Me.txtArrivedONTimeComment.Visible = True
        Me.txtArrivedONTimeComment.Tag = ComplianceQuestion1TAG
    Else
        Me.txtArrivedONTime.Text = "YES"
        Me.txtArrivedONTime.Tag = ComplianceQuestion1TAG
        Me.txtArrivedONTimeComment.Tag = "0"
        Me.txtArrivedONTimeComment.Visible = False
    End If
End Sub

Private Sub btnCalcHours_Click()
    'Calculate Hours and number of Operatives used:
    'Will have to search and collect the information off the Timesheet Record sheet per each Delivery DATE and per Delivery REF:
    'OR PER FLM in charge of EACH Delivery on that shift ?
    'GetTimes()
    'CAlc_Hours()
    Dim strStartDate As String
    Dim strEndDate As String
    Dim dtStartDate As Date
    Dim dtEndDate As Date
    Dim TotalHours As Double
    Dim OperativeHours() As Double
    Dim dblMins As Double
    Dim StartTimeCol As Long
    Dim EndTimeCol As Long
    Dim AllHours() As String
    Dim OPIDX As Long
    Dim OpName As String
    Dim WholeTime() As String
    Dim strHours As String
    Dim strMins As String
    
    ReDim OperativeHours(1)
    
    'StartTimeCol = get from TAG of relevant TEXT BOX
    StartTimeCol = 18
    EndTimeCol = 98
    TotalHours = 0
    ReDim AllHours(1)
    If TotalFields = 0 Then
        MsgBox ("Problem getting total columns")
        Exit Sub
    End If
    'AllHours = GetTimes_From_Access(WB_MainTimesheetData, MainWorksheet, CurrentRecordRow, TotalFields, StartTimeCol, EndTimeCol)
    OPIDX = 0
    If UBound(AllHours) > 0 Then
        Do While OPIDX <= UBound(AllHours)
            
            WholeTime = Split(AllHours(OPIDX), ",")
            If Not IsArrayEmpty(WholeTime) Then
                OpName = WholeTime(0)
                strStartDate = WholeTime(1)
                strEndDate = WholeTime(2)
            
                TotalHours = TotalHours + Calc_Hours(strStartDate, strEndDate, strHours, strMins)
            End If
    
        'dtStartDate = Convert_strDateToDate(strStartDate, True)
        'dtEndDate = Convert_strDateToDate(strEndDate, True)
            OPIDX = OPIDX + 1
        Loop
    End If
    Me.txtTotalHours.Text = CStr(TotalHours)
    
End Sub

Private Sub btnCompleted_Click()
    'BUTTON: Completed ?
    Me.txtCompleted.Font.Name = "Tahoma"
    If UCase(Me.txtCompleted.Text) = "YES" Then
        Me.txtCompleted.Text = "NO"
        Me.txtCompleted.Tag = "0"
        Me.txtCompletedComment.Visible = True
        Me.txtCompletedComment.Tag = ComplianceQuestion3TAG
    Else
        Me.txtCompleted.Text = "YES"
        Me.txtCompleted.Tag = ComplianceQuestion3TAG
        Me.txtCompletedComment.Tag = "0"
        Me.txtCompletedComment.Visible = False
    End If
End Sub


Private Sub btnDeleteOperative_Click()
    'DELETE OPERATIVES - including - reduction of array indexes:
    If OperativeCount > 0 Then
        Call DeleteControls(frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, OperativeCount - 1, 0)
        OperativeCount = OperativeCount - 1
    End If
    Me.txtTotalOperatives.Text = CStr(OperativeCount)
End Sub

Private Sub btnDeleteShort_Click()
    Call RemoveAllControls(frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls)
    
End Sub

Private Sub btnSafe_Click()
    'BUTTON: Is It Safe ?
    Me.txtIsItSafe.Font.Name = "Tahoma"
    If UCase(Me.txtIsItSafe.Text) = "YES" Then
        Me.txtIsItSafe.Text = "NO"
        Me.txtIsItSafe.Tag = "0"
        Me.txtIsItSafeComment.Tag = ComplianceQuestion2TAG
        Me.txtIsItSafeComment.Visible = True
    Else
        Me.txtIsItSafe.Text = "YES"
        Me.txtIsItSafe.Tag = ComplianceQuestion2TAG
        Me.txtIsItSafeComment.Tag = "0"
        Me.txtIsItSafeComment.Visible = False
    End If
End Sub

Private Sub btnSelectRecordSheetLocation_Click()
    Dim ConnectWB As Workbook
    

    'RemoteFilePath = Connect_RecordSheet("Prefs", False, ConnectWB, False)
    AccessDBpath = Connect_ACCESS_DB("Prefs")
        
End Sub


Private Sub btnUpdateEmployees_Click()
    Dim Employees() As String
    Dim XMLFilename As String
    Dim strFilter As String
    Dim strCaption As String
    
    
    strFilter = "XML Files (*.xml),*.xml,CSV Files (*.csv),*.csv,Excel Files (*.xlsx),*.xlsx,All Files (*.*),*.*"
    strCaption = "Please Select the data file "
    XMLFilename = Application.GetOpenFilename(strFilter, 1, strCaption)

    If Len(XMLFilename) = 0 Then Exit Sub
    
    Employees = ReadXML(XMLFilename, 12, 5)
    MsgBox ("UPDATED EMPLOYEES")



End Sub

Private Sub comASNNo_Change()
    Dim SearchCriteria As String
    Dim strDeliveryRef As String
    Dim strDeliveryDate As String
    Dim strASNNumber As String
    
    Call PerformComboSearch(SearchCriteria, strDeliveryDate, strDeliveryRef, strASNNumber)
    Me.txtSelectedDeliveryDate.Text = strDeliveryDate
End Sub


Private Sub comDeliveryRef_Change()
    'runs immediately after 3 chars are typed etc.
    Dim SearchCriteria As String
    Dim strDeliveryRef As String
    Dim strDeliveryDate As String
    Dim strASNNumber As String
    
    If Len(comDeliveryRef) > 3 Then
        Call PerformComboSearch(SearchCriteria, strDeliveryDate, strDeliveryRef, strASNNumber)
        Me.txtSelectedDeliveryDate.Text = strDeliveryDate
    End If
End Sub

Private Sub comFLMs_Change()
    'test for employee number and find in Employees:
    Dim Employees() As Variant
    Dim DBTable As String
    Dim SearchCriteria As String
    Dim FieldPos
    
    DBTable = "tblEmployees"
    
    
    
End Sub

Private Sub CommandButton3_Click()
'Unload UserForm1
'UserForm1.Show
End Sub


Private Sub btnAddInboundData_Click()
    'Add INBOUND DATA ONLY to the Timesheet Record sheet:
    Dim ListArr() As String
    Dim IDX As Long
    Dim DontSave As Boolean
    
    txtDeliveryDate.BackColor = vbWhite
    txtDeliveryRef.BackColor = vbWhite
    Me.comSuppliers.BackColor = vbWhite
    Me.txtASNNum.BackColor = vbWhite
    Me.txtPalletsDue.BackColor = vbWhite
    
    DontSave = False
    If Len(Me.txtDeliveryDate.Text) = 0 Then
        txtDeliveryDate.BackColor = vbRed
        DontSave = True
    End If
    If Len(Me.txtDeliveryRef.Text) = 0 Then
        txtDeliveryRef.BackColor = vbRed
        DontSave = True
    End If
    If Len(Me.comSuppliers.Text) = 0 Then
        Me.comSuppliers.BackColor = vbRed
        DontSave = True
    End If
    If Len(Me.txtASNNum.Text) = 0 Then
        Me.txtASNNum.BackColor = vbRed
        DontSave = True
    End If
    If Len(Me.txtPalletsDue.Text) = 0 Then
        Me.txtPalletsDue.BackColor = vbRed
        DontSave = True
    End If
    
    If DontSave = False Then
        Me.txtDeliveryDate.Text = Format(DeliveryDate, "dd/mm/yyyy")
        
        'WB_MainTimesheetData does not seem to get populated correctly ?
        Call InsertEntry(WB_MainTimesheetData, "", 0, 1, 16) '0 inserted to create a NEW RECORD.
        
        Me.comASNNo.Clear
        ListArr = PopulateDropdowns("Timesheet Records", 4, 0, True, WB_MainTimesheetData)
        For IDX = 0 To UBound(ListArr)
            If Len(ListArr(IDX)) > 0 Then
                Me.comASNNo.AddItem (ListArr(IDX))
            End If
        Next
        Call ClearEntry(1, TotalFields)
        Call ClearEntry(200, 298)
        Call EnableDisableControls(False, 17, TotalFields, "COMBOBOX")
        Call EnableDisableControls(False, 17, TotalFields, "TEXTBOX")
        Call EnableDisableControls(False, 8, 61, "BUTTONS")
        
        CurrentRecordRow = 0
    Else
        MsgBox ("Please enter values for the items shown in red")
        Exit Sub
    End If
End Sub

Private Sub btnAddOperationData_Click()
    'ADD OPERATION DATA to EXISTING record ONLY to the Timesheet Record sheet - do not add new record:
    'SAVE AND CONTINUE - do not erase all the fields. do NOT disable the controls - To ALLOW USER to Continue with Entry into same ASN but SAVES.
    Dim ListArr() As String
    Dim IDX As Long
    Dim EntryError As Boolean
    Dim ErrCol As Long
    Dim LowerTAGID As Long
    Dim UpperTAGID As Long
    Dim DBTable As String
    Dim AccessPath As String
    Dim DeliveryDate As String
    Dim DeliveryRef As String
    Dim TotalOpRows As Long
    Dim ReturnAnything As Boolean
    Dim ReturnUpperTag As String
    Dim ReturnLowerTag As String
    Dim ErrMessage As String
    Dim TotalFrameRows As Long
    Dim SavedOK As Boolean
    
    SaveAndContinue = True
    EntryError = False
    'If CurrentRecordRow > 0 Then
        'If CheckFinishTimeError(WB_MainTimesheetData, "Timesheet Records", CurrentRecordRow, 0, ErrCol) Then
        '    MsgBox ("Error: Finish Time is LESS than Start Time, Last Col = " & CStr(ErrCol))
        '    EntryError = True
        'End If
    If CheckComplianceQuestionError(Me.txtArrivedONTime, Me.txtArrivedONTimeComment) Then
        MsgBox ("Error: Nothing has been entered for the Compliance Question: Arrived ON Time")
        EntryError = True
    End If
    If CheckComplianceQuestionError(Me.txtIsItSafe, Me.txtIsItSafeComment) Then
        MsgBox ("Error: Nothing has been entered for the Compliance Question: Is It Safe ?")
        EntryError = True
    End If
    If CheckComplianceQuestionError(Me.txtCompleted, Me.txtCompletedComment) Then
        MsgBox ("Error: Nothing has been entered for the Compliance Question: Completed ?")
        EntryError = True
    End If
            
        
    If EntryError = False Then
        Me.txtLastSaved.Text = Format(Now(), "dd/mm/yyyy hh:nn:ss")
        'Call InsertEntry(WB_MainTimesheetData, "", CurrentRecordRow, 13, TotalFields, 82) '200 is a selection combos to search for the record only.
        AccessPath = AccessDBpath
        DBTable = "tblDeliveryInfo"
        DeliveryDate = Me.txtDeliveryDate.Text
        DeliveryRef = Me.txtDeliveryRef.Text
        'SAVE all normal entries first that go into the tblDeliveryInfo:
        'OBJECTIVE - TO ONLY CYCLE THROUGH ALL OF THE FORM CONTROLS ONCE AND GATHER ALL CONTROL INFO FOR ALL 4 TABLES !
        
        SavedOK = SaveAllControls(DeliveryDate, DeliveryRef)
        
        'Call InsertEntry_Into_ACCESS(DBTable, AccessPath, DeliveryDate, DeliveryRef, "", 1, 28)
        'SAVE all OPERATIVES to ACCESS DB - including the FLM entry:
        DBTable = "tblLabourHours"
        'Need UpperTAGID = LowerTAG+2+(TotalRows*4)
        ReturnAnything = False
        'ReturnAnything = ReturnControlInfo(DeliveryDate, DeliveryRef, "", "1", "", "TAG", "", "EndTag", ReturnUpperTag, ErrMessage)
        LowerTAGID = 40
        'TotalFrameRows = GetTotalFrameRows(frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, 43, UpperTAGID, 4, 400) 'depends on how many operative rows have been selected / entered
        'Call InsertEntry_Into_ACCESS(DBTable, AccessPath, DeliveryDate, DeliveryRef, "", LowerTAGID, UpperTAGID)
        'ReturnAnything = False
        
        'TEST if ANY SHORT or EXTRA parts have been entered:
        DBTable = "tblShortsAndExtraParts"
        'TotalFrameRows = GetTotalFrameRows(frmGI_TimesheetEntry2_1060x630.Frame_ShortParts.Controls, 1001, UpperTAGID, 2, 1000)
        LowerTAGID = 1001
        'UpperTAGID = 1002 'depends on how many SHORTS rows have been selected / entered
        'Call InsertEntry_Into_ACCESS(DBTable, AccessPath, DeliveryDate, DeliveryRef, "", LowerTAGID, UpperTAGID)
        'SAVE Supplier Compliance entries:
        DBTable = "tblSupplierCompliance"
        LowerTAGID = 801
        UpperTAGID = 803 'depends on how many Supplier Compliance entries there are:
        'Call InsertEntry_Into_ACCESS(DBTable, AccessPath, DeliveryDate, DeliveryRef, "", LowerTAGID, UpperTAGID)
        
        'THIS IS ONLY REQUIRED IF NEW DELIVERY REFERENCES WERE ADDED DURING SAVE :
        Me.comDeliveryRef.Clear
        'ListArr = PopulateDropdowns("Timesheet Records", 2) 'refresh the ASN dropdown - read in all the ACCESS RECORDS
        DBTable = "tblDeliveryInfo"
        ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessPath, 2, "", "", True)
        If Not IsArrayEmpty(ListArr) Then
            For IDX = 0 To UBound(ListArr)
                If Len(ListArr(IDX)) > 0 Then
                    Me.comDeliveryRef.AddItem (ListArr(IDX))
                End If
            Next
        End If
        Me.comASNNo.Clear
        
        'ListArr = PopulateDropdowns("Timesheet Records", 4) 'refresh the DeliveryRef = from ACCESS
        ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessPath, 4, "", "", True)
        If Not IsArrayEmpty(ListArr) Then
            For IDX = 0 To UBound(ListArr)
                If Len(ListArr(IDX)) > 0 Then
                    Me.comASNNo.AddItem (ListArr(IDX))
                End If
            Next
        End If
        UpperTAGID = 42
        Me.txtLastSaved.Text = Format(Now(), "dd/mm/yyyy hh:nn:ss")
        If SaveAndContinue = False Then
            Call ClearEntry(1, UpperTAGID)
            Call ClearEntry(300, 399) 'TIME Value Boxes
            Call EnableDisableControls(False, 17, UpperTAGID, "COMBOBOX")
            Call EnableDisableControls(False, 17, UpperTAGID, "TEXTBOX")
            Call EnableDisableControls(False, 8, 61, "BUTTONS")
            Me.comASNNo.Clear
            Me.comDeliveryRef.Clear 'do not do otherwise
            CurrentRecordRow = 0
        End If
    End If
    'Else
        'MsgBox ("Cannot get current row.")
    'End If
    
End Sub

Private Sub btnAddSplitterControl_Click()
    'ADD NEW CONTROL TO BOTTOM OF EXISTING CONTROLS:
    'Need to know what TAG number to give it?
    'Need to know the placement - what was the previous HEIGHT of the last control ? - then ADD 18 (height of controls)
    
End Sub

Private Sub btnBaggingRequire_Click()
    'BUTTON: Bagging Require:
    If UCase(Me.txtCBYesNoBagging.Text) = "YES" Then
        Me.txtCBYesNoBagging.Text = "NO"
    Else
        Me.txtCBYesNoBagging.Text = "YES"
    End If
End Sub

Private Sub btnCollar_Click()
    'BUTTON: Collar:
    If UCase(Me.txtCBYesNoCollar.Text) = "YES" Then
        Me.txtCBYesNoCollar.Text = "NO"
    Else
        Me.txtCBYesNoCollar.Text = "YES"
    End If
End Sub

Private Sub btnExit_Click()
    Unload frmGI_TimesheetEntry2_1060x630
    
    'Application.Quit - closes ALL EXCEL instances
    ActiveWorkbook.Close savechanges:=True
    
End Sub

Private Sub btnFLMFinish_Click()
    Dim SearchControl As Control
    Dim SearchControlName As String
    Dim TimeOut As Date
    Dim TAGNumber As String
    Dim CBTimeControl As Control
    Dim CBTimeControlName As String
    
    CBTimeControlName = "txtCBFLMNameFinish"
    Set CBTimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", CBTimeControlName)
    SearchControlName = "txtFLMFinishTime"
    Set SearchControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", SearchControlName)
    If SearchControl Is Nothing Then
        'could not find control
        MsgBox ("Could not find Time Control")
    End If
    If CBTimeControl.Text = Chr(82) Then
        'Time already in the textbox so remove it.
        CBTimeControl.Text = " "
        SearchControl.Text = " "
    Else
        If Not SearchControl Is Nothing Then
            TAGNumber = SearchControl.Tag
            TimeOut = Now()
            Call InsertTimeIntoControl(0, SearchControl.Name, TimeOut)
            'Need to update the Control Collective:
            
            CBTimeControl.Text = Chr(82)
        End If
    End If
End Sub

Private Sub btnFLMStart_Click()
    'BUTTON: FLM start:
    'Find Row to insert into: CurrentRecordRow should be set when SEACH is clicked
    'Call InsertTime(WB_MainTimesheetData, txtCBFLMNameStart, txtFLMStartTime, CurrentRecordRow, CLng(txtCBFLMNameStart.Tag), CLng(txtCBFLMNameStart.Tag), 82)
    Dim SearchControl As Control
    Dim SearchControlName As String
    Dim TimeOut As Date
    Dim TAGNumber As String
    Dim CBTimeControl As Control
    Dim CBTimeControlName As String
    
    CBTimeControlName = "txtCBFLMNameStart"
    Set CBTimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", CBTimeControlName)
    SearchControlName = "txtFLMStartTime"
    Set SearchControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", SearchControlName)
    If SearchControl Is Nothing Then
        'could not find control
        MsgBox ("Could not find Time Control")
    End If
    If CBTimeControl.Text = Chr(82) Then
        'Time already in the textbox so remove it.
        CBTimeControl.Text = " "
        SearchControl.Text = " "
    Else
        If Not SearchControl Is Nothing Then
            TAGNumber = SearchControl.Tag
            TimeOut = Now()
            Call InsertTimeIntoControl(0, SearchControl.Name, TimeOut)
            CBTimeControl.Text = Chr(82)
        End If
    End If
    
End Sub

Private Sub btnImportData_Click()
    'Allow user to BROWSE for the Goods Inwards data sheet and insert into current workbook:
    Call Import_Into_ACCESS
End Sub

Sub Import_Into_ACCESS()
    Dim Newsheet As String
    Dim ExtractedRows() As String
    Dim DateCol As Long
    Dim ASNCol As Long
    Dim DelRefCol As Long
    Dim SupplierCol As Long
    Dim SupplierCodeCol As Long
    Dim ExpectedCasesCol As Long
    Dim ExpectedLinesCol As Long
    Dim EstimatedPalletsCol As Long
    Dim EstimatedCagesCol As Long
    Dim EstimatedTotesCol As Long
    Dim CartonsDueCol As Long
    Dim PalletsDueCol As Long
    Dim REadyLabelCol As Long
    Dim ManHoursCol As Long
    Dim OriginCol As Long
    Dim DueTimeCol As Long
    Dim SHIFTCol As Long
    Dim ActualCasesCol As Long
    
    Dim DelRefText As String
    Dim SearchText As String
    Dim ColSearch As String
    Dim ArrIDX As Long
    Dim TotalRows As Long
    Dim WholeRow As String
    Dim TitleRow As String
    Dim TitleRow2 As String
    Dim ColIDX As Long
    Dim tempArr() As String
    Dim TotalCols As Long
    Dim DateEntry As String
    Dim Entry As String
    Dim ASNEntry As String
    Dim GIStartRow As Long
    Dim MainSheetStartRow As Long
    Dim dtImportDate As Date
    Dim ListArr() As String
    Dim TotalRowsExtracted As Long
    Dim DuplicateRows As Long
    Dim Message As String
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim strDeliveryDate As String
    Dim strDeliveryRef As String
    Dim strSupplier As String
    Dim strASNNumber As String
    Dim strExpectedCases As String
    Dim strExpectedLines As String
    Dim strEstimatedPallets As String
    Dim strEstimatedCages As String
    Dim strEstimatedTotes As String
    Dim strCartonsDue As String
    Dim strPalletsDue As String
    Dim strReadyLabel As String
    Dim strActualCages As String
    Dim strActualCases As String
    Dim strManHours As String
    Dim strShift As String
    Dim strORIGIN As String
    Dim strDueTime As String
    Dim FoundID As String
    Dim FoundID2 As String
    Dim InsertIntoAccessOK As Boolean
    Dim Criteria As String
    Dim ExcludeFields As String
    Dim ErrMessages As String
    Dim TSColName As String
    Dim TSColNumber As Long
    Dim DBDeliveryInfoTable As String
    Dim DBLabourHoursTable As String
    Dim DBShortExtraTable As String
    Dim DBComplianceTable As String
    Dim DBName As String
    Dim dtDateDue As Date
    Dim strDeliveryComments As String
    
    'Get search date into a string:
    SearchText = Me.txtImportDate.Text
    'Column G has the Delivery Date:
    'TURN OFF SCREEN UPDATE !
    'INSERT DO EVENTS in middle of PLOT LOOP !
    'SearchCol = 6
    '**********************************************************************************************
    '* For shared version - import the DAILY sheet to extract data from - into the LOCAL workbook:
    '**********************************************************************************************
    Fieldnames = ""
    FieldValues = ""
    
    Application.ScreenUpdating = False
    If Len(Me.txtImportDate.Text) > 0 Then
        dtImportDate = CDate(Me.txtImportDate.Text)
        SearchText = CStr(dtImportDate)
    Else
        SearchText = Date 'set to today
        
    End If
    TotalBlanks = 0
    Application.ScreenUpdating = False
    Call DG_InsertStuff_v5.DeleteSheet("Daily", ThisWorkbook)
    Application.StatusBar = "Please Wait ... Loading Data Sheet"
    Call getworkbook(Newsheet, "", ThisWorkbook) 'Copies the FIRST TAB of the selected workbook - and paste into current workbook. CLOSE ORIGINAL WORKBOOK.
    If DG_InsertStuff_v5.sheetExists(Newsheet, ThisWorkbook) Then
    'MsgBox (Newsheet) 'Search for Import Date and then insert into linked workbook:
    'Linked Workbook stored in : WB_MainTimesheetData
        TotalCols = ThisWorkbook.Worksheets(Newsheet).Cells(2, Columns.Count).End(xlToLeft).Column
        ColIDX = 1
        TitleRow = ConsolidateRow(Newsheet, 1, TotalCols, ";", " ", " ", ThisWorkbook)
        TitleRow2 = ConsolidateRow(Newsheet, 2, TotalCols, ";", " ", " ", ThisWorkbook)
        Call ConvertColNames("DEL. DATE", TSColName, TSColNumber, DateCol)
        Call ConvertColNames("YPO", TSColName, TSColNumber, ASNCol)
        Call ConvertColNames("REF NO.", TSColName, TSColNumber, DelRefCol)
        Call ConvertColNames("SUPPLIER CODE", TSColName, TSColNumber, SupplierCodeCol)
        Call ConvertColNames("SUPPLIER NAME", TSColName, TSColNumber, SupplierCol)
        Call ConvertColNames("CASES", TSColName, TSColNumber, ExpectedCasesCol)
        Call ConvertColNames("LINES", TSColName, TSColNumber, ExpectedLinesCol)
        Call ConvertColNames("EXPECTED PUT AWAY PALLETS", TSColName, TSColNumber, EstimatedPalletsCol)
        Call ConvertColNames("EXPECTED CAGES", TSColName, TSColNumber, EstimatedCagesCol)
        Call ConvertColNames("EXPECTED TOTES", TSColName, TSColNumber, EstimatedTotesCol)
        Call ConvertColNames("CTNS", TSColName, TSColNumber, CartonsDueCol)
        Call ConvertColNames("PLTS", TSColName, TSColNumber, PalletsDueCol)
        Call ConvertColNames("DELIVERY TYPE", TSColName, TSColNumber, REadyLabelCol)
        Call ConvertColNames("MANHR", TSColName, TSColNumber, ManHoursCol)
        Call ConvertColNames("DUE TIME", TSColName, TSColNumber, DueTimeCol)
        Call ConvertColNames("SHIFT", TSColName, TSColNumber, SHIFTCol)
        Call ConvertColNames("ORIGIN", TSColName, TSColNumber, OriginCol)
        Call ConvertColNames("ACTUAL CASES", TSColName, TSColNumber, ActualCasesCol)
        
        'Extract ALL ROWS from DAILY sheet where the specified date matches:
        'SORT the LOCAL DAILY sheet FIRST - in reverse Delivery Date order - to get latest DeliveryRef version (maybe 2 - if delivery turned away).
        Call SortSheet(ThisWorkbook, Newsheet, DateCol, DelRefCol, True)
        ExtractedRows = ExtractRows(ThisWorkbook, WB_MainTimesheetData, Newsheet, DateCol, SearchText, TotalRowsExtracted, DuplicateRows, DelRefCol)
    'Then the extracted rows need to be placed on the next available row on the records sheet
    'Each field separated by semi-colons:
    'Need to loop through array and extract each field and then place on GI_DATA sheet - row by row until array completly read:
    'Then DELETE the sheet that was copied as no longer needed.
        
        If TotalRowsExtracted > 0 Then
            Application.StatusBar = "Please Wait ... Ploting Rows"
            TotalRows = UBound(ExtractedRows) '
            tempArr = Split(TitleRow, ";")
            'TotalCols = UBound(tempArr)
            Set WB_MainTimesheetData = ThisWorkbook
            ColIDX = 1
            'INSERTS COLUMN TITLES into GI DATA sheet on first row:
            Do While ColIDX <= TotalCols
                Entry = GetFieldValue(TitleRow, ColIDX - 1, ";")
                WB_MainTimesheetData.Worksheets("GI DATA").Cells(1, ColIDX).value = Entry
                Entry = GetFieldValue(TitleRow2, ColIDX - 1, ";")
                WB_MainTimesheetData.Worksheets("GI DATA").Cells(2, ColIDX).value = Entry
                ColIDX = ColIDX + 1
            Loop
            'Merge TITLE cells:
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(1, 21).MergeCells = True
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(1, 21).HorizontalAlignment = xlCenter
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(22, 22 + 11).MergeCells = True
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(22, 22 + 11).HorizontalAlignment = xlCenter
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(22 + 11, 22 + 11 + 14).MergeCells = True
            WB_MainTimesheetData.Worksheets("GI DATA").Cells(22 + 11, 22 + 11 + 14).HorizontalAlignment = xlCenter
            ArrIDX = 0
            GIStartRow = MainGIModule_v1_1.GetNextAvailablerow(WB_MainTimesheetData, "GI DATA", 1, 1)
            MainSheetStartRow = GetNextAvailablerow(WB_MainTimesheetData, MainWorksheet, 1, 1)
            DuplicateRows = 0
            Do While ArrIDX <= TotalRows
                WholeRow = ExtractedRows(ArrIDX) 'FROM DAILY sheet - just loaded in. Is removed after data extracted.
                DateEntry = GetFieldValue(WholeRow, DateCol - 1, ";")
                ASNEntry = GetFieldValue(WholeRow, ASNCol - 1, ";")
                DelRefText = GetFieldValue(WholeRow, DelRefCol - 1, ";")
                
                'Check Entry if the DELIVERY DATE is blank. if not then plot.
                If Len(DateEntry) > 0 And Not DateEntry = "0" Then
                    ColIDX = 1
                    If SearchRows(ThisWorkbook, "GI DATA", DateCol, DateEntry) > 0 And SearchRows(ThisWorkbook, "GI DATA", DelRefCol, DelRefText) > 0 And SearchRows(ThisWorkbook, "GI DATA", ASNCol, ASNEntry) > 0 Then
                        'ROW ALREADY EXISTS: Delivery Date and ASN and Delivery REF match DAILY sheet to GI DATA.
                        DuplicateRows = DuplicateRows + 1
                    Else
                        Do While ColIDX <= TotalCols
                            Entry = GetFieldValue(WholeRow, ColIDX - 1, ";") 'Entry is the value from the column. wholerow from the DAILY sheet now in array.
                            Application.EnableEvents = False
                            '**********************************************************************
                            'NOT ABLE TO QUIT MACRO WHILE THIS LOOP IS RUNNING !!!!!!!!!!!!!!!
                            '***********************************************************************
                            'HERE - change to a DATE - European HERE before it goes to the GI DATA sheet:
                            If IsDate(Entry) Then
                                dtImportDate = CDate(Entry)
                                
                                WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).NumberFormat = "dd/mmm/yyyy"
                                WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).value = dtImportDate
                            Else
                                If IsNumeric(Entry) Then
                                    WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).NumberFormat = "General"
                                End If
                                WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).value = Entry
                            End If
                            If Len(Entry) = 0 Then
                                Entry = " "
                            End If
                            If Len(FieldValues) = 0 Then
                                'FieldVALUES = Entry
                            Else
                                'FieldVALUES = FieldVALUES & "," & Entry
                            End If
                            'Call InsertFormulaIntoRecordSheet(MainWorksheet, MainSheetStartRow + ArrIDX)
                            DoEvents
                            ColIDX = ColIDX + 1
                        Loop
                        Call DrawGrid("GI DATA", GIStartRow + ArrIDX, GIStartRow + ArrIDX, 1, TotalCols, "FULL", False, WB_MainTimesheetData)
                    End If
                    'MsgBox (FieldVALUES)
                    strDeliveryDate = GetFieldValue(WholeRow, DateCol - 1, ";")
                    strDeliveryRef = GetFieldValue(WholeRow, DelRefCol - 1, ";")
                    strSupplier = GetFieldValue(WholeRow, SupplierCodeCol - 1, ";") & " " & GetFieldValue(WholeRow, SupplierCol - 1, ";")
                    strASNNumber = GetFieldValue(WholeRow, ASNCol - 1, ";")
                    strExpectedCases = GetFieldValue(WholeRow, ExpectedCasesCol - 1, ";")
                    strExpectedLines = GetFieldValue(WholeRow, ExpectedLinesCol - 1, ";")
                    strEstimatedPallets = GetFieldValue(WholeRow, EstimatedPalletsCol - 1, ";")
                    strEstimatedCages = GetFieldValue(WholeRow, EstimatedCagesCol - 1, ";")
                    strEstimatedTotes = GetFieldValue(WholeRow, EstimatedTotesCol - 1, ";")
                    strCartonsDue = GetFieldValue(WholeRow, CartonsDueCol - 1, ";")
                    strPalletsDue = GetFieldValue(WholeRow, PalletsDueCol - 1, ";")
                    strReadyLabel = GetFieldValue(WholeRow, REadyLabelCol - 1, ";")
                    If UCase(strReadyLabel) = "PRE-LABELLED" Then
                        strReadyLabel = "YES"
                    Else
                        strReadyLabel = "NO"
                    End If
                    strActualCases = GetFieldValue(WholeRow, ActualCasesCol - 1, ";")
                    
                    strManHours = GetFieldValue(WholeRow, ManHoursCol - 1, ";")
                    strShift = GetFieldValue(WholeRow, SHIFTCol - 1, ";")
                    strORIGIN = GetFieldValue(WholeRow, OriginCol - 1, ";")
                    strDueTime = GetFieldValue(WholeRow, DueTimeCol - 1, ";")
                    strDeliveryComments = ""
                    If IsNumeric(strDueTime) Then
                        dtDateDue = CDate(strDueTime)
                        strDueTime = CStr(dtDateDue)
                    Else
                        strDeliveryComments = strDueTime
                    End If
                    'Any blank values could be dealt with here really:
                    
                    Fieldnames = "DeliveryDate,DeliveryReference,Supplier,ASNNumber,ExpectedCases,ExpectedLines,EstimatedPallets"
                    Fieldnames = Fieldnames & "," & "EstimatedCages,EstimatedTotes,CartonsDue,PalletsDue,ReadyLabel,ActualCases"
                    Fieldnames = Fieldnames & "," & "CalcHours,Shift,Origin,DueTime,DeliveryComments"
                    FieldValues = Chr(34) & strDeliveryDate & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strDeliveryRef & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strSupplier & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strASNNumber & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strExpectedCases & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strExpectedLines & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strEstimatedPallets & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strEstimatedCages & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strEstimatedTotes & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strCartonsDue & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strPalletsDue & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strReadyLabel & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strActualCases & Chr(34)
                    
                    FieldValues = FieldValues & ";" & Chr(34) & strManHours & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strShift & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strORIGIN & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strDueTime & Chr(34)
                    FieldValues = FieldValues & ";" & Chr(34) & strDeliveryComments & Chr(34)
                    InsertIntoAccessOK = False
                    ErrMessages = ""
                    Criteria = ""
                    ExcludeFields = ""
                    'MsgBox (FieldVALUES)
                    If Len(AccessDBpath) = 0 Then
                        MsgBox ("No Path specified to ACCESS Database")
                        Exit Sub
                    End If
                    DBDeliveryInfoTable = "tblDeliveryInfo"
                    DBLabourHoursTable = "tblLabourHours"
                    DBShortExtraTable = "tblShortsAndExtraParts"
                    DBComplianceTable = "tblSupplierCompliance"
                    If SearchAccessDB(AccessDBpath, DBDeliveryInfoTable, "DeliveryDate", strDeliveryDate, "DATE", "=", FoundID) And _
                        SearchAccessDB(AccessDBpath, DBDeliveryInfoTable, "DeliveryReference", strDeliveryRef, "STRING", "=", FoundID2) Then
                        'FOUND ROW
                    Else
                        'INSERT INTO ACCESS DB Table: tblDeliveryInfo - specific fields.
                        'InsertIntoAccessOK = InsertUpdateRecords(False, AccessDBpath, DBDeliveryInfoTable, Fieldnames, FieldValues, Criteria, ExcludeFields, ErrMessages)
                        InsertIntoAccessOK = InsertUpdateRecords_Using_Parameters(False, "0", AccessDBpath, DBDeliveryInfoTable, Fieldnames, FieldValues, _
                            Criteria, ExcludeFields, ErrMessages, False, ";")
                        If Len(ErrMessages) > 0 Then
                            MsgBox ("Insert Error Message: " & ErrMessages)
                            
                        End If
                    End If
                    'Call InsertFormulaIntoRecordSheet(WB_MainTimesheetData, MainWorksheet, MainSheetStartRow + ArrIDX, False, WholeRow)
                    'Draw Grid around each record in GI DATA:
                    
                End If
                DoEvents
                WB_MainTimesheetData.Worksheets(MainWorksheet).Range("A:A").NumberFormat = "dd/mmm/yyyy"
                WB_MainTimesheetData.Worksheets("GI DATA").Range("M:M").NumberFormat = "hh:mm:ss"
                ArrIDX = ArrIDX + 1
            Loop
            Message = "Total Rows Imported: " & CStr(ArrIDX - 1)
            If DuplicateRows > 0 Then
                Message = Message & " , Duplicate Rows: " & CStr(DuplicateRows)
            End If
            MsgBox (Message)
        Else
            Message = "No Rows Imported"
            If DuplicateRows > 0 Then
                Message = Message & ", Duplicated Rows: " & CStr(DuplicateRows)
            Else
                Message = Message & ", no duplicates so DATE NOT EXIST"
            End If
            MsgBox (Message)
            Exit Sub
        End If
        
        '************************************** SORT SHEETS **************************************************
        'Call SortSheet(WB_MainTimesheetData, "GI DATA", 6)
        'Call SortSheet(WB_MainTimesheetData, "Timesheet Records", 1)
        '******************************************************************************************************
    Else
        MsgBox ("Could not find data sheet")
        Exit Sub
    End If
    Call DeleteSheet("Daily", ThisWorkbook)
    If TotalRows < 2 Then
        MsgBox ("NO ROWS WERE ADDED")
    End If
    'ADD TOTAL INSERTED also:
    'MsgBox ("TOTAL BLANKS = " & CStr(TotalBlanks))
    
    Me.comDeliveryRef.Clear
    Criteria = ""
    DBName = ""
    ListArr = PopulateDropdowns_From_ACCESS(DBDeliveryInfoTable, AccessDBpath, 2, DBName, Criteria, True)
    'ListArr = PopulateDropdowns(MainWorksheet, 2, 0, True, WB_MainTimesheetData)
    'get dropdowns from ACCESS TABLE:
    If Not IsArrayEmpty(ListArr) Then
        For IDX = 0 To UBound(ListArr)
            If Len(ListArr(IDX)) > 0 Then
                Me.comDeliveryRef.AddItem (ListArr(IDX))
            End If
        Next
    End If
    Me.comASNNo.Clear
    ListArr = PopulateDropdowns_From_ACCESS(DBDeliveryInfoTable, AccessDBpath, 4, DBName, Criteria, True)
    'ListArr = PopulateDropdowns(MainWorksheet, 4, 0, True, WB_MainTimesheetData)
    If Not IsArrayEmpty(ListArr) Then
        For IDX = 0 To UBound(ListArr)
            If Len(ListArr(IDX)) > 0 Then
                Me.comASNNo.AddItem (ListArr(IDX))
            End If
        Next
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Call TestTodayImport
    
End Sub

Sub Import_Into_Worksheet()
    Dim Newsheet As String
    Dim ExtractedRows() As String
    Dim DateCol As Long
    Dim ASNCol As Long
    Dim DelRefCol As Long
    Dim DelRefText As String
    Dim SearchText As String
    Dim ColSearch As String
    Dim ArrIDX As Long
    Dim TotalRows As Long
    Dim WholeRow As String
    Dim TitleRow As String
    Dim ColIDX As Long
    Dim tempArr() As String
    Dim TotalCols As Long
    Dim DateEntry As String
    Dim Entry As String
    Dim ASNEntry As String
    Dim GIStartRow As Long
    Dim MainSheetStartRow As Long
    Dim dtImportDate As Date
    Dim ListArr() As String
    Dim TotalRowsExtracted As Long
    Dim DuplicateRows As Long
    Dim Message As String
    
    'Get search date into a string:
    SearchText = Me.txtImportDate.Text
    'Column G has the Delivery Date:
    'TURN OFF SCREEN UPDATE !
    'INSERT DO EVENTS in middle of PLOT LOOP !
    'SearchCol = 6
    '**********************************************************************************************
    '* For shared version - import the DAILY sheet to extract data from - into the LOCAL workbook:
    '**********************************************************************************************
    Application.ScreenUpdating = False
    If Len(Me.txtImportDate.Text) > 0 Then
        dtImportDate = CDate(Me.txtImportDate.Text)
        SearchText = CStr(dtImportDate)
    Else
        SearchText = Date 'set to today
        
    End If
    TotalBlanks = 0
    Call DG_InsertStuff_v4.DeleteSheet("Daily", ThisWorkbook)
    Application.StatusBar = "Please Wait ... Loading Data Sheet"
    Call getworkbook(Newsheet, "", ThisWorkbook) 'Copies the FIRST TAB of the selected workbook - and paste into current workbook. CLOSE ORIGINAL WORKBOOK.
    If DG_InsertStuff_v4.sheetExists(Newsheet, ThisWorkbook) Then
    'MsgBox (Newsheet) 'Search for Import Date and then insert into linked workbook:
    'Linked Workbook stored in : WB_MainTimesheetData
        TotalCols = ThisWorkbook.Worksheets(Newsheet).Cells(2, Columns.Count).End(xlToLeft).Column
        ColIDX = 1
        Do While ColIDX <= TotalCols
            ColSearch = ThisWorkbook.Worksheets(Newsheet).Cells(2, ColIDX).value
            If UCase(ColSearch) = "DEL. DATE" Then
                DateCol = ColIDX 'Gets the DeliveryDate Column position from the LOCAL Daily Sheet
            End If
            If UCase(ColSearch) = "YPO" Then
                ASNCol = ColIDX 'Gets the ASN column position from the LOCAL Daily sheet
            End If
            If UCase(ColSearch) = "REF NO." Then
                DelRefCol = ColIDX 'Gets the Delivery Reference Column position from the LOCAL Daily sheet.
            End If
            ColIDX = ColIDX + 1
        Loop
        'Extract ALL ROWS from DAILY sheet where the specified date matches:
        ExtractedRows = ExtractRows(ThisWorkbook, WB_MainTimesheetData, Newsheet, DateCol, SearchText, TotalRowsExtracted, DuplicateRows, DelRefCol)
    'Then the extracted rows need to be placed on the next available row on the records sheet
    'Each field separated by semi-colons:
    'Need to loop through array and extract each field and then place on GI_DATA sheet - row by row until array completly read:
    'Then DELETE the sheet that was copied as no longer needed.
        
        If TotalRowsExtracted > 0 Then
            Application.StatusBar = "Please Wait ... Ploting Rows"
            TotalRows = UBound(ExtractedRows) '
            TitleRow = ExtractedRows(0) 'FIRST ROW - which will be the titles.
            tempArr = Split(WholeRow, ";")
            'TotalCols = UBound(tempArr)
            
            ColIDX = 1
            Do While ColIDX <= TotalCols
                Entry = GetFieldValue(WholeRow, ColIDX - 1, ";")
                WB_MainTimesheetData.Worksheets("GI DATA").Cells(1, ColIDX).value = Entry
                ColIDX = ColIDX + 1
            Loop
            ArrIDX = 1
            GIStartRow = MainGIModule_v1_1.GetNextAvailablerow(WB_MainTimesheetData, "GI DATA", 1, 1)
            MainSheetStartRow = GetNextAvailablerow(WB_MainTimesheetData, MainWorksheet, 1, 1)
            Do While ArrIDX <= TotalRows
                WholeRow = ExtractedRows(ArrIDX) 'FROM DAILY sheet - just loaded in. Is removed after data extracted.
                DateEntry = GetFieldValue(WholeRow, DateCol - 1, ";")
                'Check Entry if the DELIVERY DATE is blank. if not then plot.
                If Len(DateEntry) > 0 And Not DateEntry = "0" Then
                    ColIDX = 1
                    Do While ColIDX <= TotalCols
                        Entry = GetFieldValue(WholeRow, ColIDX - 1, ";")
                        Application.EnableEvents = False
                        '**********************************************************************
                        'NOT ABLE TO QUIT MACRO WHILE THIS LOOP IS RUNNING !!!!!!!!!!!!!!!
                        '***********************************************************************
                        'HERE - change to a DATE - European HERE before it goes to the GI DATA sheet:
                        If IsDate(Entry) Then
                            dtImportDate = CDate(Entry)
                            WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).NumberFormat = "dd/mmm/yyyy"
                            WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).value = dtImportDate
                        Else
                            WB_MainTimesheetData.Worksheets("GI DATA").Cells(GIStartRow + ArrIDX, ColIDX).value = Entry
                        End If
                        
                        'Call InsertFormulaIntoRecordSheet(MainWorksheet, MainSheetStartRow + ArrIDX)
                        DoEvents
                        ColIDX = ColIDX + 1
                    Loop
                    
                    'HAVE to prefix with TARGET WORKBOOK:
                    
                    'Draw Grid around each record in GI DATA:
                    Call DrawGrid("GI DATA", GIStartRow + ArrIDX, GIStartRow + ArrIDX, 1, TotalCols, "FULL", False, WB_MainTimesheetData)
                    Call InsertFormulaIntoRecordSheet(WB_MainTimesheetData, MainWorksheet, MainSheetStartRow + ArrIDX, False, WholeRow)
                End If
                DoEvents
                WB_MainTimesheetData.Worksheets(MainWorksheet).Range("A:A").NumberFormat = "dd/mmm/yyyy"
                ArrIDX = ArrIDX + 1
            Loop
            Message = "Total Rows Imported: " & CStr(ArrIDX - 1)
            If DuplicateRows > 0 Then
                Message = Message & " , Duplicate Rows: " & CStr(DuplicateRows)
            End If
            MsgBox (Message)
        Else
            Message = "No Rows Imported"
            If DuplicateRows > 0 Then
                Message = Message & ", Duplicated Rows: " & CStr(DuplicateRows)
            Else
                Message = Message & ", no duplicates so DATE NOT EXIST"
            End If
            MsgBox (Message)
            Exit Sub
        End If
        
        '************************************** SORT SHEETS **************************************************
        Call SortSheet(WB_MainTimesheetData, "GI DATA", 6)
        Call SortSheet(WB_MainTimesheetData, "Timesheet Records", 1)
        '******************************************************************************************************
    Else
        MsgBox ("Could not find data sheet")
        Exit Sub
    End If
    Call DeleteSheet("Daily", ThisWorkbook)
    If TotalRows < 2 Then
        MsgBox ("NO ROWS WERE ADDED")
    End If
    'ADD TOTAL INSERTED also:
    MsgBox ("TOTAL BLANKS = " & CStr(TotalBlanks))
    
    Me.comDeliveryRef.Clear
    ListArr = PopulateDropdowns(MainWorksheet, 2, 0, True, WB_MainTimesheetData)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comDeliveryRef.AddItem (ListArr(IDX))
        End If
    Next
    Me.comASNNo.Clear
    ListArr = PopulateDropdowns(MainWorksheet, 4, 0, True, WB_MainTimesheetData)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comASNNo.AddItem (ListArr(IDX))
        End If
    Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub ConvertColNames(ByRef DailyColName As String, ByRef TimesheetColName As String, ByRef TimesheetColNumber As Long, Optional ByRef DailyCol As Long = 0)
    Dim OldName As String
    Dim NewName As String
    
    OldName = UCase(DailyColName)
    
    If OldName = "DEL. DATE" Or DailyCol = 6 Then
        TimesheetColName = "DELIVERY DATE"
        TimesheetColNumber = 1
        DailyCol = 6
        DailyColName = "DEL. DATE"
    End If
    If OldName = "REF NO." Or DailyCol = 7 Then
        TimesheetColName = "DELIVERY REFERENCE"
        TimesheetColNumber = 2
        DailyCol = 7
        DailyColName = "REF NO."
    End If
    If UCase(OldName) = "SUPPLIER CODE" Or DailyCol = 2 Then
        TimesheetColName = "SUPPLIER"
        TimesheetColNumber = 3
        DailyCol = 2
        DailyColName = "SUPPLIER CODE"
    End If
    If UCase(OldName) = "SUPPLIER NAME" Or DailyCol = 3 Then
        TimesheetColName = "SUPPLIER"
        TimesheetColNumber = 3
        DailyCol = 3
        DailyColName = "SUPPLIER NAME"
    End If
    If UCase(OldName) = "YPO" Or DailyCol = 10 Then
        TimesheetColName = "ASN NUMBER"
        TimesheetColNumber = 4
        DailyCol = 10
        DailyColName = "YPO"
    End If
    If UCase(OldName) = "CASES" Or DailyCol = 15 Then
        TimesheetColName = "EXPECTED CASES"
        TimesheetColNumber = 5
        DailyCol = 15
        
    End If
    If UCase(OldName) = "LINES" Or DailyCol = 14 Then
        TimesheetColName = "EXPECTED LINES"
        TimesheetColNumber = 6
        DailyCol = 14
    End If
    If UCase(OldName) = "EXPECTED PUT AWAY PALLETS" Or DailyCol = 19 Then
        TimesheetColName = "ESTIMATED PALLETS"
        TimesheetColNumber = 7
        DailyCol = 19
    End If
    If UCase(OldName) = "EXPECTED CAGES" Or DailyCol = 18 Then
        TimesheetColName = "ESTIMATED CAGES"
        TimesheetColNumber = 8
        DailyCol = 18
    End If
    If UCase(OldName) = "EXPECTED TOTES" Or DailyCol = 17 Then
        TimesheetColName = "ESTIMATED TOTES"
        TimesheetColNumber = 9
        DailyCol = 17
    End If
    If UCase(OldName) = "PLTS" Or DailyCol = 4 Then
        TimesheetColName = "PALLETS DUE"
        TimesheetColNumber = 11
        DailyCol = 4
    End If
    If UCase(OldName) = "DELIVERY TYPE" Or DailyCol = 1 Then
        TimesheetColName = "READY LABEL"
        TimesheetColNumber = 12
        DailyCol = 1
    End If
    If UCase(OldName) = "CTNS" Or DailyCol = 5 Then
        TimesheetColName = "CARTONS DUE"
        TimesheetColNumber = 0
        DailyCol = 5
    End If
    If UCase(OldName) = "MANHR" Or DailyCol = 20 Then
        TimesheetColName = "MAN HOURS"
        TimesheetColNumber = 0
        DailyCol = 20
    End If
    If UCase(OldName) = "DUE TIME" Or DailyCol = 13 Then
        TimesheetColName = "DUE TIME"
        TimesheetColNumber = 0
        DailyCol = 13
    End If
    If UCase(OldName) = "SHIFT" Or DailyCol = 11 Then
        TimesheetColName = "SHIFT"
        TimesheetColNumber = 0
        DailyCol = 11
    End If
    If UCase(OldName) = "ORIGIN" Or DailyCol = 8 Then
        TimesheetColName = "ORIGIN"
        TimesheetColNumber = 0
        DailyCol = 8
    End If
    If UCase(OldName) = "ACTUAL CASES" Or DailyCol = 16 Then
        TimesheetColName = "ACTUAL CASES"
        TimesheetColNumber = 0
        DailyCol = 16
    End If


End Sub


Sub InsertFormulaIntoRecordSheet(WB As Workbook, WorksheetName As String, Rownum As Long, Optional InsertAsFormula As Boolean = True, Optional WholeRow As String = "")
    'Populate 11 columns on main record sheet with the formula needed to look up the values from GI DATA
    Dim ColIDX As Long
    Dim InsertValue As String
    Dim dtImportDate As Date
    Dim TimesheetCol As Long
    Dim DailyCol As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    If InsertAsFormula Then
        ColIDX = 1
        InsertValue = "='GI DATA'!F" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 2
        InsertValue = "='GI DATA'!G" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 3
        InsertValue = "='GI DATA'!B" & CStr(Rownum) & " & " & Chr(34) & Chr(32) & Chr(34) & " & " & "'GI DATA'!C" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 4
        InsertValue = "='GI DATA'!J" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 5
        InsertValue = "='GI DATA'!O" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 6
        InsertValue = "='GI DATA'!N" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 7
        InsertValue = "='GI DATA'!S" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 8
        InsertValue = "='GI DATA'!R" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 9
        InsertValue = "='GI DATA'!Q" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 11
        InsertValue = "='GI DATA'!D" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 10
        InsertValue = "='GI DATA'!E" & CStr(Rownum)
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
        ColIDX = 12
        InsertValue = "=IF('GI DATA'!A" & CStr(Rownum) & "=" & Chr(34) & "PRE-LABELLED" & Chr(34) & "," & Chr(34) & "YES" & Chr(34) & "," & Chr(34) & "NO" & Chr(34) & ")"
        ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).Formula = InsertValue
    Else
        If Len(WholeRow) > 0 Then
            'WHOLEROW is ZERO INDEX based so will be ONE LESS than the column it represents.
            ColIDX = 1
            TimesheetCol = 1
            
            InsertValue = GetFieldValue(WholeRow, 5, ";")
            dtImportDate = CDate(InsertValue)
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).NumberFormat = "dd/mmm/yyyy"
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = dtImportDate
            ColIDX = 2
            InsertValue = GetFieldValue(WholeRow, 6, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 3
            InsertValue = GetFieldValue(WholeRow, 1, ";") & " " & GetFieldValue(WholeRow, 2, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 4
            InsertValue = GetFieldValue(WholeRow, 9, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 5
            InsertValue = GetFieldValue(WholeRow, 14, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 6
            InsertValue = GetFieldValue(WholeRow, 13, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 7
            InsertValue = GetFieldValue(WholeRow, 18, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 8
            InsertValue = GetFieldValue(WholeRow, 17, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 9
            'InsertValue = "='GI DATA'!Q" & CStr(Rownum)
            InsertValue = GetFieldValue(WholeRow, 16, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 11
            InsertValue = GetFieldValue(WholeRow, 3, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 12
            If UCase(GetFieldValue(WholeRow, 0, ";")) = "PRE-LABELLED" Then
                InsertValue = "YES"
            Else
                InsertValue = "NO"
            End If
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            ColIDX = 10
            InsertValue = GetFieldValue(WholeRow, 4, ";")
            ThisWB.Worksheets(WorksheetName).Cells(Rownum, ColIDX).value = InsertValue
            
        Else
            MsgBox ("Insert Formula: ROW passed is BLANK")
        End If
    End If
    ThisWB.Worksheets(WorksheetName).Cells.NumberFormat = "General"
    
End Sub


Private Sub btnPalletise_Click()
    'BUTTON: Palletise:
    If UCase(Me.txtCBYesNoPalletise.Text = "YES") Then
        Me.txtCBYesNoPalletise.Text = "NO"
    Else
        'Me.txtCBPalletise.Font = "Wingdings 2"
        Me.txtCBYesNoPalletise.Text = "YES"
    End If
End Sub

Private Sub btnPRINT_Click()
    'need new code.
    'Application.Dialogs(xlDialogPrinterSetup).Show
    'Me.PrintForm
    DoEvents
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + _
        KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + _
        KEYEVENTF_KEYUP, 0
    DoEvents
    Workbooks.Add
    Application.Wait Now + TimeValue("00:00:02")
    'ActiveSheet.PasteSpecial Format:="Bitmap", Link:=False, _
        DisplayAsIcon:=False
    ActiveSheet.Range("A1").Select
    'added to force landscape
    ActiveSheet.PageSetup.Orientation = xlLandscape
    
   
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With

    ActiveSheet.PageSetup.PrintArea = ""
    
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        '.PrintQuality = 300
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWorkbook.Close False
    Application.Dialogs(xlDialogPrinterSetup).Show
    Me.PrintForm
End Sub

Private Sub btnReadyLabel_Click()
    'BUTTON: Ready Label:
    If Len(Me.txtCBReadyLabel.Text) > 0 Then
        Me.txtCBReadyLabel.Text = ""
    Else
        Me.txtCBReadyLabel.Font.Name = "WINGDINGS 2"
        Me.txtCBReadyLabel.Text = Chr(80)
    End If
End Sub

Private Sub btnResetInboundData_Click()
    'CLEAR ALL FIELDS from Inbound AREA:
    Call ClearEntry(FirstInboundTAG, LastInboundTAG)
End Sub

Private Sub btnResetOperationData_Click()
    Call ClearEntry(FirstOperationTAG, LastOperationTAG)
    Call ClearEntry(FirstTimeTAG, LastTimeTAG)
    Call ClearEntry(FirstSCTAG, LastSCTAG)
    
End Sub


Private Sub btnSearchInboundData_Click()
   Call SearchInboundData_From_ACCESS
    
End Sub

Sub Search_InboundData()
     'Find ROW on which ASN Number entry is:
    'Times in record need to be displayed in special textboxes TAGnumber from 441 - 645 (allows 50 rows=200 controls)
    Dim LastRow As Long
    Dim RowIDX As Long
    Dim DeliveryRef As String
    Dim ASNNo As String
    Dim RowNum1 As Long
    Dim RowNum2 As Long
    Dim Rownum As Long
    Dim ASNCol As Long
    Dim DeliveryCol As Integer
    Dim SearchCol As Long
    Dim SearchText As String
    Dim WorksheetName As String
    
    WorksheetName = "Timesheet Records"
    
    'If Not sheetExists(WorksheetName) Then
        'LastRow = WB_MainTimesheetData.Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).row
    'End If
    RowNum1 = 0
    RowNum2 = 0
    TotalFields = 42
    ASNCol = 4
    DeliveryCol = 2
    Call ClearEntry(FirstInboundTAG, LastInboundTAG)
    Call ClearEntry(25, 27)
    Call ClearEntry(FirstOperationTAG, LastOperationTAG)
    Call ClearEntry(FirstTimeTAG, LastTimeTAG) 'FLM controls
    Call ClearEntry(FirstSCTAG, LastSCTAG) 'Supplier Compliance controls
    Me.txtLastSaved.Text = ""
    Me.txtArrivedONTime.Text = ""
    Me.txtArrivedONTimeComment.Text = ""
    Me.txtIsItSafe.Text = ""
    Me.txtIsItSafeComment.Text = ""
    Me.txtCompleted.Text = ""
    Me.txtCompletedComment.Text = ""
    
    'If Len(Me.comDeliveryRef.Text) > 0 Then
    '    'BOTH used as search:
    '    SearchCol = 2 'DeliveryRef
    '    SearchText = Me.comDeliveryRef.Text
    '    RowNum1 = SearchRows(WorksheetName, SearchCol, SearchText)
    '    RowNum = RowNum1
    'End If
    If Me.rbASN.value = True Then
        If Len(Me.comASNNo.Text) > 0 Then
            'BOTH used as search:
            SearchCol = ASNCol 'ASN NO.
            SearchText = Me.comASNNo.Text
            RowNum2 = SearchRows(WB_MainTimesheetData, WorksheetName, SearchCol, SearchText)
            Rownum = RowNum2
        Else
            MsgBox ("Please select an ASN No")
            Exit Sub
        End If
    End If
    If Me.rbDeliveryRef.value = True Then
        If Len(Me.comDeliveryRef.Text) > 0 Then
            'BOTH used as search:
            SearchCol = DeliveryCol 'Delivery REF
            SearchText = Me.comDeliveryRef.Text
            RowNum2 = SearchRows(WB_MainTimesheetData, WorksheetName, SearchCol, SearchText)
            Rownum = RowNum2
        Else
            MsgBox ("Please select a Delivery No")
            Exit Sub
        End If
    End If
    'If Len(Me.comDeliveryRef.Text) > 0 And Len(Me.comASNNO.Text) > 0 Then
    '    'BOTH used as search:
    '    SearchCol = 2 'DeliveryRef
    '    SearchText = Me.comDeliveryRef.Text
    '    RowNum1 = SearchRows(WorksheetName, SearchCol, SearchText)
    '
    '    SearchCol = 4 'ASN NO.
    '    SearchText = Me.comASNNO.Text
    '    RowNum2 = SearchRows(WorksheetName, SearchCol, SearchText)
    '
    '    If RowNum1 = RowNum2 Then
    '        RowNum = RowNum1
     '   Else
     '       MsgBox ("Delivery Ref Row and ASN No Row are different")
      '      Exit Sub
      '  End If
    'End If
    If Rownum > 0 Then
        'Now we have the correct row:
        Call EnableDisableControls(True, 17, TotalFields, "TEXTBOX")
        Call EnableDisableControls(True, 17, TotalFields, "COMBOBOX")
        Call EnableDisableControls(True, 3, TotalFields, "BUTTONS") 'Buttons btn156 - btn160 are the bottom buttons.
        Call EnableDisableControls(False, 1, 1, "BUTTONS")
        Call EnableDisableControls(False, 31, 31, "TEXTBOX")
        
        'TotalFields needs to be set first:
        Call PopulateUserformControls(WB_MainTimesheetData, WorksheetName, Rownum)
        CurrentRecordRow = Rownum
    Else
        MsgBox ("Did not find Row Number. Please select an ASN No or Delivery Ref")
        Exit Sub
    End If
End Sub

Sub SearchInboundData_From_ACCESS()
    Dim RowIDX As Long
    Dim strDeliveryRef As String
    Dim ASNNo As String
    Dim RowNum1 As Long
    Dim RowNum2 As Long
    Dim Rownum As Long
    Dim ASNCol As Long
    Dim DeliveryCol As Integer
    Dim SearchCol As Long
    Dim strSearchText As String
    Dim WorksheetName As String
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim DBTable_DeliveryInfo As String
    Dim DBTable_LabourHours As String
    Dim DBTable_ShortAndExtra As String
    Dim DBTable_Compliance As String
    Dim SearchCriteria As String
    Dim FoundSelection As Boolean
    Dim strDeliveryDate As String
    Dim ValueArr() As String
    Dim TotalDBFields As Long
    
    RowNum1 = 0
    RowNum2 = 0
    ASNCol = 4
    DeliveryCol = 2
    Call ClearEntry(FirstInboundTAG, LastInboundTAG)
    Call ClearEntry(25, 27)
    Call ClearEntry(FirstOperationTAG, 30)
    Call ClearEntry(34, LastOperationTAG)
    Call ClearEntry(FirstTimeTAG, LastTimeTAG) 'FLM controls
    Call ClearEntry(FirstSCTAG, LastSCTAG) 'Supplier Compliance controls
    Me.txtLastSaved.Text = ""
    Me.txtArrivedONTime.Text = ""
    Me.txtArrivedONTimeComment.Text = ""
    Me.txtIsItSafe.Text = ""
    Me.txtIsItSafeComment.Text = ""
    Me.txtCompleted.Text = ""
    Me.txtCompletedComment.Text = ""
    
    DBTable_DeliveryInfo = "tblDeliveryInfo"
    DBTable_LabourHours = "tblLabourHours"
    DBTable_ShortAndExtra = "tblShortsAndExtraParts"
    DBTable_Compliance = "tblSupplierCompliance"
    
    Call PerformComboSearch(SearchCriteria, strDeliveryDate, strDeliveryRef, ASNNo) 'SearchText being the criteria used to search for the DeliveryRef OR the ASN.
    
    'Call EnableDisableControls(False, FirstInboundTAG, LastInboundTAG, "ALL")
    Call EnableDisableControls(True, FirstOperationTAG, LastOperationTAG, "TEXTBOX")
    Call EnableDisableControls(True, FirstOperationTAG, LastOperationTAG, "COMBOBOX")
    Call EnableDisableControls(True, 2, LastOtherBtnTAG, "BUTTONS") 'Buttons btn100 - btn104 are the bottom buttons.
    Call EnableDisableControls(False, 1, 1, "BUTTONS")
    'TotalFields needs to be set first:
    'Call PopulateUserformControls(WB_MainTimesheetData, WorksheetName, Rownum)
    'Call PopulateUserformControls_From_Access(strDeliveryRef, ASNNo, strDeliveryDate, 1, 27, AccessDBPath, DBTable_DeliveryInfo, TotalDBFields)
    Call RemoveAllControls(frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls)
    Call RemoveAllControls(frmGI_TimesheetEntry2_1060x630.Frame_ShortParts.Controls)
    Call RemoveAllControls(frmGI_TimesheetEntry2_1060x630.Frame_ExtraParts.Controls)
    Set ctrlCollection = Nothing
    Call Test_PopulationOfControls(strDeliveryRef, ASNNo, strDeliveryDate, 1, 27, AccessDBpath, DBTable_DeliveryInfo, TotalDBFields)
    
    
End Sub

Sub PerformComboSearch(ByRef SearchCriteria As String, ByRef strDeliveryDate As String, ByRef strDeliveryRef As String, ByRef ASNNo As String)
    Dim RowNum1 As Long
    Dim RowNum2 As Long
    Dim Rownum As Long
    Dim ASNCol As Long
    Dim DeliveryCol As Integer
    Dim SearchCol As Long
    Dim WorksheetName As String
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim DBTable_DeliveryInfo As String
    Dim DBTable_LabourHours As String
    Dim DBTable_Operatives As String
    Dim DBTable_ShortAndExtra As String
    Dim DBTable_Compliance As String
    Dim FoundSelection As Boolean
    Dim ValueArr() As String
    Dim SortFields As String
    Dim Reversed As Boolean

    SearchText = ""
    strDeliveryDate = ""
    strDeliveryRef = ""
    ASNNo = ""
    SearchCriteria = ""
    Reversed = False
    SortFields = ""
    
    DBTable_DeliveryInfo = "tblDeliveryInfo"
    DBTable_LabourHours = "tblLabourHours"
    DBTable_Operatives = "tblOperatives"
    DBTable_ShortAndExtra = "tblShortsAndExtraParts"
    DBTable_Compliance = "tblSupplierCompliance"
    'If Len(Me.comDeliveryRef.Text) > 0 Then
    '    'BOTH used as search:
    '    SearchCol = 2 'DeliveryRef
    '    SearchText = Me.comDeliveryRef.Text
    '    RowNum1 = SearchRows(WorksheetName, SearchCol, SearchText)
    '    RowNum = RowNum1
    'End If
    If Me.rbASN.value = True Then
        If Len(Me.comASNNo.Text) > 0 Then
            'BOTH used as search:
            SearchCol = ASNCol 'ASN NO.
            ASNNo = Me.comASNNo.Text
            SearchText = ASNNo
            SearchCriteria = "ASNNumber = " & Chr(34) & SearchText & Chr(34)
        Else
            MsgBox ("Please select an ASN No")
            Exit Sub
        End If
    End If
    If Me.rbDeliveryRef.value = True Then
        If Len(Me.comDeliveryRef.Text) > 0 Then
            'BOTH used as search:
            SearchCol = DeliveryCol 'Delivery REF
            strDeliveryRef = Me.comDeliveryRef.Text
            SearchText = strDeliveryRef
            SearchCriteria = "DeliveryReference = " & Chr(34) & SearchText & Chr(34)
        Else
            MsgBox ("Please select a Delivery No")
            Exit Sub
        End If
    End If
    FoundSelection = LoadAccessDBTable(DBTable_DeliveryInfo, AccessDBpath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues)
    'Fieldnames is populated with all of the fields from the table, FieldValues populated with all the values from the record as comma-delim string
    'Retrieve the Delivery Date in the database for the dropdown just selected:
    If FoundSelection Then
        ValueArr = Split(FieldValues, ",") 'FieldValues has quotes around them !!!!!!!!!!!!!!!!!!!!
        strDeliveryDate = ValueArr(1)
        Me.txtSelectedDeliveryDate.Text = Format(Convert_strDateToDate(strDeliveryDate, False), "dd/mm/yyyy")
    Else
        Me.txtSelectedDeliveryDate.Text = ""
    End If
End Sub

Private Sub btnSearchInboundData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call Search_InboundData
    End If
End Sub


Private Sub btnTestACCESS_Click()
    Dim InsertOK As Boolean
    
    'InsertOK = InsertUpdateRecords()
    'If InsertOK Then
    '    MsgBox ("OK inserted Test Record")
    'End If
    'Call Load_From_Access_Test
    
    
End Sub


Private Sub btnWrapStrap_Click()
    'BUTTON: WRAP / STRAP:
    If UCase(Me.txtCBYesNoWrapStrap.Text) = "YES" Then
        Me.txtCBYesNoWrapStrap.Text = "NO"
    Else
        Me.txtCBYesNoWrapStrap.Text = "YES"
    End If
End Sub

Private Sub CommandButton4_Click()
Dim X As Long
Dim Y As Long

  
  

End Sub

Private Sub Label49_Click()

End Sub

Private Sub lblReadyLabel_Click()
    
End Sub

Private Sub TextBox72_Change()
    MsgBox ("BANG")
End Sub

Private Sub imgCalFrom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'SHOW CALANDER and allow user to choose date:
    DeliveryDate = CalendarForm.GetDate(SelectedDate:=DeliveryDate, FirstDayOfWeek:=Monday, RangeOfYears:=5, DateFontSize:=11, TodayButton:=True, OkayButton:=False, ShowWeekNumbers:=True, _
        PositionTop:=1, PositionLeft:=1, BackgroundColor:=RGB(243, 249, 251), HeaderColor:=RGB(147, 205, 2221), HeaderFontColor:=RGB(255, 255, 255), SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), DateColor:=RGB(243, 249, 251), DateFontColor:=RGB(31, 78, 120), SaturdayFontColor:=RGB(0, 0, 0), SundayFontColor:=RGB(1, 1, 1), _
        DateBorder:=True, DateSpecialEffect:=fmSpecialEffectRaised, DateHoverColor:=RGB(223, 240, 245), DateSelectedColor:=RGB(202, 223, 242), TrailingMonthFontColor:=RGB(155, 194, 230), _
        TodayFontColor:=RGB(0, 176, 80), FirstWeekOfYear:=FirstFourDays)
    
    Me.txtDeliveryDate.Text = Format(DeliveryDate, "dd/mmm/yyyy")
End Sub

Private Sub Label44_Click()
    Dim IDX As Long
    Dim ErrCol As Long
    
    IDX = 1
    Do While IDX <= TotalFields
    
        'Call SetControlBackgroundColour(CStr(IDX), vbRed)
        Call SetControlBackgroundColour(CStr(IDX), vbWhite)
        IDX = IDX + 1
    Loop
    
    
    
End Sub

Private Sub imgCalandar_DeliveryDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'SHOW Delivery Date CALANDAR
    Dim SearchCriteria As String
    Dim DBTable As String
    
    DeliveryDate = CalendarForm.GetDate(SelectedDate:=DeliveryDate, FirstDayOfWeek:=Monday, RangeOfYears:=5, DateFontSize:=11, TodayButton:=True, OkayButton:=False, ShowWeekNumbers:=True, _
        PositionTop:=1, PositionLeft:=1, BackgroundColor:=RGB(243, 249, 251), HeaderColor:=RGB(147, 205, 2221), HeaderFontColor:=RGB(255, 255, 255), SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), DateColor:=RGB(243, 249, 251), DateFontColor:=RGB(31, 78, 120), SaturdayFontColor:=RGB(0, 0, 0), SundayFontColor:=RGB(1, 1, 1), _
        DateBorder:=True, DateSpecialEffect:=fmSpecialEffectRaised, DateHoverColor:=RGB(223, 240, 245), DateSelectedColor:=RGB(202, 223, 242), TrailingMonthFontColor:=RGB(155, 194, 230), _
        TodayFontColor:=RGB(0, 176, 80), FirstWeekOfYear:=FirstFourDays)
    
    Me.txtSelectedDeliveryDate.Text = Format(DeliveryDate, "dd/mmm/yyyy")
    'me.comASNNo
    'Me.comDeliveryRef
    'Populate the combos - comDeliveryRef and comASN
    SearchCriteria = ""
    SearchCriteria = "[DeliveryDate] = " & "#" & Format(DeliveryDate, "yyyy/mm/dd") & "#"
    Me.comDeliveryRef.Clear
    DBTable = "tblDeliveryinfo"
    ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessDBpath, 2, "", SearchCriteria, True)
    If IsArrayEmpty(ListArr) Then
        'Nothing Returned - Database Not updated for this Date
        MsgBox ("Database is empty for this date - SELECT IMPORT DATA")
    Else
        For IDX = 0 To UBound(ListArr)
            If Len(ListArr(IDX)) > 0 Then
                Me.comDeliveryRef.AddItem (ListArr(IDX))
            End If
        Next
        Me.comASNNo.Clear
        ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessDBpath, 4, "", SearchCriteria, True)
        If IsArrayEmpty(ListArr) Then
            MsgBox ("Could not get ASN List ???")
            Exit Sub
        Else
            For IDX = 0 To UBound(ListArr)
                If Len(ListArr(IDX)) > 0 Then
                    Me.comASNNo.AddItem (ListArr(IDX))
                End If
            Next
        End If
    End If
    
End Sub

Private Sub imgCalandar_ImportDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ImportDate = CalendarForm.GetDate(SelectedDate:=DeliveryDate, FirstDayOfWeek:=Monday, RangeOfYears:=5, DateFontSize:=11, TodayButton:=True, OkayButton:=False, ShowWeekNumbers:=True, _
        PositionTop:=1, PositionLeft:=1, BackgroundColor:=RGB(243, 249, 251), HeaderColor:=RGB(147, 205, 2221), HeaderFontColor:=RGB(255, 255, 255), SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), DateColor:=RGB(243, 249, 251), DateFontColor:=RGB(31, 78, 120), SaturdayFontColor:=RGB(0, 0, 0), SundayFontColor:=RGB(1, 1, 1), _
        DateBorder:=True, DateSpecialEffect:=fmSpecialEffectRaised, DateHoverColor:=RGB(223, 240, 245), DateSelectedColor:=RGB(202, 223, 242), TrailingMonthFontColor:=RGB(155, 194, 230), _
        TodayFontColor:=RGB(0, 176, 80), FirstWeekOfYear:=FirstFourDays)
    
    Me.txtImportDate.Text = Format(ImportDate, "dd/mmm/yyyy")
    
    
End Sub

Private Sub rbASN_Click()
    Me.comASNNo.Visible = True
    Me.comDeliveryRef.Visible = False
    Me.rbASN.value = True
    Me.rbDeliveryRef.value = False
    Me.lblSelectSearch.Caption = "Select ASN:"
End Sub

Private Sub rbDeliveryRef_Click()
    Me.comDeliveryRef.Visible = True
    Me.comASNNo.Visible = False
    Me.rbASN.value = False
    Me.rbDeliveryRef.value = True
    Me.lblSelectSearch.Caption = "Select Del Ref:"
End Sub

Private Sub ScrollBar1_Change()
    'CHANGE SIZE OF WHOLE USERFORM - using ZOOM property set to value of scrollbar
    Dim ZoomWidth As Double
    Dim ZoomHeight As Double
    Dim lngZoomWidth As Long
    Dim lngZoomHeight As Long
    Dim IncrementWidth As Long
    Dim IncrementHeight As Long
    
    LastWidth = Me.Width
    LastHeight = Me.Height
    Me.Zoom = Me.ScrollBar1.value
    lblZoom = "Zoom: " & CStr(ScrollBar1.value) & " %"
    'IncrementWidth = ((1 / 100) * Me.Width)
    'IncrementHeight = ((1 / 100) * Me.Height)
    
    ZoomWidth = ((Me.ScrollBar1.value / 100) * CLng(Me.Width)) - 9
    ZoomHeight = ((Me.ScrollBar1.value / 100) * CLng(Me.Height)) - 6
    lngZoomWidth = CLng(ZoomWidth)
    lngZoomHeight = CLng(ZoomHeight)
    'Me.Width = lngZoomWidth
    'Me.Height = lngZoomHeight
    'IncrementWidth = lngZoomWidth
    'IncrementHeight = lngZoomHeight
    'IncrementWidth = CLng(Me.txtEstimatedCages.Text)
    'IncrementHeight = CLng(Me.txtEstimatedPallets.Text)
    IncrementWidth = 9
    IncrementHeight = 6
    Me.lblInfo.Caption = CStr(lngZoomWidth) & "," & CStr(lngZoomHeight)
    Me.lblInfo2.Caption = CStr(Me.Width) & "," & CStr(Me.Height) '800x492 at 76%
    'Me.Width = IncrementWidth
    'Me.Height = IncrementHeight
    
    
    'Me.Width = ZoomWidth
    'Me.Height = ZoomHeight
    'IncrementWidth = CLng(WorksheetFunction.Ceiling((1 / 100) * ZoomWidth, 0))
    'IncrementHeight = CLng(WorksheetFunction.Ceiling((1 / 100) * ZoomHeight, 0))
    'Me.lblResolution.Caption = CStr(ZoomWidth) & "," & CStr(ZoomHeight)
    
    If LastZoomValue < Me.ScrollBar1.value Then
        'User has increased size
        'Me.lblInfo.Caption = "Increse"
        'Me.Width = CLng(Me.Width + IncrementWidth)
        'Me.Height = CLng(Me.Height + IncrementHeight)
        'IncrementWidth = Me.Width - LastWidth
        'IncrementHeight = Me.Height - LastHeight
        Me.Width = Me.Width + IncrementWidth
        Me.Height = Me.Height + IncrementHeight
        'Me.lblInfo2.Caption = "FrameSize:OpBUttons: " & CStr(Me.frameOperationButtons.Left) & "," & CStr(Me.frameOperationButtons.Top)
        'lblInfo.Caption = ">" & CStr(Me.Width) & "," & CStr(Me.Height)
        'lblInfo.Caption = ">" & CStr(IncrementWidth) & "," & CStr(IncrementHeight)
        
        'lblInfo2.Caption = CStr(IncrementWidth) & "," & CStr(IncrementHeight)
    End If
    If LastZoomValue > Me.ScrollBar1.value Then
        'Me.lblInfo.Caption = "Reduced"
        'Me.Width = CLng(Me.Width - IncrementWidth)
        'Me.Height = CLng(Me.Height - IncrementHeight)
        'lblInfo.Caption = "<" & CStr(Me.Width) & "," & CStr(Me.Height)
        'IncrementWidth = LastWidth - Me.Width
        'IncrementHeight = LastHeight - Me.Height
        Me.Width = Me.Width - IncrementWidth
        Me.Height = Me.Height - IncrementHeight
        'lblInfo.Caption = "<" & CStr(IncrementWidth) & "," & CStr(IncrementHeight)
        'lblInfo2.Caption = CStr(IncrementWidth) & "," & CStr(IncrementHeight)
        'Me.lblInfo2.Caption = "FrameSize:OpBUttons: " & CStr(Me.Frame_Questions.Left) & "," & CStr(Me.Frame_Questions.Top)
    End If
        
    LastZoomValue = Me.ScrollBar1.value
    
End Sub

Private Sub txtCBCollar_Change()
    '
End Sub

Private Sub txtCBPalletise_Change()
    '
   
End Sub

Private Sub txtCBSplittersSorterStart5_Change()
    '
End Sub

Private Sub txtCBWrapStrap_Change()
    '
    
End Sub

Function AddControl(ControlType As String, ControlIndex As Long, ControlName As String, strTAG As String, _
        intLeft As Integer, intTop As Integer, intWidth As Integer, strValue As String, Optional ByRef ComboArray As Variant = Nothing, _
        Optional intHeight As Integer = 0, Optional lngBackColor As Long = vbWhite, Optional SelMargin As Boolean = True, _
        Optional MakeBold As Boolean = False, _
        Optional intFontSize As Integer = 10, Optional intTabIndex As Integer = 0, Optional strFontName As String = "Tahoma", _
        Optional MakeVisible As Boolean = True) As Long
    Dim IDX As Long
    Dim comboArr() As String
    Dim CTRL As Control
    
    If UCase(ControlType) = "COMBO" Then
        Set CTRL = Me.Frame_Operatives.Controls.Add("Forms.ComboBox.1", ControlName, MakeVisible)
        Set Lots_Combo(ControlIndex) = CTRL
        
        Lots_Combo(ControlIndex).Top = intTop
        Lots_Combo(ControlIndex).Left = intLeft
        Lots_Combo(ControlIndex).Width = intWidth
        Lots_Combo(ControlIndex).Font.Name = strFontName
        Lots_Combo(ControlIndex).Font.Size = intFontSize
        Lots_Combo(ControlIndex).Font.Bold = MakeBold
        Lots_Combo(ControlIndex).Tag = strTAG
        Lots_Combo(ControlIndex).Text = strValue
        Lots_Combo(ControlIndex).SelectionMargin = SelMargin
        If intTabIndex > 0 Then
            Lots_Combo(ControlIndex).TabIndex = intTabIndex
        End If
        If intHeight > 0 Then
            Lots_Combo(ControlIndex).Height = intHeight
        End If
        Lots_Combo(ControlIndex).BackColor = lngBackColor
        If Not IsArrayEmpty(ComboArray) Then
            For IDX = 0 To UBound(ComboArray)
                If Len(ComboArray(IDX)) > 0 Then
                    Lots_Combo(ControlIndex).AddItem (ComboArray(IDX))
                End If
            Next
        End If
        'Add AFterUpdate Event:
        
        CTRL = Lots_Combo(ControlIndex)
        Set AfterUpdateArr(ControlIndex).comboAfterUpdate = CTRL
        'DELETE BUTTON is NOT removing the control names from the collection ???
        'TESTING - ADD Operative and then DELETE OPERATIVE and then ADD OPERATIVE - giving error that control name already exists
        If Not InCollection("MISSING", CTRLS, CTRL.Name) Then
            CTRLS.Add CTRL, CTRL.Name
            ReDim Preserve Lots_Combo(UBound(Lots_Combo) + 1)
        End If
        If Not InCollection("EMPTY", CTRLS, CTRL.Name) Then
            'CTRLS.Add CTrl, CTrl.Name 'complains that NAME already exists in the collection.
            'ReDim Preserve Lots_Combo(UBound(Lots_Combo) + 1)
        End If
        
        
    End If
    
    
    If UCase(ControlType) = "BTN" Then
        Set CTRL = Me.Frame_Operatives.Controls.Add("Forms.CommandButton.1", ControlName, MakeVisible)
        Set Lots_CmdBtn(ControlIndex) = CTRL
        
        With Lots_CmdBtn(ControlIndex)
            .Top = intTop
            .Left = intLeft
            .Width = intWidth
            .Font.Name = strFontName
            .Font.Size = intFontSize
            .Font.Bold = MakeBold
            .Caption = strValue
            .Tag = strTAG
            .BackColor = lngBackColor
            If intTabIndex > 0 Then
                .TabIndex = intTabIndex
            End If
            If intHeight > 0 Then
                .Height = intHeight
            End If
        End With
        CTRL = Lots_CmdBtn(ControlIndex)
        
        'Getting error here - OBJECT REQUIRED:
        If InStr(1, UCase(CTRL.Name), "START", vbTextCompare) > 0 Then
            'ReDim Preserve cmdbtnTimeStartArr(UBound(cmdbtnTimeStartArr) + 1)
            'Set cmdbtnTimeStartArr(ControlIndex).cbTimeStartEvent = CTrl
            Set btnArray(ControlIndex).cbTimeStartEvent = CTRL
        End If
        If InStr(1, UCase(CTRL.Name), "END", vbTextCompare) > 0 Then
            'ReDim Preserve cmdbtnTimeEndArr(UBound(cmdbtnTimeEndArr) + 1)
            'Set cmdbtnTimeEndArr(ControlIndex).cbTimeEndEvent = CTrl
            Set btnArray(ControlIndex).cbTimeEndEvent = CTRL
        End If
        CTRLS.Add CTRL, CTRL.Name
        ReDim Preserve Lots_CmdBtn(UBound(Lots_CmdBtn) + 1)
        ReDim Preserve btnArray(UBound(btnArray) + 1)
        
    End If
    
    If UCase(ControlType) = "TEXTBOX" Then
        Set CTRL = Me.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
        Set Lots_txtBox(ControlIndex) = CTRL
        With Lots_txtBox(ControlIndex)
            .Top = intTop
            .Left = intLeft
            .Width = intWidth
            .Font.Name = strFontName
            .Font.Size = intFontSize
            .Font.Bold = MakeBold
            .Text = strValue
            .Tag = strTAG
            .SelectionMargin = SelMargin
            If intTabIndex > 0 Then
                .TabIndex = intTabIndex
            End If
            If intHeight > 0 Then
                .Height = intHeight
            End If
        End With
        CTRL = Lots_txtBox(ControlIndex)
        'SET AfterUpdate event:
        Set AfterUpdateArr(ControlIndex).TxtBoxAfterUpdate = CTRL
        CTRLS.Add CTRL, CTRL.Name 'NEED to test if NOT already in the CTRLS collection
        ReDim Preserve Lots_txtBox(UBound(Lots_txtBox) + 1)
    End If
    ReDim Preserve AfterUpdateArr(UBound(AfterUpdateArr) + 1)
    
    AddControl = ControlIndex + 1
    
End Function

Sub AddOperative(ByRef OpID As Long, ByRef TextTAGID As Long, ByVal TimeTAGStart As Long, ByRef ButtonTAGID As Long, ByRef btnIndex As Long, _
    ByRef txtBoxIndex As Long, ByRef comboIndex As Long)
    Dim RowGap As Long
    Dim TopPos As Integer
    Dim ScrollBarHeight As Long
    Dim ComboArray() As String
    Dim TimeTAGID As Long
    
    RowGap = 19
    If OpID = 1 Then
        TopPos = 1
    Else
        TopPos = (OpID - 1) * RowGap
    End If
    'Each Combo and Textbox and Command Button have to be uniquely numbered consequtively -
    'as a limited amount of indexes for the array have been declared for each type of control.
    ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
    
    comboIndex = AddControl("combo", comboIndex, "comOperativeName" & CStr(OpID), CStr(TextTAGID), _
        0, TopPos, 175, "Select Name", ComboArray, 0, vbWhite, True, False, 9, 0, "Tahoma", True)
    TextTAGID = TextTAGID + 1
    ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
    comboIndex = AddControl("combo", comboIndex, "comOperativeActivity" & CStr(OpID), CStr(TextTAGID), _
        175, TopPos, 130, "Select Activity", ComboArray, 0, vbWhite, True, True, 10, 0, "Tahoma", True)
    TextTAGID = TextTAGID + 1
    btnIndex = AddControl("btn", btnIndex, "btnOperativeTimeStart" & CStr(OpID), CStr("btn" & ButtonTAGID), _
        310, TopPos, 20, "@", Nothing, 18, RGB(255, 255, 0), True, True, 8, 0, "Tahoma", True)
        'Associate the button event here to the class: clsTimesheetButtons
        
        
        
    ButtonTAGID = ButtonTAGID + 1
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtCBOperativeTimeStart" & CStr(OpID), CStr(TextTAGID), _
        335, TopPos, 20, "P", Nothing, 0, RGB(255, 255, 0), True, True, 10, 0, "Wingdings 2", True)
    TimeTAGID = TextTAGID + TimeTAGStart
    TextTAGID = TextTAGID + 1
    
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtOperativeTimeStart" & CStr(OpID), CStr(TimeTAGID), _
        360, TopPos, 50, "00:00:00", Nothing, 0, RGB(255, 255, 10), False, True, 10, 0, "Cambria", True)
    
    btnIndex = AddControl("btn", btnIndex, "btnOperativeTimeEnd" & CStr(OpID), CStr("btn" & ButtonTAGID), _
        415, TopPos, 20, "@", Nothing, 18, RGB(255, 255, 20), True, True, 8, 0, "Tahoma", True)
    ButtonTAGID = ButtonTAGID + 1
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtCBOperativeTimeEnd" & CStr(OpID), CStr(TextTAGID), _
        440, TopPos, 20, "P", Nothing, 0, RGB(255, 255, 0), True, True, 10, 0, "Wingdings 2", True)
    TimeTAGID = TextTAGID + TimeTAGStart
    TextTAGID = TextTAGID + 1
    
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtOperativeTimeEnd" & CStr(OpID), CStr(TimeTAGID), _
        465, TopPos, 50, "99:59:59", Nothing, 0, RGB(255, 255, 10), False, True, 10, 0, "Cambria", True)
    'TextTAGID = TextTAGID + 1
    OpID = OpID + 1
    ScrollBarHeight = Me.Frame_Operatives.ScrollHeight
    If OpID > 1 Then
        'roughly 100 = 5 rows
        ScrollBarHeight = ScrollBarHeight + (100 / 5)
        Me.Frame_Operatives.ScrollHeight = ScrollBarHeight
    End If
    'If all indexes are set to 1 initially - then after execution the following values result:
    'txtBox_Index = 5 (1+4)
    'combo_Index = 3 (1+2)
    'cmd_Index = 3 (1+2)
    
End Sub

Sub AddNewOperative(ByRef OpID As Long, ByRef TagID As Long, strDeliveryDate As String, strDeliveryRef As String, ASN As String, _
    ByVal TimeTAGStart As Long, ByRef ButtonTAGID As Long, _
    ByRef btnIndex As Long, ByRef txtBoxIndex As Long, ByRef comboIndex As Long, ByVal ControlFieldname As String)
    
    Dim RowGap As Long
    Dim TopPos As Integer
    Dim ScrollBarHeight As Long
    Dim ComboArray() As String
    Dim TimeTAGID As Long
    Dim ControlText As String
    Dim ControlType As String
    Dim ControlTAG As String
    Dim ControlDate As Date
    Dim ControlLeft As Integer
    Dim ControlTop As Integer
    Dim ControlWidth As Integer
    Dim ControlHeight As Integer
    Dim ControlDeliveryDate As Date
    Dim ControlDeliveryRef As String
    Dim ControlASN As String
    Dim ControlOBJCount As Long
    Dim ControlStartTAG As String
    Dim ControlEndTAG As String
    Dim Dic_Collection As Object
    Dim ControlRowNumber As Long
    Dim ControlTotalRows As Long
    Dim MakeVisible As Boolean
    Dim BackColor As Long
    Dim ControlLeftMargin As Boolean
    Dim ControlFieldnames() As String
    
    RowGap = 19
    If OpID = 1 Then
        TopPos = 1
    Else
        TopPos = (OpID - 1) * RowGap
    End If
    
    If IsDate(strDeliveryDate) Then
        ControlDeliveryDate = CDate(strDeliveryDate)
    Else
        'MsgBox ("Need to pass proper delivery date")
        'Exit Sub
        ControlDeliveryDate = CDate("01/01/1970")
    End If
    ControlDeliveryRef = strDeliveryRef
    
    'controlfieldnames = getfields_from_access()
    
    MakeVisible = True
    ControlType = "COMBOBOX"
    ControlText = "Select Employee"
    ControlTAG = CStr(TagID)
    ControlDate = Now()
    ControlLeft = 0
    ControlTop = TopPos
    ControlWidth = 175
    ControlHeight = 0
    'ControlDeliveryDate = strDeliveryDate
    ControlDeliveryRef = strDeliveryRef
    ControlASN = ASN
    ControlOBJCount = OpID
    ControlStartTAG = "0"
    ControlEndTAG = "0"
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.CompareMode = vbTextCompare
    ControlRowNumber = 0
    ControlTotalRows = 0
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
    
    comboIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "comOperativeName" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    TagID = TagID + 1
    
    ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
    
    ControlType = "COMBOBOX"
    ControlText = "Select Activity"
    ControlTAG = CStr(TagID)
    ControlLeft = 175
    ControlWidth = 130
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    comboIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "comOperativeActivity" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    TagID = TagID + 1
    
    ControlType = "BTN"
    ControlText = "@"
    ControlTAG = "BTN" & CStr(ButtonTAGID)
    ControlLeft = 310
    ControlWidth = 20
    ControlHeight = 20
    BackColor = RGB(255, 255, 0) 'YELLOW ?
    ControlLeftMargin = False
    
    btnIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "btnOperativeTimeStart" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    ButtonTAGID = ButtonTAGID + 1
    
    'txtOperativeTimeStart :
    
    TimeTAGID = TagID + TimeTAGStart
    
    ControlType = "TEXTBOX"
    ControlText = "00:00:00"
    ControlTAG = CStr(TimeTAGID)
    ControlLeft = 335
    ControlWidth = 50
    'ControlHeight = 20
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    txtBoxIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "txtOperativeTimeStart" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    
    ControlType = "BTN"
    ControlText = "@"
    ControlTAG = "BTN" & CStr(ButtonTAGID)
    ControlLeft = 415
    ControlWidth = 20
    'ControlHeight = 20
    ControlText = "@"
    BackColor = RGB(255, 255, 20)
    ControlLeftMargin = False
    
    btnIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "btnOperativeTimeEnd" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    ButtonTAGID = ButtonTAGID + 1
    
    TimeTAGID = TagID + TimeTAGStart
    'TAGID = TAGID + 1
    
    ControlType = "TEXTBOX"
    ControlTAG = CStr(TimeTAGID)
    ControlText = "00:00:00"
    ControlLeft = 465
    ControlWidth = 50
    'ControlHeight = 20
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    txtBoxIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", _
        Nothing, "txtOperativeTimeEnd" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    OpID = OpID + 1
    ScrollBarHeight = Me.Frame_Operatives.ScrollHeight
    If OpID > 1 Then
        'roughly 100 = 5 rows
        'ScrollBarHeight = ScrollBarHeight + (100 / 5)
        ScrollBarHeight = ScrollBarHeight + (OpID * 20)
        Me.Frame_Operatives.ScrollHeight = ScrollBarHeight
    End If
    OperativeCount = OpID
    TextTAGID = TagID


End Sub

Sub AddOperativeToCollection(ByRef OpID As Long, ByRef TextTAGID As Long, ByVal TimeTAGStart As Long, ByRef ButtonTAGID As Long, ByRef btnIndex As Long, _
    ByRef txtBoxIndex As Long, ByRef comboIndex As Long)
    Dim RowGap As Long
    Dim TopPos As Integer
    Dim ScrollBarHeight As Long
    Dim ComboArray() As String
    Dim TimeTAGID As Long
    
    RowGap = 19
    If OpID = 1 Then
        TopPos = 1
    Else
        TopPos = (OpID - 1) * RowGap
    End If
    'Each Combo and Textbox and Command Button have to be uniquely numbered consequtively -
    'as a limited amount of indexes for the array have been declared for each type of control.
    
    'AddNewControl(TheUserform As UserForm, IDPrefix As String, TheControl As Control, ControlName As String, ControlText As String, _
    'ControlType As String, ControlTag As String, ControlDate As Date, _
    'ControlLeft As Integer, ControlTop As Integer, ControlWidth As Integer, ControlHeight As Integer, _
    'ControlDeliveryDate As Date, ControlDeliveryRef As String, Optional ControlASN As String = "", Optional ControlObjCount As Long, _
    'Optional ControlStartTAG As String = "", Optional ControlEndTag As String = "", Optional ByRef Dic_Collection As Scripting.Dictionary, _
    'Optional ControlRowNumber As Long, Optional ControlTotalRows As Long, Optional MakeVisible As Boolean = True, _
    'Optional ByRef ListArray As Variant = Nothing)
    
    
    ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
    
    'comboIndex = AddtheControl(frmGI_TimesheetEntry2_1060x630, "combo", comboIndex, "comOperativeName" & CStr(OpID), CStr(TextTAGID), _
    '    0, TopPos, 175, "Select Name", ComboArray, 0, vbWhite, True, False, 9, 0, "Tahoma", True)
    comboIndex = AddControl("combo", comboIndex, "comOperativeName" & CStr(OpID), CStr(TextTAGID), _
        0, TopPos, 175, "Select Name", ComboArray, 0, vbWhite, True, False, 9, 0, "Tahoma", True)
    TextTAGID = TextTAGID + 1
    
    ComboArray = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
    comboIndex = AddControl("combo", comboIndex, "comOperativeActivity" & CStr(OpID), CStr(TextTAGID), _
        175, TopPos, 130, "Select Activity", ComboArray, 0, vbWhite, True, True, 10, 0, "Tahoma", True)
    TextTAGID = TextTAGID + 1
    
    btnIndex = AddControl("btn", btnIndex, "btnOperativeTimeStart" & CStr(OpID), CStr("btn" & ButtonTAGID), _
        310, TopPos, 20, "@", Nothing, 18, RGB(255, 255, 0), True, True, 8, 0, "Tahoma", True)
        'Associate the button event here to the class: clsTimesheetButtons
        
    ButtonTAGID = ButtonTAGID + 1
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtCBOperativeTimeStart" & CStr(OpID), CStr(TextTAGID), _
        335, TopPos, 20, "P", Nothing, 0, RGB(255, 255, 0), True, True, 10, 0, "Wingdings 2", True)
    TextTAGID = TextTAGID + 1
    TimeTAGID = TextTAGID + TimeTAGStart
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtOperativeTimeStart" & CStr(OpID), CStr(TimeTAGID), _
        360, TopPos, 50, "00:00:00", Nothing, 0, RGB(255, 255, 10), False, True, 10, 0, "Cambria", True)
    
    btnIndex = AddControl("btn", btnIndex, "btnOperativeTimeEnd" & CStr(OpID), CStr("btn" & ButtonTAGID), _
        415, TopPos, 20, "@", Nothing, 18, RGB(255, 255, 20), True, True, 8, 0, "Tahoma", True)
    ButtonTAGID = ButtonTAGID + 1
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtCBOperativeTimeEnd" & CStr(OpID), CStr(TextTAGID), _
        440, TopPos, 20, "P", Nothing, 0, RGB(255, 255, 0), True, True, 10, 0, "Wingdings 2", True)
    TextTAGID = TextTAGID + 1
    TimeTAGID = TextTAGID + TimeTAGStart
    txtBoxIndex = AddControl("textbox", txtBoxIndex, "txtOperativeTimeEnd" & CStr(OpID), CStr(TimeTAGID), _
        465, TopPos, 50, "99:59:59", Nothing, 0, RGB(255, 255, 10), False, True, 10, 0, "Cambria", True)
    'TextTAGID = TextTAGID + 1
    OpID = OpID + 1
    ScrollBarHeight = Me.Frame_Operatives.ScrollHeight
    If OpID > 1 Then
        'roughly 100 = 5 rows
        ScrollBarHeight = ScrollBarHeight + (100 / 5)
        Me.Frame_Operatives.ScrollHeight = ScrollBarHeight
    End If
    'If all indexes are set to 1 initially - then after execution the following values result:
    'txtBox_Index = 5 (1+4)
    'combo_Index = 3 (1+2)
    'cmd_Index = 3 (1+2)


End Sub

Sub DeleteControls(TheControls As Controls, OpID As Long, Optional IDX As Long = 0)
    Dim CTRL As Control
    Dim FoundCtrl As Boolean
    Dim OperativeName As String
    Dim CtrlName As String
    Dim CTRLIDX As Long
    Dim varItem As Variant
    Dim ControlID As Variant
    
    FoundCtrl = False
    If IDX > 0 Then
        TheControls.Remove CTRLS.Item(IDX).Name
    End If
    
    If OpID > 0 Then
        
        For CTRLIDX = ctrlCollection.Count To 1 Step -1
            'Getting Catastrophic failure - when adding a control then delete and add and delete again:
            '2 rows left - so total = 16 controls . BUT CTRLS had 24 control ITEMS allocated. CTRlIDX = 24 BUT no control object present in these 16-24.
            'so the CTRLS.COUNT does not mean that there are ACTUALLY that many controls - some are EMPTY elements not defined.
            'so for example CTRLS.COUNT reports 24 elements. BUT only 16 are actually allocated with valid OBJECTS.
            If Not InCollection("MISSING", ctrlCollection, ctrlCollection.Item(CTRLIDX)) Then 'Not assigned so remove whole item from collection.
                Set varItem = ctrlCollection(CTRLIDX)
                CtrlName = varItem.ControlName
                ControlID = varItem.ControlID
                If InStr(1, CtrlName, CStr(OpID), vbTextCompare) > 0 Then
                    FoundCtrl = True
                    'On Error Resume Next
                    If Not TheControls.Item(CtrlName) Is Nothing Then
                        TheControls.Remove CtrlName
                        ctrlCollection.Remove ControlID
                    End If
                End If
            End If
        Next
    End If
End Sub

Sub RemoveAllControls(TheControls As Controls)
    On Error Resume Next
    TheControls.Clear
End Sub

Sub TestTodayImport(Optional ByRef SearchCriteria As String = "", Optional DBTable As String = "tblDeliveryInfo")
    Dim ListArr() As String
    Dim TodaySearch As String
    Dim dtCurrentDate As Date
    
    dtCurrentDate = Now()
    TodaySearch = "[DeliveryDate] = " & "#" & Format(dtCurrentDate, "yyyy/mm/dd") & "#"
    
    
    If Len(SearchCriteria) = 0 Then
        SearchCriteria = TodaySearch
    End If
    ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessDBpath, 2, "", SearchCriteria, True)
    If IsArrayEmpty(ListArr, 1) Then
        'Today Data not available - need to update
        Me.btnImportData.BackColor = vbRed
    Else
        'make searchcriteria search and return all references for the current month / week ???
        Me.btnImportData.BackColor = vbGreen
    End If
End Sub

Private Sub txtDeliveryRef_AfterUpdate()
    'How could we capture any AFTER UPDATE event for all the comboboxes and TextBoxes ?
    'Needs to be one event procedure - that captures which control has just been updated.
    'Then it can be written into the collection.
    'Otherwise - it would mean adding code to every AfterUpdate event for every control on the form manually !
    
    'YES USE THE CLASS MODULE - THE TIME BUTTON CLICK EVENTS ARE THE SAME TECNIQUE - EXCEPT AFTER UPDATE EVENT FOR BOTH TEXTBOXES AND COMBOBOXES
    ' - BUT TREATED SEPARATELY AS THEY ARE THEIR OWN CONTROLS WITH THEIR OWN SLIGHTLY DIFFERENT PROPERTIES.
    
End Sub

Private Sub txtPalletsArrived_Change()
    Dim ErrMessage As String
    Dim NewEntry As String
    Dim FoundControl As Boolean
    Dim ReturnValue As Variant
    
    NewEntry = Me.txtPalletsArrived.Text
    
    If Len(Me.txtPalletsArrived) > 3 Then
        Call UpdateControlCollection(Me.txtDeliveryDate, Me.txtDeliveryRef, "", "txtPalletsArrived", ctrlCollection, txtPalletsArrived, NewEntry, ErrMessage)
        If Len(ErrMessage) > 0 Then
            MsgBox ("Error in UPdATe: " & ErrMessage)
        End If
        'FoundControl = ReturnControlInfo(Me.txtDeliveryDate, Me.txtDeliveryRef, "", "", Me.txtPalletsArrived.Name, "NAME", "", "ControlValue", ReturnValue)
        
        MsgBox ("return value = " & ReturnValue)
        
    End If
    
    
End Sub

Private Sub UserForm_Initialize()
    '**************************************** INITIALIZE USERFORM HERE **************************************************************
    'SO when the userform first pops up : ALL FIELDS DISABLED except txtASNNO and txtDeliveryReference.
    'These fields are used to search for an existing record.
    'When the user tabs out of one of these fields - or clicks another field / or exits the field as an event :
    'A search will be performed on the entry.
    'Will return either - "Record does not exist / ASN does not exist" or it will display the top part or both if they exist.
    'Call EnableDisableControls(False) 'Disable ALL Controls.
    'Call EnableDisableControls(True, 200, 201) 'Enable 2 combo boxes
    'Call EnableDisableControls(True, 8, 11, "COMBOBOX")
    'Call EnableDisableControls(True, 38, 41, "BUTTONS") 'Enable Buttons - Operation
    'Need to know the screen resolution so that the correct userform size can be set to start with.
    'Original size => Width  = 1060, Height = 650
    
    Dim ListArr() As String
    Dim Activities() As String
    Dim IDX As Long
    Dim Screens As Integer
    Dim Screenwidth As Long
    Dim ScreenHeight As Long
    Dim WorkingWidth As Integer
    Dim WorkingHeight As Integer
    Dim ErrMessage As String
    Dim CTRLIDX As Long
    Dim StartHeight As Long
    Dim StartTAG As Long
    Dim btnTOP1 As Long
    Dim btnTOP2 As Long
    Dim txtTOP1 As Long
    Dim txtTOP2 As Long
    Dim txtTOP3 As Long
    Dim txtTOP4 As Long
    Dim ret As String
    Dim NewPrefsSheet As Worksheet
    Dim RecordSheet As Worksheet
    Dim ConnectWB As Workbook
    Dim NextControl As Long
    Dim DBTable As String
    Dim cmdbtnTimeStartArr() As New clsTimesheetButtons
    Dim cmdbtnTimeEndArr() As New clsTimesheetButtons
    Dim SearchCriteria As String
    Dim dtCurrentDate As Date
    Dim strToday As String
    Dim strFormCaption As String
    Dim strTitleCaption As String
    
    'Listarr = PopulateDropdowns("Timesheet Records", 2)
    'For Idx = 0 To UBound(Listarr)
    '    If Len(Listarr(Idx)) > 0 Then
    '        Me.comDeliveryRef.AddItem (Listarr(Idx))
    '    End If
    'Next
    Set CTRLS = New Collection
    
    Call GetVersion(strTitleCaption, strFormCaption)
    
    Me.lblTitle_GoodsInTimesheet.Caption = "Goodsin Timesheet " & strTitleCaption
    Me.Caption = "Goods In Timesheet " & strFormCaption
    
    DBTable = "tblDeliveryInfo"
    Me.txtImportDate.Text = Format(Now(), "dd/mm/yyyy")
    
    'Set dic_ControlCollection = CreateObject("Scripting.Dictionary")
    dic_ControlCollection.RemoveAll
    
    Me.lblProgress.Width = 0
    MainWorksheet = "Timesheet Records"
    CopiedDataSheet = "GI Data"
    ComplianceQuestion1TAG = "142"
    ComplianceQuestion2TAG = "143"
    ComplianceQuestion3TAG = "144"
    ComplianceQuestion4TAG = ""
    ComplianceQuestion5TAG = ""
    FurtherCommentsTAG = "145"
    Me.txtFurtherComments.Tag = FurtherCommentsTAG
    Me.txtArrivedONTime.Tag = ComplianceQuestion1TAG
    Me.txtArrivedONTimeComment.Tag = "0"
    Me.txtIsItSafe.Tag = ComplianceQuestion2TAG
    Me.txtIsItSafeComment.Tag = "0"
    Me.txtCompleted.Tag = ComplianceQuestion3TAG
    Me.txtCompletedComment.Tag = "0"
    
    'To get Current number of textboxes and combos needed for X operatives:
    'Get total number of child records in the tblLabourHours table for the selected Delivery Ref.
    'First method - add dynamically to the form - with tag as read from each record -
    ' - each record has the NAME of the operative. ACTIVITY and START TIME and END TIME and TAG NUMBER with DELIVERY REF.
    '2nd method - create 30 textboxes at design time - but make them hidden and make the scrollbar value change according
    ' - to the current number of combos / operatives selected - roughly 100 per 5 controls - so 400-500 for 20 controls.
    ' make them visible as needed.
    'keypreview = True
    
    'Call SortComboBox(Me.comASNNO)
    Call ClearEntry(1, 16)
    Call ClearEntry(17, TotalFields)
    Call ClearEntry(441, 442)
    Call EnableDisableControls(False, 17, TotalFields, "COMBOBOX")
    Call EnableDisableControls(False, 17, TotalFields, "TEXTBOX")
    Call EnableDisableControls(False, 8, 61, "BUTTONS") 'not correct - check.
    'Clear array first ?
    Me.comSuppliers.Clear
    ListArr = PopulateDropdowns("Suppliers", 2, 0, False, WB_MainTimesheetData)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comSuppliers.AddItem (ListArr(IDX))
        End If
    Next
    'Call SortComboBox(Me.comFLMs)
    Me.comFLMs.Clear
    ListArr = PopulateDropdowns("FLM", 2, 0, False, WB_MainTimesheetData) 'GET data from the FLM sheet on column 2
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            'Me.comFLMs.AddItem (Listarr(IDX))
        End If
    Next
    'Me.comUnloadTipperName.Clear
    ListArr = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comFLMs.AddItem (ListArr(IDX))
        End If
    Next
    Activities = PopulateDropdowns("Activities", 1, 0, True, WB_MainTimesheetData)
    
    txtBox_Index = 1
    combo_Index = 1
    cmd_Index = 1
    OperativeCount = 1
    TextTAGID = 43
    TimeTAGID = 400
    btnTAGID = 10
    ShortTAGID = 1
    ExtraTAGID = 1
    ShortCount = 1
    ExtraCount = 1
    
    ReDim Lots_Combo(1)
    ReDim Lots_CmdBtn(1)
    ReDim Lots_txtBox(1)
    ReDim cmdbtnTimeStartArr(1)
    ReDim cmdbtnTimeEndArr(1)
    ReDim btnArray(1)
    ReDim AfterUpdateArr(1)
    
    'Call AddOperative(OperativeCount, TextTAGID, TimeTAGID, btnTAGID, cmd_Index, txtBox_Index, combo_Index)
    'Call AddNewOperatives(OperativeCount, TextTAGID, "01/01/1970", "0", "", TimeTAGID, btnTAGID)
    
    'Multiply Height x 19
    'NextIndex = AddControl(controlType As String, ControlIndex As Long, ControlName As String, strTAG As String, _
        intLeft As Integer, intTop As Integer, intWidth As Integer, strValue As String, Optional ByRef comboArray As Variant = Nothing, _
        Optional intHeight As Integer = 0, Optional lngBackColor As Long = vbWhite, Optional SelMargin As Boolean, Optional MakeBold As Boolean = False, _
        Optional intFontSize As Integer = 10, Optional intTabIndex As Integer = 0, Optional strFontName As String = "Tahoma", _
        Optional MakeVisible As Boolean = True) As Long
    Call GetScreenResolution(Screenwidth, ScreenHeight)
    Me.lblResolution.Caption = "Res: " & CStr(Screenwidth) & "x" & CStr(ScreenHeight)
    
    Me.ScrollBar1.Max = 200
    Me.ScrollBar1.Min = 10
    Me.ScrollBar1.value = 100
    Me.ScrollBar1.SmallChange = 1
    Me.ScrollBar1.LargeChange = 2
    
    LastZoomValue = 100
    
    If MainGIModule_v1_1.sheetExists("Prefs") Then
        If ScreenHeight < 1080 Then
            SetZoom = ThisWorkbook.Worksheets("Prefs").Cells(2, 2).value
            SetWidth = ThisWorkbook.Worksheets("Prefs").Cells(3, 2).value
            SetHeight = ThisWorkbook.Worksheets("Prefs").Cells(4, 2).value
            
            If Len(SetZoom) > 0 Then
                'Me.Zoom = SetZoom
                'Need to test the current resolution:
                
                Me.ScrollBar1.value = SetZoom
            Else
                Me.Zoom = Me.ScrollBar1.value
            End If
            If Len(SetWidth) > 0 Then
                Me.Width = SetWidth
            End If
            If Len(SetHeight) > 0 Then
                Me.Height = SetHeight
            End If
        End If
    End If
    
    Me.Zoom = ScrollBar1.value
    LastZoomValue = ScrollBar1.value
    lblZoom = "Zoom: " & CStr(ScrollBar1.value) & " %"
    
    If MainGIModule_v1_1.sheetExists("Prefs", ThisWorkbook) Then
        ret = ThisWorkbook.Worksheets("Prefs").Cells(1, 2).value
        If Len(ret) > 0 Then
            'TEST if file exists:
            If Test_File_Exist_With_Dir(ret) Then
                AccessDBpath = ret
                Application.DisplayAlerts = False
                Me.btnSelectRecordSheetLocation.BackColor = vbGreen
            Else
                MsgBox ("Cannot Find DATA file - click the RE-Select Data button at the top")
                Exit Sub
            End If
        Else
            Me.btnSelectRecordSheetLocation.BackColor = vbRed
            MsgBox ("Remote shared ACCESS database NOT specified.")
            'create local database - if not exist on local machine ?
            
        End If
    Else
        
        'create new PREFS sheet and ask user to locate data:
        Set NewPrefsSheet = ThisWorkbook.Worksheets.Add
        NewPrefsSheet.Name = "Prefs"
        MsgBox ("Prefs sheet recreated: Please choose location for ACCESS DATABASE")
        RemoteFilePath = Connect_ACCESS_DB("Prefs") 'Sets the workbook to the shared database file.
        If Len(RemoteFilePath) = 0 Then
            MsgBox ("Cancelled Getting Remote database")
        Else
            If Test_File_Exist_With_Dir(RemoteFilePath) Then
                'MsgBox ("Please close form and re-open")
                'Exit Sub
                Me.btnSelectRecordSheetLocation.BackColor = vbGreen
            Else
                MsgBox ("Could not find path")
                Me.btnSelectRecordSheetLocation.BackColor = vbRed
                Exit Sub
            End If
        End If
        'Call MakeShared(WB_MainTimesheetData)
    End If
    
    'LOAD the ASN selection combo with ALL the ASNs entered into the record sheet.
    dtCurrentDate = Now()
    SearchCriteria = "[DeliveryDate] = " & "#" & Format(dtCurrentDate, "yyyy/mm/dd") & "#"
    Me.comDeliveryRef.Clear
    
    Call TestTodayImport(SearchCriteria)
    
    ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessDBpath, 2, "", "", True)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comDeliveryRef.AddItem (ListArr(IDX))
        End If
    Next
    Me.comASNNo.Clear
    'ListArr = PopulateDropdowns(MainWorksheet, 4, 0, True, WB_MainTimesheetData)
    ListArr = PopulateDropdowns_From_ACCESS(DBTable, AccessDBpath, 4, "", "", True)
    For IDX = 0 To UBound(ListArr)
        If Len(ListArr(IDX)) > 0 Then
            Me.comASNNo.AddItem (ListArr(IDX))
        End If
    Next
    
    'Call Execute_SaveTAblesAndFieldsToLookup
    'Call Execute_UpdateControlNamesInLookup
    
    
End Sub
