Attribute VB_Name = "MainGIModule_v1_1"
Option Explicit
'Goods In Timesheet Main Program - written by Daniel Goss 2018 version RP7.32
Public CurrentRecordRow As Long
Public DeliveryDate As Date
Public ImportDate As Date
Public TotalFields As Long
Public MainWorksheet As String
Public CopiedDataSheet As String
Public LastWidth As Long
Public LastZoomValue As Integer
Public LastHeight As Long
Public TotalBlanks As Long
Public btnSplitters(10) As MSForms.CommandButton
Public comSplitters(5) As MSForms.ComboBox
Public txtSplitters(20) As MSForms.TextBox
Public PercentDone As Long
Public SaveAndContinue As Boolean
Public TESTING As Boolean
Public WB_MainTimesheetData As Workbook
Public RemoteFilePath As String
Public AccessDBpath As String
Public FurtherCommentsTAG As String
Public ComplianceQuestion1TAG As String
Public ComplianceQuestion2TAG As String
Public ComplianceQuestion3TAG As String
Public ComplianceQuestion4TAG As String
Public ComplianceQuestion5TAG As String
Public ComplianceQuestion6TAG As String
Public Lots_Combo() As MSForms.ComboBox
Public Lots_CmdBtn() As MSForms.CommandButton
Public Lots_txtBox() As MSForms.TextBox
Public btnArray() As New clsTimesheetButtons
Public AfterUpdateArr() As New clsAfterUpdate
Public txtBox_Index As Long
Public combo_Index As Long
Public cmd_Index As Long
Public OperativeCount As Long
Public ShortCount As Long
Public ShortTAGID As Long
Public ExtraCount As Long
Public ExtraTAGID As Long
Public TextTAGID As Long
Public TimeTAGID As Long
Public btnTAGID As Long
Public CTRLS As Collection
Public dic_ControlCollection As New Scripting.Dictionary
Public ctrlCollection As New Collection
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
Public Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Public Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_SNAPSHOT = 44
Public Const VK_LMENU = 164
Public Const KEYEVENTF_KEYUP = 2
Public Const KEYEVENTF_EXTENDEDKEY = 1
Public Const USERVERSION = "v1.1"
Public Const SYSTEM_VERSION = "RP7.3"
Public Const FirstInboundTAG As Long = 1
Public Const LastInboundTAG As Long = 18
Public Const FirstOperationTAG As Long = 19
Public Const LastOperationTAG As Long = 46
Public Const FirstTimeTAG As Long = 441
Public Const LastTimeTAG As Long = 446 'includes one row in the Operatives Frame.
Public Const FirstSCTAG As Long = 801
Public Const LastSCTAG As Long = 807
Public Const FirstStaticBtnTAG As Long = 1
Public Const LastStaticBtnTAG As Long = 7
Public Const FirstOpBtnTAG As Long = 8
Public Const LastOpBtnTAG As Long = 149
Public Const FirstOtherBtnTAG As Long = 150
Public Const LastOtherBtnTAG As Long = 169

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

Sub Auto_open()

    Call StartEntry
 
End Sub

Public Function InCollection(ErrType As String, col As Collection, key As Variant, Optional ByRef IsInCollection As Boolean = False) As Boolean
    Dim var As Variant
    Dim errNumber As Long

    InCollection = False
    Set var = Nothing

    Err.Clear
    On Error Resume Next
    var = IsObject(col(key))
    IsInCollection = var
    
    'var = col.Item(key)
    errNumber = CLng(Err.Number)
    On Error GoTo 0
    'hmmm yea 5 if KEY totally NOT in collection but what if the ITEM element still exists with nothing in it - not defined ?
    '5 is not in, 0 and 438 represent incollection
    If UCase(ErrType) = "EMPTY" Then 'ITEM not set to any object - control has been removed
        If errNumber = -2147418113 Then  ' it is 5 if not in collection
            InCollection = False
        Else
            InCollection = True
        End If
    End If
    If UCase(ErrType) = "MISSING" Then 'VAR = NOTHING ; the key - being a string = empty ; len=0 - refers to TIME textbox - is still in collection.
        If errNumber = 5 Then ' it is 5 if not in collection
            InCollection = False
        ElseIf errNumber = 91 Then
            InCollection = False
        Else
            InCollection = True
        End If
    End If
End Function

Sub GetVersion(ByRef strUserVersion As String, ByRef strSystemVersion As String)

strUserVersion = USERVERSION
strSystemVersion = SYSTEM_VERSION

End Sub

Sub AddControlInfo(ControlLastSaved As Date, ControlTable As String, ControlFieldname As String, TotalObjects As Long, StartIDX As Long, IDPrefix As String, _
    TheControl As Control, ControlName As String, ControlText As String, ControlType As String, ControlTAG As String, ControlDate As Date, _
    ControlLeft As Integer, ControlTop As Integer, ControlWidth As Integer, ControlHeight As Integer, _
    ControlDeliveryDate As Date, ControlDeliveryRef As String, Optional ControlASN As String = "", Optional ControlOBJCount As Long, _
    Optional ControlStartTAG As String = "", Optional ControlEndTAG As String = "", Optional ByRef Dic_Collection As Scripting.Dictionary, _
    Optional ControlRowNumber As Long, Optional ControlTotalRows As Long, Optional MakeVisible As Boolean = True, _
    Optional ByRef ListArray As Variant = Nothing, Optional clsControlObject As clsControls = Nothing)
    
    Dim tempControl As clsControls
    Dim i As Long
    
    Dim CollectionKey As Variant
    
    'Set dic_ControlCollection = CreateObject("Scripting.Dictionary")
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Set dic_ControlCollection = New Scripting.Dictionary
    'Set Dic_Collection = New Scripting.Dictionary
    dic_ControlCollection.CompareMode = TextCompare
    Dic_Collection.CompareMode = TextCompare
    'Dic_Collection.RemoveAll
    
    'This procedure will save the Normal Controls
    'Objective - to have all code for the Control COLLECTION / Script Dictionary OBJECT all in one place.
    ' - maybe all the above parameters could also just be collected under one array and passed in ?????????? just a number rather than individual variables ???
    '.SelectionMargin = SelMargin
    '.TabIndex = intTabIndex
    '.BackColor = lngBackColor
    
    For i = 1 To TotalObjects
        Set tempControl = New clsControls
        If Not clsControlObject Is Nothing Then
            tempControl = clsControlObject
        Else
            tempControl.ControlDBTable = ControlTable
            tempControl.ControlFieldname = ControlFieldname
            tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
            tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
            tempControl.TheControl = TheControl
            tempControl.ControlName = ControlName
            tempControl.ControlValue = ControlText
            tempControl.ControlType = ControlType
            tempControl.ControlTAG = ControlTAG
            tempControl.ControlDate = ControlDate
            tempControl.ControlLeftPos = ControlLeft
            tempControl.ControlTopPos = ControlTop
            tempControl.ControlWidthPos = ControlWidth
            tempControl.ControlHeightPos = ControlHeight
            tempControl.ControlDeliveryDate = ControlDeliveryDate
            tempControl.ControlDeliveryRef = ControlDeliveryRef
            tempControl.ControlASNNumber = ControlASN
            tempControl.ControlObjNumber = ControlOBJCount + i
            tempControl.ControlStartTAG = ControlStartTAG
            tempControl.ControlEndTAG = ControlEndTAG
            tempControl.ControlRowNumber = ControlRowNumber
            tempControl.ControlTotalRows = ControlTotalRows
            If UCase(ControlType) = "TEXTBOX" Then
                Set tempControl.TxtBoxAfterUpdate = TheControl
            End If
            If UCase(ControlType) = "COMBOBOX" Then
                Set tempControl.comboAfterUpdate = TheControl
            End If
            If Len(ControlLastSaved) = 0 Then
                ControlLastSaved = CDate("01/01/1970")
            Else
                tempControl.ControlLastSaved = ControlLastSaved
            End If
        End If
                
        CollectionKey = tempControl.ControlTAG
        If Len(ControlDeliveryRef) > 0 Then
            CollectionKey = ControlDeliveryDate & "_" & ControlDeliveryRef & "_" & tempControl.ControlTAG
        ElseIf Len(ControlASN) > 0 Then
            CollectionKey = ControlDeliveryDate & "_" & ControlASN & "_" & tempControl.ControlTAG
        End If
        'ctrlCollection.Add tempControl, tempControl.ControlTag
        If Not InCollection("MISSING", ctrlCollection, CollectionKey) Then
            ctrlCollection.Add tempControl, CollectionKey
        End If
        If Not Dic_Collection.Exists(CollectionKey) Then
            Set Dic_Collection(CollectionKey) = tempControl
        End If
        If Not dic_ControlCollection.Exists(CollectionKey) Then
            Set dic_ControlCollection(CollectionKey) = tempControl
        End If
    Next i
    'could add a method in the class to find a control - by its ID or by its NAME or any property listed above.
    'The method would return either the actual control OR the control NAME OR the control TAG.
    'hmmm would have to be at Module level - as this would involve the COLLECTION - and its INDEX to the object / control being searched.
    
    
    
End Sub

Function AddNewControl(IsNewControl As Boolean, TheControls As Controls, ControlFieldname As String, IDPrefix As String, TheControl As Control, ControlName As String, _
    ControlText As String, ControlType As String, ControlTAG As String, ControlDate As Date, _
    ControlLeft As Integer, ControlTop As Integer, ControlWidth As Integer, ControlHeight As Integer, _
    ControlDeliveryDate As Date, ControlDeliveryRef As String, Optional ControlASN As String = "", Optional ControlOBJCount As Long, _
    Optional ControlStartTAG As String = "", Optional ControlEndTAG As String = "", Optional ByRef Dic_Collection As Scripting.Dictionary, _
    Optional ControlRowNumber As Long, Optional ControlTotalRows As Long, Optional MakeVisible As Boolean = True, _
    Optional ByRef ListArray As Variant = Nothing, Optional ControlBACKCOLOUR As Long = vbWhite, Optional ControlAddLeftMargin As Boolean = True) As Long
    
    Dim tempControl As clsControls
    Dim NewCtrl As Control
    Dim IDX As Long
    Dim CollectionKey As Variant
    
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.RemoveAll
    Set tempControl = New clsControls
    
    If UCase(ControlType) = "COMBOBOX" Then
        If IsNewControl Then
            Set NewCtrl = TheControls.Add("Forms.ComboBox.1", ControlName, MakeVisible)
            'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.ComboBox.1", ControlName, MakeVisible)
            NewCtrl.Name = ControlName
            NewCtrl.Top = ControlTop
            NewCtrl.Left = ControlLeft
            NewCtrl.Width = ControlWidth
            NewCtrl.Font.Name = "Cambria"
            NewCtrl.Font.Size = 11
            NewCtrl.Tag = ControlTAG
            NewCtrl.Text = ControlText
            NewCtrl.BackColor = ControlBACKCOLOUR
            NewCtrl.SelectionMargin = ControlAddLeftMargin
            If ControlHeight > 0 Then
                NewCtrl.Height = ControlHeight
            End If
            If Not IsArrayEmpty(ListArray) Then
                For IDX = 0 To UBound(ListArray)
                    If Len(ListArray(IDX)) > 0 Then
                        NewCtrl.AddItem (ListArray(IDX))
                    End If
                Next
            End If
            'DELETE BUTTON is NOT removing the control names from the collection ???
            'TESTING - ADD Operative and then DELETE OPERATIVE and then ADD OPERATIVE - giving error that control name already exists
            If Not InCollection("MISSING", CTRLS, NewCtrl.Name) Then
                'CTRLS.Add ctrl, NewCtrl.Name
                'ReDim Preserve Lots_Combo(UBound(Lots_Combo) + 1)
            End If
            If Not InCollection("EMPTY", CTRLS, NewCtrl.Name) Then
                'CTRLS.Add CTrl, CTrl.Name 'complains that NAME already exists in the collection.
                'ReDim Preserve Lots_Combo(UBound(Lots_Combo) + 1)
            End If
            Set tempControl.comboAfterUpdate = NewCtrl
            tempControl.TheControl = NewCtrl 'SET already included in the LET procedure within the class module.
        End If 'might need else here instead ?
        'For Each Param In tempControl
         '   MsgBox (Param)
        'Next
        tempControl.ControlName = ControlName
        tempControl.ControlFieldname = ControlFieldname
        tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
        tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
        tempControl.ControlValue = ControlText
        tempControl.ControlType = ControlType
        tempControl.ControlTAG = ControlTAG
        tempControl.ControlDate = ControlDate
        tempControl.ControlLeftPos = ControlLeft
        tempControl.ControlTopPos = ControlTop
        tempControl.ControlWidthPos = ControlWidth
        tempControl.ControlHeightPos = ControlHeight
        tempControl.ControlDeliveryDate = ControlDeliveryDate
        tempControl.ControlDeliveryRef = ControlDeliveryRef
        tempControl.ControlASNNumber = ControlASN
        tempControl.ControlObjNumber = ControlOBJCount
        tempControl.ControlStartTAG = ControlStartTAG
        tempControl.ControlEndTAG = ControlEndTAG
        tempControl.ControlRowNumber = ControlRowNumber
        tempControl.ControlTotalRows = ControlTotalRows
        tempControl.ControlBACKCOLOUR = ControlBACKCOLOUR
        tempControl.ControlLeftMargin = ControlAddLeftMargin
        
        'end if
    End If
    
    
    If UCase(ControlType) = "BTN" Then
        'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.CommandButton.1", ControlName, MakeVisible)
        Set NewCtrl = TheControls.Add("Forms.CommandButton.1", ControlName, MakeVisible)
        'Set Lots_CmdBtn(ControlIndex) = ctrl
        NewCtrl.Top = ControlTop
        NewCtrl.Left = ControlLeft
        NewCtrl.Width = ControlWidth
        NewCtrl.Font.Name = "Cambria"
        NewCtrl.Font.Size = 11
        NewCtrl.Tag = ControlTAG
        NewCtrl.Caption = ControlText
        NewCtrl.BackColor = ControlBACKCOLOUR
        'NewCtrl.SelectionMargin = ControlAddLeftMargin 'NOT applicable to BUttons !
        
        If ControlHeight > 0 Then
            NewCtrl.Height = ControlHeight
        End If
        'Getting error here - OBJECT REQUIRED:
        If InStr(1, UCase(NewCtrl.Name), "START", vbTextCompare) > 0 Then
            'ReDim Preserve cmdbtnTimeStartArr(UBound(cmdbtnTimeStartArr) + 1)
            'Set cmdbtnTimeStartArr(ControlIndex).cbTimeStartEvent = CTrl
            'Set btnArray(ControlIndex).cbTimeStartEvent = ctrl
            Set tempControl.cbTimeStartEvent = NewCtrl
        End If
        If InStr(1, UCase(NewCtrl.Name), "END", vbTextCompare) > 0 Then
            'ReDim Preserve cmdbtnTimeEndArr(UBound(cmdbtnTimeEndArr) + 1)
            'Set cmdbtnTimeEndArr(ControlIndex).cbTimeEndEvent = CTrl
            'Set btnArray(ControlIndex).cbTimeEndEvent = ctrl
            Set tempControl.cbTimeEndEvent = NewCtrl
        End If
        tempControl.TheControl = NewCtrl 'SET NOT NEEDED HERE - part of internal procedure.
        tempControl.ControlName = ControlName
        tempControl.ControlFieldname = ControlFieldname
        tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
        tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
        tempControl.ControlValue = ControlText
        tempControl.ControlType = ControlType
        tempControl.ControlTAG = ControlTAG
        tempControl.ControlDate = ControlDate
        tempControl.ControlLeftPos = ControlLeft
        tempControl.ControlTopPos = ControlTop
        tempControl.ControlWidthPos = ControlWidth
        tempControl.ControlHeightPos = ControlHeight
        tempControl.ControlDeliveryDate = ControlDeliveryDate
        tempControl.ControlDeliveryRef = ControlDeliveryRef
        tempControl.ControlASNNumber = ControlASN
        tempControl.ControlObjNumber = ControlOBJCount
        tempControl.ControlStartTAG = ControlStartTAG
        tempControl.ControlEndTAG = ControlEndTAG
        tempControl.ControlRowNumber = ControlRowNumber
        tempControl.ControlTotalRows = ControlTotalRows
        tempControl.ControlBACKCOLOUR = ControlBACKCOLOUR
        tempControl.ControlLeftMargin = ControlAddLeftMargin
        
    End If
    
    If UCase(ControlType) = "TEXTBOX" Then
        'Set NewCtrl = TheUserform.Frame_Operatives.Controls.Add("Forms.Textbox.1", ControlName, MakeVisible)
        Set NewCtrl = TheControls.Add("Forms.Textbox.1", ControlName, MakeVisible)
        NewCtrl.Top = ControlTop
        NewCtrl.Left = ControlLeft
        NewCtrl.Width = ControlWidth
        NewCtrl.Font.Name = "Cambria"
        NewCtrl.Font.Size = 11
        NewCtrl.Tag = ControlTAG
        NewCtrl.Text = ControlText
        NewCtrl.BackColor = ControlBACKCOLOUR
        NewCtrl.SelectionMargin = ControlAddLeftMargin
        If ControlHeight > 0 Then
            NewCtrl.Height = ControlHeight
        End If
        tempControl.TheControl = NewCtrl
        tempControl.ControlName = ControlName
        tempControl.ControlFieldname = ControlFieldname
        tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
        tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
        tempControl.ControlValue = ControlText
        tempControl.ControlType = ControlType
        tempControl.ControlTAG = ControlTAG
        tempControl.ControlDate = ControlDate
        tempControl.ControlLeftPos = ControlLeft
        tempControl.ControlTopPos = ControlTop
        tempControl.ControlWidthPos = ControlWidth
        tempControl.ControlHeightPos = ControlHeight
        tempControl.ControlDeliveryDate = ControlDeliveryDate
        tempControl.ControlDeliveryRef = ControlDeliveryRef
        tempControl.ControlASNNumber = ControlASN
        tempControl.ControlObjNumber = ControlOBJCount
        tempControl.ControlStartTAG = ControlStartTAG
        tempControl.ControlEndTAG = ControlEndTAG
        tempControl.ControlRowNumber = ControlRowNumber
        tempControl.ControlTotalRows = ControlTotalRows
        tempControl.ControlBACKCOLOUR = ControlBACKCOLOUR
        tempControl.ControlLeftMargin = ControlAddLeftMargin
        
        Set tempControl.TxtBoxAfterUpdate = NewCtrl
    End If
    
    CollectionKey = tempControl.ControlTAG
    If Len(ControlDeliveryDate) > 0 Then
        CollectionKey = ControlDeliveryDate & "_" & ControlDeliveryRef & "_" & tempControl.ControlTAG
    ElseIf Len(ControlASN) > 0 Then
        CollectionKey = ControlDeliveryDate & "_" & ControlASN & "_" & tempControl.ControlTAG
    End If
    'ctrlCollection.Add tempControl, tempControl.ControlTag
    If Not InCollection("MISSING", ctrlCollection, CollectionKey) Then
        ctrlCollection.Add tempControl, CollectionKey
    End If
    If Not Dic_Collection.Exists(CollectionKey) Then
        Dic_Collection.Add CollectionKey, tempControl
    End If
    
    AddNewControl = AddNewControl + 1
    
End Function


Sub AddLabourInfo(TotalObjects As Long, StartIDX As Long, IDPrefix As String, TheControl As Control, ControlName As String, ControlText As String, _
    ControlType As String, ControlTAG As String, ControlDate As Date, _
    ControlDeliveryDate As Date, ControlDeliveryRef As String, ControlASN As String, ControlOBJCount As Long, Optional ByRef Dic_Collection As Scripting.Dictionary, _
    Optional ControlStartTAG As String = "", Optional ControlEndTAG As String = "", Optional ControlRowNumber As Long = 0, _
    Optional ControlTotalRows As Long = 0, _
    Optional ControlOpName As String, _
    Optional ControlActivity As String, _
    Optional ControlStartDateTime As Date, _
    Optional ControlEndDateTime As Date, _
    Optional ControlPartNo As String, _
    Optional ControlQty As Long, _
    Optional ControlExtraShort As String)
    
    Dim tempControl As clsControls
    Dim i As Long
    Dim CollectionKey As Variant
    
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.RemoveAll
    
    
    For i = 1 To TotalObjects
        Set tempControl = New clsControls
        tempControl.ControlID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlTAG
        tempControl.ControlAltID = CStr(ControlDeliveryDate) & "_" & ControlDeliveryRef & "_" & ControlName
        tempControl.TheControl = TheControl
        tempControl.ControlName = ControlName
        tempControl.ControlValue = ControlText
        tempControl.ControlType = ControlType
        tempControl.ControlTAG = ControlTAG
        tempControl.ControlDate = ControlDate
        tempControl.ControlDeliveryDate = ControlDeliveryDate
        tempControl.ControlDeliveryRef = ControlDeliveryRef
        tempControl.ControlObjNumber = ControlOBJCount + i
        tempControl.ControlStartTAG = ControlStartTAG
        tempControl.ControlEndTAG = ControlEndTAG
        tempControl.ControlRowNumber = ControlRowNumber
        tempControl.ControlTotalRows = ControlTotalRows
        
        tempControl.ControlOpName = ControlOpName
        tempControl.ControlOpActivity = ControlActivity
        tempControl.ControlOpStartDateTime = ControlStartDateTime
        tempControl.ControlOpEndDateTime = ControlEndDateTime
        tempControl.ControlPartNo = ControlPartNo
        tempControl.ControlQty = ControlQty
        tempControl.ControlExtraShort = ControlExtraShort
        Set tempControl.TxtBoxAfterUpdate = TheControl
        Set tempControl.comboAfterUpdate = TheControl
        'tempControl.ControlList = ListArray 'if combobox
            
        'ctrlCollection.Add tempControl, tempControl.ControlName
        
    
    
    CollectionKey = tempControl.ControlTAG
    If Len(ControlDeliveryRef) > 0 Then
        CollectionKey = ControlDeliveryDate & "_" & ControlDeliveryRef & "_" & tempControl.ControlTAG
    ElseIf Len(ControlASN) > 0 Then
        CollectionKey = ControlDeliveryDate & "_" & ControlASN & "_" & tempControl.ControlTAG
    End If
    'ctrlCollection.Add tempControl, tempControl.ControlTag
    If Not InCollection("MISSING", ctrlCollection, CollectionKey) Then
        ctrlCollection.Add tempControl, CollectionKey
    End If
    If Not Dic_Collection.Exists(CollectionKey) Then
        Dic_Collection.Add CollectionKey, tempControl
    End If

    Next i

End Sub

Sub UpdateControlCollection(DeliveryDate As String, DeliveryRef As String, TAGNumber As String, ControlName As String, _
        ControlCollection As Collection, ControlToUpdate As Control, ValueToInsert As Variant, Optional ByRef ErrMessage As String = "")
    Dim varControlKey As Variant
    Dim ctrlProperty As Variant
    Dim ctrlUpdate As Control
    Dim ControlType As String
    Dim strTagNumber As String
    Dim ControlValue As String
    Dim tempControl As clsControls
    Dim NewCtrl As Control
    Dim IDX As Long
    Dim CollectionKey As Variant
    
    If Len(DeliveryDate) = 0 Then
        ErrMessage = "Delivery Date Not Specified in UPDATE"
        Exit Sub
    End If
    If Len(DeliveryRef) = 0 Then
        ErrMessage = "Delivery Ref Not Specified in UPDATE"
        Exit Sub
    End If
    If Len(TAGNumber) = 0 Then
        ErrMessage = "TAG NUMBER not specified in UPDATE"
        Exit Sub
    End If
    Set varControlKey = Nothing
    If Len(TAGNumber) > 0 Then
        varControlKey = DeliveryDate & "_" & DeliveryRef & "_" & TAGNumber
    ElseIf Len(ControlName) > 0 Then
        varControlKey = DeliveryDate & "_" & DeliveryRef & "_" & ControlName
    Else
        'Invalid.
    End If
    If InCollection("MISSING", ControlCollection, varControlKey) Then
        Set ctrlProperty = ControlCollection.Item(varControlKey)
        ControlName = ctrlProperty.ControlName
        strTagNumber = ctrlProperty.ControlTAG
        ControlValue = ctrlProperty.ControlValue
        ControlType = ctrlProperty.ControlType
        MsgBox ("Control Name = " & ControlName & ", TAG=" & strTagNumber & " ,value=" & ControlValue & ", " & ControlType)
        ControlCollection(varControlKey) = ValueToInsert
    End If


End Sub

Function UpdateCollection(ByRef coll As Collection, varkey As Variant, ValueToChange As Variant, _
    Optional DeliveryDate As Date, Optional ASNNumber As String = "", Optional DeliveryRef As String, Optional TAGNumber As String) As Collection
    Dim NewKey As Variant
    Dim tempControl As clsControls
    
    On Error GoTo Err_UpdateCollection
    
    If Len(varkey) > 0 Then
        NewKey = varkey
    Else
        If Len(CStr(DeliveryDate)) > 0 Then
            If Len(DeliveryRef) > 0 Then
                NewKey = DeliveryDate & "_" & DeliveryRef & "_" & TAGNumber
            End If
        ElseIf Len(ASNNumber) > 0 Then
            If Len(DeliveryRef) > 0 Then
                NewKey = DeliveryDate & "_" & ASNNumber & "_" & TAGNumber
            End If
        End If
        NewKey = varkey
    End If
    Set tempControl = coll.Item(NewKey)
    tempControl.ControlValue = ValueToChange
    coll.Remove NewKey
    coll.Add tempControl, NewKey
    
    Set UpdateCollection = coll

Exit Function

Err_UpdateCollection:

    Call Error_Report("UpdateCollection")

End Function


Function ReturnControlInfo(SearchCollectionKey As Variant, NewValue As Variant, SearchDeliveryDate As String, SearchDeliveryRef As String, SearchASNNumber As String, _
    SearchTAGNumber As String, SearchControlName As String, SearchProperty As String, SearchValue As Variant, _
    ByVal ReturnProperty As String, ByRef ReturnValue As Variant, Optional ByRef ErrMessage As String = "", _
        Optional ByRef ReturnclsControl As clsControls, Optional SendBlank As Boolean = False, Optional FoundControl As Control, _
        Optional ISFoundControl As Boolean = False) As Boolean
    Dim tempControl As New clsControls
    Dim CollectionKey As Variant
    
    'Set tempControl = New clsControls
    ReturnValue = ""
    ReturnControlInfo = False
    If Len(SearchCollectionKey) > 0 Then
        CollectionKey = SearchCollectionKey
        
    Else
    'must pass the whole key info : Delivery Date _ Delivery Ref _ TAG NUMBER
        If Len(SearchDeliveryDate) = 0 Then
            ErrMessage = "No Delivery Date Specified"
            Exit Function
        Else
            If Not IsDate(SearchDeliveryDate) Then
                ErrMessage = "Delivery Date is NOT Valid"
                Exit Function
            End If
        End If
        If Len(SearchTAGNumber) > 0 Then
            CollectionKey = SearchTAGNumber
        ElseIf Len(SearchControlName) > 0 Then
            CollectionKey = SearchControlName
        Else
            ErrMessage = "No TAG or Control Name Specified"
            Exit Function
        End If
        If Len(SearchDeliveryRef) > 0 Then
            CollectionKey = SearchDeliveryDate & "_" & SearchDeliveryRef & "_" & CollectionKey
        ElseIf Len(SearchASNNumber) > 0 Then
            CollectionKey = SearchDeliveryDate & "_" & SearchASNNumber & "_" & CollectionKey
        Else
            ErrMessage = "No Delivery Ref or ASN Specified"
            Exit Function
        End If
    End If
    For Each tempControl In ctrlCollection
        Set ReturnValue = Nothing
        If UCase(SearchProperty) = "TAG" Then
            If Len(SearchDeliveryRef) > 0 Then
                If UCase(tempControl.ControlID) = UCase(CollectionKey) Then 'Delivery REF
                
                    ReturnValue = ReturnCollectionPropertyValue(ReturnProperty, tempControl, NewValue, SearchValue, FoundControl, ISFoundControl, SendBlank)
                    
                    If UCase(ReturnProperty) = UCase("ControlName") Then
                        ReturnValue = tempControl.ControlName
                    End If
                    If UCase(ReturnProperty) = UCase("ControlFieldname") Then
                        ReturnValue = tempControl.ControlFieldname
                        
                    End If
                    If UCase(ReturnProperty) = UCase("ControlValue") Then
                        ReturnValue = tempControl.ControlValue
                    End If
                    If UCase(ReturnProperty) = UCase("ControlDate") Then 'As in the Date and Time of when the information was saved to the control collection object
                        ReturnValue = tempControl.ControlDate
                    End If
                    If UCase(ReturnProperty) = UCase("StartTag") Then
                        ReturnValue = tempControl.ControlStartTAG
                    End If
                    If UCase(ReturnProperty) = UCase("EndTag") Then
                        ReturnValue = tempControl.ControlEndTAG
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryDate") Then
                        ReturnValue = tempControl.ControlDeliveryDate
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryRef") Then
                        ReturnValue = tempControl.ControlDeliveryRef
                    End If
                    If UCase(ReturnProperty) = UCase("Type") Then
                        ReturnValue = tempControl.ControlType
                    End If
                    If UCase(ReturnProperty) = UCase("TotalRows") Then
                        ReturnValue = tempControl.ControlTotalRows
                    End If
                    If UCase(ReturnProperty) = UCase("TheControl") Then
                        Set ReturnValue = tempControl.TheControl
                    End If
                    
                    If Not IsEmpty(ReturnValue) Then
                        ReturnControlInfo = True
                    End If
                End If
            End If
            If Len(SearchASNNumber) > 0 Then
                If UCase(tempControl.ControlASNID) = UCase(CollectionKey) Then 'ASN NUMBER with TAG
                
                    ReturnValue = ReturnCollectionPropertyValue(ReturnProperty, tempControl, NewValue, SearchValue, FoundControl, ISFoundControl)
                
                    If UCase(ReturnProperty) = UCase("ControlName") Then
                        ReturnValue = tempControl.ControlName
                    End If
                    If UCase(ReturnProperty) = UCase("ControlFieldname") Then
                        ReturnValue = tempControl.ControlFieldname
                    End If
                    If UCase(ReturnProperty) = UCase("ControlValue") Then
                        ReturnValue = tempControl.ControlValue
                    End If
                    If UCase(ReturnProperty) = UCase("ControlDate") Then 'As in the Date and Time of when the information was saved to the control collection object
                        ReturnValue = tempControl.ControlDate
                    End If
                    If UCase(ReturnProperty) = UCase("StartTag") Then
                        ReturnValue = tempControl.ControlStartTAG
                    End If
                    If UCase(ReturnProperty) = UCase("EndTag") Then
                        ReturnValue = tempControl.ControlEndTAG
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryDate") Then
                        ReturnValue = tempControl.ControlDeliveryDate
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryRef") Then
                        ReturnValue = tempControl.ControlDeliveryRef
                    End If
                    If UCase(ReturnProperty) = UCase("Type") Then
                        ReturnValue = tempControl.ControlType
                    End If
                    If UCase(ReturnProperty) = UCase("TotalRows") Then
                        ReturnValue = tempControl.ControlTotalRows
                    End If
                    If UCase(ReturnProperty) = UCase("TheControl") Then
                        Set ReturnValue = tempControl.TheControl
                    End If
                    
                    If Not IsEmpty(ReturnValue) Then
                        ReturnControlInfo = True
                    End If
                End If
            End If
        End If
        If UCase(SearchProperty) = "NAME" Then
            If Len(SearchDeliveryRef) > 0 Then
                If UCase(tempControl.ControlAltID) = UCase(CollectionKey) Then
                    
                    ReturnValue = ReturnCollectionPropertyValue(ReturnProperty, tempControl, NewValue, SearchValue, FoundControl, ISFoundControl)
                    
                    If UCase(ReturnProperty) = UCase("ControlName") Then
                        ReturnValue = tempControl.ControlName
                    End If
                    If UCase(ReturnProperty) = UCase("ControlFieldname") Then
                        ReturnValue = tempControl.ControlFieldname
                    End If
                    If UCase(ReturnProperty) = UCase("ControlTag") Then
                        ReturnValue = tempControl.ControlTAG
                    End If
                    If UCase(ReturnProperty) = UCase("ControlValue") Then
                        ReturnValue = tempControl.ControlValue
                    End If
                    If UCase(ReturnProperty) = UCase("ControlDate") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDate
                    End If
                    If UCase(ReturnProperty) = UCase("StartTag") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlStartTAG
                    End If
                    If UCase(ReturnProperty) = UCase("EndTag") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlEndTAG
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryDate") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDeliveryDate
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryRef") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDeliveryRef
                    End If
                    If UCase(ReturnProperty) = UCase("Type") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlType
                    End If
                    If UCase(ReturnProperty) = UCase("TotalRows") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlTotalRows
                    End If
                    If UCase(ReturnProperty) = UCase("TheControl") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        Set ReturnValue = tempControl.TheControl
                    End If
                    
                    
                    
                    If Not IsEmpty(ReturnValue) Then
                        ReturnControlInfo = True
                    End If
                End If
            End If
            If Len(SearchASNNumber) > 0 Then
                If UCase(tempControl.ControlASNAltID) = UCase(CollectionKey) Then
                
                    ReturnValue = ReturnCollectionPropertyValue(ReturnProperty, tempControl, NewValue, SearchValue, FoundControl, ISFoundControl)
                    
                    If UCase(ReturnProperty) = UCase("ControlName") Then
                        ReturnValue = tempControl.ControlName
                    End If
                    If UCase(ReturnProperty) = UCase("ControlFieldname") Then
                        ReturnValue = tempControl.ControlFieldname
                    End If
                    If UCase(ReturnProperty) = UCase("ControlTag") Then
                        ReturnValue = tempControl.ControlTAG
                    End If
                    If UCase(ReturnProperty) = UCase("ControlValue") Then
                        ReturnValue = tempControl.ControlValue
                    End If
                    If UCase(ReturnProperty) = UCase("ControlDate") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDate
                    End If
                    If UCase(ReturnProperty) = UCase("StartTag") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlStartTAG
                    End If
                    If UCase(ReturnProperty) = UCase("EndTag") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlEndTAG
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryDate") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDeliveryDate
                    End If
                    If UCase(ReturnProperty) = UCase("DeliveryRef") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlDeliveryRef
                    End If
                    If UCase(ReturnProperty) = UCase("Type") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlType
                    End If
                    If UCase(ReturnProperty) = UCase("TotalRows") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        ReturnValue = tempControl.ControlTotalRows
                    End If
                    If UCase(ReturnProperty) = UCase("TheControl") Then 'As in the Date and Time of when an activity starts or finishes in the Operatives Frame
                        Set ReturnValue = tempControl.TheControl
                    End If
                    
                    If Not IsEmpty(ReturnValue) Then
                        ReturnControlInfo = True
                    End If
                End If
            End If
        End If
    Next
    
    Set ReturnclsControl = tempControl
End Function

Function ReturnCollectionPropertyValue(ByVal ReturnProperty As String, ByRef CollectionControl As clsControls, NewValue As Variant, _
    Optional ByVal SearchValue As Variant, Optional ByRef FoundControl As Control = Nothing, Optional ByRef IsControlFound As Boolean = False, _
    Optional SendBlank As Boolean = False) As Variant
    Dim RetValue As Variant
    Dim ControlName As String
    Dim ControlType As String
    Dim CTRL As Control
    
    Set ReturnCollectionPropertyValue = Nothing
    Set RetValue = Nothing
    ControlType = CollectionControl.ControlType
    ControlName = CollectionControl.ControlName
    
    'if combobox changed - could be something like "Nathan Kirkpatrick"
    'This is NOT a fieldname !
    
    If UCase(ReturnProperty) = UCase("ControlName") Then
        RetValue = CollectionControl.ControlName
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlName = NewValue
        End If
        If SendBlank Then CollectionControl.ControlName = ""
    End If
    If UCase(ReturnProperty) = UCase("ControlFieldname") Then
        RetValue = CollectionControl.ControlFieldname
        If Not IsEmpty(NewValue) Then
            'GET NEW FIELDNAME first:
            CollectionControl.ControlFieldname = NewValue
        End If
        If SendBlank Then CollectionControl.ControlFieldname = ""
    End If
    If UCase(ReturnProperty) = UCase("ControlValue") Then
        RetValue = CollectionControl.ControlValue 'old value
        
        If Len(SearchValue) > 0 Then
            If UCase(RetValue) = UCase(SearchValue) Then
                Set FoundControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, ControlType, "", ControlName)
                If Not FoundControl Is Nothing Then
                    IsControlFound = True
                End If
            End If
        End If
        
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlValue = NewValue
        End If
        If SendBlank Then CollectionControl.ControlValue = ""
    End If
    If UCase(ReturnProperty) = UCase("ControlDate") Then 'As in the Date and Time of when the information was saved to the control collection object
        RetValue = CollectionControl.ControlDate
    End If
    If UCase(ReturnProperty) = UCase("StartTag") Then
        RetValue = CollectionControl.ControlStartTAG
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlStartTAG = NewValue
        End If
        If SendBlank Then CollectionControl.ControlStartTAG = ""
    End If
    If UCase(ReturnProperty) = UCase("EndTag") Then
        RetValue = CollectionControl.ControlEndTAG
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlEndTAG = NewValue
        End If
        If SendBlank Then CollectionControl.ControlEndTAG = ""
    End If
    If UCase(ReturnProperty) = UCase("DeliveryDate") Then
        RetValue = CollectionControl.ControlDeliveryDate
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlDeliveryDate = NewValue
        End If
        If SendBlank Then CollectionControl.ControlDeliveryDate = CDate("01/01/1970")
    End If
    If UCase(ReturnProperty) = UCase("DeliveryRef") Then
        RetValue = CollectionControl.ControlDeliveryRef
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlDeliveryRef = NewValue
        End If
        If SendBlank Then CollectionControl.ControlDeliveryRef = ""
    End If
    If UCase(ReturnProperty) = UCase("Type") Then
        RetValue = CollectionControl.ControlType
        If Not NewValue Is Nothing Then
            CollectionControl.ControlType = NewValue
        End If
        If SendBlank Then CollectionControl.ControlType = ""
    End If
    If UCase(ReturnProperty) = UCase("TotalRows") Then
        RetValue = CollectionControl.ControlTotalRows
        If Not IsEmpty(NewValue) Then
            CollectionControl.ControlTotalRows = NewValue
        End If
        If SendBlank Then CollectionControl.ControlTotalRows = 0
    End If
    If UCase(ReturnProperty) = UCase("TheControl") Then
        Set RetValue = CollectionControl.TheControl
        If Not IsEmpty(NewValue) Then
            Set CollectionControl.TheControl = NewValue
        End If
        If SendBlank Then Set CollectionControl.TheControl.value = ""
    End If
    
    If Len(SearchValue) > 0 Then
        'Search TempControl for value
        'For Each CTRL In CollectionControl
            
        'Next
    End If
    
    If Not IsEmpty(RetValue) Then
        ReturnCollectionPropertyValue = RetValue
    End If

End Function

Function FindControlInfo(TheCollection As Collection, SearchField As String, SearchText As String, ByRef ReturnInfo As String) As Boolean
    Dim CTRL As Control
    Dim IDX As Long
    
    FindControlInfo = False
    IDX = 0
    For Each CTRL In TheCollection
        If UCase(SearchField) = "TAG" Then
            If UCase(CTRL.Tag) = UCase(SearchText) Then
                'returninfo =
                FindControlInfo = True
            End If
        End If
    Next
    


End Function

Sub StartEntry()
Attribute StartEntry.VB_Description = "MACRO to show the userform to enter Timesheet details and add to Timesheet Records sheet as new record / row.\n"
Attribute StartEntry.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim Timesheet As Worksheet
'
' StartEntry Macro by Daniel Goss 2018. Version 7.1 04-JUN-2018 09:00
'
    TESTING = False
    Range("A2").Select
    CurrentRecordRow = 0
    
    'SET TAG ID:
    ComplianceQuestion1TAG = "801"
    ComplianceQuestion2TAG = "802"
    ComplianceQuestion3TAG = "803"
    ComplianceQuestion4TAG = "804"
    ComplianceQuestion5TAG = "805"
    ComplianceQuestion6TAG = "806"
    FurtherCommentsTAG = "807"
    
    'MainWorksheet = "Timesheet Records"
    
    SaveAndContinue = True
    'TotalFields = ActiveWorkbook.Worksheets(MainWorksheet).Cells(1, Columns.Count).End(xlToLeft).Column
    frmGI_TimesheetEntry2_1060x630.Show 'unlimited operatives.
    
    
    
End Sub

Sub workbook_open()

    Call StartEntry
    

End Sub

Sub HideTimesheetRecords(WB As Workbook)
    Dim ThisWB As Workbook

On Error GoTo Err_HideTimesheetRecords
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    If Not TESTING Then
        'MAKE IT HIDDEN !
        'MAKE Sheet1 visible - with just a begin button on it.
        If MainGIModule_v1_1.sheetExists(MainWorksheet, ThisWB) Then
            ThisWB.Sheets(MainWorksheet).Visible = xlSheetVeryHidden
        End If
        If MainGIModule_v1_1.sheetExists("Employees", ThisWB) Then
            ThisWB.Sheets("Employees").Visible = xlSheetVeryHidden
        End If
        If MainGIModule_v1_1.sheetExists("Suppliers", ThisWB) Then
            ThisWB.Sheets("Suppliers").Visible = xlSheetVeryHidden
        End If
        If MainGIModule_v1_1.sheetExists("FLM", ThisWB) Then
            ThisWB.Sheets("FLM").Visible = xlSheetVeryHidden
        End If
        
    End If
Exit Sub

Err_HideTimesheetRecords:
    Call Error_Report("HideTimesheetRecords")
    
End Sub

Public Function GetScreenResolution_Actual(ByVal ScreenNumber As Integer, ByRef Screenwidth As Integer, ByRef ScreenHeight As Integer, _
                                            Optional ByRef Message As String = "", Optional ByRef WorkingWidth As Integer, Optional ByRef WorkingHeight As Integer) As Integer
    Dim NumberOfScreens As Integer
        
    'WorkingHeight = Application.Windows.forms.screen.primaryscreen.workingarea.Height
    'WorkingWidth = Application.Windows.forms.screen.primaryscreen.workingarea.Width
    'NumberOfScreens = Application.screen.AllScreens.Count
    
    'Not working in EXCEL.
    If ScreenNumber < 2 Then
        Screenwidth = Application.screen.primaryscreen.Bounds.Width
        ScreenHeight = Application.screen.primaryscreen.Bounds.Height
    Else
        If ScreenNumber <= NumberOfScreens Then
            Screenwidth = Application.screen.AllScreens(ScreenNumber).Bounds.Width
            ScreenHeight = Application.screen.AllScreens(ScreenNumber).Bounds.Height
        Else
            Message = "Error: Passed Parameter Exceeds Number of Screens Available"
        End If
    End If
    GetScreenResolution_Actual = NumberOfScreens
End Function


Public Sub GetScreenResolution(ByRef Screenwidth As Long, ByRef ScreenHeight As Long)
    
    Screenwidth = GetSystemMetrics(SM_CXSCREEN)
    ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Sub


Public Sub Unhide_Timesheet_Records()
Attribute Unhide_Timesheet_Records.VB_Description = "Unihide"
Attribute Unhide_Timesheet_Records.VB_ProcData.VB_Invoke_Func = "U\n14"
    Dim PWEntry As String
    Dim UsePasswordForm As usrfrmPassword
    Dim DATA_WORKBOOK As Workbook
    Dim Data_Worksheet As String
    
    Set UsePasswordForm = New usrfrmPassword
    With UsePasswordForm
        .Show
        PWEntry = .Password
    End With
    Set DATA_WORKBOOK = Nothing
    Data_Worksheet = "Timesheet Records"
    If DATA_WORKBOOK Is Nothing Then
    'pword = InputBox("Enter Password:", "Password Required to View Timesheet Records sheet")
        If PWEntry = "blue" Then
            'MainWorksheet = "Timesheet Records"
            If MainGIModule_v1_1.sheetExists(Data_Worksheet, ActiveWorkbook) Then
                ActiveWorkbook.Sheets(Data_Worksheet).Visible = True
            End If
            'ActiveWorkbook.Sheets(Data_Worksheet).Visible = True
            If MainGIModule_v1_1.sheetExists("Employees", ActiveWorkbook) Then
                ActiveWorkbook.Sheets("Employees").Visible = True
            End If
            If MainGIModule_v1_1.sheetExists("Suppliers", ActiveWorkbook) Then
                ActiveWorkbook.Sheets("Suppliers").Visible = True
            End If
            If MainGIModule_v1_1.sheetExists("FLM", ActiveWorkbook) Then
                ActiveWorkbook.Sheets("FLM").Visible = True
            End If
        Else
            MsgBox ("Password Invalid")
        End If
    Else
        If PWEntry = "yes" Then
            'MainWorksheet = "Timesheet Records"
            DATA_WORKBOOK.Sheets(Data_Worksheet).Visible = True
            DATA_WORKBOOK.Sheets("Employees").Visible = True
            DATA_WORKBOOK.Sheets("Suppliers").Visible = True
            DATA_WORKBOOK.Sheets("FLM").Visible = True
        Else
            MsgBox ("Password Invalid")
        End If
    End If
End Sub

Function ConvertBadChars(Entry As String, REplaceWith As String, Optional IncludeComma As Boolean = False, Optional IncludeSpeechMarks As Boolean = False, Optional IncludeSPACE As Boolean = False) As String
    Dim GoodEntry As Variant
    Dim BadChars As Variant
    Dim UglyEntry As Variant
    Dim BadCharArr As Variant
    Dim IDX As Integer
        Dim BadCharList As Variant
    
    On Error GoTo Err_ConvertBadChars
    
    ConvertBadChars = ""
    UglyEntry = Entry
        BadCharList = "["
    BadCharList = BadCharList & "," & "]"
    BadCharList = BadCharList & "," & "/"
    BadCharList = BadCharList & "," & "*"
    BadCharList = BadCharList & "," & "\"
    BadCharList = BadCharList & "," & "?"
    If IncludeComma Then
        BadCharList = BadCharList & "," & ","
    End If
    If IncludeSpeechMarks Then
        BadCharList = BadCharList & "," & Chr(34)
    End If
    If IncludeSPACE Then
        BadCharList = BadCharList & "," & Chr(32)
    End If
    UglyEntry = Entry
    BadChars = Array(BadCharList)
    BadCharArr = Split(BadCharList, ",")
    If Len(UglyEntry) > 0 Then
        For IDX = LBound(BadCharArr) To UBound(BadCharArr)
            GoodEntry = Replace(UglyEntry, BadCharArr(IDX), REplaceWith, 1)
        Next
    End If
    ConvertBadChars = GoodEntry
    
Exit Function

Err_ConvertBadChars:
Call Error_Report("ConvertBadChars")

End Function

Function sheetExists(sheetToFind As String, Optional WB As Workbook) As Boolean
Dim Sheet As Worksheet
Dim ThisWB As Workbook

On Error GoTo Err_sheetExists
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    sheetExists = False
    For Each Sheet In ThisWB.Worksheets
        If UCase(sheetToFind) = UCase(Sheet.Name) Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet

Exit Function

Err_sheetExists:

Call Error_Report("SheetExists")
    
End Function

Sub getworkbook(ByRef Newsheet As String, Optional SpecFilename As String, Optional DestWB As Workbook)
    ' Get workbook...
    Dim WS As Worksheet
    Dim strFilter As String
    Dim strCaption As String
    Dim targetWorkbook As Workbook
    Dim ThisWB As Workbook
    Dim OpenWB As Workbook
    Dim ret As Variant
    Dim Sheetname As String
    Dim PickedPos As Integer
    Dim IDX As Integer
    
    On Error GoTo Err_GetWorkbook
    
    If DestWB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = DestWB
    End If
    Set targetWorkbook = ThisWB

    ' get the customer workbook
    
    Application.ScreenUpdating = False
    If Len(SpecFilename) > 0 Then
        'Filter = "CSV Files (*.csv),*.csv|Excel Files (*.xlsx),*.xlsx|All Files (*.*),*.*"
        'Caption = "Please Select the data file: " & SpecFilename
        'ret = Application.GetOpenFilename(Filter, , Caption)

        'If ret = False Then Exit Sub
    
            Set OpenWB = Workbooks.Open(SpecFilename)
            OpenWB.Sheets(1).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
            'Sheetname = SpecFilename
            'If Not sheetExists(Sheetname) Then
            '    NewSheet = "Data_"
            '    ActiveSheet.Name = Sheetname
            '    NewSheet = Sheetname
            'End If
    Else
        strFilter = "Excel b Files (*.xlsb),*.xlsb,Excel Files (*.xlsx),*.xlsx,All Files (*.*),*.*"
        strCaption = "Please Select the data file "
        ret = Application.GetOpenFilename(strFilter, 1, strCaption)

        If ret = False Then Exit Sub
        Set OpenWB = Workbooks.Open(ret, False, True) 'Open Selected workbook - where the DAILY sheet is - as read only.
        OpenWB.Sheets(1).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
        PickedPos = InStr(ret, "Picked")
        'Sheetname = "DATA_" & Mid(ret, PickedPos, 16)
        Sheetname = OpenWB.Sheets(1).Name
        OpenWB.Close
        
        Newsheet = Sheetname
    End If
Exit Sub
Err_GetWorkbook:
    Error_Report ("GetWorkbook()")
   
End Sub

Sub Connect_Data_Workbook(ByRef DataWorkbook As Workbook, ByRef DataWorksheet As String, Optional SpecFilename As String = "")
    Dim WS As Worksheet
    Dim strFilter As String
    Dim strCaption As String
    Dim targetWorkbook As Workbook
    'Dim wb As Workbook
    Dim ret As Variant
    Dim Sheetname As String
    Dim PickedPos As Integer
    Dim IDX As Integer
    
    On Error GoTo Err_GetWorkbook
    
    Set targetWorkbook = Application.ActiveWorkbook

    ' get the customer workbook
    
        
    If Len(SpecFilename) > 0 Then
        'Filter = "CSV Files (*.csv),*.csv|Excel Files (*.xlsx),*.xlsx|All Files (*.*),*.*"
        'Caption = "Please Select the data file: " & SpecFilename
        'ret = Application.GetOpenFilename(Filter, , Caption)

        'If ret = False Then Exit Sub
    
            Set DataWorkbook = Workbooks.Open(SpecFilename)
            DataWorkbook.Sheets(1).Move After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
            'Sheetname = SpecFilename
            'If Not sheetExists(Sheetname) Then
            '    NewSheet = "Data_"
            '    ActiveSheet.Name = Sheetname
            '    NewSheet = Sheetname
            'End If
    Else
        strFilter = "Excel b Files (*.xlsb),*.xlsb,Excel Files (*.xlsx),*.xlsx,All Files (*.*),*.*"
        strCaption = "Please Select the data file in the shared workbook."
        ret = Application.GetOpenFilename(strFilter, 1, strCaption)
        
        If ret = False Then Exit Sub
        Set DataWorkbook = Workbooks.Open(ret, False, True)
        'wb.Sheets(1).Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
        'Sheetname = "DATA_" & Mid(ret, PickedPos, 16)
        DataWorksheet = DataWorkbook.Sheets(1).Name 'get first worksheet in workbook.
        'wb.Close
        
    End If
Exit Sub
Err_GetWorkbook:
    Error_Report ("Connect_Data_Workbook()")
   
    
    
    
End Sub

Function Connect_RecordSheet(ByVal PreferenceSheet As String, _
        Optional OpenTheWorkbook As Boolean = False, Optional ByRef ConnectWB As Workbook, Optional MakeReadOnly As Boolean = False, _
        Optional ByRef IsValidPath As Boolean = True) As String
    Dim Newsheet As Worksheet
    Dim strFilter As String
    Dim strCaption As String
    Dim ret As Variant

    Connect_RecordSheet = ""
    'SELECT the location of the main record data sheet to capture each Goods In receipt per GRID:
    strFilter = "Excel b Files (*.xlsb),*.xlsb,ALL Excel Files (*.xls?),*.xls?,All Files (*.*),*.*"
    strCaption = "Please Select the location of the GI record data sheet"
    ret = Application.GetOpenFilename(strFilter, 2, strCaption)

    If ret = False Then Exit Function
    
    If OpenTheWorkbook Then
        Set ConnectWB = Workbooks.Open(ret, False, MakeReadOnly) 'path,updatelinks,readonly.
    End If
    If MainGIModule_v1_1.sheetExists(PreferenceSheet, ThisWorkbook) Then
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 2).value = ret
    Else
        Set Newsheet = ThisWorkbook.Worksheets.Add
        Newsheet.Name = "Prefs"
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 1).value = "Data Workbook:"
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 2).value = ret
    End If
    Connect_RecordSheet = ret
    'Test if the path is valid - but for now assume it is:
    IsValidPath = True

End Function

Function Connect_ACCESS_DB(ByVal PreferenceSheet As String) As String
    Dim DBPath As String
    Dim Newsheet As Worksheet
    Dim strFilter As String
    Dim strCaption As String
    Dim ret As Variant
    
    Connect_ACCESS_DB = ""
    strFilter = "ACCESS Files (*.accdb),*.accdb,Old Access Files (*.MDB),*.MDB,All Files (*.*),*.*"
    strCaption = "Please Select the location of the GI ACCESS Database"
    ret = Application.GetOpenFilename(strFilter, 1, strCaption)
    
    If ret = False Then Exit Function
    
    If MainGIModule_v1_1.sheetExists(PreferenceSheet, ThisWorkbook) Then
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 2).value = ret
    Else
        Set Newsheet = ThisWorkbook.Worksheets.Add
        Newsheet.Name = "Prefs"
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 1).value = "GI ACCESS Database:"
        ThisWorkbook.Worksheets(PreferenceSheet).Cells(1, 2).value = ret
    End If
    
    Connect_ACCESS_DB = ret
    
End Function


Function Test_File_Exist_With_Dir(FileAndPath As String) As Boolean
'Updateby Extendoffice 20161109
    Dim FilePath As String
    
    Test_File_Exist_With_Dir = False
    Application.ScreenUpdating = False
    FilePath = ""
    On Error Resume Next
    FilePath = Dir(FileAndPath)
    On Error GoTo 0
    If FilePath = "" Then
        Test_File_Exist_With_Dir = False
    Else
        Test_File_Exist_With_Dir = True
    End If
    Application.ScreenUpdating = True
End Function

Sub MakeShared(WB As Workbook)
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If

    If Not WB.MultiUserEditing Then
        Application.DisplayAlerts = False
        'WB.SaveAs ThisWB.Name, AccessMode:=xlShared, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        'WB.SaveAs ThisWB.Name, AccessMode:=xlShared
        Application.DisplayAlerts = True
        'MsgBox "Now Shared"
    End If
End Sub

Public Sub Error_Report(ProcedureName As String)

If Err > 0 Then
        If Err = 13 Then
            MsgBox "Error in " & ProcedureName & " - mismatch :" & vbCrLf & vbCrLf & "Err = " & Err.Number & _
            vbCrLf & "Description: " & Err.Description & vbCrLf & " : Source:" & Err.Source
        Else
            MsgBox "Normal Error in " & ProcedureName & " :" & vbCrLf & vbCrLf & "Err = " & Err.Number & _
            vbCrLf & "Description: " & Err.Description & vbCrLf & " : Source:" & Err.Source
                        
        End If
End If
If Err < 0 Then
    MsgBox "Strange Error in " & ProcedureName & " :" & vbCrLf & vbCrLf & "Err = " & Err.Number & _
            vbCrLf & "Description: " & Err.Description & vbCrLf & " : Source:" & Err.Source
End If
Err.Clear

End Sub

Sub SetupTimesheetColumns(WB As Workbook, WorksheetName As String)
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    ThisWB.Worksheets(WorksheetName).Cells(1, 1).value = "Delivery Date"
    ThisWB.Worksheets(WorksheetName).Cells(1, 2).value = "Delivery Reference"
    

End Sub

Function GetNextAvailablerow(WB As Workbook, WorksheetName As String, StartRow As Long, CheckCol As Long) As Long
    Dim IDX As Long
    Dim LastRow As Long
    Dim sht As Worksheet
    Dim BlankRow As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    Set sht = ThisWB.Sheets(WorksheetName)
    GetNextAvailablerow = StartRow
    IDX = StartRow
    'LastRow = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    'BlankRow = sht.Cells.Find(" ", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    LastRow = ThisWB.Worksheets(WorksheetName).Cells(Rows.Count, CheckCol).End(xlUp).row

    
    'Do While IDX < LastRow
        
    '    If Len(Worksheets(WorksheetName).Cells(StartRow, CheckCol).value) > 0 Then
            
    '        Exit Do
            
            
    '    End If
    '    IDX = IDX + 1
    'Loop
    GetNextAvailablerow = LastRow + 1
    CurrentRecordRow = LastRow + 1
End Function

Sub SetControlBackgroundColour(TAGNumber As String, Colour As Long)
    Dim CTRL As Control
    
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        If UCase(CTRL.Tag) = UCase(TAGNumber) Then
            CTRL.BackColor = Colour
            Exit For
        End If
    Next
End Sub


Function CheckFinishTimeError(WB As Workbook, WorksheetName As String, RowNumber As Long, Optional SpecificColumnNumber As Long = 0, Optional ByRef ErrColumn As Long = 0) As Boolean
    Dim ColNum As Long
    Dim CellEntry1 As String
    Dim CellEntry2 As String
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    CheckFinishTimeError = False
    ColNum = 1
    ErrColumn = 0
    If RowNumber = 0 Then
        MsgBox ("Error - Row passed is 0")
        CheckFinishTimeError = True
        Exit Function
    End If
    Do While ColNum <= TotalFields
        If SpecificColumnNumber > 0 Then
            CellEntry1 = ThisWB.Worksheets(WorksheetName).Cells(RowNumber, SpecificColumnNumber).value
            CellEntry2 = ThisWB.Worksheets(WorksheetName).Cells(RowNumber, SpecificColumnNumber + 1).value
        Else
            CellEntry1 = ThisWB.Worksheets(WorksheetName).Cells(RowNumber, ColNum).value
            CellEntry2 = ThisWB.Worksheets(WorksheetName).Cells(RowNumber, ColNum + 1).value
        End If
        If IsDate(CellEntry1) And IsDate(CellEntry2) Then
            'OK so both dates side-by-side are valid in the record:
            'Test if Finish Time is LESS than Start Time = ERROR
            If CDate(CellEntry2) < CDate(CellEntry1) Then
                CheckFinishTimeError = True
                Call SetControlBackgroundColour(CStr(ColNum + 1), vbRed)
                ErrColumn = ColNum + 1 'last control found to have error
                If SpecificColumnNumber > 0 Then
                    Exit Do
                End If
            Else
                'DATES are OK - set control background to white:
                Call SetControlBackgroundColour(CStr(ColNum + 1), vbWhite)
            End If
        End If
        ColNum = ColNum + 1
    Loop

End Function

Function CheckFinishTimeError_In_FORM(StartTimeValue As String, EndTimeValue As String, Optional SpecificColumnNumber As Long = 0, Optional ByRef ErrColumn As Long = 0) As Boolean
    Dim ColNum As Long
    Dim CellEntry1 As String
    Dim CellEntry2 As String
    Dim StartTimeFieldValue As String
    Dim EndTimeFieldValue As String
    
    On Error GoTo Err_CheckFinishTimeError_In_FORM
    
    CheckFinishTimeError_In_FORM = False
    ColNum = 1
    ErrColumn = 0
    If Len(StartTimeValue) = 0 Then
        MsgBox ("Nothing Passed for Start Time")
        CheckFinishTimeError_In_FORM = True
        Exit Function
    End If
    If Len(EndTimeValue) = 0 Then
        MsgBox ("Nothing Passed for End Time")
        CheckFinishTimeError_In_FORM = True
        Exit Function
    End If
    'Do While ColNum <= TotalFields
        If SpecificColumnNumber > 0 Then
            'CellEntry1 = RS.Fields.INDEX(StartTimeField).Value
            'CellEntry2 = RS.Fields.INDEX(StartTimeField).Value
            
            CellEntry1 = StartTimeValue
            CellEntry2 = EndTimeValue
        Else
            CellEntry1 = StartTimeValue
            CellEntry2 = EndTimeValue
        End If
        If IsDate(CellEntry1) And IsDate(CellEntry2) Then
            'OK so both dates side-by-side are valid in the record:
            'Test if Finish Time is LESS than Start Time = ERROR
            If CDate(CellEntry2) < CDate(CellEntry1) Then
                CheckFinishTimeError_In_FORM = True
                If SpecificColumnNumber > 0 Then
                    Call SetControlBackgroundColour(CStr(SpecificColumnNumber + 1), vbRed)
                End If
                'ErrColumn = ColNum + 1 'last control found to have error
            Else
                'DATES are OK - set control background to white:
                If SpecificColumnNumber > 0 Then
                    Call SetControlBackgroundColour(CStr(ColNum + 1), vbWhite)
                End If
            End If
        End If
        'ColNum = ColNum + 1
    'Loop
    
    Exit Function
Err_CheckFinishTimeError_In_FORM:

Call Error_Report("CheckFinishTimeError_In_FORM()")

End Function

Function FindFormControl(UserFormName As UserForm, ControlType As Variant, TAGNumber As String, Optional ControlName As String = "", _
    Optional FrameControls As Controls) As Control
    Dim CTRL As Control
    Dim UserFormControls As Controls
    'frmGI_TimesheetEntry2_1060x630
    
    On Error GoTo Err_FindFormControl
    
    Set FindFormControl = Nothing
    If Not UserFormName Is Nothing Then
        Set UserFormControls = UserFormName.Controls
    ElseIf Not FrameControls Is Nothing Then
        Set UserFormControls = FrameControls
    Else
        MsgBox ("Controls not passed in FindFormControl")
        Exit Function
    End If
    
    If Len(TAGNumber) > 0 Then
        For Each CTRL In UserFormControls
            If Len(ControlType) > 0 Then
                If UCase(TypeName(CTRL)) = UCase(ControlType) Then
                    If UCase(CTRL.Tag) = UCase(TAGNumber) Then
                        Set FindFormControl = CTRL
                        Exit For
                    End If
                End If
            Else
                If UCase(CTRL.Tag) = UCase(TAGNumber) Then
                    Set FindFormControl = CTRL
                    Exit For
                End If
            End If 'len(controltype)
        Next
    End If

    If Len(ControlName) > 0 Then
        For Each CTRL In UserFormControls
            If Len(ControlType) > 0 Then
                If UCase(TypeName(CTRL)) = UCase(ControlType) Then
                    If UCase(CTRL.Name) = UCase(ControlName) Then
                        Set FindFormControl = CTRL
                        Exit For
                    End If
                End If
            Else
                If UCase(CTRL.Name) = UCase(ControlName) Then
                    Set FindFormControl = CTRL
                    Exit For
                End If
            End If
        Next
    End If

Exit Function

Err_FindFormControl:

Call Error_Report("FindFormControl()")

End Function

Function CheckComplianceQuestionError(CBQuestion As Control, CBAnswerBox As Control) As Boolean
    Dim QuestionTAG As Long
    Dim CellEntry1 As String
    Dim CellEntry2 As String
    Dim QuestionCtrl As Control
    Dim AnswerCtrl As Control
    Dim ControlCount As Long
    
    CheckComplianceQuestionError = False
    
    'Check when Checkbox = NO that there has been something entered in the box below it:
        Set QuestionCtrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "Textbox", CBQuestion.Tag)
        Set AnswerCtrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "Textbox", CBAnswerBox.Tag)
        If UCase(CBQuestion.Text) = "YES" Then
            Call SetControlBackgroundColour(CBAnswerBox.Tag, vbWhite)
        Else
            If UCase(CBQuestion.Text) = "NO" Then
                If Len(CBAnswerBox.Text) = 0 Then
                    Call SetControlBackgroundColour(CBAnswerBox.Tag, vbRed)
                    CheckComplianceQuestionError = True
                Else
                    Call SetControlBackgroundColour(CBAnswerBox.Tag, vbWhite)
                End If
            End If
        End If

End Function

Sub InsertEntry(WB As Workbook, Optional SpecificEntry As String = "", Optional SpecificRow As Long = 0, Optional TagLowRange As Long = 0, Optional TagUpperRange As Long = 0, _
        Optional TimeSymbol As Long = 136)
    'Author: DANIEL GOSS - MAY 2018
    Dim CTRL As Control
    Dim myRow As Long
    Dim myCol As Long
    Dim Entry As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim DontSave As Boolean
    Dim ControlCount As Long
    Dim ThisWB As Workbook
    Dim dtDateEntry As Date
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If

    ThisWB.Worksheets(MainWorksheet).Protect Password:="yes", userinterfaceonly:=True
    
    
    If SpecificRow > 0 Then
        myRow = SpecificRow
        'need to record row here in public var ????
    Else
        myRow = MainGIModule_v1_1.GetNextAvailablerow(ThisWB, MainWorksheet, 2, 1) 'StartRow = 2 and check col = 1
        'Need to record row here in public var CurrentRow ????
    End If
    ControlCount = 0
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        myCol = 0
        DontSave = False
        ControlCount = ControlCount + 1
        If TypeName(CTRL) = "TextBox" Then
            Entry = CTRL
            If IsDate(Entry) Then
                dtDateEntry = CDate(Entry)
            End If
            'MsgBox ("ASCII=" & CStr(Asc(Entry)))
            'txtCtrl = ctrl
            If Len(CTRL.Tag) > 0 Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    myCol = CLng(CTRL.Tag)
                    'Entry = txtCtrl.Text
                    
                End If
            End If
        End If
        If TypeName(CTRL) = "ComboBox" Then
            Entry = CTRL
            'txtCtrl = ctrl
            If IsNumeric(CLng(CTRL.Tag)) Then
                myCol = CLng(CTRL.Tag)
                'Entry = txtCtrl.Text
                
            End If
        End If
        If myRow > 0 And myCol > 0 Then
            If Len(SpecificEntry) > 0 Then
                FinalEntry = SpecificEntry
            Else
                FinalEntry = Entry
                If IsDate(Entry) Then
                    dtDateEntry = CDate(Entry)
                    FinalEntry = dtDateEntry
                End If
            End If
            If Len(FinalEntry) = 0 Then
                If InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    FinalEntry = "NO"
                    
                End If
                    'All other blank text boxes
                If InStr(1, CTRL.Name, "Start", vbTextCompare) > 0 Then
                    FinalEntry = ""
                End If
                If InStr(1, CTRL.Name, "Finish", vbTextCompare) > 0 Then
                    FinalEntry = ""
                End If
                'FinalEntry = ""
                'End If
            End If
            If Len(FinalEntry) = 1 Then
                If Asc(FinalEntry) = 80 And InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    'Found a TICK in a TEXTBOX !
                    FinalEntry = "YES"
                End If
                'only works if the timesymbol matches the final entry symbol passed - single char from the check box.
                If Asc(FinalEntry) = TimeSymbol And InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    'Found a CLOCK - 6 O Clock icon in a TEXTBOX !
                    'FinalEntry = Format(Now(), "dd/mm/YYYY HH:MM:ss")
                    DontSave = True
                    
                End If
            End If
            If myRow > 0 And myCol > 0 Then
                If TagLowRange > 0 And myCol <= TagUpperRange And myCol >= TagLowRange Then
                    If DontSave = False Then
                        If IsDate(FinalEntry) Then
                            MsgBox ("DATE = " & FinalEntry)
                            'RemoteFilePath is the public variable holding the full path to the remote workbook with data.
                            dtDateEntry = Convert_strDateToDate(FinalEntry, True)
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).NumberFormat = "dd/mmm/yyyy HH:mm"
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).value = dtDateEntry
                        Else
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).value = FinalEntry
                        End If
                        If Len(SpecificEntry) > 0 Then Exit For
                    End If
                End If
                If TagLowRange = 0 And TagUpperRange = 0 Then
                    If DontSave = False Then
                        If IsDate(FinalEntry) Then
                            dtDateEntry = Convert_strDateToDate(FinalEntry, True)
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).NumberFormat = "dd/mmm/yyyy HH:mm"
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).value = dtDateEntry
                        Else
                            ThisWB.Worksheets(MainWorksheet).Cells(myRow, myCol).value = FinalEntry
                        End If
                        If Len(SpecificEntry) > 0 Then Exit For
                    End If
                End If
            End If
        End If
    Next
    ThisWB.Save

End Sub

Sub RemoveEntry(WB As Workbook, ReplaceDateWith As String, VisTimeBoxCtrl As Control, RowNumber As Long, LowerLimit As Long)
    Dim FinalEntry As String
    Dim ThisWB As Workbook
    Dim CTRL As Control
    
    Set CTRL = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TextBox", CStr(LowerLimit), "")
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
        
    Else
        Set ThisWB = WB
    End If

    FinalEntry = ReplaceDateWith
    VisTimeBoxCtrl.Text = ""
    'ThisWB.Worksheets(MainWorksheet).Cells(RowNumber, LowerLimit).value = FinalEntry
    CTRL.Text = FinalEntry
End Sub

Function GetTotalFrameRows(TheFrame As Controls, ByVal LowestTAG As Long, ByRef HighestTag As Long, ByVal TotalControls As Long, VTTAG As Long) As Long
    Dim FrameCtrl As Control
    Dim IDX As Long
    Dim strTAGID As String
    Dim TagID As Long
    Dim TotalRows As Long
    
    TotalRows = 0
    HighestTag = 0
    For Each FrameCtrl In TheFrame
        strTAGID = FrameCtrl.Tag
        If IsNumeric(strTAGID) Then
            TagID = CLng(strTAGID)
            If TagID > VTTAG Then
                TagID = TagID - VTTAG
            End If
            If TagID > HighestTag Then
                HighestTag = TagID
            End If
        End If
    Next
    If HighestTag > 0 Then
        TotalRows = (HighestTag - (LowestTAG - 1)) / TotalControls
    End If
    GetTotalFrameRows = TotalRows

End Function

Sub Save_Controls_To_Dic(DeliveryDate As Variant, DeliveryRef As String, ByRef ReturnID As Variant, SearchInFrame As Boolean, DBTable As String, FrameName As String, FrameControls As Controls, FormControls As Variant, _
    ControlTypes As Variant, Dic_ControlsAndFields As Object, ByRef Dic_SavedInfo As Object, ByVal Dic_SavedOtherInfo As Object, ByVal LowestTAG As Long, ByRef HighestTag As Long, _
        ByVal NumControlsPerRow As Long, VTTAG As Long, TotalFormControls As Long, Optional StartFieldIDX As Long = 0)
    Dim RowIDX As Long
    Dim CTRL As Control
    Dim FieldIDX As Long
    Dim TotalFrameRows As Long
    Dim FieldName As String
    Dim FieldValue As String
    Dim strControlName As String
    Dim StartFieldIndex As Long
    Dim FoundItem As Boolean
    Dim UpdateDB As Boolean
    Dim UpdateCriteria As String
    Dim varkey As Variant
    Dim NewFieldNames As String
    Dim strFieldValues As String
    Dim ExcludeFields As String
    Dim ErrMessages As String
    Dim SavedOK As Boolean
    Dim REcordID As String
    Dim ReturnID2 As Variant
    
    If StartFieldIDX > 0 Then
        StartFieldIndex = StartFieldIDX
    Else
        StartFieldIndex = 1
    End If
    NewFieldNames = ""
    strFieldValues = ""
    FoundItem = SearchAccessDB(AccessDBpath, DBTable, "DeliveryDate", DeliveryDate, "DATE", ">=", ReturnID2, _
            "DeliveryReference", DeliveryRef, "STRING", "=", "DeliveryDate", True)
    If FoundItem Then
        UpdateDB = True
        UpdateCriteria = "ID = " & CLng(ReturnID2)
        REcordID = CStr(ReturnID2)
    Else
        UpdateDB = False
        REcordID = "0"
    End If
    Dic_SavedInfo.RemoveAll
    If SearchInFrame Then
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, NumControlsPerRow, VTTAG)
        For RowIDX = 1 To TotalFrameRows
            For FieldIDX = StartFieldIndex To TotalFormControls
                strControlName = FormControls(FieldIDX)
                Set CTRL = FindFormControl(Nothing, ControlTypes(FieldIDX), "", strControlName & CStr(RowIDX), FrameControls)
                If Not CTRL Is Nothing Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(FieldIDX))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        'Sometimes may skip ?
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                Else
                    MsgBox ("Could NOT find Control: " & FormControls(FieldIDX))
                End If
            Next 'INNER to get each field value
            'The other BITS needed to be saved to the same table - tblOperatives:
            NewFieldNames = ""
            strFieldValues = ""
            'NOW SAVE the ROW to DB :
            REcordID = CStr(ReturnID)
            If Len(ReturnID) > 0 Then
                'OK we have a parent ID field value:
                For Each varkey In Dic_SavedInfo
                    If Len(varkey) = 0 Then
                        GoTo NEXT_varkey_IN_Dic_SavedInfo_IN_SAVE_CONTROLS
                    End If
                    FieldName = varkey
                    FieldValue = Dic_SavedInfo(varkey)
                    If Len(NewFieldNames) = 0 Then
                        NewFieldNames = FieldName
                        strFieldValues = FieldValue
                    Else
                        NewFieldNames = NewFieldNames & ";" & FieldName
                        strFieldValues = strFieldValues & ";" & FieldValue
                    End If
NEXT_varkey_IN_Dic_SavedInfo_IN_SAVE_CONTROLS:
                Next 'EACH field in Dic_SavedInfo
                'Dic_SavedOtherInfo should include the LINKID field to the parent ID
                For Each varkey In Dic_SavedOtherInfo
                    If Len(varkey) = 0 Then
                        GoTo NEXT_varkey_IN_Dic_SavedOtherInfo_IN_SAVE_CONTROLS
                    End If
                    FieldName = varkey
                    FieldValue = Dic_SavedOtherInfo(varkey)
                    If Len(NewFieldNames) = 0 Then
                        NewFieldNames = FieldName
                        strFieldValues = FieldValue
                    Else
                        NewFieldNames = NewFieldNames & ";" & FieldName
                        strFieldValues = strFieldValues & ";" & FieldValue
                    End If
NEXT_varkey_IN_Dic_SavedOtherInfo_IN_SAVE_CONTROLS:
                Next 'EACH field in Dic_SavedOtherInfo
                ExcludeFields = ""
                SavedOK = InsertUpdateRecords_Using_Parameters(UpdateDB, REcordID, AccessDBpath, DBTable, NewFieldNames, strFieldValues, UpdateCriteria, ExcludeFields, _
                    ErrMessages, False, ";")
                If Len(ErrMessages) > 0 Then
                    MsgBox ErrMessages
                End If
            Else
                MsgBox ("NO ID returned from parent table")
                Exit Sub
            End If
        Next 'Outer next per row
    Else 'PARENT TABLE - tblLabourHours etc:
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, NumControlsPerRow, VTTAG)
        For FieldIDX = 1 To TotalFormControls
            strControlName = FormControls(FieldIDX)
            Set CTRL = FindFormControl(Nothing, ControlTypes(FieldIDX), "", strControlName, FrameControls)
            If Not CTRL Is Nothing Then
                FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(FieldIDX))
                FieldValue = CTRL
                If Not Dic_SavedInfo.Exists(FieldName) Then
                    'Sometimes this gets skipped because mysteriously the field is already in the SCRIPT DICTIONARY ????
                    Dic_SavedInfo(FieldName) = FieldValue
                End If
            Else
                MsgBox ("Could NOT find Control: " & FormControls(FieldIDX))
            End If
        Next
        'NOW SAVE TO THE DATABASE !
        NewFieldNames = ""
        strFieldValues = ""
        For Each varkey In Dic_SavedInfo
            If Len(varkey) = 0 Then
                GoTo NEXT_varkey_IN_Dic_SavedInfo2_IN_SAVE_CONTROLS
            End If
            If Not IsEmpty(varkey) Then
                FieldName = varkey
                FieldValue = Dic_SavedInfo(varkey)
                If Len(NewFieldNames) = 0 Then
                    NewFieldNames = FieldName
                    strFieldValues = FieldValue
                Else
                    NewFieldNames = NewFieldNames & ";" & FieldName
                    strFieldValues = strFieldValues & ";" & FieldValue
                End If
            End If
NEXT_varkey_IN_Dic_SavedInfo2_IN_SAVE_CONTROLS:
        Next
        'Dic_SavedOtherInfo should include the LINKID field to the parent ID
        For Each varkey In Dic_SavedOtherInfo
            If Len(varkey) = 0 Then
                GoTo NEXT_varkey_IN_Dic_SavedOtherInfo2_IN_SAVE_CONTROLS
            End If
            If Not IsEmpty(varkey) Then
                FieldName = varkey
                FieldValue = Dic_SavedOtherInfo(varkey)
                If Len(NewFieldNames) = 0 Then
                    NewFieldNames = FieldName
                    strFieldValues = FieldValue
                Else
                    NewFieldNames = NewFieldNames & ";" & FieldName
                    strFieldValues = strFieldValues & ";" & FieldValue
                End If
            End If
NEXT_varkey_IN_Dic_SavedOtherInfo2_IN_SAVE_CONTROLS:
        Next 'EACH field in Dic_SavedOtherInfo
        
        ExcludeFields = ""
        'Need to save to the NEW InsertUpdateRecords function which uses parameters.
        'Do we pass the script dictionary also ?
        'Needs to locate each fieldname in the database - as ALL the fields may NOT be needed to save or have values ?
        'Need the FieldType of  each Fieldname.
        SavedOK = InsertUpdateRecords_Using_Parameters(UpdateDB, REcordID, AccessDBpath, DBTable, NewFieldNames, strFieldValues, UpdateCriteria, ExcludeFields, _
                    ErrMessages, False, ";")
        If Len(ErrMessages) > 0 Then
            MsgBox ErrMessages
        End If
        If SavedOK Then
            'IF INSERT PERFORMED - GET THE NEW ID:
            If UpdateDB = False Then
                FoundItem = SearchAccessDB(AccessDBpath, DBTable, "DeliveryDate", DeliveryDate, "DATE", ">=", ReturnID, _
                    "DeliveryReference", DeliveryRef, "STRING", "=", "DeliveryDate", True)
                If FoundItem Then
                    UpdateDB = True
                    UpdateCriteria = "ID = " & CLng(ReturnID)
                Else
                    UpdateDB = False
                End If
            End If
        End If
    End If
    
End Sub


Sub PrepareAndSaveFrameControls(FrameName As String, DeliveryDate As String, DeliveryRef As String, DBTable As String, ByRef Dic_ControlsAndFields As Object, VTTAG As Long)
    Dim FieldnamesArr() As String
    Dim FormControls() As String
    Dim ControlTypes() As String
    Dim ControlIDX As Long
    Dim FormControlKey As Variant
    Dim SearchCriteria As String
    Dim SaveCriteria As String
    Dim ExcludeFields As String
    Dim FieldIDX As Long
    Dim Dic_SavedInfo As Object
    Dim CTRL As Control
    Dim FLMName As String
    Dim FLMCBStartTime As String
    Dim FLMCBEndTime As String
    Dim strFLMStartTime As String
    Dim dtFLMStartTime As Date
    Dim strFLMEndTime As String
    Dim dtFLMEndTime As Date
    Dim TheUserform As frmGI_TimesheetEntry2_1060x630
    Dim FrameControls As Controls
    Dim Entry As String
    Dim ControlCount As Long
    Dim strTAGID As String
    Dim TagID As Long
    Dim OpName As String
    Dim OpActivity As String
    Dim strOpStartTime As String
    Dim OpStartTime As Date
    Dim strOpEndTime As String
    Dim OpEndTime As Date
    Dim FieldName As String
    Dim FieldValue As String
    Dim FieldControlKey As Variant
    Dim RowIndex As Long
    Dim SearchRow As Variant
    Dim varkey As Variant
    Dim NewFieldNames As String
    Dim strFieldValues As String
    Dim ReturnID As Variant
    Dim FoundItem As Boolean
    Dim UpdateDB As Boolean
    Dim ErrMessages As String
    Dim SavedOK As Boolean
    Dim TotalFrameRows As Long
    Dim LowestTAG As Long
    Dim HighestTag As Long
    Dim Dic_SavedOtherInfo As Object
    
    Set Dic_ControlsAndFields = New Scripting.Dictionary
    Dic_ControlsAndFields.RemoveAll
    Set Dic_SavedInfo = New Scripting.Dictionary
    Dic_SavedInfo.RemoveAll
    Set Dic_SavedOtherInfo = New Scripting.Dictionary
    Dic_SavedOtherInfo.RemoveAll
    SearchCriteria = ""
    LowestTAG = 0
    HighestTag = 0
    
    FieldnamesArr = strToStringArray(GetFieldnames_From_ACCESS(DBTable, AccessDBpath, SearchCriteria, ""), ",", 1, False, False, False, "_", False)
    'totalframerows =
    ReDim FormControls(16)
    ReDim ControlTypes(16)
    If UCase(FrameName) = "OPERATIVES" Then
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls
        LowestTAG = 43
        
        FormControls(1) = "txtDeliveryDate"
        FormControls(2) = "txtDeliveryRef"
        FormControls(3) = "comFLMs"
        FormControls(4) = "txtFLMStartTime"
        FormControls(5) = "txtFLMFinishTime"
        FormControls(6) = "comOperativeName"
        FormControls(7) = "comOperativeActivity"
        FormControls(8) = "txtCBOperativeTimeStart"
        FormControls(9) = "txtOperativeTimeStart"
        FormControls(10) = "txtCBOperativeTimeEnd"
        FormControls(11) = "txtOperativeTimeEnd"
        
        ControlTypes(1) = "TEXTBOX"
        ControlTypes(2) = "TEXTBOX"
        ControlTypes(3) = "COMBOBOX"
        ControlTypes(4) = "TEXTBOX"
        ControlTypes(5) = "TEXTBOX"
        ControlTypes(6) = "COMBOBOX"
        ControlTypes(7) = "COMBOBOX"
        ControlTypes(8) = "TEXTBOX"
        ControlTypes(9) = "TEXTBOX"
        ControlTypes(10) = "TEXTBOX"
        ControlTypes(11) = "TEXTBOX"
        
        ControlIDX = 1
        FieldIDX = 1 '0 is the first field
        Do While ControlIDX < 12
            FormControlKey = DBTable & ";" & FrameName & "_" & FormControls(ControlIDX)
            If Not Dic_ControlsAndFields.Exists(FormControlKey) Then
                Dic_ControlsAndFields(FormControlKey) = FieldnamesArr(FieldIDX)
                FieldIDX = FieldIDX + 1
            End If
            ControlIDX = ControlIDX + 1
        Loop
        Dic_SavedInfo.RemoveAll
        
        'dic_SavedInfo - will contain the field and value pairs.
        'GET FLM CONTROL DETAILS:
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, 4, VTTAG)
        Dic_SavedOtherInfo.RemoveAll
        'Search Database to check if this DeliveryRef already exists (for today): - in Save_Controls_To_Dic()
        'Dic_SavedOtherInfo - populate here to include the OTHER bits for the PARENT TABLE:
        'SAVE StartTagID:
            If Not Dic_SavedOtherInfo.Exists("StartTagID") Then
                Dic_SavedOtherInfo("StartTagID") = CStr(LowestTAG)
            End If
            'SAVE EndTagID:
            If Not Dic_SavedOtherInfo.Exists("EndTagID") Then
                Dic_SavedOtherInfo("EndTagID") = CStr(HighestTag)
            End If
            
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Controls
        Call Save_Controls_To_Dic(DeliveryDate, DeliveryRef, ReturnID, False, DBTable, FrameName, FrameControls, FormControls, ControlTypes, Dic_ControlsAndFields, _
            Dic_SavedInfo, Dic_SavedOtherInfo, LowestTAG, HighestTag, 4, VTTAG, 5, 1)
        'GET FRAME OPERATIVE CONTROL DETAILS:
        DBTable = "tblOperatives"
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, 4, VTTAG)
        Dic_SavedOtherInfo.RemoveAll
        'Dic_SavedOtherInfo - populate here to include the OTHER bits for the CHILD TABLE:
        'SAVE FrameRowNumber: - needs to be inside the loop in Save_Controls_To_Dic ???
        If Len(ReturnID) > 0 Then
            If Not Dic_SavedOtherInfo.Exists("LinkID") Then
                Dic_SavedOtherInfo("LinkID") = ReturnID
                MsgBox ("New Operative Rec ID = " & ReturnID)
            End If
        Else
            MsgBox ("NO DELIVERY ID returned from Parent")
            Exit Sub
        End If
        'SAVE Delivery Date
            If Not Dic_SavedOtherInfo.Exists("DeliveryDate") Then
                Dic_SavedOtherInfo("DeliveryDate") = DeliveryDate
            End If
        'SAVE Delivery Reference
            If Not Dic_SavedOtherInfo.Exists("DeliveryReference") Then
                Dic_SavedOtherInfo("DeliveryReference") = DeliveryRef
            End If
        'SET DBTable to tblOperatives HERE.
        
        Call Save_Controls_To_Dic(DeliveryDate, DeliveryRef, ReturnID, True, DBTable, FrameName, FrameControls, FormControls, ControlTypes, Dic_ControlsAndFields, _
            Dic_SavedInfo, Dic_SavedOtherInfo, LowestTAG, HighestTag, 4, VTTAG, 11, 1)
        
        For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
            If UCase(TypeName(CTRL)) = "COMBOBOX" Or UCase(TypeName(CTRL)) = "TEXTBOX" Then
                If UCase(CTRL.Name) = UCase(FormControls(1) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(1))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(2) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(2))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(3) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(3))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
            End If
            If Len(FLMName) > 0 And Len(strFLMStartTime) > 0 And Len(strFLMEndTime) > 0 Then
                Exit For
            End If
        Next
        ControlCount = 0
        RowIndex = 1
        For Each CTRL In FrameControls
            ControlCount = ControlCount + 1
            Dic_SavedInfo.RemoveAll
            'Only seems to save ONE of the controls here
            ' Due to NOT going back to get the next control before saving !
            ' Need an INNER loop - to loop the controls on the SAME row - ie use instr(ctrl.name,cstr(RowIndex))>0
            If UCase(TypeName(CTRL)) = "COMBOBOX" Or UCase(TypeName(CTRL)) = "TEXTBOX" Then
                
                If UCase(CTRL.Name) = UCase(FormControls(4) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(4))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(5) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(5))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(6) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(6))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(7) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(7))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                
                If UCase(CTRL.Name) = UCase(FormControls(8) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(8))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase(FormControls(9) & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & FormControls(9))
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                'Still in ComboBoxes or Textboxes here
            End If
            
            'Save Operative to DB table:
            For Each varkey In Dic_SavedInfo
                If Len(varkey) = 0 Then
                    GoTo NEXT_varkey_IN_PREPAREANDSAVEFRAMECONTROLS
                End If
                FieldName = varkey
                FieldValue = Dic_SavedInfo(varkey)
                If Len(NewFieldNames) = 0 Then
                    NewFieldNames = FieldName
                    strFieldValues = FieldValue
                Else
                    NewFieldNames = NewFieldNames & ";" & FieldName
                    strFieldValues = strFieldValues & ";" & FieldValue
                End If
NEXT_varkey_IN_PREPAREANDSAVEFRAMECONTROLS:
            Next
            'Just check if everything is saved and hunky dory !
            
            'OK Assume that all the Delivery References ARE UNIQUE then we do NOT need to search for the Delivery Date also
            ' - only the DeliveryRef and The FrameRowNumber:
            SearchRow = RowIndex
            FoundItem = SearchAccessDB(AccessDBpath, DBTable, "DeliveryReference", DeliveryRef, "STRING", "=", ReturnID, _
                "FrameRowNumber", SearchRow, "INTEGER", "=", "", False)
            SaveCriteria = ""
            If FoundItem Then
                UpdateDB = True
                SaveCriteria = "ID = " & CLng(ReturnID)
            Else
                UpdateDB = False
            End If
            'ExcludeFields = "DeliveryComments,TAGID"
            ExcludeFields = ""
            'Need to filter out the bad chars in the values. Even the operative names selected from the combo box seem to have bad chars.
            'SavedOK = InsertUpdateRecords(UpdateDB, AccessDBpath, DBTable, NewFieldNames, strFieldValues, SaveCriteria, ExcludeFields, ErrMessages)
            If Len(ErrMessages) > 0 Then
                MsgBox ErrMessages
            End If
        Next
    End If
    
    If UCase(FrameName) = "SHORTS" Then
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Frame_ShortParts.Controls
        LowestTAG = 1001
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, 2, VTTAG)
        FormControls(1) = "txtShortPartNo"
        FormControls(2) = "txtShortQty"
        ControlTypes(1) = "TEXTBOX"
        ControlTypes(2) = "TEXTBOX"
        
        ControlIDX = 1
        FieldIDX = 3 '0 is the first field
        Do While ControlIDX < 3
            FormControlKey = DBTable & ";" & FrameName & "_" & FormControls(ControlIDX)
            If Not Dic_ControlsAndFields.Exists(FormControlKey) Then
                Dic_ControlsAndFields(FormControlKey) = FieldnamesArr(FieldIDX)
                FieldIDX = FieldIDX + 1
            End If
            ControlIDX = ControlIDX + 1
        Loop
        
        ControlCount = 0
        RowIndex = 1
        For Each CTRL In FrameControls
            ControlCount = ControlCount + 1
            If UCase(TypeName(CTRL)) = "COMBOBOX" Or UCase(TypeName(CTRL)) = "TEXTBOX" Then
                If UCase(CTRL.Name) = UCase("txtShortPartNo" & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & CTRL.Name)
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase("txtShortQty" & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & CTRL.Name)
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
            End If
        Next
        
        
    End If
    
    If UCase(FrameName) = "EXTRAS" Then
        Set FrameControls = frmGI_TimesheetEntry2_1060x630.Frame_ExtraParts.Controls
        LowestTAG = 2001
        TotalFrameRows = GetTotalFrameRows(FrameControls, LowestTAG, HighestTag, 2, VTTAG)
        FormControls(1) = "txtExtraPartNo"
        FormControls(2) = "txtExtraQty"
        ControlTypes(1) = "TEXTBOX"
        ControlTypes(2) = "TEXTBOX"
        
        ControlIDX = 1
        FieldIDX = 3 '0 is the first field
        Do While ControlIDX < 3
            FormControlKey = DBTable & ";" & FrameName & "_" & FormControls(ControlIDX)
            If Not Dic_ControlsAndFields.Exists(FormControlKey) Then
                Dic_ControlsAndFields(FormControlKey) = FieldnamesArr(FieldIDX)
                FieldIDX = FieldIDX + 1
            End If
            ControlIDX = ControlIDX + 1
        Loop
        
        ControlCount = 0
        RowIndex = 1
        For Each CTRL In FrameControls
            ControlCount = ControlCount + 1
            If UCase(TypeName(CTRL)) = "COMBOBOX" Or UCase(TypeName(CTRL)) = "TEXTBOX" Then
                If UCase(CTRL.Name) = UCase("txtExtraPartNo" & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & CTRL.Name)
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
                If UCase(CTRL.Name) = UCase("txtExtraQty" & CStr(RowIndex)) Then
                    FieldName = Dic_ControlsAndFields(DBTable & ";" & FrameName & "_" & CTRL.Name)
                    FieldValue = CTRL
                    If Not Dic_SavedInfo.Exists(FieldName) Then
                        Dic_SavedInfo(FieldName) = FieldValue
                    End If
                End If
            End If
        Next
    End If
    
    If UCase(FrameName) = UCase("SUPPLIER COMPLIANCE") Then
        FormControls(1) = "txtArrivedOnTime"
        FormControls(2) = "txtArrivedOnTimeComment"
        FormControls(3) = "txtIsItSafe"
        FormControls(4) = "txtIsItSafeComment"
        FormControls(5) = "txtCompleted"
        FormControls(6) = "txtCompletedComment"
        ControlTypes(1) = "TEXTBOX"
        ControlTypes(2) = "TEXTBOX"
        ControlTypes(3) = "TEXTBOX"
        ControlTypes(4) = "TEXTBOX"
        ControlTypes(5) = "TEXTBOX"
        ControlTypes(6) = "TEXTBOX"
        
        ControlIDX = 1
        FieldIDX = 3 '0 is the first field
        Do While ControlIDX < 3
            FormControlKey = DBTable & ";" & FrameName & "_" & FormControls(ControlIDX)
            If Not Dic_ControlsAndFields.Exists(FormControlKey) Then
                Dic_ControlsAndFields(FormControlKey) = FieldnamesArr(FieldIDX)
                FieldIDX = FieldIDX + 1
            End If
            ControlIDX = ControlIDX + 1
        Loop
        
        
        
    End If

End Sub

Sub SaveFrameInfo(ByRef dic_Saved As Object, DBTable As String)

End Sub

Function SaveAllControls(ByVal DeliveryDate As String, ByVal DeliveryRef As String) As Boolean
    Dim VarControl As Variant
    Dim ctrlProperty As Variant
    Dim TagName As String
    Dim ControlType As String
    Dim ControlName As String
    Dim ControlLowerTag As Long
    Dim ControlUpperTag As Long
    Dim ControlFieldname As String
    Dim ControlDBTable As String
    Dim ControlValue As Variant
    Dim PropertyName As String
    Dim ControlID As Variant
    Dim FrameRow As Long
    Dim TotalFrameRows As Long
    Dim TotalFrames As Long
    Dim TotalControlsPerRow As Long
    Dim VTTAG As Long
    Dim OpLowerTag As Long
    Dim RowNumber As Long
    Dim RowIDX As Long
    Dim SaveOK() As Boolean
    Dim IsUpdate As Boolean
    Dim EncaseFields As Boolean
    Dim RecID As Variant
    Dim Fieldnames() As String
    Dim FieldValues() As String
    Dim UpdateCriteria As String
    Dim ExcludeFields As String
    Dim ErrMessages As String
    Dim ValueDelim As String
    Dim ControlDeliveryDate As Date
    Dim ControlDeliveryRef As String
    Dim DBTables() As String
    Dim FoundRecord As Boolean
    Dim ControlFieldsTable As String
    Dim JustTAG As Boolean
    Dim SearchCriteria As String
    Dim allLookupFields As Variant
    Dim dict_ReturnLookupFields As New Scripting.Dictionary
    Dim LookupFields As String
    Dim LookupValues As String
    Dim FieldsTableLoadedOK As Boolean
    Dim Messages As String
    
    ReDim DBTables(5)
    ReDim Fieldnames(5)
    ReDim FieldValues(5)
    ReDim SaveOK(5)
    
    If Len(DeliveryDate) = 0 Then
        MsgBox ("No Delivery Date Passed to SAVE procedure")
        Exit Function
    End If
    If Len(DeliveryRef) = 0 Then
        MsgBox ("No Delivery Reference Passed to SAVE procedure")
        Exit Function
    End If
    SaveAllControls = False
    
    DBTables(1) = "tblDeliveryInfo"
    DBTables(2) = "tblSupplierCompliance"
    DBTables(3) = "tblLabourHours"
    DBTables(4) = "tblOperatives"
    DBTables(5) = "tblShortsAndExtraParts"
    
    ControlFieldsTable = "tblFieldsAndTAGS"
    
    Fieldnames(1) = ""
    Fieldnames(2) = ""
    Fieldnames(3) = ""
    Fieldnames(4) = ""
    Fieldnames(5) = ""
    FieldValues(1) = ""
    FieldValues(2) = ""
    FieldValues(3) = ""
    FieldValues(4) = ""
    FieldValues(5) = ""
    
    Messages = ""
    For Each VarControl In ctrlCollection
        If VarControl Is Nothing Then
            GoTo NEXT_SAVE_CONTROLS_ITERATION
        End If
        Set ctrlProperty = VarControl
        ControlDeliveryDate = ctrlProperty.ControlDeliveryDate
        ControlDeliveryRef = ctrlProperty.ControlDeliveryRef
        
        If UCase(CStr(ControlDeliveryDate)) = UCase(DeliveryDate) Then
            If UCase(ControlDeliveryRef) = UCase(DeliveryRef) And UCase(ctrlProperty.ControlType) = "COMBOBOX" And UCase(ctrlProperty.ControlType) = "TEXTBOX" Then
                ControlFieldname = ctrlProperty.ControlFieldname
                
                TagName = ctrlProperty.ControlTAG
                ControlID = ctrlProperty.ControlID 'DeliveryDate and DeliveryRef and TAG NUMBER - for actual ctrlCollection KEY.
                ControlName = ctrlProperty.ControlName
                ControlLowerTag = ctrlProperty.ControlStartTAG
                ControlUpperTag = ctrlProperty.ControlEndTAG
                ControlDBTable = ctrlProperty.ControlDBTable
                
                ControlValue = ctrlProperty.ControlValue
                
                If Len(ControlFieldname) = 0 Then
                    'GET it from the Lookup table ! - will need some form of ID though - pass the Control Name:
                    Messages = Messages & vbCrLf & " No Fieldname for : " & ControlName
                    GoTo NEXT_SAVE_CONTROLS_ITERATION
                End If
                
                'OK WARNING - THE COLLECTION MAY CONTAIN OTHER KEYS NOT ASSOCIATED WITH THE CURRENT DELIVERY DATE OR REFERENCE.
                
                If UCase(ControlDBTable) = UCase("tblDeliveryInfo") Then
                    'THIS WILL BE AN UPDATE:
                    If Len(Fieldnames(1)) = 0 Then
                        Fieldnames(1) = ControlFieldname
                    Else
                        Fieldnames(1) = Fieldnames(1) & "," & ControlFieldname
                    End If
                    If Len(FieldValues(1)) = 0 Then
                        FieldValues(1) = ControlValue
                    Else
                        FieldValues(1) = FieldValues(1) & "," & ControlValue
                    End If
                    
                    EncaseFields = False
                    ErrMessages = ""
                    ValueDelim = ","
                    ExcludeFields = ""
                    'test for IsUpdate: Search Database first for Delivery Date and Reference:
                    FoundRecord = SearchAccessDB(AccessDBpath, DBTables(1), "DeliveryDate", DeliveryDate, "DATE", "=", RecID, _
                        "DeliveryReference", DeliveryRef, "STRING", "=")
                    If FoundRecord Then
                        IsUpdate = True
                        UpdateCriteria = "ID = " & RecID
                    Else
                        IsUpdate = False
                        UpdateCriteria = ""
                    End If
                    
                    
                End If
                If UCase(ControlDBTable) = UCase("tblSupplierCompliance") Then
                    'STATIC DATA - NOW ROWS. WILL BE INSERT TO START WITH - THEN USER MAY CHANGE THE NOs TO YES AND REMOVE THE COMMENTS.
                    
                    If Len(Fieldnames(2)) = 0 Then
                        Fieldnames(2) = ControlFieldname
                    Else
                        Fieldnames(2) = Fieldnames(2) & "," & ControlFieldname
                    End If
                    If Len(FieldValues(2)) = 0 Then
                        FieldValues(2) = ControlValue
                    Else
                        FieldValues(2) = FieldValues(2) & "," & ControlValue
                    End If
                    
                    EncaseFields = False
                    ErrMessages = ""
                    ValueDelim = ","
                    ExcludeFields = ""
                    'test for IsUpdate: Search Database first for Delivery Date and Reference:
                    FoundRecord = SearchAccessDB(AccessDBpath, DBTables(2), "DeliveryDate", DeliveryDate, "DATE", "=", RecID, _
                        "DeliveryReference", DeliveryRef, "STRING", "=")
                    If FoundRecord Then
                        IsUpdate = True
                        UpdateCriteria = "ID = " & RecID
                    Else
                        IsUpdate = False
                        UpdateCriteria = ""
                    End If
                    
                    
                    
                End If
                If UCase(ControlDBTable) = UCase("tblLabourHours") Then
                    'STATIC DATA - INSERT first and then maybe UPDATE later is needed.
                    'CHECK if RECORD EXISTS FIRST.
                    
                    If Len(Fieldnames(3)) = 0 Then
                        Fieldnames(3) = ControlFieldname
                    Else
                        Fieldnames(3) = Fieldnames(3) & "," & ControlFieldname
                    End If
                    If Len(FieldValues(3)) = 0 Then
                        FieldValues(3) = ControlValue
                    Else
                        FieldValues(3) = FieldValues(3) & "," & ControlValue
                    End If
                    
                    EncaseFields = False
                    ErrMessages = ""
                    ValueDelim = ","
                    ExcludeFields = ""
                    'test for IsUpdate: Search Database first for Delivery Date and Reference:
                    FoundRecord = SearchAccessDB(AccessDBpath, DBTables(3), "DeliveryDate", DeliveryDate, "DATE", "=", RecID, _
                        "DeliveryReference", DeliveryRef, "STRING", "=")
                    If FoundRecord Then
                        IsUpdate = True
                        UpdateCriteria = "ID = " & RecID
                    Else
                        IsUpdate = False
                        UpdateCriteria = ""
                    End If
                    
                    
                End If
                If UCase(ControlDBTable) = UCase("tblOperatives") Then
                    'NEED TO LOOP THROUGH EACH ROW IN THE FRAME_OPERATIVES CONTROLS HERE AND SAVE.
                    'THIS WILL MORE LIKELY BE AN INSERT - BUT CAN ALSO BE UPDATE - IF USER IS CHANGING ANY INFO WITHIN THE ROWS.
                    
                    TotalControlsPerRow = 4
                    VTTAG = 400
                    RowNumber = Get_NumericPartOfString(ControlName)
                    ctrlProperty.ControlRowNumber = RowNumber
                    TotalFrames = GetTotalFrameRows(frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlLowerTag, ControlUpperTag, _
                        TotalControlsPerRow, VTTAG)
                    'ok so we have total frames and the highest tag now.
                    ctrlProperty.ControlEndTAG = ControlUpperTag
                    'Turn all the values into fields and values to save:
                    
                    'if the RowNumber is different to the PREVIOUS ROW NUMBER then SAVE.
                    
                    'RowIDX = 1
                    'Do While RowIDX <= TotalFrameRows
                    
                    'Loop
                    
        
                End If
                If UCase(ControlDBTable) = UCase("tblShortsAndExtraParts") Then
                    'AGAIN NEED TO LOOP THROUGH EACH ROW OF CONTROLS IN FRAME_SHORTSANDEXTRAS AND SAVE EACH ROW:
                    'THIS WILL MOST LIKELY BE AN INSERT - BUT COULD ALSO BE AN UPDATE IF ANY PART NUMBER OR QUANTITY NEEDS CHANGING.
                    ' -  HEY WE ALL MAKE MISTAKES !
                    
                End If
            End If
        End If
NEXT_SAVE_CONTROLS_ITERATION:
    Next
            
    If Len(Messages) > 0 Then
        MsgBox ("Fieldnames Missing: ")
    Else
    
            
        SaveOK(1) = InsertUpdateRecords_Using_Parameters(IsUpdate, RecID, AccessDBpath, ControlDBTable, Fieldnames(1), FieldValues(1), UpdateCriteria, _
                    ExcludeFields, ErrMessages, EncaseFields, ValueDelim)
        SaveOK(2) = InsertUpdateRecords_Using_Parameters(IsUpdate, RecID, AccessDBpath, ControlDBTable, Fieldnames(2), FieldValues(2), UpdateCriteria, _
                            ExcludeFields, ErrMessages, EncaseFields, ValueDelim)
        SaveOK(3) = InsertUpdateRecords_Using_Parameters(IsUpdate, RecID, AccessDBpath, ControlDBTable, Fieldnames(3), FieldValues(3), UpdateCriteria, _
                            ExcludeFields, ErrMessages, EncaseFields, ValueDelim)
            
    End If

    
    'BUT we need the record ID of the parent record first ?????????????????? so the LINKID is populated.
    
    'OK so we have two loops then. The first to establish and save the STATIC control records and setup record IDs.
    'The second loop will go through and filter on just the CHILD TABLES and will have an inner loop that will detect when a row change happens.
    
End Function

Function Get_NumericPartOfString(TheString As String) As Long
    Dim IDX As Long
    Dim strPart As String
    Dim strNumber As String
    Dim NumericPart As Long
    
    NumericPart = 0
    strPart = TheString
    IDX = 1
    Do While IDX < Len(TheString)
        If IsNumeric(Mid(TheString, IDX, 1)) Then
            strNumber = strNumber & Mid(TheString, IDX, 1)
        End If
        IDX = IDX + 1
    Loop
    If Len(strNumber) > 0 Then
        NumericPart = CLng(strNumber)
    End If
    
    Get_NumericPartOfString = NumericPart
    
End Function

Sub InsertEntry_Into_ACCESS(DBTable As String, AccessPath As String, ByVal DeliveryDate As String, ByVal DeliveryRef As String, _
        Optional SpecificEntry As String = "", Optional TagLowRange As Long = 0, Optional TagUpperRange As Long = 0, _
        Optional TimeSymbol As Long = 136, Optional VTTAG As Long = 400, Optional scTAG As Long = 800, _
        Optional ShortTAG As Long = 1000, Optional ExtraTAG As Long = 2000)
    'Author: DANIEL GOSS - MAY 2018
    Dim FieldName As String
    Dim Fieldnames As String
    Dim NewFieldNames As String
    Dim NewFieldValues As String
    Dim FieldValues() As String
    Dim FieldValue As String
    Dim strFieldValues As String
    Dim FieldnameArr() As String
    Dim TestFields() As String
    Dim TestValues() As String
    Dim SavedOK As Boolean
    Dim UpdateDB As Boolean
    Dim UpdateCriteria As String
    Dim ExcludeFields As String
    Dim ErrMessages As String
    Dim FoundDate As Boolean
    Dim ReturnID As String
    Dim dic_ControlInfo As Scripting.Dictionary
    Dim dic_Frame_Controls As Scripting.Dictionary
    Dim varkey As Variant
    Dim REcordID As String
    
    Set dic_ControlInfo = CreateObject("Scripting.Dictionary")
    Set dic_Frame_Controls = CreateObject("Scripting.Dictionary")
    dic_ControlInfo.RemoveAll
    
    ReDim FieldnameArr(1)
    ReDim TestFields(1)
    ReDim TestValues(1)
    
    FoundDate = False
    UpdateCriteria = ""
    Fieldnames = GetFieldnames_From_ACCESS(DBTable, AccessPath)
    FieldnameArr = strToStringArray(Fieldnames, ",", 1) 'Assume Fields in the database are in the same order as the tag numbers.
    
    If UCase(DBTable) = UCase("tblDeliveryInfo") Then
        FieldValues = ExtractInfo_From_Controls("DELIVERYINFO", dic_ControlInfo, FieldnameArr, DeliveryDate, DeliveryRef, SpecificEntry, TagLowRange, TagUpperRange, _
            TimeSymbol, VTTAG, scTAG, ShortTAG, ExtraTAG)
        FoundDate = False
        UpdateCriteria = ""
        'FieldValues may NOT be correct. but Dic_ControlInfo is good.
        NewFieldNames = ""
        strFieldValues = ""
        For Each varkey In dic_ControlInfo
            If Len(varkey) = 0 Then
                GoTo NEXT_varkey_IN_INSERTENTRY_INTO_ACCESS
            End If
            FieldName = varkey
            FieldValue = dic_ControlInfo(varkey)
            If Len(NewFieldNames) = 0 Then
                NewFieldNames = FieldName
                strFieldValues = FieldValue
            Else
                NewFieldNames = NewFieldNames & "," & FieldName
                strFieldValues = strFieldValues & ";" & FieldValue
            End If
NEXT_varkey_IN_INSERTENTRY_INTO_ACCESS:
        Next
        TestFields = strToStringArray(NewFieldNames, ",", 1)
        TestValues = strToStringArray(strFieldValues, ";", 1)
        
        FoundDate = SearchAccessDB(AccessDBpath, DBTable, "DeliveryDate", DeliveryDate, "DATE", ">=", ReturnID, _
            "DeliveryReference", DeliveryRef, "STRING", "=", "DeliveryDate", True)
        If FoundDate Then
            UpdateDB = True
            UpdateCriteria = "ID = " & CLng(ReturnID)
        Else
            UpdateDB = False
        End If
        ExcludeFields = "DeliveryComments,TAGID"
        REcordID = CStr(ReturnID)
        'New version of InsertupdateRecords_with_parameters - required here.
        SavedOK = InsertUpdateRecords_Using_Parameters(UpdateDB, REcordID, AccessDBpath, DBTable, NewFieldNames, strFieldValues, UpdateCriteria, ExcludeFields, ErrMessages, _
            False, ";")
        'SavedOK = InsertUpdateRecords(UpdateDB, AccessPath, DBTable, NewFieldNames, strFieldValues, Criteria, ExcludeFields, ErrMessages)
        If Len(ErrMessages) > 0 Then
            MsgBox ErrMessages
        End If
    End If
    
    If UCase(DBTable) = UCase("tblLabourHours") Then
        Call PrepareAndSaveFrameControls("OPERATIVES", DeliveryDate, DeliveryRef, DBTable, dic_Frame_Controls, VTTAG)
        'calculate Total Hours per Delivery REF:
        
    End If
    
    If UCase(DBTable) = UCase("tblShortsAndExtraParts") Then
        Call PrepareAndSaveFrameControls("SHORTS", DeliveryDate, DeliveryRef, DBTable, dic_Frame_Controls, ShortTAG)
    End If
    
    If UCase(DBTable) = UCase("tblShortsAndExtraParts") Then
        Call PrepareAndSaveFrameControls("EXTRAS", DeliveryDate, DeliveryRef, DBTable, dic_Frame_Controls, ExtraTAG)
    End If
    
    If UCase(DBTable) = UCase("tblSupplierCompliance") Then
        Call PrepareAndSaveFrameControls("SUPPLIER COMPLIANCE", DeliveryDate, DeliveryRef, DBTable, dic_Frame_Controls, scTAG)
    End If
    
End Sub

Function ExtractInfo_From_Controls(Section As String, ByRef dic_ControlInfo As Object, ByVal Fieldnames As Variant, _
    ByVal DeliveryDate As String, ByVal DeliveryRef As String, Optional SpecificEntry As String = "", _
    Optional TagLowRange As Long = 0, Optional TagUpperRange As Long = 0, _
        Optional TimeSymbol As Long = 136, Optional VTTAG As Long = 400, Optional scTAG As Long = 800, Optional ShortTAG = 1000, Optional ExtraTAG = 2000) As String()
    'ExtractInfo_From_Controls = Nothing
    Dim CTRL As Control
    Dim myRow As Long
    Dim TagID As Long
    Dim Entry As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim DontSave As Boolean
    Dim ControlCount As Long
    Dim ThisWB As Workbook
    Dim dtDateEntry As Date
    Dim strDate As String
    Dim FieldValues() As String
    Dim FieldIndex As Long
    Dim FieldName As String
    Dim Increment As Long
    Dim NumberOfRows As Long
    Dim dic_OperativeControlFields As Object
    'Dim AccessDBPath As String
    
    Set dic_ControlInfo = CreateObject("Scripting.Dictionary")
    dic_ControlInfo.RemoveAll
    Set dic_OperativeControlFields = New Scripting.Dictionary
    dic_ControlInfo.CompareMode = vbTextCompare
    dic_OperativeControlFields.CompareMode = vbTextCompare
    ControlCount = 0
    ReDim FieldValues(60)
    ExtractInfo_From_Controls = FieldValues
    If Len(DeliveryDate) = 0 Then
        MsgBox ("NO DELIVERY DATE Passed to INSERT")
        Exit Function
    End If
    If Len(DeliveryRef) = 0 Then
        MsgBox ("NO DELIVERY REFERENCE passed to INSERT")
        Exit Function
    End If
    FieldIndex = 1
    
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        'Either - Find the control KEY in the collection and update the TEXT property with the NEW values for each control
        ' - requiring New function - which will use FIND FORM CONTROL in a loop and replace with the new value.
        ' OR - REMOVE the entry from the collection (but not the dynamic control?) and call AddControlInfo() again ???
        'Since the AfterUpdate event does not seem to work - this will be the only way the controls in the COLLECTION can be updated.
        TagID = 0
        DontSave = False
        ControlCount = ControlCount + 1
        If UCase(TypeName(CTRL)) = "TEXTBOX" Or UCase(TypeName(CTRL)) = "COMBOBOX" Then
            Entry = CTRL
            If IsDate(Entry) Then 'how is the date and time stored on the form ?
                'or is the collection maintained dynamically as the entries are entered into the boxes ?
                dtDateEntry = CDate(Entry) 'Only records the time and NOT the date - needs to be dtTimeEntry maybe ?
            End If
            'MsgBox ("ASCII=" & CStr(Asc(Entry)))
            'txtCtrl = ctrl
            If Len(CTRL.Tag) > 0 Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    TagID = CLng(CTRL.Tag)
                    'BUT Supplier Compliance = sc1 to sc6 ??? - no 800 to 805
                End If
                If UCase(Section) = "DELIVERYINFO" Then
                    Increment = 1
                    FieldIndex = (TagID - TagLowRange) + Increment
                ElseIf UCase(Section) = "LABOUR" Then
                    Increment = 3
                    
                    If TagID > VTTAG And TagID < VTTAG + 350 Then
                        NumberOfRows = TagUpperRange / 4
                    End If
                    FieldIndex = (TagID - TagLowRange) + Increment
                ElseIf UCase(Section) = "SUPPLIER COMPLIANCE" Then
                    Increment = 1
                    FieldIndex = (TagID - TagLowRange) + Increment
                ElseIf UCase(Section) = "EXTRA SHORTS" Then
                    Increment = 1
                    FieldIndex = (TagID - TagLowRange) + Increment
                End If
            End If
        End If
        
        If TagID > 0 Then
            If Len(SpecificEntry) > 0 Then
                FinalEntry = SpecificEntry
            Else
                FinalEntry = Entry
                If IsDate(Entry) Then
                    dtDateEntry = CDate(Entry) 'if FinalEntry just contains a TIME - search for colon - then add the delivery date to it.
                    FinalEntry = dtDateEntry
                End If
            End If
            
            If Len(FinalEntry) = 0 Then
                If InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    FinalEntry = "NO"
                    
                End If
                    'All other blank text boxes
                If InStr(1, CTRL.Name, "Start", vbTextCompare) > 0 Then
                    FinalEntry = ""
                End If
                If InStr(1, CTRL.Name, "Finish", vbTextCompare) > 0 Then
                    FinalEntry = ""
                End If
                'FinalEntry = ""
                'End If
            End If
            If Len(FinalEntry) = 1 Then
                If Asc(FinalEntry) = 80 And InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    'Found a TICK in a TEXTBOX !
                    FinalEntry = "YES"
                End If
                'only works if the timesymbol matches the final entry symbol passed - single char from the check box.
                If Asc(FinalEntry) = TimeSymbol And InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                    'Found a CLOCK - 6 O Clock icon in a TEXTBOX !
                    'FinalEntry = Format(Now(), "dd/mm/YYYY HH:MM:ss")
                    DontSave = True
                    
                End If
            End If
            'TAGID is the TAG number.
            'Need a way to associate this tag number with the fieldname.
            
            If TagID > 0 Then
                If TagLowRange > 0 And TagID <= TagUpperRange And TagID >= TagLowRange Then
                    If DontSave = False Then
                        If IsDate(FinalEntry) Then
                            'MsgBox ("DATE = " & FinalEntry)
                            'RemoteFilePath is the public variable holding the full path to the remote ACCESS DATABASE with data.
                            dtDateEntry = Convert_strDateToDate(FinalEntry, True)
                            strDate = CStr(dtDateEntry)
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).NumberFormat = "dd/mmm/yyyy HH:mm"
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).value = dtDateEntry
                        Else
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).value = FinalEntry
                            'NEED A LOOKUP from TAGID to Fieldname:
                            'NOW is the looup a SPREADSHEET or is it another ACCESS TABLE ????
                            'spread over 3 tables.
                        End If
                        'TAGID not necesarily follows in order. How do we link the FIELD NAME to the Field Value / TAG ID ???
                        
                        'FieldIndex = (TAGID - TagLowRange) + 1 'calculated earlier !
                        
                        If FieldIndex > 0 Then
                            FieldValues(FieldIndex) = FinalEntry
                            FieldName = UCase(Fieldnames(FieldIndex))
                            If Not dic_ControlInfo.Exists(FieldName) Then
                                dic_ControlInfo.Add FieldName, FinalEntry
                                'dic_ControlInfo(Fieldname) = FinalEntry
                            End If
                        End If
                        'FieldValues(TAGID) = FinalEntry
                        'Fieldvalues = accumulate the fieldvalues in order of TAGID.
                        If Len(SpecificEntry) > 0 Then Exit For
                    End If
                End If
                If TagLowRange = 0 And TagUpperRange = 0 Then
                    If DontSave = False Then
                        If IsDate(FinalEntry) Then
                            dtDateEntry = Convert_strDateToDate(FinalEntry, True)
                            strDate = CStr(dtDateEntry)
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).NumberFormat = "dd/mmm/yyyy HH:mm"
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).value = dtDateEntry
                        Else
                            'ThisWB.Worksheets(MainWorksheet).Cells(myRow, TAGID).value = FinalEntry
                            
                        End If
                        
                        If FieldIndex >= 0 Then
                            FieldValues(FieldIndex) = FinalEntry
                            FieldName = UCase(Fieldnames(FieldIndex))
                            If Not dic_ControlInfo.Exists(FieldName) Then
                                dic_ControlInfo(FieldName) = FinalEntry
                            End If
                        End If
                        If Len(SpecificEntry) > 0 Then Exit For
                    End If
                End If
            End If
        End If
    Next
    
    ExtractInfo_From_Controls = FieldValues

End Function

Sub SaveTAblesAndFieldsToLookup(DBTable As String)
Dim FieldIDX As Long
Dim FieldArr() As String
Dim Fieldnames As String
Dim Criteria As String
Dim FoundField As Boolean
Dim REcordID As Long
Dim FieldName As String
Dim SavedOK As Boolean
Dim TheFields As String
Dim TheValues As String
Dim ExcludeFields As String
Dim UpdateCriteria As String
Dim ErrMessage As String
Dim EncaseFields As Boolean
Dim LookupTable As String

Criteria = ""

Fieldnames = GetFieldnames_From_ACCESS(DBTable, AccessDBpath, Criteria, "")
'check each field. Does it already exist in the database ?
FieldArr = strToStringArray(Fieldnames, ",", 1, False, False, False, "_", False)
EncaseFields = False
ExcludeFields = ""

LookupTable = "tblFieldsAndTAGS"

For FieldIDX = 1 To UBound(FieldArr)
    FieldName = ""
    FieldName = FieldArr(FieldIDX)
    TheFields = "TableName,TableField"
    TheValues = DBTable & "," & FieldName
    FoundField = SearchAccessDB(AccessDBpath, LookupTable, "TableName", DBTable, "STRING", "=", REcordID, "TableField", FieldName, "STRING", "=")
    If Not FoundField Then
        'Insert into tblLookupFieldsandTAGS
        SavedOK = InsertUpdateRecords_Using_Parameters(False, REcordID, AccessDBpath, LookupTable, TheFields, TheValues, _
            UpdateCriteria, ExcludeFields, ErrMessage, EncaseFields, ",")
        
    End If
Next

End Sub

Sub SaveControlNamesToLookup(DBTable As String)
    Dim CTRL As Control
    Dim ControlTAG As String
    Dim ControlName As String
    Dim REcordID As Long
    Dim UpdateCriteria As String
    Dim SaveTable As String
    Dim FoundTAG As Boolean
    Dim SavedOK As Boolean
    Dim Fields As String
    Dim Values As String
    Dim ExcludeFields As String
    Dim ErrMessage As String
    
    SaveTable = "tblFieldsAndTAGs"
    'Loop through all the controls on the USERFORM to get the TAG NUMBERS.
    ' - Then search through the table for each tag to get the RECORD NUMBER to update with the control name.
    
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        If UCase(TypeName(CTRL)) = "TEXTBOX" Or UCase(TypeName(CTRL)) = "COMBOBOX" Then
            ControlTAG = CTRL.Tag
            ControlName = CTRL.Name
            FoundTAG = SearchAccessDB(AccessDBpath, SaveTable, "TAGID", ControlTAG, "STRING", "=", REcordID, "TableName", DBTable, "STRING", "=")
            If FoundTAG Then
                Fields = "controlName"
                UpdateCriteria = ""
                ExcludeFields = ""
                Values = ControlName
                ErrMessage = ""
                SavedOK = InsertUpdateRecords_Using_Parameters(True, REcordID, AccessDBpath, SaveTable, Fields, Values, UpdateCriteria, _
                    ExcludeFields, ErrMessage, False, ",")
                'Exit For - this would stop the search for whole table.
            End If
        End If
        
    Next
    
    
End Sub


Sub Execute_SaveTAblesAndFieldsToLookup()

Call SaveTAblesAndFieldsToLookup("tblDeliveryinfo")
Call SaveTAblesAndFieldsToLookup("tblLabourHours")
Call SaveTAblesAndFieldsToLookup("tblOperatives")
Call SaveTAblesAndFieldsToLookup("tblShortsAndExtraParts")
Call SaveTAblesAndFieldsToLookup("tblSupplierCompliance")

End Sub

Sub Execute_UpdateControlNamesInLookup()

Call SaveControlNamesToLookup("tblDeliveryinfo")
Call SaveControlNamesToLookup("tblLabourHours")
Call SaveControlNamesToLookup("tblOperatives")
Call SaveControlNamesToLookup("tblShortsAndExtraParts")
Call SaveControlNamesToLookup("tblSupplierCompliance")


End Sub

Function ExtractFieldsWithoutValues(Fieldnames As String, FieldValues As String, ByRef NewFieldValues As String) As String
    Dim IDX As Long
    Dim FieldsArr() As String
    Dim ValuesArr() As String
    Dim FieldValue As String
    Dim NewFieldArr() As String
    Dim NewValueArr() As String
    Dim FieldName As String
    Dim NewFieldNames As String
    Dim NewIDX As Long
    
    ReDim FieldsArr(1)
    
    ExtractFieldsWithoutValues = ""
    FieldsArr = strToStringArray(Fieldnames, ",")
    ReDim ValuesArr(UBound(FieldsArr) + 1)
    ReDim NewFieldArr(1)
    ReDim NewValueArr(1)
    ValuesArr = strToStringArray(FieldValues, ",")
    'loop through the values - save only those that contain values
    IDX = 0
    NewIDX = 1
    Do While IDX < UBound(FieldsArr)
        FieldName = FieldsArr(IDX)
        FieldValue = ValuesArr(IDX)
        If Len(ValuesArr(IDX)) > 0 Then
            If Len(NewFieldNames) = 0 Then
                NewFieldNames = FieldName
                NewFieldValues = FieldValue
            Else
                NewFieldNames = NewFieldNames & ";" & FieldName
                NewFieldValues = NewFieldValues & ";" & FieldValue
            End If
            
        End If
        ReDim Preserve ValuesArr(UBound(ValuesArr) + 1)
        ReDim Preserve NewFieldArr(UBound(NewFieldArr) + 1)
        ReDim Preserve NewValueArr(UBound(NewValueArr) + 1)
        
        IDX = IDX + 1
    Loop
    
    ExtractFieldsWithoutValues = NewFieldNames


End Function

Sub InsertTime(WB As Workbook, CBctrl As Control, VisTimeBoxCtrl As Control, RowNumber As Long, LowerLimit As Long, UpperLimit As Long, symbol As Long)
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    If Len(CBctrl.Text) > 0 Then
        CBctrl.Text = ""
        Call RemoveEntry(ThisWB, "NO", VisTimeBoxCtrl, RowNumber, LowerLimit)
        Call SetControlBackgroundColour(CStr(LowerLimit), vbWhite)
        
    Else
        'CBctrl.Font.Name = "Tahoma"
        CBctrl.Font.Name = "Wingdings 2"
        CBctrl.Text = Chr(symbol)
        If RowNumber > 0 Then
            Call InsertEntry(ThisWB, Format(Now(), "dd/mm/YYYY HH:MM:ss"), RowNumber, LowerLimit, UpperLimit, symbol)
            VisTimeBoxCtrl.Text = Format(Now(), "HH:MM:ss")
        End If
        
    End If
End Sub


Sub InsertTimeIntoControl(SearchTAG As Long, SearchControlName As String, ByRef DATETIMEOUT As Date, Optional AddVTAG As Boolean = True, _
    Optional VTTAG As Long = 400)
Dim ReturnTagNumber As Long
Dim TimeControl As Control

'REDUNDANT !
'PERFORMED WITHIN THE EVENT PROCEDURE OF clsControls instead .

If AddVTAG Then
    ReturnTagNumber = SearchTAG + VTTAG
Else
    ReturnTagNumber = SearchTAG
End If
'TICK or AT symbol not being displayed - or not dissapearing ??
'Now search for the ReturnTagNumber CONTROL on the form and insert the DATE and TIME into it !
If SearchTAG > 0 Then
    Set TimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", CStr(ReturnTagNumber), "")
    'Also add the current datetime into the appropriate control from TIMEOUT
End If
If Len(SearchControlName) > 0 Then
    Set TimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", SearchControlName)
    'Also add the current datetime into the appropriate control from TIMEOUT
End If
TimeControl.Text = Format(DATETIMEOUT, "HH:MM:ss")

End Sub

Sub InsertValueIntoControl(ControlType As String, SearchTAG As Long, SearchControlName As String, ByVal InsertValue As String)
Dim OutputControl As Control
Dim tempControl As clsControls

'During SEARCH - all Controls collection will be populated.
If SearchTAG > 0 Then
    Set OutputControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, ControlType, CStr(SearchTAG), "")
End If
If Len(SearchControlName) > 0 Then
    Set OutputControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, ControlType, "", SearchControlName)
End If

OutputControl.Text = InsertValue

'Need to search the clsControls / Controls Collection OBJECt for this Control and update the value.


End Sub

Sub InsertValueIntoControlCollection(ControlType As String, SearchTAG As Long, SearchControlName As String, ByVal NewValue As String)
Dim OutputControl As Control
Dim tempControl As clsControls

If SearchTAG > 0 Then
    Set OutputControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, ControlType, CStr(SearchTAG), "")
End If
If Len(SearchControlName) > 0 Then
    Set OutputControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, ControlType, "", SearchControlName)
End If

OutputControl.Text = NewValue
'Now insert / update in the Control Collection:


End Sub


Function InsertUpdateRecords(ByVal Update As Boolean, ByVal DBPath As String, ByVal DBTable As String, ByVal Fieldnames As String, ByVal FieldValues As String, _
    Optional ByVal Criteria As String = "", Optional ByVal ExcludeFields As String = "", Optional ByRef ErrMessages As String, _
    Optional EncaseFields As Boolean = False, Optional FieldValueDelim As String = ";") As Boolean
Dim cn As Object
Dim rs As Object
Dim myConn As String
Dim provider As String
Dim con As OLEDBConnection
'Dim cmd As oledbcommand
'Dim da As oledbdataadapter
Dim strSQL As String
Dim strDeliveryRef As String
Dim dtDeliveryDate As Date
Dim strDeliveryDate As String
Dim ExcludedFields As String
Dim FieldTypeArr() As String
Dim FieldType As Integer
Dim Dic As Object

InsertUpdateRecords = False

On Error GoTo Err_InsertUpdateRecords

'myConn = "G:\MIS\Goodsin Timesheet Data\GoodsInTimesheetRecords.accdb"
If Len(DBPath) > 0 Then
    myConn = DBPath
Else
    myConn = ""
    ErrMessages = "DATABASE PATH NOT SPECIFIED"
    Exit Function
End If

'Set cn = ADODB.Connection

Set cn = New ADODB.Connection
    
    Application.ScreenUpdating = False

    'Set TargetRange = ThisWorkbook.Sheets("Prefs").Range("A10")

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & myConn & ";"
    
    If Len(DBTable) > 0 Then
        'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
        'strSQL = strSQL & " VALUES (""" & Fieldvalues & """);"
    Else
        MsgBox ("DATABASE TABLE NOT SPECIFIED")
        cn.Close
        Set cn = Nothing
        Exit Function
    End If
    ExcludedFields = ExcludeFields
    'NEED to use SearchAccessDB() as boolean
    Set Dic = CreateObject("Scripting.Dictionary")
    FieldTypeArr = Get_FieldTypes(DBPath, DBTable, Dic)
    'GETTING AUTHENTICATION ERROR when the above function is run.
    
    If Update = True Then
        strSQL = PrepareUpdate(DBPath, DBTable, Fieldnames, FieldValues, Criteria, ExcludedFields, True, FieldValueDelim, FieldValueDelim)
    Else
        strSQL = PrepareInsert(DBPath, DBTable, Fieldnames, FieldValues, Dic, ExcludedFields, True, FieldValueDelim, FieldValueDelim)
    End If
    'strDeliveryRef = "TEST-DAN-YES"
    'Fieldnames = "DeliveryDate,DeliveryReference"
    'dtDeliveryDate = CDate("21/05/2018")
    'strDeliveryDate = Format(dtDeliveryDate, "dd/MM/yyyy")
    'FieldValues = Chr(34) & strDeliveryDate & Chr(34) & "," & Chr(34) & strDeliveryRef & Chr(34)
    'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
    'strSQL = strSQL & " VALUES (" & FieldValues & ")"
    'MsgBox ("SQL= " & strSQL)
    cn.Execute strSQL
    cn.Close
    Set cn = Nothing
    Set Dic = Nothing
    InsertUpdateRecords = True

Exit Function
Err_InsertUpdateRecords:

Call Error_Report("InsertUpdateRecords()")
     'ErrMessages = "Error in InsertUpdateRecords"

End Function

Function InsertUpdateRecords_Using_Parameters(ByVal UpdateRecords As Boolean, ByVal REcordID As String, ByVal DBPath As String, ByVal DBTable As String, ByVal Fieldnames As String, ByVal FieldValues As String, _
    Optional ByVal UpdateCriteria As String = "", Optional ByVal ExcludeFields As String = "", Optional ByRef ErrMessages As String, _
    Optional EncaseFields As Boolean = False, Optional FieldValueDelim As String = ";") As Boolean

    Dim cn As Object
    Dim rs As Object
    Dim cmd As Object
    Dim myConn As String
    Dim provider As String
    Dim con As OLEDBConnection
    'Dim cmd As oledbcommand
    'Dim da As oledbdataadapter
    Dim strSQL As String
    Dim strDeliveryRef As String
    Dim dtDeliveryDate As Date
    Dim strDeliveryDate As String
    Dim ExcludedFields As String
    Dim FieldTypeArr() As String
    Dim FieldType As Integer
    Dim ParamInfo As String
    Dim ParamLength As Long
    Dim Dic_FieldTypes As Object
    Dim Dic_ParamInfo As Object
    Dim ParamInfoArr() As String
    Dim FieldIDX As Integer
    Dim ParamArr() As Object
    Dim Param As Object
    Dim ParamKey As Variant
    Dim ParamName As String
    Dim ParamValue As Variant
    Dim FieldLength As Integer
    Dim FieldValue As String
    
    InsertUpdateRecords_Using_Parameters = False
    
    On Error GoTo Err_InsertUpdateRecords_Using_Parameters
    
    'myConn = "G:\MIS\Goodsin Timesheet Data\GoodsInTimesheetRecords.accdb"
    If Len(DBPath) > 0 Then
        myConn = DBPath
    Else
        myConn = ""
        ErrMessages = "DATABASE PATH NOT SPECIFIED"
        Exit Function
    End If
    
    If Len(DBTable) > 0 Then
            'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
            'strSQL = strSQL & " VALUES (""" & Fieldvalues & """);"
    Else
        MsgBox ("DATABASE TABLE NOT SPECIFIED")
        cn.Close
        Set cn = Nothing
        Exit Function
    End If
    
    'Set cn = ADODB.Connection
    Set cmd = New ADODB.Command
    Set cn = New ADODB.Connection
    Set Dic_ParamInfo = CreateObject("Scripting.Dictionary")
    Dic_ParamInfo.RemoveAll
    Application.ScreenUpdating = False
    Set Dic_FieldTypes = CreateObject("Scripting.Dictionary")
    Dic_FieldTypes.RemoveAll
        'Set TargetRange = ThisWorkbook.Sheets("Prefs").Range("A10")
    
        'Set cn = CreateObject("ADODB.Connection")
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & myConn & ";"
        cn.CursorLocation = 3 'Client Side
        cmd.ActiveConnection = cn
        
    ExcludedFields = ExcludeFields
    
    'NEED to use SearchAccessDB() as boolean
    'Set Dic = CreateObject("Scripting.Dictionary")
    ReDim FieldTypeArr(1)
    ReDim ParamArr(1)
    FieldTypeArr = Get_FieldTypes(DBPath, DBTable, Dic_FieldTypes)
    'GETTING AUTHENTICATION ERROR when the above function is run.
    
    If UpdateRecords = True Then
        strSQL = PrepareUpdate_With_Parameters(REcordID, DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, UpdateCriteria, ExcludedFields, False, ",", FieldValueDelim)
    Else
        strSQL = PrepareInsert_With_Parameters(REcordID, DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, ExcludedFields, False, ",", FieldValueDelim)
    End If
    'strSQL = "UPDATE " & DBTable & " SET [DeliveryDate] = ?"
        'strDeliveryRef = "TEST-DAN-YES"
        'Fieldnames = "DeliveryDate,DeliveryReference"
        'dtDeliveryDate = CDate("21/05/2018")
        'strDeliveryDate = Format(dtDeliveryDate, "dd/MM/yyyy")
        'FieldValues = Chr(34) & strDeliveryDate & Chr(34) & "," & Chr(34) & strDeliveryRef & Chr(34)
        'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
        'strSQL = strSQL & " VALUES (" & FieldValues & ")"
        'MsgBox ("SQL= " & strSQL)
    'cmd.Parameters.Clear
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    'cmd.CreatePArameter(ParamName,FieldType,Direction,Size,Value)
    'where direction = 1) adParamInput, 2) adParamOutput, 3) adParamInputOutput, 4) adParamReturnValue
    'where size = Field Size - now i think use 0 for all types except the VARCHAR / Short Text
    FieldIDX = 1
    
    For Each ParamKey In Dic_ParamInfo
        'ParamKey = "@" & fieldname
        Set Param = Nothing
        If Len(ParamKey) = 0 Then
            GoTo NEXT_ParamKey
        End If
        ParamName = ParamKey
        ParamInfo = Dic_ParamInfo(ParamKey)
        ParamInfoArr = Split(ParamInfo, "_")
        ParamLength = Len(ParamInfoArr(2))
        If IsNumeric(ParamInfoArr(0)) Then
            FieldType = CInt(ParamInfoArr(0))
        Else
            'bugger
            FieldType = 0
        End If
        If IsNumeric(ParamInfoArr(1)) Then
            FieldLength = CLng(ParamInfoArr(1))
            If FieldType = 202 Or FieldType = 203 Then
                If FieldLength = 0 Then
                    FieldLength = 1
                End If
            End If
        Else
            'bugger
            FieldLength = Len(ParamValue)
        End If
        If Len(ParamInfoArr(2)) > 0 And FieldType > 0 Then
            'get fieldvalue
            If IsDate(ParamInfoArr(2)) Then
                ParamValue = CDate(ParamInfoArr(2))
            ElseIf IsNumeric(ParamInfoArr(2)) Then 'INTEGER / LONG / DOUBLE / DECIMAL / CURRENCY, a fieldtype of 3 or 5 .
                'ALSO Delivery Ref ends up here - as it looks like a numeric value !
                ParamValue = ParamInfoArr(2)
            Else 'TEXT / STRING value:
            
                ParamValue = ParamInfoArr(2)
            End If
            
            Set Param = cmd.CreateParameter(ParamName, FieldType, adParamInput, FieldLength, ParamValue)
            cmd.Parameters.Append Param
        Else
            'Empty Value
            
            'ParamValue = " " - was causing error in Numeric Fields.
        End If
        'TREAT DATES as STRINGS before Insertion:
        
        FieldIDX = FieldIDX + 1
NEXT_ParamKey:
    Next
    
    cmd.Execute
    cn.Close
    Set cmd = Nothing
    Set cn = Nothing
    Set Dic_FieldTypes = Nothing
    Set Dic_ParamInfo = Nothing
    'Set ParamArr = Nothing
    InsertUpdateRecords_Using_Parameters = True
    
    Exit Function
Err_InsertUpdateRecords_Using_Parameters:
    
    Call Error_Report("InsertUpdateRecords_Using_Parameters()")
         'ErrMessages = "Error in InsertUpdateRecords"



End Function

Function PrepareParameters(ByVal REcordID As String, ByRef InsertSQL As String, ByRef UpdateSQL As String, ByRef Dic_ParameterInfo As Object, _
    ByVal DBTable As String, ByRef ParamFieldnames As String, ByRef ParamValues As Variant, _
    ByVal Dic_FieldsAndValues As Object, Optional Encase_Fields As Boolean = False, Optional UpdateCriteria As String = "") As Long
    'Dim FieldArr() As String
    Dim ValueArr As Variant
    Dim FieldName As String
    Dim fldValue As Variant
    Dim ParamKey As Variant
    Dim ParamItem As Variant
    Dim Dic_FieldTypes As Object
    Dim FieldType As Integer
    Dim FieldTypesArr() As String
    Dim varkey As Variant
    Dim fldLength As Long
    Dim ParamFldLength As Long
    Dim DontSave As Boolean
    Dim IncludeComma As Boolean
    Dim IncludeQuotes As Boolean
    Dim IncludeSpaces As Boolean
    Dim DeliveryComments As String
    Dim FieldCount As Long
    
    'THIS FUNCTION ONLY SAVES 1 RECORD.
    
    PrepareParameters = 0
    If Dic_FieldsAndValues.Count = 0 Then
        Exit Function
    End If
    DontSave = False
    'ReDim FieldArr(1)
    ReDim FieldTypesArr(1)
    fldLength = 0
    FieldType = 0
    FieldCount = 0
    InsertSQL = ""
    UpdateSQL = "UPDATE " & DBTable & " SET "
    DeliveryComments = "ISSUES: "
    ParamFldLength = 0
    ParamFieldnames = ""
    ParamValues = ""
    
    
    Set Dic_FieldTypes = New Scripting.Dictionary
    Set Dic_ParameterInfo = New Scripting.Dictionary
    FieldTypesArr = Get_FieldTypes(AccessDBpath, DBTable, Dic_FieldTypes)
    For Each varkey In Dic_FieldsAndValues
        FieldCount = FieldCount + 1
        If IsEmpty(varkey) Or Len(varkey) = 0 Then
            GoTo NEXT_ITERATION_IN_PREPAREPARAMETERS
        End If
        'VarKey = UCase(VarKey)
        
        FieldName = UCase(varkey)
        fldValue = UCase(Dic_FieldsAndValues(varkey))
        FieldType = UCase(Dic_FieldTypes(FieldName))
        DontSave = False 'for EACH field.
        fldLength = Len(fldValue)
        ParamFldLength = 0
        If FieldType = 202 Then
            'SHORT TEXT
            ParamFldLength = Len(fldValue)
            If fldLength > 0 Then
                If Encase_Fields Then
                    fldValue = Chr(34) & fldValue & Chr(34)
                End If
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = False
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
            Else
                'BLANK / EMPTY field value:
                'dontsave = True ??????????????????????????????? what if user wants to blank out a field ??
                fldValue = " "
                ParamFldLength = 1
                DeliveryComments = DeliveryComments & " Field: " & FieldName & " =EMPTY,"
            End If
        ElseIf FieldType = 203 Then
            'LONG TEXT
            ParamFldLength = Len(fldValue)
            If fldLength > 0 Then
                ParamFldLength = Len(fldValue)
                If Encase_Fields Then
                    fldValue = Chr(34) & fldValue & Chr(34)
                End If
                IncludeComma = True
                IncludeQuotes = True
                IncludeSpaces = False
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
            Else
                'BLANK / EMPTY field value:
                'dontsave = True ??????????????????????????????? what if user wants to blank out a field ??
                fldValue = " "
                ParamFldLength = 1
                DeliveryComments = DeliveryComments & " Field: " & FieldName & " =EMPTY,"
            End If
        ElseIf FieldType = 3 Then
            'INTEGER or LONG - NUMERIC.
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength > 0 Then
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                If IsNumeric(fldValue) Then
                    'OK SAVE PARAMETER
                Else
                    'TEXT appears in this numeric field:
                    DeliveryComments = DeliveryComments & " Field: " & FieldName & ": " & CStr(fldValue) & ","
                    fldValue = 0
                    
                End If
            Else
                DeliveryComments = DeliveryComments & " Field: " & FieldName & ": EMPTY,"
                fldValue = 0
            End If
        ElseIf FieldType = 5 Then
            'DOUBLE - will have PRECISION and NUMBER of DECIMAL PLACES
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength > 0 Then
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                If IsNumeric(fldValue) Then
                    'OK SAVE PARAMETER
                Else
                    'TEXT appears in this numeric field:
                    DeliveryComments = DeliveryComments & " Field: " & FieldName & ": " & CStr(fldValue) & ","
                    fldValue = 0
                    
                End If
            Else
                DeliveryComments = DeliveryComments & " Field: " & FieldName & ": EMPTY,"
                fldValue = 0 'Could put 0.00f instead ????
            End If
            
            
        ElseIf FieldType = 6 Then
            'CURRENCY
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength > 0 Then
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                If IsNumeric(fldValue) Then
                    'OK SAVE PARAMETER
                Else
                    'TEXT appears in this numeric field:
                    DeliveryComments = DeliveryComments & " Field: " & FieldName & ": " & CStr(fldValue) & ","
                    fldValue = 0
                    
                End If
            Else
                DeliveryComments = DeliveryComments & " Field: " & FieldName & ": EMPTY,"
                fldValue = 0 'Could put 0.00f instead ????
            End If
            
        ElseIf FieldType = 7 Then
            'DATEs
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength < 6 Then
                fldValue = "1/1/1970"
                DeliveryComments = DeliveryComments & " Field:" & FieldName & " Invalid Date,"
            Else
                fldValue = ConvertBadChars(CStr(fldValue), "", IncludeComma, IncludeQuotes, IncludeSpaces)
                fldValue = Convert_strDateToDate(CStr(fldValue), False)
            End If
        ElseIf FieldType = 0 Then
            'Field Type is Unrecognised or some error here ? = 0
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength < 2 Then
                DontSave = True
            Else
                DontSave = True
            End If
            DeliveryComments = DeliveryComments & " Field:" & FieldName & ": Type0= " & fldValue
        Else
            'UNKNOWN FIELD TYPE:
            ParamFldLength = 0
            IncludeComma = True
            IncludeQuotes = True
            IncludeSpaces = True
            If fldLength < 2 Then
                DontSave = True
            Else
                DontSave = True
            End If
            DeliveryComments = DeliveryComments & " Field:" & FieldName & ": Unknown Type= " & fldValue
            
        End If
        If DontSave = False Then
            'Create Parameter for this field:
            ParamKey = "@" & FieldName
            If FieldCount < 2 Then
                ParamFieldnames = FieldName
                ParamValues = ParamKey
                UpdateSQL = UpdateSQL & "[" & FieldName & "]" & " = " & ParamKey
            Else
                ParamFieldnames = ParamFieldnames & "," & FieldName
                ParamValues = ParamValues & "," & ParamKey
                UpdateSQL = UpdateSQL & ",[" & FieldName & "]" & " = " & ParamKey
            End If
            ParamItem = CStr(FieldType) & "_" & CStr(ParamFldLength) & "_" & fldValue & "_" & DeliveryComments
            If Not Dic_ParameterInfo.Exists(ParamKey) Then
                Dic_ParameterInfo(ParamKey) = ParamItem
            End If
            
        Else
            'DONT CREATE PARAMETER - INVALID value:
            'BUT still save / update the DeliveryComments:
            DeliveryComments = DeliveryComments & " Field:" & FieldName & ": Not Saved= " & fldValue
        End If
NEXT_ITERATION_IN_PREPAREPARAMETERS:
    Next
    If Len(REcordID) > 0 Then
        If CLng(REcordID) > 0 Then
            UpdateSQL = UpdateSQL & " WHERE ID = " & CLng(REcordID)
        End If
    Else
        If Len(UpdateCriteria) > 0 Then
            UpdateSQL = UpdateSQL & " WHERE " & UpdateCriteria
        End If
    End If
    InsertSQL = "INSERT INTO " & DBTable & " (" & ParamFieldnames & ")" & " VALUES " & "(" & ParamValues & ")"
    PrepareParameters = Dic_ParameterInfo.Count
    

End Function

Function PrepareUpdate_With_Parameters(ByRef REcordID As String, ByVal DBPath As String, ByVal DBTable As String, ByRef Fieldnames As String, ByVal FieldValues As String, _
        ByRef Dic_ParameterInfo As Object, ByVal UpdateCriteria As String, Optional ByRef ExcludeFields As String = "", Optional Encase_Fields As Boolean = False, _
        Optional FieldDelim As String = ",", Optional ValueDelim As String = ";") As String
    'PrepareUpdate_With_Parameters(DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, UpdateCriteria, ExcludedFields, False, FieldValueDelim, FieldValueDelim)
    Dim FieldNameArray() As String
    Dim IgnoreFieldsArray() As String
    Dim ValueArray() As String
    Dim FinalCMD As String
    Dim IDX As Integer
    Dim fldName As String
    Dim fldValue As String
    Dim NumFields As Integer
    Dim UpdateCmd As String
    Dim IncludeComma As Boolean
    Dim IncludeSpeechMarks As Boolean
    Dim Operator As String
    Dim ParamKey As Variant
    Dim ParamItem As Variant
    Dim Dic_FieldTypes As Object
    Dim FieldType As Integer
    Dim FieldTypesArr() As String
    Dim varkey As Variant
    Dim Dic_FieldsAndValues As Object
    Dim InsertSQL As String
    Dim UpdateSQL As String
    Dim ParamFieldnames As String
    Dim ParamValues As String
    Dim TotalParameters As Long
    
    FinalCMD = ""
    ReDim FieldNameArray(1)
    ReDim IgnoreFieldsArray(1)
    ReDim ValueArray(1)
    
    On Error GoTo Err_PrepareUpdate_With_Parameters
    
    Set Dic_ParameterInfo = CreateObject("Scripting.Dictionary")
    Dic_ParameterInfo.RemoveAll
    Dic_ParameterInfo.CompareMode = TextCompare
    Set Dic_FieldTypes = CreateObject("Scripting.Dictionary")
    Dic_FieldTypes.RemoveAll
    Dic_FieldTypes.CompareMode = TextCompare
    Set Dic_FieldsAndValues = CreateObject("Scripting.Dictionary")
    Dic_FieldsAndValues.RemoveAll
    Dic_FieldsAndValues.CompareMode = TextCompare
    FieldTypesArr = Get_FieldTypes(AccessDBpath, DBTable, Dic_FieldTypes)
    
    PrepareUpdate_With_Parameters = ""
    If Len(ExcludeFields) > 0 Then
        IgnoreFieldsArray = strToStringArray(ExcludeFields, ",")
    End If
    If Len(DBTable) = 0 Then
        MsgBox ("Error in PrepareUpdate_With_Parameters: No Database Table specified")
        PrepareUpdate_With_Parameters = ""
        Exit Function
    End If
    If Len(Fieldnames) = 0 Then
        'NumFields = GetNumFields(connString, "SELECT * FROM " & TableName, DBName, Fieldnames)
        Fieldnames = GetFieldnames_From_ACCESS(DBTable, AccessDBpath, UpdateCriteria, "")
    End If
    
    FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 1, False, True, False, "_", False)
    If Len(FieldValues) > 0 Then
        IncludeComma = True
        IncludeSpeechMarks = True
        ValueArray = strToStringArray(FieldValues, ValueDelim, 1, False, True, IncludeComma, "_", IncludeSpeechMarks)
    Else
        MsgBox ("Error in PrepareUpdate_With_Parameters: No values specified")
        PrepareUpdate_With_Parameters = ""
        Exit Function
    End If
    'BUT check that the values are removed too if they corresponded with those fields removed ????
    Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray, 1, 1) 'rebuilds whole list without the extracted fields
    
    If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
        If UBound(FieldNameArray) < UBound(ValueArray) Then
            MsgBox ("Error in PrepareUpdate_With_Parameters: Number of Fields Passed are LESS than Number of VALUES passed.")
            PrepareUpdate_With_Parameters = ""
            Exit Function
        End If
        If UBound(FieldNameArray) > UBound(ValueArray) Then
            MsgBox ("Error in PrepareUpdate_With_Parameters: Number of Fields Passed are GREATER than Number of VALUES passed.")
            PrepareUpdate_With_Parameters = ""
            Exit Function
        End If
    End If
    'Prepare and insert parameter INFO for the script dictionary
    
    
    'IDX = 1
    'UpdateCmd = "UPDATE " & DBTable & " SET "
    
    Dic_FieldsAndValues.RemoveAll
    For IDX = 1 To UBound(FieldNameArray)
        fldName = UCase(FieldNameArray(IDX))
        
        If Not Dic_FieldsAndValues.Exists(fldName) Then
            Dic_FieldsAndValues(fldName) = ValueArray(IDX)
        End If
    Next
    'Prepare and insert parameter INFO for the script dictionary
    
    TotalParameters = PrepareParameters(REcordID, InsertSQL, UpdateSQL, Dic_ParameterInfo, DBTable, ParamFieldnames, ParamValues, Dic_FieldsAndValues, False)
    
    FinalCMD = UpdateSQL
    
    PrepareUpdate_With_Parameters = FinalCMD
Exit Function

Err_PrepareUpdate_With_Parameters:

    Call Error_Report("Error in PrepareUpdate_With_Parameters()")

End Function

Function PrepareInsert_With_Parameters(ByVal REcordID As String, ByVal DBPath As String, ByVal DBTable As String, ByRef Fieldnames As String, ByVal FieldValues As String, _
        ByRef Dic_ParameterInfo As Object, _
        Optional ByRef ExcludeFields As String = "", Optional Encase_Fields As Boolean = False, Optional FieldDelim As String = ",", _
        Optional ValueDelim As String = ";") As String
    'PrepareInsert_With_Parameters(DBPath, DBTable, Fieldnames, FieldValues, Dic_ParamInfo, ExcludedFields, False, FieldValueDelim, FieldValueDelim)
    Dim FieldNameArray() As String
    Dim IgnoreFieldsArray() As String
    Dim ValueArray() As String
    Dim FinalCMD As String
    Dim IDX As Integer
    Dim fldName As String
    Dim fldValue As Variant
    Dim fldValues As Variant
    Dim NumFields As Integer
    Dim UpdateCmd As String
    Dim IncludeComma As Boolean
    Dim IncludeSpeechMarks As Boolean
    Dim Operator As String
    Dim ParamKey As Variant
    Dim ParamItem As Variant
    Dim Dic_FieldTypes As Object
    Dim FieldType As Integer
    Dim FieldTypesArr() As String
    Dim varkey As Variant
    Dim fldLength As Long
    Dim TotalParameters As Long
    Dim InsertSQL As String
    Dim UpdateSQL As String
    Dim ParamFieldnames As String
    Dim ParamValues As String
    Dim Dic_FieldsAndValues As Object
    
    
    FinalCMD = ""
    fldValues = ""
    ReDim FieldNameArray(1)
    ReDim IgnoreFieldsArray(1)
    ReDim ValueArray(1)
    
    On Error GoTo Err_PrepareInsert_With_Parameters
    
    Set Dic_ParameterInfo = CreateObject("Scripting.Dictionary")
    Dic_ParameterInfo.CompareMode = TextCompare
    Set Dic_FieldsAndValues = New Scripting.Dictionary
    Dic_FieldsAndValues.RemoveAll
    Dic_FieldsAndValues.CompareMode = TextCompare
    'Dic_ParameterInfo.RemoveAll
    Set Dic_FieldTypes = CreateObject("Scripting.Dictionary")
    Dic_FieldTypes.RemoveAll
    Dic_FieldTypes.CompareMode = TextCompare
    FieldTypesArr = Get_FieldTypes(AccessDBpath, DBTable, Dic_FieldTypes)
    
    PrepareInsert_With_Parameters = ""
    
    If Len(ExcludeFields) > 0 Then
        IgnoreFieldsArray = strToStringArray(ExcludeFields, FieldDelim, 0, False, True, False, "_", False)
    End If
    If Len(DBTable) = 0 Then
        MsgBox ("Error in PrepareInsert_With_Parameters: No Database Table specified")
        PrepareInsert_With_Parameters = ""
        Exit Function
    End If
    
    If Len(Fieldnames) = 0 Then
        Fieldnames = GetFieldnames_From_ACCESS(DBTable, AccessDBpath, "", "")
    End If
    'SO what if Dic_ParameterInfo.count > 0 then ????????????
    'This means that its been passed back in - after evaluation from the parameters section - before the EXECUTE SQL.
    ' - means that number of fields do not match number of parameter @ values passed.
    ' so we need to resolve this. maybe instead of @P1 .... we need to use the actual fieldname - @DeliveryDate etc.
    ' This way we can remove the @ symbol and know which fields have been passed and which ARE MISSING !
    If Dic_ParameterInfo.Count > 0 Then
        'OK so we can get all of the fields from the table to start with.
        'Then we need to eliminate the fields that are NOT in the Dictionary:
        ' - this will form the fields that are left to put into the final SQL query.
        
    End If
    '
    'FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True) 'EACH fieldname passed must have square brackets around it.
    FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 1, False, True, False, "_", False)
    If Len(FieldValues) > 0 Then
        IncludeComma = True
        IncludeSpeechMarks = True
        ValueArray = strToStringArray(FieldValues, ValueDelim, 1, False, True, IncludeComma, "", IncludeSpeechMarks)
    Else
        MsgBox ("Error in PrepareInsert_With_Parameters: No values specified")
        PrepareInsert_With_Parameters = ""
        Exit Function
    End If
    Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, FieldDelim, FieldNameArray, 1, 1)
    If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
        If UBound(FieldNameArray) < UBound(ValueArray) Then
            MsgBox ("Error in PrepareInsert_With_Parameters: Number of Fields Passed are LESS than Number of VALUES passed.")
            PrepareInsert_With_Parameters = ""
            Exit Function
        End If
        If UBound(FieldNameArray) > UBound(ValueArray) Then
            MsgBox ("Error in PrepareInsert_With_Parameters: Number of Fields Passed are GREATER than Number of VALUES passed.")
            PrepareInsert_With_Parameters = ""
            Exit Function
        End If
    End If
    Dic_FieldsAndValues.RemoveAll
    For IDX = 1 To UBound(FieldNameArray)
        fldName = UCase(FieldNameArray(IDX))
        If Not Dic_FieldsAndValues.Exists(fldName) Then
            Dic_FieldsAndValues(fldName) = ValueArray(IDX)
        End If
    Next
    'Prepare and insert parameter INFO for the script dictionary
    
    TotalParameters = PrepareParameters(REcordID, InsertSQL, UpdateSQL, Dic_ParameterInfo, DBTable, ParamFieldnames, ParamValues, Dic_FieldsAndValues, False)
    
    FinalCMD = InsertSQL
    PrepareInsert_With_Parameters = FinalCMD
    Exit Function
    
Err_PrepareInsert_With_Parameters:
    
    Call Error_Report("Error in PrepareInsert_With_Parameters()")



End Function

Function Get_FieldTypes(DBPath As String, DBTable As String, ByRef Dic_FieldTypes As Scripting.Dictionary) As String()
    Dim cn As Object
    Dim ADOSET As Object
    Dim objAccess As Object
    Dim myConn As String
    Dim provider As String
    Dim IDX As Integer
    Dim FieldType As Integer
    Dim FieldTypesArr() As String
    Dim i As Integer
    Dim strSQL As String
    Dim FieldName As String
    'Dim dict As Scripting.Dictionary
    
    On Error GoTo Get_FieldTypes
    
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBPath & ";"
    
    Set Dic_FieldTypes = CreateObject("Scripting.Dictionary")
    Dic_FieldTypes.RemoveAll
    Dic_FieldTypes.CompareMode = TextCompare
    
    If Len(DBTable) > 0 Then
        'strSQL = "INSERT INTO " & DBTable & " (" & Fieldnames & ") "
        'strSQL = strSQL & " VALUES (""" & Fieldvalues & """);"
    Else
        MsgBox ("DATABASE TABLE NOT SPECIFIED")
        cn.Close
        Set cn = Nothing
        Exit Function
    End If
    
    'Set objAccess = CreateObject("Access.Application")
    strSQL = "SELECT * FROM " & DBTable
    Set ADOSET = New ADODB.Recordset
    Set ADOSET = CreateObject("ADODB.Recordset")
    ADOSET.Open strSQL, cn, adOpenStatic, adLockOptimistic, adCmdText
    'Call objAccess.OpenCurrentDatabase(DBPath)
        'get tables data
    'Set ADOSET = objAccess.CurrentProject.Connection.Execute(DBTable)
    ReDim FieldTypesArr(ADOSET.Fields.Count)
    For i = 0 To ADOSET.Fields.Count - 1
        FieldType = ADOSET.Fields.Item(i).Type
        FieldName = UCase(ADOSET.Fields.Item(i).Name)
        If Not Dic_FieldTypes.Exists(FieldName) Then
            Dic_FieldTypes.Add FieldName, FieldType
        End If
        FieldTypesArr(i) = CStr(FieldType)
    Next i

    ADOSET.Close
    Set ADOSET = Nothing
    cn.Close
    Set cn = Nothing
    
    'objRecordset.ActiveConnection = CurrentProject.Connection
    'objRecordset.Open ("MyTable1")
    
    Get_FieldTypes = FieldTypesArr
    
Exit Function

Get_FieldTypes:
    Call Error_Report("Get_FieldTypes()")
    
End Function


Sub Load_From_Access_Test()
    Dim cn As Object, rs As Object
    Dim intColIndex As Integer
    Dim DBFullName As String
    Dim TargetRange As Range

    DBFullName = "G:\MIS\Goodsin Timesheet Data\GoodsInTimesheetRecords.accdb"

    'On Error GoTo Whoa
    
    Set cn = New ADODB.Connection
    
    Application.ScreenUpdating = False

    Set TargetRange = ThisWorkbook.Sheets("Prefs").Range("A10")

    Set cn = CreateObject("ADODB.Connection")
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBFullName & ";"
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBFullName & ";"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [tblShortsAndExtraParts] WHERE [PartNo] = 'Part02'", cn, , , adCmdText
    
    
    
    ' Write the field names
    For intColIndex = 0 To rs.Fields.Count - 1
        TargetRange.Offset(1, intColIndex).value = rs.Fields(intColIndex).Name
    Next

    ' Write recordset
    TargetRange.Offset(1, 0).CopyFromRecordset rs
    
    Application.ScreenUpdating = True
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    On Error GoTo 0
    Exit Sub

End Sub


Function SearchAccessDB(DBFullPath As String, DBTable As String, SearchField As String, SearchValue As Variant, SearchType As String, _
    Optional Operation As String = "=", Optional ByRef FoundID As Variant, Optional SearchField2 As String = "", Optional SearchValue2 As Variant, _
    Optional SearchType2 As String = "", Optional Operation2 As String = "=", Optional SortField As String = "", _
    Optional Reversed As Boolean = False, Optional ByVal LinkID As String = "", Optional ReturnField As String = "", Optional ReturnValue As Variant) As Boolean
    Dim cn As Object, rs As Object
    Dim intColIndex As Integer
    Dim DBFullName As String
    Dim TargetRange As Range
    Dim SQLqry As String
    Dim NewSearchValue As String
    Dim NewSearchValue2 As String
    'Dim TheID As String
    
    SearchAccessDB = False
    'DBFullName = "G:\MIS\Goodsin Timesheet Data\GoodsInTimesheetRecords.accdb"
    DBFullName = DBFullPath
    'On Error GoTo Whoa
    
    Set cn = New ADODB.Connection
    Application.ScreenUpdating = False

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & DBFullName & ";"
    Operation = " " & Operation & " "
    Operation2 = " " & Operation2 & " "
    Set rs = New ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    If UCase(SearchType) = "STRING" Then
        'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " " & Operation & " '" & SearchText & "')"
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & "'" & SearchValue & "')"
        
    ElseIf UCase(SearchType) = "INTEGER" Then
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & SearchValue & ")"
        
    ElseIf UCase(SearchType) = "DATE" Then
        
        NewSearchValue = Format(CDate(SearchValue), "dd/mm/yyyy")
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & "#" & NewSearchValue & "#" & ")"
    ElseIf UCase(SearchType) = "DATETIME-reverse" Then
        NewSearchValue = Format(CDate(SearchValue), "yyyy/mm/dd hh:nn:ss")
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & "#" & NewSearchValue & "#" & ")"
    ElseIf UCase(SearchType) = "DATETIME" Then
        NewSearchValue = Format(CDate(SearchValue), "dd/mm/yyyy hh:nn:ss")
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & "#" & NewSearchValue & "#" & ")"
    ElseIf UCase(SearchType) = "DATETIME-hyphon" Then
        NewSearchValue = Format(CDate(SearchValue), "yyyy-mm-dd hh:nn:ss")
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & "#" & NewSearchValue & "#" & ")"
    Else 'double type
        SQLqry = "SELECT * FROM " & DBTable & " WHERE ([" & SearchField & "] " & Operation & SearchValue & "f" & ")"
    End If
    
    If UCase(SearchType2) = "STRING" Then
        'QryStr = "SELECT * FROM " & DBTable & " WHERE (" & SearchField & " " & Operation & " '" & SearchText & "')"
        If Len(SearchField2) > 0 Then
            SQLqry = SQLqry & " AND ([" & SearchField2 & "] " & Operation2 & "'" & SearchValue2 & "')"
        End If
        
    ElseIf UCase(SearchType2) = "INTEGER" Then
        
        If Len(SearchField2) > 0 Then
            SQLqry = SQLqry & " AND ([" & SearchField2 & "] " & Operation2 & SearchValue2 & ")"
        End If
    ElseIf UCase(SearchType) = "DATE" Then
        If Len(SearchField2) > 0 Then
            NewSearchValue2 = Format(CDate(SearchValue2), "dd/mm/yyyy")
            SQLqry = " AND ([" & SearchField2 & "] " & Operation2 & "#" & NewSearchValue2 & "#" & ")"
        End If
    ElseIf UCase(SearchType) = "DATETIME" Then
        If Len(SearchField2) > 0 Then
            NewSearchValue2 = Format(CDate(SearchValue2), "dd/mm/yyyy hh:nn:ss")
            SQLqry = " AND ([" & SearchField2 & "] " & Operation2 & "#" & NewSearchValue2 & "#" & ")"
        End If
    Else
        If Len(SearchField2) > 0 Then
            SQLqry = " AND ([" & SearchField2 & "] " & Operation2 & SearchValue2 & "f" & ")"
        End If
    End If
    
    If Len(SortField) > 0 Then
        SQLqry = SQLqry & " ORDER BY " & SortField
        If Reversed Then
            SQLqry = SQLqry & " DESC"
        Else
            SQLqry = SQLqry & " ASC"
        End If
    End If
    'SQLqry = SQLqry & ";"
    'The following gives an error - No Value given for one or more required parameters:
    
    'rs.Open SQLqry, cn, adOpenDynamic, adLockOptimistic, adCmdText
    rs.Open SQLqry, cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        SearchAccessDB = True
        FoundID = rs.Fields("ID").value 'This works !
        If Len(ReturnField) > 0 Then
            ReturnValue = rs.Fields(ReturnField).value
        End If
        'FoundID = CStr(rs!["ID"]) 'returns FIRST record it finds - which is what we want - because records sorted by Delivery Date in reverse so latest first.
        
    End If
    ' TEST if rs contains anything
    'rs.Fields(intColIndex).Name
    
    ' Write recordset
    
    
    Application.ScreenUpdating = True
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    On Error GoTo 0
End Function

Function LoadAccessDBTable(ByVal DBTable As String, AccessDBpath As String, _
    Optional JustTAG As Boolean = False, Optional Criteria As String = "", Optional SortFields As String = "", Optional ReversedSort As Boolean = False, _
    Optional AddQuotesAroundFieldValues As Boolean = False, Optional ByRef Fieldnames As String = "", _
    Optional ByRef FieldValues As String = "", Optional ByRef dict_ReturnValues As Scripting.Dictionary = Nothing, _
    Optional ByRef AllrecsArr As Variant = Nothing, Optional LowerTag As Long = 0, Optional UpperTag As Long = 0, _
    Optional StartTAG As Long = 0, Optional AddFieldToValue As Boolean = False) As Boolean
'Using ADO to Import data from an Access Database Table to an Excel worksheet

    Dim strMyPath As String
    Dim strDBName As String
    Dim strDB As String
    Dim strSQL As String
    Dim IDX As Long
    Dim FieldIDX As Long
    Dim TAGNumber As Long
    Dim n As Long
    Dim TotalFields As Long
    Dim TotalRECS As Long
    Dim rng As Range
    Dim ADOSET As Object
    Dim connDB As Object
    Dim AllFields() As String
    Dim FieldValue As String
    Dim FieldName As String
    Dim PrimaryKeyValue As String
    Dim Entry As String
    Dim FieldType As String
    Dim varkey As Variant
    Dim FinalEntry As Variant
    Dim ExtractDate As String
    Dim FirstHashPos As Integer
    Dim LastHashPos As Integer
    Dim NewDate As Date
    Dim NewSearchDate As String
    
    '--------------
    'THE CONNECTION OBJECT - check if default DB exists first:
    
    On Error GoTo Err_LoadAccessDBTable
    
    LoadAccessDBTable = False
    
    If Len(AccessDBpath) > 0 Then
        strDB = AccessDBpath
    Else
        MsgBox ("No Path to database supplied")
        LoadAccessDBTable = False
        Exit Function
    End If
    
    If Len(DBTable) = 0 Then
        'MsgBox ("Please specify DB Table to load")
        Application.StatusBar = "No DB Table Specified"
        'usrHeatMapControlPanel.txtOutput.Text = "No DB Table Specified"
        Exit Function
    End If
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    If Len(strDB) > 0 Then
        Set connDB = New ADODB.Connection
        connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    Else
        Exit Function
    End If
    
    'if DATE type then use: NewSearchValue2 = Format(CDate(SearchValue2), "dd/mm/yyyy"), enclosed with ##
    FirstHashPos = InStr(Criteria, "#")
    LastHashPos = InStr(FirstHashPos + 1, Criteria, "#")
    'ExtractDate = Mid(Criteria, FirstHashPos + 1, (LastHashPos - FirstHashPos) - 1)
    'NewDate = CDate(ExtractDate)
    
    'Set dict_ReturnValues = New Scripting.Dictionary
    Set dict_ReturnValues = CreateObject("Scripting.Dictionary")
    dict_ReturnValues.RemoveAll
    dict_ReturnValues.CompareMode = vbTextCompare
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    'Set the ADO Recordset object:
    Set ADOSET = New ADODB.Recordset
    strSQL = "SELECT * FROM " & DBTable
    If Len(Criteria) > 0 Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
    If Len(SortFields) > 0 Then
        strSQL = strSQL & " ORDER BY " & SortFields
        If ReversedSort Then
            strSQL = strSQL & " DESC"
        Else
            strSQL = strSQL & " ASC"
        End If
    End If
    
    '--------------
    dict_ReturnValues.CompareMode = TextCompare
    ADOSET.Open strSQL, connDB, adOpenStatic, adLockOptimistic, adCmdText
    
    TotalFields = ADOSET.Fields.Count
    TotalRECS = ADOSET.RecordCount
    If TotalRECS = 0 Then
        'MsgBox ("CANNOT FIND ENTRY")
        LoadAccessDBTable = False
        Exit Function
    End If
    ReDim AllrecsArr(TotalFields + 1, TotalRECS + 1)
    
    ADOSET.MoveFirst
    IDX = 0
    
    Do While Not ADOSET.EOF
        TAGNumber = StartTAG
        For FieldIDX = 0 To TotalFields - 1
            Entry = "0"
            FinalEntry = "0"
            'If Len(ADOSET.Fields(0).value) > 0 Then
            '    'TEST the field type for primary key - on the rare occaison that the ID / Primary Key is NOT the first field in the table.
            '    PrimaryKeyValue = ADOSET.Fields(0).value
            '    FieldNAme = ADOSET.Fields(0).Name
            '    Entry = ADOSET.Fields(0).value
            'End If
                'Fieldname = ADOSET.Fields(FieldIDX).Name & "_" & PrimaryKeyValue
            FieldName = ADOSET.Fields(FieldIDX).Name
            FieldType = ADOSET.Fields(FieldIDX).Type
            If FieldType = "3" Then 'NUMBER - INTEGER or LONG
                If Len(ADOSET.Fields(FieldIDX).value) > 0 Then
                    'check for primary key here:
                    If IsNumeric(ADOSET.Fields(FieldIDX).value) Then
                        Entry = ADOSET.Fields(FieldIDX).value
                    Else
                        Entry = "0"
                    End If
                Else
                    Entry = "0"
                End If
            End If
            If FieldType = "5" Then 'NUMBER - DOUBLE
                If Len(ADOSET.Fields(FieldIDX).value) > 0 Then
                    'check for primary key here:
                    If IsNumeric(ADOSET.Fields(FieldIDX).value) Then
                    
                        Entry = ADOSET.Fields(FieldIDX).value
                    Else
                        Entry = "0.00f"
                    End If
                Else
                    Entry = "0.00f"
                End If
            End If
            If FieldType = "202" Or FieldType = "203" Then 'STRING - SHORT AND LONG TEXT
                If Len(ADOSET.Fields(FieldIDX).value) > 0 Then
                    Entry = ADOSET.Fields(FieldIDX).value
                Else
                    Entry = " "
                End If
            End If
            If FieldType = "7" Then 'DATE
                If Len(ADOSET.Fields(FieldIDX).value) > 0 Then
                    Entry = ADOSET.Fields(FieldIDX).value
                Else
                    Entry = "1901/01/01"
                End If
            End If
            If AddQuotesAroundFieldValues Then
                Entry = Chr(34) & Entry & Chr(34)
            End If
            
            If JustTAG = True Then
                varkey = CStr(TAGNumber)
                If LowerTag > 0 And UpperTag > 0 Then
                    varkey = varkey & "_" & UCase(CStr(LowerTag)) & "_" & UCase(CStr(UpperTag))
                End If
            Else
                varkey = UCase(DBTable) & "_" & UCase(FieldName) & "_" & CStr(TAGNumber)
                If LowerTag > 0 And UpperTag > 0 Then
                    varkey = varkey & "_" & UCase(CStr(LowerTag)) & "_" & UCase(CStr(UpperTag))
                End If
            End If
            FinalEntry = Entry
            If AddFieldToValue Then
                FinalEntry = Entry & "_" & UCase(DBTable) & "_" & FieldName
            End If
            If dict_ReturnValues.Exists(varkey) Then
                If Len(dict_ReturnValues(varkey)) = 0 Then
                
                    dict_ReturnValues(varkey) = FinalEntry
                        
                End If
            Else
                'dict_ReturnValues.Add varKey, Entry
                dict_ReturnValues(varkey) = FinalEntry
            End If
            
            FieldValue = FinalEntry
            If FieldIDX = 0 Then
                'dict_ReturnValues.Item(FieldNAme) = PrimaryKeyValue
                FieldValues = FieldValue
                Fieldnames = FieldName
            Else
                FieldValues = FieldValues & "," & FieldValue
                'dict_ReturnValues.Item(FieldNAme) = FieldValue
                Fieldnames = Fieldnames & "," & FieldName
            End If
            
            AllrecsArr(FieldIDX, IDX) = FinalEntry
            'End If
            TAGNumber = TAGNumber + 1
            'FieldIDX = FieldIDX + 1
        Next
        IDX = IDX + 1 'RECORD INDEX.
        ADOSET.MoveNext
    Loop
    If Len(Fieldnames) > 0 Then
        LoadAccessDBTable = True
    End If
         'ADOSET.Open strSQL, connDB, adOpenForwardOnly, adLockOptimistic, adCmdText
    ADOSET.Close
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set ADOSET = Nothing
    Set connDB = Nothing
    
    Exit Function
    
Err_LoadAccessDBTable:
    
    Call Error_Report("Error in LoadAccessDBTable()")

End Function

Function GetFieldnames_From_ACCESS(ByVal DBTable As String, ByVal DBPathWithDBName As String, Optional Criteria As String = "", Optional DBName As String = "") As String
    Dim strMyPath As String
    Dim strDBName As String
    Dim strDB As String
    Dim strSQL As String
    Dim IDX As Long
    Dim FieldIDX As Long
    Dim n As Long
    Dim TotalFields As Long
    Dim rng As Range
    Dim ADOSET As Object
    Dim connDB As Object
    Dim AllFields() As String
    Dim FieldValue As String
    Dim FieldName As String
    Dim Fieldnames As String
    Dim PrimaryKeyValue As String
    Dim dict_ReturnValues As Scripting.Dictionary

    strDBName = DBName
    If Len(DBPathWithDBName) > 0 Then
        strMyPath = DBPathWithDBName
    Else
        strMyPath = ThisWorkbook.Path
    End If
    
    If Len(DBName) = 0 Then
        strDB = DBPathWithDBName
    Else
        strDB = strMyPath & "\" & strDBName
    End If
    If Len(DBTable) = 0 Then
        Application.StatusBar = "No DB Table Specified"
        Exit Function
    End If
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    If Len(strDB) > 0 Then
        Set connDB = New ADODB.Connection
        connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    Else
        Exit Function
    End If

    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    'Set the ADO Recordset object:
    Set ADOSET = New ADODB.Recordset
    strSQL = "SELECT * FROM " & DBTable
    If Len(Criteria) > 0 Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
        
    GetFieldnames_From_ACCESS = ""
    
    ADOSET.Open strSQL, connDB, adOpenStatic, adLockOptimistic, adCmdText

    'Set rng = ws.Range("A1")
    TotalFields = ADOSET.Fields.Count
    'ADOSET.MoveFirst
    IDX = 0
    'dict_ReturnValues.RemoveAll
    'Do While Not ADOSET.EOF
        FieldIDX = 1 'Exclude the ID
        While FieldIDX < TotalFields
            FieldName = ADOSET.Fields(FieldIDX).Name
            If FieldIDX = 1 Then
                'dict_ReturnValues.Item(FieldName) = FieldValue
                Fieldnames = FieldName
            Else
                'dict_ReturnValues.Item(FieldName) = FieldValue
                Fieldnames = Fieldnames & "," & FieldName
            End If
            FieldIDX = FieldIDX + 1
        Wend
        'ADOSET.MoveNext
    'Loop
    GetFieldnames_From_ACCESS = Fieldnames
    ADOSET.Close
    Set ADOSET = Nothing
    connDB.Close
    Set connDB = Nothing


End Function

Function RemoveExtractedFields(ByRef OldArray() As String, ByRef ExtractedArray() As String, ByVal substring As String, ByRef NewArray() As String, _
            Optional ByVal ElementsOldStartIDX As Integer = 0, Optional ByVal ElementsNewStartIDX As Integer = 0) As String
        Dim TotalExtracted As Integer
        Dim FieldIDX As Integer
        Dim ElementsOld As Integer
        Dim ElementsNew As Integer
        Dim commaPos As Integer
        Dim tempArr() As String
        Dim NewString As String
        Dim FinalString As String
        Dim EX As Integer
        
        RemoveExtractedFields = ""
        
        TotalExtracted = 0
        FinalString = ""
        commaPos = 1
        ElementsOld = ElementsOldStartIDX
        ElementsNew = ElementsNewStartIDX
        
            ReDim tempArr(1)
            For EX = 0 To UBound(ExtractedArray)
                NewString = ""
                For FieldIDX = ElementsOld To UBound(OldArray)
                    If UCase(ExtractedArray(EX)) = UCase(OldArray(FieldIDX)) Then
                        'no operation needed
                    Else
                        If Len(NewString) = 0 Then
                            NewString = OldArray(FieldIDX)
                        Else
                            NewString = NewString & substring & OldArray(FieldIDX)
                        End If
                    End If
                Next
                OldArray = strToStringArray(NewString, substring, ElementsNewStartIDX, False, True, False, "_", False)
                

            Next
            NewArray = OldArray
            FinalString = StringArrayToString(NewArray, substring, ElementsNewStartIDX, False, False, "_")
        
            'MsgBox("Exception Error in RemoveExtractedFields: " & exc.ToString, vbOK, "Exception Error in Asset Register")

        RemoveExtractedFields = FinalString
    End Function


Function PrepareInsert(ByVal DBName As String, ByVal DBTable As String, ByRef Fieldnames As String, ByVal FieldValues As String, _
        ByVal FieldTypes As Scripting.Dictionary, _
        Optional ByRef ExcludeFields As String = "", Optional Encase_Fields As Boolean = False, Optional FieldDelim As String = ",", _
        Optional ValueDelim As String = ";") As String
    Dim FieldNameArray() As String
    Dim IgnoreFieldsArray() As String
    Dim ValueArray() As String
    Dim FinalCMD As String
    Dim IDX As Integer
    Dim fldName As String
    Dim NumFields As Integer
    Dim NewFieldsArray() As String
    Dim IncludeComma As Boolean
        Dim IncludeSpeechMarks As Boolean
    Dim FieldType As Integer
    
    On Error GoTo Err_PrepareInsert
    
    FinalCMD = ""
    ReDim FieldNameArray(1)
    ReDim NewFieldsArray(1)
    ReDim IgnoreFieldsArray(1)
    On Error GoTo Err_PrepareInsert
        If Len(ExcludeFields) > 0 Then
            IgnoreFieldsArray = strToStringArray(ExcludeFields, ",", 0, False, True)
        End If
        If Len(DBTable) = 0 Then
            MsgBox ("Error in PrepareInsert: No Database Table specified")
            PrepareInsert = ""
            Exit Function
        End If
        If Len(Fieldnames) = 0 Then
            'NumFields = GetNumFields(connString, "SELECT * FROM " & DBTable, DBName, Fieldnames)
        End If
        FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True) 'EACH fieldname passed must have square brackets around it.
        If Len(FieldValues) > 0 Then
            IncludeComma = True
                        IncludeSpeechMarks = True
            ValueArray = strToStringArray(FieldValues, ValueDelim, 0, Encase_Fields, True, IncludeComma, "_", IncludeSpeechMarks)
        Else
            MsgBox ("Error in PrepareInsert: No values specified")
            PrepareInsert = ""
            Exit Function
        End If
        Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray)
        If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then

            If UBound(FieldNameArray) < UBound(ValueArray) Then
                'MsgBox("Error in PrepareInsert: Number of Fields passed are LESS THAN Number of Values Passed.", vbOK, "MISS-MATCH Error in Asset Register")
                'PrepareInsert = ""
                'Exit Function
            End If
            If UBound(FieldNameArray) > UBound(ValueArray) Then
                'MsgBox("Error in PrepareInsert: Number of Fields passed are GREATER THAN Number of Values Passed.", vbOK, "MISS-MATCH Error in Asset Register")
                'PrepareInsert = ""
                'Exit Function
            End If
        End If
        IDX = 0
        Do While IDX < UBound(ValueArray)
            If FieldTypes(FieldNameArray(IDX)) = 3 Or FieldTypes(FieldNameArray(IDX)) = 5 Then
                If Len(ValueArray(IDX)) > 1 Then
                    If Asc(Mid(ValueArray(IDX), 2, 1)) = 32 Then
                        'If FieldTypes.Items(IDX) = 3 Or FieldTypes.Items(IDX) = 5 Then
                        'item should be an integer and still a string thats empty
                            ValueArray(IDX) = "0"
                        'End If
                    End If
                End If
            End If
            IDX = IDX + 1
        Loop
        IncludeComma = False
        FieldValues = StringArrayToString(ValueArray, ",", 0, True, IncludeComma, "_")
        
        FinalCMD = "INSERT INTO " & DBTable & " (" & Fieldnames & ") VALUES (" & FieldValues & ")"
        'MsgBox ("FinalCMD = " & FinalCMD)
        PrepareInsert = FinalCMD
    Exit Function
    
Err_PrepareInsert:
    
    Call Error_Report("Error in PrepareInsert()")
End Function

Function PrepareUpdate(ByVal DBName As String, ByVal TableName As String, ByRef Fieldnames As String, ByVal FieldValues As String, _
        Optional ByVal Criteria As String = "", Optional ByRef ExcludeFields As String = "", Optional Encase_Fields As Boolean = False, _
        Optional FieldDelim As String = ";", Optional ValueDelim As String = ";") As String
    Dim FieldNameArray() As String
    Dim IgnoreFieldsArray() As String
    Dim ValueArray() As String
    Dim FinalCMD As String
    Dim IDX As Integer
    Dim fldName As String
    Dim fldValue As String
    Dim NumFields As Integer
    Dim UpdateCmd As String
    Dim IncludeComma As Boolean
    Dim IncludeSpeechMarks As Boolean

    FinalCMD = ""
    ReDim FieldNameArray(1)
    ReDim IgnoreFieldsArray(1)
    ReDim ValueArray(1)
    
    On Error GoTo Err_PrepareUpdate

        'If Len(IgnoreFields) > 0 Then
        'IgnoreFieldsArray = strToStringArray(IgnoreFields, ",")
        'End If
        If Len(TableName) = 0 Then
            MsgBox ("Error in PrepareUpdate: No Database Table specified")
            PrepareUpdate = ""
            Exit Function
        End If
        If Len(Fieldnames) = 0 Then
            'NumFields = GetNumFields(connString, "SELECT * FROM " & TableName, DBName, Fieldnames)
            'Fieldnames = GetFields(DBName, TableName)
        End If
        FieldNameArray = strToStringArray(Fieldnames, FieldDelim, 0, False, True)
        If Len(FieldValues) > 0 Then
            IncludeComma = True
                        IncludeSpeechMarks = True
            ValueArray = strToStringArray(FieldValues, ValueDelim, 0, True, True, IncludeComma, "_", IncludeSpeechMarks)
        Else
            MsgBox ("Error in PrepareUpdate: No values specified")
            PrepareUpdate = ""
            Exit Function
        End If
        'BUT check that the values are removed too if they corresponded with those fields removed ????
        Fieldnames = RemoveExtractedFields(FieldNameArray, IgnoreFieldsArray, ",", FieldNameArray, 0, 0) 'rebuilds whole list without the extracted fields
        
        If UBound(FieldNameArray) > 0 And UBound(ValueArray) > 0 Then
            If UBound(FieldNameArray) < UBound(ValueArray) Then
                MsgBox ("Error in PrepareUpdate: Number of Fields Passed are LESS than Number of VALUES passed.")
                PrepareUpdate = ""
                Exit Function
            End If
            If UBound(FieldNameArray) > UBound(ValueArray) Then
                MsgBox ("Error in PrepareUpdate: Number of Fields Passed are GREATER than Number of VALUES passed.")
                PrepareUpdate = ""
                Exit Function
            End If
        End If
        IDX = 0
        UpdateCmd = "UPDATE " & TableName & " SET "
        For IDX = 0 To UBound(FieldNameArray)
            fldName = FieldNameArray(IDX)
            fldValue = ValueArray(IDX)
            If Len(fldValue) > 0 Then
                If Encase_Fields Then
                    If IDX = 0 Then
                        FinalCMD = fldName & " = " & Chr(34) & fldValue & Chr(34)
                    Else
                        FinalCMD = FinalCMD & "," & fldName & " = " & Chr(34) & fldValue & Chr(34)
                    End If
                Else
                    If IDX = 0 Then
                        FinalCMD = fldName & " = " & fldValue
                    Else
                        FinalCMD = FinalCMD & "," & fldName & " = " & fldValue
                    End If
                End If
            Else
                'No value
            End If
        Next
        FinalCMD = UpdateCmd & FinalCMD
        If Len(Criteria) > 0 Then
            FinalCMD = FinalCMD & " WHERE " & Criteria
        End If
        
        PrepareUpdate = FinalCMD
Exit Function

Err_PrepareUpdate:
    
    Call Error_Report("Error in PrepareUpdate()")

End Function

Function strToStringArray(TheString As String, substring As String, Optional ByVal ElementStartIDX As Integer = 0, _
    Optional Encase_Fields As Boolean = False, Optional RemoveBadChars As Boolean = False, Optional IncludeCommaInBadChars As Boolean = False, _
            Optional REplaceWith As String = "_", Optional IncludeSpeechMarksInBadChars As Boolean = False) As String()
        Dim tempArr() As String
        Dim IDX As Integer
        Dim Elements As Integer
        Dim commaPos As Integer
        Dim Extract As String

        commaPos = 1
        IDX = 0
        Elements = ElementStartIDX
        ReDim tempArr(1)
        On Error GoTo err_strToStringArray

        Do Until commaPos = 0
            commaPos = InStr(IDX + 1, TheString, substring)
            If commaPos > 0 Then
                Extract = Mid(TheString, IDX + 1, (commaPos - (IDX + 1))) 'from delim to next delim
            Else
                Extract = Mid(TheString, IDX + 1, Len(TheString)) 'From 1 to end of string
            End If
            ReDim Preserve tempArr(UBound(tempArr) + 1)
            If RemoveBadChars Then
                Extract = ConvertBadChars(Extract, REplaceWith, IncludeCommaInBadChars, IncludeSpeechMarksInBadChars)
            End If
            If Encase_Fields Then
                Extract = Chr(34) & Extract & Chr(34)
            End If
            tempArr(Elements) = Extract
            IDX = commaPos
            Elements = Elements + 1
        Loop
            'Array.Copy(tempArr, strToStringArray, UBound(tempArr))
        strToStringArray = tempArr
        Exit Function
err_strToStringArray:
        Call Error_Report("Error in strToStringArray()")

End Function


Function StringArrayToString(ByRef theArray() As String, substring As String, Optional ByVal ElementStartIDX As Integer = 0, _
    Optional RemoveBadChars As Boolean = False, Optional IncludeCommaInBadChars As Boolean = False, Optional REplaceWith As String = "_") As String
    Dim tempArr() As String
    Dim IDX As Integer
    Dim Elements As Integer
    Dim FinalString As String
    Dim RawEntry As String
    
    IDX = 0
    Elements = ElementStartIDX
    FinalString = ""
    StringArrayToString = ""
    ReDim tempArr(1)
    On Error GoTo Err_StringArrayToString
    For Elements = ElementStartIDX To UBound(theArray)
        If Len(theArray(Elements)) = 0 Then
            'dont copy into string
        Else
            RawEntry = theArray(Elements)
            If RemoveBadChars Then
                RawEntry = ConvertBadChars(RawEntry, REplaceWith, IncludeCommaInBadChars)
            End If
            If Elements = ElementStartIDX Then
                
                FinalString = RawEntry
            Else
                FinalString = FinalString & substring & RawEntry
            End If
        End If
    Next
    
            'Array.Copy(tempArr, strToStringArray, UBound(tempArr))
    StringArrayToString = FinalString
    Exit Function
Err_StringArrayToString:
    Call Error_Report("Error in StringArrayToString()")

    End Function

Function Convert_strDateToDate(strDate As String, IncludeTime As Boolean) As Date
Dim FinalDate As Date
Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim intHour As Integer
Dim intMinute As Integer
Dim intSecond As Integer
Dim TheDate As String
Dim TheTime As String
Dim FinalTime As Date
Dim dtDateAndTime As Date
Dim strDateAndTime As String

If Not IsDate(strDate) Then
    Exit Function
End If
TheDate = Format(CDate(strDate), "yyyy-mm-dd")
TheTime = Format(CDate(strDate), "HH:MM:ss")
'so we have 2018-MAY-10
'21:48:56
'
intYear = CInt(Left(TheDate, 4))
intMonth = CInt(Mid(TheDate, 6, 2))
intDay = CInt(Mid(TheDate, 9, 2))

If intYear = 1899 Then
    intYear = 1901
End If
FinalDate = DateSerial(intYear, intMonth, intDay)
dtDateAndTime = FinalDate
If IncludeTime = True Then
    intHour = CInt(Left(TheTime, 2))
    intMinute = CInt(Mid(TheTime, 4, 2))
    intSecond = CInt(Mid(TheTime, 7, 2))
    FinalTime = TimeSerial(intHour, intMinute, intSecond) 'might need hh:nn:ss here ???
    strDateAndTime = Format(CStr(FinalDate) & " " & CStr(FinalTime), "yyyy/mmm/dd HH:MM:ss")
    dtDateAndTime = CDate(strDateAndTime)
End If

Convert_strDateToDate = dtDateAndTime

End Function

Sub ClearEntry(Optional TagLowRange As Long = 0, Optional TagUpperRange As Long = 0)
    Dim CTRL As Control
    Dim TagNo As Long
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        TagNo = 0
        If Len(CTRL.Tag) > 0 Then
            If TypeName(CTRL) = "TextBox" Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    TagNo = CLng(CTRL.Tag)
                End If
            End If
            If TypeName(CTRL) = "ComboBox" Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    TagNo = CLng(CTRL.Tag)
                End If
            End If
        End If
        FinalEntry = ""
        If TagNo > 0 Then
            If InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                FinalEntry = ""
                'set background to white:
                Call SetControlBackgroundColour(CStr(TagNo), vbWhite)
            End If
            If TagLowRange > 0 And TagNo <= TagUpperRange And TagNo >= TagLowRange Then
                CTRL.value = FinalEntry
            End If
            If TagLowRange = 0 And TagUpperRange = 0 Then
                CTRL.value = FinalEntry
            End If
        End If
    Next


End Sub

Sub PopulateUserformControls(WB As Workbook, WorksheetName As String, RecordRow As Long, Optional StartViewTimeTextboxTag As Long = 219, Optional Diff_Between_VTTAG As Long = 400)
    Dim CTRL As Control
    Dim myRow As Long
    Dim myCol As Long
    Dim Entry As String
    Dim Entry2 As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim ctrlArr() As String
    Dim TAGNumber As Long
    Dim strFinDate As String
    Dim strStartDate As String
    Dim TimeViewControlTagNumber As String
    Dim TimeViewArr() As String
    Dim TimeViewArrElement As Long
    Dim DateEntry As String
    Dim TimeEntry As String
    Dim DataTagNumber As Long
    Dim ControlCount As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    'NOT POPULATIng all controls ? - top missed out ??? and from TAG 63 onwards ?
    myCol = 1
    myRow = RecordRow
    'TotalFields = ActiveWorkbook.Worksheets(WorksheetName).Cells(1, Columns.Count).End(xlToLeft).Column
    ReDim ctrlArr(TotalFields + 1)
    Do While myCol <= TotalFields
        'ctrlArr(myCol) = ThisWB.Worksheets(WorksheetName).Cells(myRow, myCol).value
        
        myCol = myCol + 1
    Loop
    TimeViewArrElement = 0
    ReDim TimeViewArr(2)
    ControlCount = 0
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        'Gather the times for the TIME VIEW controls = TAG = 219 to 298
        ControlCount = ControlCount + 1
        If Len(CTRL.Tag) > 0 Then
            TimeEntry = ""
            If TypeName(CTRL) = "TextBox" Then
                If IsNumeric(CTRL.Tag) Then
                    TAGNumber = CLng(CTRL.Tag)
                    If TAGNumber >= StartViewTimeTextboxTag Then
                        DataTagNumber = TAGNumber - Diff_Between_VTTAG 'To get ViewTimeControl Tag to correspond to its corresponding DATA COLUMN/RECORD Position.
                        If DataTagNumber <= TotalFields And DataTagNumber > 0 Then
                            DateEntry = ctrlArr(DataTagNumber)
                            If Len(DateEntry) > 0 Then
                                If IsDate(DateEntry) Then 'might need hh:nn:ss here ???
                                    TimeEntry = Format(CDate(DateEntry), "HH:MM:ss")
                                Else
                                    'Entry not a date:
                                    TimeEntry = ""
                                End If
                            Else 'Date Entry is blank: BUT NO is still being put into the CB TEXT BOXES instead of blanks or times.
                                TimeEntry = ""
                            End If 'DataTagNumber <= TotalFields
                        End If 'DataTagNumber <= TotalFields
                        CTRL.Text = TimeEntry
                    End If 'If TagNumber > 209
                End If 'IsNumeric(ctrl.Tag)
            End If 'Typename() = Textbox
        End If 'if LEN(TAG)>0
    Next
    ControlCount = 0
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        ControlCount = ControlCount + 1
        If Len(CTRL.Tag) > 0 Then
            Entry = ""
            If TypeName(CTRL) = "TextBox" Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    TAGNumber = CLng(CTRL.Tag)
                    If TAGNumber <= TotalFields Then
                        Entry = ctrlArr(TAGNumber) 'GET DATA FROM ARRAY.
                        If Len(Entry) > 0 Then
                            If TAGNumber = ComplianceQuestion1TAG Then
                                If UCase(Entry) = "YES" Then
                                    'PUT YES in the Checkbox and disable the lower comment box:
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with the reason in Entry. If Entry is blank - this code is skipped.
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            If TAGNumber = ComplianceQuestion2TAG Then
                                If UCase(Entry) = "YES" Then
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with reason - even if left blank.
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            If TAGNumber = ComplianceQuestion3TAG Then
                                If UCase(Entry) = "YES" Then
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with reason - even if left blank.
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            
                            
                            'IF the control is a checkbox textbox then test if it has the right symbol in it:
                            If InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                                'ITS A CHECKBOX
                                If IsDate(Entry) Then 'Is it a start date or end date ? or just a NORMAL Date ???
                                    'Get next date:
                                    'if there is an empty date - wont be saved in the array - therfore out of sync.
                                    
                                    Entry2 = ctrlArr(TAGNumber + 1)
                                    strStartDate = Entry
                                    strFinDate = Entry2
                                    'TimeViewControlTagNumber = TagNumber + 191
                                    'TimeViewArr(TimeViewArrElement) = ctrlArr(TagNumber) 'Time View controls start at TAG NUMBER 210.
                                    'TimeViewArr(TimeViewArrElement + 1) = ctrlArr(TagNumber + 1) 'Time Values are recorded in pairs - DateTime start,DAteTime finish
                                    'ReDim Preserve TimeViewArr(UBound(TimeViewArr) + 2)
                                    'TimeViewArrElement = TimeViewArrElement + 2
                                    CTRL.Text = Chr(82)
                                Else
                                    'CONTROL is still a checkbox but not DATE:
                                    If UCase(Entry) = "YES" Then
                                        CTRL.Text = Chr(80)
                                    Else
                                        CTRL.Text = ""
                                    End If
                                    If InStr(1, CTRL.Name, "CBYesNo", vbTextCompare) > 0 Then
                                        If UCase(Entry) = "YES" Then
                                            CTRL.Text = "YES"
                                        Else
                                            CTRL.Text = "NO"
                                        End If
                                    End If
                                End If
                                
                                
                            Else 'ITS NOT A CHECKBOX - normal FULL TEXTBOX
                                CTRL.Text = Entry
                            End If
                        Else ' Len(Entry) = 0 EMPTY cell in RECORD:
                            If InStr(1, CTRL.Name, "CB", vbTextCompare) > 0 Then
                                'populate the checkbox with NO or a empty.
                                CTRL.Text = ""
                                If InStr(1, CTRL.Name, "CBYesNo", vbTextCompare) > 0 Then
                                    If Len(CTRL.Text) = 0 Then
                                        CTRL.Text = "NO"
                                    End If
                                End If
                                
                            Else
                                'OTHER TEXT BOXES that are empty - including the Compliance Questions:
                                'SO really we need to cast a default entry for the compliance questions -
                                ' or at least give the choice to the business logic user:
                                If TAGNumber = ComplianceQuestion1TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = True
                                End If
                                If TAGNumber = ComplianceQuestion2TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = True
                                End If
                                If TAGNumber = ComplianceQuestion3TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = True
                                End If
                                
                                CTRL.Text = ""
                            End If
                        End If
                    Else 'TAG NUMBER > TotalFields:
                        
                        'Populate the TIME VIEW CONTROLS - from 219 onwards:
                        
                    End If ' TagNumber <= TotalFields
                End If 'IsNumeric
            End If 'TypeName(ctrl)
            If TypeName(CTRL) = "ComboBox" Then
                If IsNumeric(CLng(CTRL.Tag)) Then
                    TAGNumber = CLng(CTRL.Tag)
                    If TAGNumber <= TotalFields Then
                        Entry = ctrlArr(TAGNumber) 'GET DATA FROM RECORD CELL ON SPREADSHEET: Timesheet Records
                        CTRL.Text = Entry
                    Else
                        'combo box has entries outside the visible range of the stored record:
                        
                    End If
                End If
            End If 'TypeName(ctrl)
        End If 'IF LEN(TAG)>0
NextIteration:
    Next
    
End Sub

Sub PopulateUserformControls_From_Access(strDeliveryRef As String, ASN As String, strDeliveryDate As String, StartTAG As Long, EndTag As Long, DBAccessPath As String, DBTable As String, Optional ByRef TotalDBFields As Long = 0, Optional StartViewTimeTextboxTag As Long = 441, Optional Diff_Between_VTTAG As Long = 400)
    Dim Formctrl As Control
    Dim myRow As Long
    Dim myCol As Long
    Dim Entry As String
    Dim Entry2 As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim ctrlArr() As String
    Dim TAGNumber As Long
    Dim strFinDate As String
    Dim strStartDate As String
    Dim TimeViewControlTagNumber As String
    Dim TimeViewArr() As String
    Dim TimeViewArrElement As Long
    Dim DateEntry As String
    Dim TimeEntry As String
    Dim DataTagNumber As Long
    Dim ControlCount As Long
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim SearchCriteria As String
    Dim SearchDateCriteria As String
    Dim dict_Wholerec As Scripting.Dictionary
    Dim strSupplier As String
    Dim strASNNO As String
    Dim strExpectedCases As String
    Dim IDX As Long
    Dim DictKey As Variant
    Dim DictItem As String
    Dim FieldnameArr() As String
    Dim TheFieldname As String
    Dim SearchText As String
    Dim ControlLastSaved As Date
    Dim SortFields As String
    Dim Reversed As Boolean
    'use LoadAccessDBTable()
    
    Set Formctrl = Nothing
    
    'NOT POPULATIng all controls ? - top missed out ??? and from TAG 63 onwards ?
    myCol = 1
    'TotalFields = ActiveWorkbook.Worksheets(WorksheetName).Cells(1, Columns.Count).End(xlToLeft).Column
    ReDim ctrlArr(TotalDBFields + 1) 'not used here.
    DBTable = "tblDeliveryInfo"
    '5 tables that contain 145 fields of info that have to be retrieved.
    '1 - 16 = INBOUND DATA, 17 - 145 OPERATIONS and SUPPLIER COMPLIANCE
    'But the number of Operatives is variable. Each control has a TAG number - unique.
    'variable amount of records - one record per Operative with Start and End Times.
    SearchCriteria = ""
    SearchDateCriteria = ""
    SortFields = ""
    Reversed = False
    'We NEED the DeliveryReference or the ASN and the DeliveryDate here:
    If Len(strDeliveryDate) > 0 Then
        SearchDateCriteria = "[DeliveryDate] = " & "#" & Format(strDeliveryDate, "yyyy/mm/dd") & "#"
    End If
    If Len(strDeliveryRef) > 0 Then
        SearchText = strDeliveryRef
        SearchCriteria = "[DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34)
    End If
    If Len(ASN) > 0 Then
        SearchText = ASN
        SearchCriteria = "[ASNNumber] = " & Chr(34) & ASN & Chr(34)
    End If
    
    If Len(SearchDateCriteria) > 0 Then
        SearchCriteria = SearchCriteria & " AND " & SearchDateCriteria
    End If
    
    
    Set dict_Wholerec = CreateObject("Scripting.Dictionary")
    dict_Wholerec.RemoveAll
    If LoadAccessDBTable(DBTable, DBAccessPath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues, dict_Wholerec) Then
        
        IDX = StartTAG
        'POPULATE all the form controls with the found record - which the user has just selected - either the DeliveryRef or the ASN Number.
        
        FieldnameArr = Split(Fieldnames, ",")
        Do While IDX <= EndTag
            'Use the function to find the correct database table based on the control name passed ?
            'ctrl.name = txtDeliveryRef for example:
            TheFieldname = FieldnameArr(IDX - 1)
            Set Formctrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TextBox", CStr(IDX))
            
            DictKey = dict_Wholerec(TheFieldname)
            DictItem = dict_Wholerec(TheFieldname).Item 'gives the VALUE from the Fieldname in the DB record.
            Call AddControlInfo(ControlLastSaved, DBTable, TheFieldname, 1, StartTAG, CStr(IDX), Formctrl, Formctrl.Name, DictItem, TypeName(Formctrl), Formctrl.Tag, Now(), _
                Formctrl.Left, Formctrl.Top, Formctrl.Width, Formctrl.Height, CDate(strDeliveryDate), strDeliveryRef, ASN, IDX)
            IDX = IDX + 1
        Loop
    End If
    
    'Now use all of the Control Info based in the CONTROLS COLLECTION to populate ALL of the controls on the userform :
    
    TimeViewArrElement = 0
    ReDim TimeViewArr(2)
    ControlCount = 0
    
    For Each Formctrl In frmGI_TimesheetEntry2_1060x630.Controls
        'Gather the times for the TIME VIEW controls = TAG = 219 to 298
        ControlCount = ControlCount + 1
        If Len(Formctrl.Tag) > 0 Then
            TimeEntry = ""
            If TypeName(Formctrl) = "TextBox" Then
                If IsNumeric(Formctrl.Tag) Then
                    TAGNumber = CLng(Formctrl.Tag)
                    If TAGNumber >= StartViewTimeTextboxTag Then
                        DataTagNumber = TAGNumber - Diff_Between_VTTAG 'To get ViewTimeControl Tag to correspond to its corresponding DATA COLUMN/RECORD Position.
                        If DataTagNumber <= TotalFields Then
                            DateEntry = ctrlArr(DataTagNumber)
                            If Len(DateEntry) > 0 Then
                                If IsDate(DateEntry) Then 'might need hh:nn:ss instead ?
                                    TimeEntry = Format(CDate(DateEntry), "HH:MM:ss")
                                Else
                                    'Entry not a date:
                                    TimeEntry = ""
                                End If
                            Else 'Date Entry is blank: BUT NO is still being put into the CB TEXT BOXES instead of blanks or times.
                                TimeEntry = ""
                            End If 'DataTagNumber <= TotalFields
                        End If 'DataTagNumber <= TotalFields
                        Formctrl.Text = TimeEntry 'Will show only the TIME PART - but will need full date and time to work out total hours.
                    End If 'If TagNumber > 209
                End If 'IsNumeric(ctrl.Tag)
            End If 'Typename() = Textbox
        End If 'if LEN(TAG)>0
    Next
    'Dont forget to populate all of the Short and EXTRA dynamic controls with the SKU numbers from the tblShortAndExtras
    
    ControlCount = 0
    For Each Formctrl In frmGI_TimesheetEntry2_1060x630.Controls
        ControlCount = ControlCount + 1
        If Len(Formctrl.Tag) > 0 Then
            Entry = ""
            If TypeName(Formctrl) = "TextBox" Then
                If IsNumeric(CLng(Formctrl.Tag)) Then
                    TAGNumber = CLng(Formctrl.Tag)
                    If TAGNumber <= TotalFields Then 'May now need Lower TAG REF and Upper TAG ref. we have sep table for Supplier Compliance.
                        Entry = ctrlArr(TAGNumber) 'GET DATA FROM RECORD in ACCESS DATABASE
                        If Len(Entry) > 0 Then
                            If TAGNumber = ComplianceQuestion1TAG Then 'Constant TAG Reference - which will now be from 800.
                                If UCase(Entry) = "YES" Then
                                    'PUT YES in the Checkbox and disable the lower comment box:
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with the reason in Entry. If Entry is blank - this code is skipped.
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            If TAGNumber = ComplianceQuestion2TAG Then
                                If UCase(Entry) = "YES" Then
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with reason - even if left blank.
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            If TAGNumber = ComplianceQuestion3TAG Then
                                If UCase(Entry) = "YES" Then
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "YES"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = ""
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = False
                                Else
                                    'DEFAULT condition = set to NO with reason - even if left blank.
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = True
                                End If
                                GoTo NextIteration
                            End If
                            
                            
                            
                            'IF the control is a checkbox textbox then test if it has the right symbol in it:
                            If InStr(1, Formctrl.Name, "CB", vbTextCompare) > 0 Then
                                'ITS A CHECKBOX
                                If IsDate(Entry) Then 'Is it a start date or end date ? or just a NORMAL Date ???
                                    'Get next date:
                                    'if there is an empty date - wont be saved in the array - therfore out of sync.
                                    
                                    Entry2 = ctrlArr(TAGNumber + 1)
                                    strStartDate = Entry
                                    strFinDate = Entry2
                                    'TimeViewControlTagNumber = TagNumber + 191
                                    'TimeViewArr(TimeViewArrElement) = ctrlArr(TagNumber) 'Time View controls start at TAG NUMBER 210.
                                    'TimeViewArr(TimeViewArrElement + 1) = ctrlArr(TagNumber + 1) 'Time Values are recorded in pairs - DateTime start,DAteTime finish
                                    'ReDim Preserve TimeViewArr(UBound(TimeViewArr) + 2)
                                    'TimeViewArrElement = TimeViewArrElement + 2
                                    Formctrl.Text = Chr(82)
                                Else
                                    'CONTROL is still a checkbox but not DATE:
                                    If UCase(Entry) = "YES" Then
                                        Formctrl.Text = Chr(80)
                                    Else
                                        Formctrl.Text = ""
                                    End If
                                    If InStr(1, Formctrl.Name, "CBYesNo", vbTextCompare) > 0 Then
                                        If UCase(Entry) = "YES" Then
                                            Formctrl.Text = "YES"
                                        Else
                                            Formctrl.Text = "NO"
                                        End If
                                    End If
                                End If
                                
                                
                            Else 'ITS NOT A CHECKBOX - normal FULL TEXTBOX
                                Formctrl.Text = Entry
                            End If
                        Else ' Len(Entry) = 0 EMPTY cell in RECORD:
                            If InStr(1, Formctrl.Name, "CB", vbTextCompare) > 0 Then
                                'populate the checkbox with NO or a empty.
                                Formctrl.Text = ""
                                If InStr(1, Formctrl.Name, "CBYesNo", vbTextCompare) > 0 Then
                                    If Len(Formctrl.Text) = 0 Then
                                        Formctrl.Text = "NO"
                                    End If
                                End If
                                
                            Else
                                'OTHER TEXT BOXES that are empty - including the Compliance Questions:
                                'SO really we need to cast a default entry for the compliance questions -
                                ' or at least give the choice to the business logic user:
                                If TAGNumber = ComplianceQuestion1TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTime.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Tag = ComplianceQuestion1TAG
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtArrivedONTimeComment.Visible = True
                                End If
                                If TAGNumber = ComplianceQuestion2TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafe.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Tag = ComplianceQuestion2TAG
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtIsItSafeComment.Visible = True
                                End If
                                If TAGNumber = ComplianceQuestion3TAG Then
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Tag = "0"
                                    frmGI_TimesheetEntry2_1060x630.txtCompleted.Text = "NO"
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Tag = ComplianceQuestion3TAG
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Text = Entry
                                    frmGI_TimesheetEntry2_1060x630.txtCompletedComment.Visible = True
                                End If
                                
                                Formctrl.Text = ""
                            End If
                        End If
                    Else 'TAG NUMBER > TotalFields:
                        
                        'Populate the TIME VIEW CONTROLS - from 219 onwards:
                        
                    End If ' TagNumber <= TotalFields
                End If 'IsNumeric
            End If 'TypeName(ctrl)
            If TypeName(Formctrl) = "ComboBox" Then
                If IsNumeric(CLng(Formctrl.Tag)) Then
                    TAGNumber = CLng(Formctrl.Tag)
                    If TAGNumber <= TotalFields Then
                        Entry = ctrlArr(TAGNumber) 'GET DATA FROM array which would be from RECORD in ACCESS DATABASE initally.
                        Formctrl.Text = Entry
                    Else
                        'combo box has entries outside the visible range of the stored record:
                        
                    End If
                End If
            End If 'TypeName(ctrl)
        End If 'IF LEN(TAG)>0
NextIteration:
    Next

End Sub

Sub AddNewOperatives(ByRef OpID As Long, ByRef TagID As Long, strDeliveryDate As String, strDeliveryRef As String, ASN As String, _
        ByVal TimeTAGStart As Long, Fieldnames As String, TotalRows As Long, Optional ByRef NEWIndex As Long = 0)
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
    Dim ControlFieldname As String
    Dim FieldnameArr() As String
    Dim LoadFieldsOK As Boolean
    Dim ControlFieldsTable As String
    Dim SearchCriteria As String
    Dim ControlDBTable As String
    
    ReDim FieldnameArr(1)
    If Len(Fieldnames) > 0 Then
        FieldnameArr = strToStringArray(Fieldnames, ",", 1, False, False, False, "_", False)
    Else
        ReDim FieldnameArr(10)
    End If
    RowGap = 19
    If OpID = 1 Then
        TopPos = 1
    Else
        TopPos = (OpID - 1) * RowGap
    End If
    
    ControlDBTable = "tblOperatives"
    
    If IsDate(strDeliveryDate) Then
        ControlDeliveryDate = CDate(strDeliveryDate)
    Else
        'MsgBox ("Need to pass proper delivery date")
        'Exit Sub
        ControlDeliveryDate = CDate("01/01/1970")
    End If
    ControlDeliveryRef = strDeliveryRef
    
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
    ControlTotalRows = TotalRows ' need to know lowerTAG and number of fields in frame_Operatives.
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    ComboArray = PopulateDropdowns("Employees", 2, 0, False, WB_MainTimesheetData)
    
    ControlFieldname = FieldnameArr(1)
    ControlRowNumber = OpID
    
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
        "comOperativeName" & CStr(OpID), ControlText, ControlType, ControlTAG, _
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
    
    ControlFieldname = FieldnameArr(2)
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
    "comOperativeActivity" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    TagID = TagID + 1
    
    ControlType = "BTN"
    ControlText = "@"
    ControlTAG = "BTN" & CStr(btnTAGID)
    ControlLeft = 325
    ControlWidth = 20
    ControlHeight = 20
    BackColor = RGB(255, 255, 0) 'YELLOW ?
    ControlLeftMargin = False
    
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
    "btnOperativeTimeStart" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    btnTAGID = btnTAGID + 1
    
    'txtOperativeTimeStart :
    
    TimeTAGID = TagID + TimeTAGStart
    
    ControlType = "TEXTBOX"
    ControlText = "00:00:00"
    ControlTAG = CStr(TimeTAGID)
    ControlLeft = 350
    ControlWidth = 60
    'ControlHeight = 20
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    ControlFieldname = FieldnameArr(3)
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
    "txtOperativeTimeStart" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    
    ControlType = "BTN"
    ControlText = "@"
    ControlTAG = "BTN" & CStr(btnTAGID)
    ControlLeft = 430
    ControlWidth = 20
    'ControlHeight = 20
    ControlText = "@"
    BackColor = RGB(255, 255, 20)
    ControlLeftMargin = False
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
    "btnOperativeTimeEnd" & CStr(OpID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    btnTAGID = btnTAGID + 1
    
    TimeTAGID = TagID + TimeTAGStart
    'TAGID = TAGID + 1
    
    ControlType = "TEXTBOX"
    ControlTAG = CStr(TimeTAGID)
    ControlText = "00:00:00"
    ControlLeft = 455
    ControlWidth = 60
    'ControlHeight = 20
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    ControlFieldname = FieldnameArr(4)
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_Operatives.Controls, ControlFieldname, "ID", Nothing, _
    "txtOperativeTimeEnd" & CStr(OpID), _
        ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, _
        ComboArray, BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    OpID = OpID + 1
    
    ScrollBarHeight = frmGI_TimesheetEntry2_1060x630.Frame_Operatives.ScrollHeight
    If OpID > 1 Then
        'roughly 100 = 5 rows
        'ScrollBarHeight = ScrollBarHeight + (100 / 5)
        ScrollBarHeight = ScrollBarHeight + (OpID * 20)
        frmGI_TimesheetEntry2_1060x630.Frame_Operatives.ScrollHeight = ScrollBarHeight
    End If
    OperativeCount = OpID
    TextTAGID = TagID


End Sub


Sub AddNewShorts(ByRef PartID As Long, ByRef TagID As Long, strDeliveryDate As String, strDeliveryRef As String, ASN As String, _
        LowerTag As Long, UpperTag As Long, Optional TotalRows As Long = 0, Optional ByRef NEWIndex As Long = 0)
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
    Dim ControlFieldname As String
    Dim FieldnameArr() As String
    Dim Fieldnames As String
    
    RowGap = 19
    If PartID = 1 Then
        TopPos = 1
    Else
        TopPos = (PartID - 1) * RowGap
    End If
    
    If IsDate(strDeliveryDate) Then
        ControlDeliveryDate = CDate(strDeliveryDate)
    Else
        'MsgBox ("Need to pass proper delivery date")
        'Exit Sub
        ControlDeliveryDate = CDate("01/01/1970")
    End If
    ControlDeliveryRef = strDeliveryRef
    
    Fieldnames = GetFieldnames_From_ACCESS("tblShortsAndExtraParts", AccessDBpath, "", "")
    FieldnameArr = strToStringArray(Fieldnames, ",", 0, False, False, False, "_", False)
    
    ControlType = "TEXTBOX"
    MakeVisible = True
    ControlText = ""
    ControlTAG = CStr(TagID)
    ControlDate = Now()
    ControlLeft = 0
    ControlTop = TopPos
    ControlHeight = 0
    ControlWidth = 95
    'ControlDeliveryDate = strDeliveryDate
    ControlDeliveryRef = strDeliveryRef
    ControlASN = ASN
    ControlOBJCount = PartID
    ControlStartTAG = CStr(LowerTag)
    ControlEndTAG = CStr(UpperTag)
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.CompareMode = vbTextCompare
    ControlRowNumber = CStr(PartID)
    ControlTotalRows = TotalRows
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    'FOR EACH CONTROL in the SHORTS frame:
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_ShortParts.Controls, ControlFieldname, "ID", Nothing, _
    "txtShortPartNo" & CStr(PartID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, ComboArray, _
        BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    
    ControlType = "TEXTBOX"
    MakeVisible = True
    ControlText = ""
    ControlTAG = CStr(TagID)
    ControlDate = Now()
    ControlLeft = 95
    ControlTop = TopPos
    ControlHeight = 0
    ControlWidth = 50
    'ControlDeliveryDate = strDeliveryDate
    ControlDeliveryRef = strDeliveryRef
    ControlASN = ASN
    ControlOBJCount = PartID
    ControlStartTAG = CStr(LowerTag)
    ControlEndTAG = CStr(UpperTag)
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.CompareMode = vbTextCompare
    ControlRowNumber = PartID
    ControlTotalRows = TotalRows
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_ShortParts.Controls, ControlFieldname, "ID", Nothing, _
        "txtShortQty" & CStr(PartID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, ComboArray, _
        BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    
    PartID = PartID + 1
    ShortCount = PartID
    ShortTAGID = TagID
End Sub

Sub AddNewExtras(ByRef PartID As Long, ByRef TagID As Long, strDeliveryDate As String, strDeliveryRef As String, ASN As String, _
        LowerTag As Long, UpperTag As Long, Optional TotalRows As Long = 0, Optional ByRef NEWIndex As Long = 0)
    
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
    Dim ControlFieldname As String
    Dim FieldnameArr() As String
    
    RowGap = 19
    If PartID = 1 Then
        TopPos = 1
    Else
        TopPos = (PartID - 1) * RowGap
    End If
    
    If IsDate(strDeliveryDate) Then
        ControlDeliveryDate = CDate(strDeliveryDate)
    Else
        'MsgBox ("Need to pass proper delivery date")
        'Exit Sub
        ControlDeliveryDate = CDate("01/01/1970")
    End If
    ControlDeliveryRef = strDeliveryRef
    
    ControlType = "TEXTBOX"
    MakeVisible = True
    ControlText = ""
    ControlTAG = CStr(TagID)
    ControlDate = Now()
    ControlLeft = 0
    ControlTop = TopPos
    ControlHeight = 0
    ControlWidth = 95
    'ControlDeliveryDate = strDeliveryDate
    ControlDeliveryRef = strDeliveryRef
    ControlASN = ASN
    ControlOBJCount = PartID
    ControlStartTAG = CStr(LowerTag)
    ControlEndTAG = CStr(UpperTag)
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.CompareMode = vbTextCompare
    ControlRowNumber = CStr(PartID)
    ControlTotalRows = TotalRows
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_ExtraParts.Controls, ControlFieldname, "ID", Nothing, _
        "txtExtraPartNo" & CStr(PartID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, ComboArray, _
        BackColor, ControlLeftMargin)
    
    TagID = TagID + 1
    
    ControlType = "TEXTBOX"
    MakeVisible = True
    ControlText = ""
    ControlTAG = CStr(TagID)
    ControlDate = Now()
    ControlLeft = 95
    ControlTop = TopPos
    ControlHeight = 0
    ControlWidth = 50
    'ControlDeliveryDate = strDeliveryDate
    ControlDeliveryRef = strDeliveryRef
    ControlASN = ASN
    ControlOBJCount = PartID
    ControlStartTAG = CStr(LowerTag)
    ControlEndTAG = CStr(UpperTag)
    Set Dic_Collection = CreateObject("Scripting.Dictionary")
    Dic_Collection.CompareMode = vbTextCompare
    ControlRowNumber = PartID
    ControlTotalRows = TotalRows
    BackColor = RGB(240, 248, 255) 'ALICEBLUE
    ControlLeftMargin = False
    
    NEWIndex = AddNewControl(True, frmGI_TimesheetEntry2_1060x630.Frame_ExtraParts.Controls, ControlFieldname, "ID", Nothing, _
        "txtExtraQty" & CStr(PartID), ControlText, ControlType, ControlTAG, _
        ControlDate, ControlLeft, ControlTop, ControlWidth, ControlHeight, ControlDeliveryDate, ControlDeliveryRef, ControlASN, _
        ControlOBJCount, ControlStartTAG, ControlEndTAG, Dic_Collection, ControlRowNumber, ControlTotalRows, MakeVisible, ComboArray, _
        BackColor, ControlLeftMargin)
    
    PartID = PartID + 1
    ExtraCount = PartID
    ExtraTAGID = TagID

End Sub

Sub Test_PopulationOfControls(strDeliveryRef As String, ASN As String, strDeliveryDate As String, StartTAG As Long, EndTag As Long, _
        DBAccessPath As String, DBTable As String, Optional ByRef TotalDBFields As Long = 0, Optional StartViewTimeTextboxTag As Long = 441, _
        Optional Diff_Between_VTTAG As Long = 400)
    Dim Formctrl As Control
    Dim myRow As Long
    Dim myCol As Long
    Dim Entry2 As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim ctrlArr() As String
    Dim strFinDate As String
    Dim strStartDate As String
    Dim TimeViewControlTagNumber As String
    Dim TimeViewArr() As String
    Dim TimeViewArrElement As Long
    Dim TimeStartTAG As Long
    Dim TimeEndTAG As Long
    Dim TimeStartValue As String
    Dim TimeEndValue As String
    Dim TimeStartFormControl As Control
    Dim TimeEndFormControl As Control
    Dim TimeStartFieldname As String
    Dim TimeEndFieldname As String
    Dim DateEntry As String
    Dim TimeEntry As String
    Dim DataTagNumber As Long
    Dim ControlCount As Long
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim SearchCriteria As String
    Dim SearchDateCriteria As String
    Dim strSupplier As String
    Dim strASNNO As String
    Dim strExpectedCases As String
    Dim IDX As Long
    Dim DictKey As Variant
    Dim DictItem As String
    Dim FieldnameArr() As String
    Dim TheFieldname As String
    Dim SearchText As String
    Dim ControlName As String
    Dim strTagNumber As String
    Dim TagName As String
    Dim TAGNumber As Long
    Dim ControlLeft As Integer
    Dim ControlTop As Integer
    Dim ControlWidth As Integer
    Dim ControlHeight As Integer
    Dim ControlType As String
    Dim varControlKey1 As Variant
    Dim varControlKey2 As Variant
    Dim varControlKey3 As Variant
    Dim varControlKey4 As Variant
    Dim ControlValue As String
    Dim ControlValueArr() As String
    Dim ControlDBTable As String
    Dim ControlLowerTag As Long
    Dim ControlUpperTag As Long
    Dim ControlFieldname As String
    Dim ControlDate As Date
    Dim ControlLastSaved As Date
    Dim varCollectionKEY As Variant
    Dim ctrlProperty As Variant
    Dim DBTables() As String
    Dim LowerTag() As Long
    Dim UpperTag() As Long
    Dim TableIDX As Long
    Dim LoadedOK() As Boolean
    Dim FieldIDX As Long
    Dim StartIDX() As Long
    Dim EndIDX() As Long
    Dim FrameRows() As Long
    Dim DateAndRefCriteria As String
    Dim RecordIDCriteria As String
    Dim myDic_Collection As New Scripting.Dictionary
    Dim dict_ReturnLookupFields As New Scripting.Dictionary
    Dim dict_Wholerecs As New Scripting.Dictionary
    Dim ControlFieldsTable As String
    Dim AllFields As Variant
    Dim FieldNameFromFieldTable As String
    Dim FieldValueFromFieldTable As Variant
    Dim TAGNameFromFieldTable As String
    Dim FieldsTableLoadedOK As Boolean
    Dim allLookupFields As Variant
    Dim AllCurrentFields As Variant
    Dim JustTAG As Boolean
    Dim ReturnValue As Variant
    Dim LastTAG As Long
    Dim RecStartTAG As Variant
    Dim RecEndTAG As Variant
    Dim TotalRows As Long
    Dim TotalFrameControls As Long
    Dim FoundItem As Boolean
    Dim RecID As Variant
    Dim ValueArr() As String
    Dim AllShortParts As Variant
    Dim AllExtraParts As Variant
    Dim RowIDX As Long
    Dim HighestTag As Long
    Dim LowestTAG As Long
    Dim Entry As Variant
    Dim NumControls As Long
    Dim LookupFields As String
    Dim LookupValues As String
    Dim SortFields As String
    Dim Reversed As Boolean
    Dim NewTAG As Long
    
    ReDim DBTables(6)
    ReDim LowerTag(6)
    ReDim UpperTag(6)
    ReDim LoadedOK(6)
    ReDim StartIDX(6)
    ReDim EndIDX(6)
    ReDim FrameRows(6)
    
    Set dict_ReturnLookupFields = New Scripting.Dictionary
    Set dict_Wholerecs = New Scripting.Dictionary
    Set myDic_Collection = New Scripting.Dictionary
    myDic_Collection.CompareMode = TextCompare
    dict_Wholerecs.CompareMode = TextCompare
    dict_ReturnLookupFields.CompareMode = TextCompare
    
    ControlFieldsTable = "tblFieldsAndTAGs"
    'SearchCriteria = "TableName = " & Chr(34) & ControlDBTable & Chr(34)
    SearchCriteria = ""
    JustTAG = True
    SortFields = "TABLENAME,SORTFIELD"
    Reversed = False
    FieldsTableLoadedOK = LoadAccessDBTable(ControlFieldsTable, DBAccessPath, JustTAG, SearchCriteria, SortFields, Reversed, _
            False, LookupFields, LookupValues, dict_ReturnLookupFields, allLookupFields)
    
    DBTables(1) = "tblDeliveryInfo"
    LowerTag(1) = FirstInboundTAG
    UpperTag(1) = 28
    StartIDX(1) = 1
    EndIDX(1) = 28
    DBTables(2) = "tblLabourHours"
    LowerTag(2) = 40
    UpperTag(2) = 42
    StartIDX(2) = 3
    EndIDX(2) = 5
    
    strDeliveryDate = Format(CDate(strDeliveryDate), "dd/mm/yyyy")
    SearchText = strDeliveryRef
    'SearchCriteria = "([DeliveryDate] = #" & strDeliveryDate & "#) AND ([DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34) & ")"
    SearchCriteria = "([DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34) & ")"
    SortFields = ""
    Reversed = False
    FoundItem = LoadAccessDBTable(DBTables(2), AccessDBpath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues)
    DBTables(3) = "tblOperatives"
    
    'Fieldnames is populated with all of the fields from the table, FieldValues populated with all the values from the record as comma-delim string
    'Retrieve the Delivery Date in the database for the dropdown just selected:
    If FoundItem Then
        ValueArr = Split(FieldValues, ",") 'FieldValues has quotes around them !!!!!!!!!!!!!!!!!!!!
        LowerTag(3) = ValueArr(11)
        UpperTag(3) = ValueArr(12)
    Else
        LowerTag(3) = 43
        UpperTag(3) = 46
    End If
    TotalFrameControls = 4
    FrameRows(3) = (UpperTag(3) - (LowerTag(3) - 1)) / TotalFrameControls
    StartIDX(3) = 4
    EndIDX(3) = 7
    OperativeCount = 1
    NewTAG = LowerTag(3)
    IDX = 1
    Do While IDX <= FrameRows(3)
        Call AddNewOperatives(OperativeCount, NewTAG, strDeliveryDate, strDeliveryRef, ASN, Diff_Between_VTTAG, Fieldnames, FrameRows(3))
        IDX = IDX + 1
    Loop
    
    DBTables(4) = "tblShortsAndExtraParts"
    SearchText = strDeliveryRef
    SortFields = ""
    Reversed = False
    SearchCriteria = "DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34) & " AND ShortOrExtra = " & Chr(34) & "Short" & Chr(34)
    FoundItem = LoadAccessDBTable("tblShortsAndExtraParts", AccessDBpath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues, Nothing, _
        AllShortParts)
    If FoundItem Then
        HighestTag = 0
        LowestTAG = 1001
        RowIDX = 0
        Do While RowIDX < UBound(AllShortParts, 2)
            Entry = AllShortParts(9, RowIDX)
            If Entry > HighestTag Then
                HighestTag = Entry
            End If
            Entry = AllShortParts(8, RowIDX)
            If Entry < LowestTAG And Entry > 0 Then
                LowestTAG = Entry
            End If
            RowIDX = RowIDX + 1
        Loop
        ValueArr = Split(FieldValues, ",") 'FieldValues has quotes around them !!!!!!!!!!!!!!!!!!!!
        LowerTag(4) = LowestTAG
        UpperTag(4) = HighestTag
    Else
        LowerTag(4) = 1001
        UpperTag(4) = 1002
    End If
    TotalFrameControls = 2
    FrameRows(4) = (UpperTag(4) - (LowerTag(4) - 1)) / TotalFrameControls
    StartIDX(4) = 3
    EndIDX(4) = 4
    NewTAG = LowerTag(4)
    IDX = 1
    Do While IDX <= FrameRows(4)
        Call AddNewShorts(ShortCount, NewTAG, strDeliveryDate, strDeliveryRef, ASN, LowerTag(4), UpperTag(4), FrameRows(4))
        IDX = IDX + 1
    Loop
    
    DBTables(5) = "tblShortsAndExtraParts"
    SearchText = strDeliveryRef
    SortFields = ""
    Reversed = False
    SearchCriteria = "DeliveryReference = " & Chr(34) & strDeliveryRef & Chr(34) & " AND ShortOrExtra = " & Chr(34) & "Extra" & Chr(34)
    FoundItem = LoadAccessDBTable("tblShortsAndExtraParts", AccessDBpath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues, Nothing, AllExtraParts)
    If FoundItem Then
        HighestTag = 0
        LowestTAG = 2001
        RowIDX = 0
        Do While RowIDX < UBound(AllExtraParts, 2)
            Entry = AllExtraParts(9, RowIDX)
            If Entry > HighestTag Then
                HighestTag = Entry
            End If
            Entry = AllExtraParts(8, RowIDX)
            If Entry < LowestTAG And Entry > 0 Then
                LowestTAG = Entry
            End If
            RowIDX = RowIDX + 1
        Loop
        ValueArr = Split(FieldValues, ",") 'FieldValues has quotes around them !!!!!!!!!!!!!!!!!!!!
        LowerTag(5) = LowestTAG
        UpperTag(5) = HighestTag
    Else
        LowerTag(5) = 2001
        UpperTag(5) = 2002
    End If
    TotalFrameControls = 2
    FrameRows(5) = (UpperTag(5) - (LowerTag(5) - 1)) / TotalFrameControls
    StartIDX(5) = 3
    EndIDX(5) = 4
    NewTAG = LowerTag(5)
    IDX = 1
    Do While IDX <= FrameRows(5)
        Call AddNewExtras(ExtraCount, NewTAG, strDeliveryDate, strDeliveryRef, ASN, LowerTag(5), UpperTag(5), FrameRows(5))
        IDX = IDX + 1
    Loop
    'uncomment and set next array to index 6. redim as 6 now.
    DBTables(6) = "tblSupplierCompliance"
    LowerTag(6) = 801
    UpperTag(6) = 807
    StartIDX(6) = 3
    EndIDX(6) = 8
    'use LoadAccessDBTable()
    
    Set Formctrl = Nothing
    
    'NOT POPULATIng all controls ? - top missed out ??? and from TAG 63 onwards ?
    myCol = 1
    'TotalFields = ActiveWorkbook.Worksheets(WorksheetName).Cells(1, Columns.Count).End(xlToLeft).Column
    ReDim ctrlArr(TotalDBFields + 1)
    'DBTable = "tblDeliveryInfo"
    'variable amount of records - one record per Operative with Start and End Times.
    SearchCriteria = ""
    SearchDateCriteria = ""
    
    For TableIDX = 1 To UBound(DBTables)
        ControlLowerTag = LowerTag(TableIDX)
        ControlUpperTag = UpperTag(TableIDX)
        ControlDBTable = DBTables(TableIDX)
        SearchCriteria = "TableName = " & Chr(34) & ControlDBTable & Chr(34)
        JustTAG = True
        SortFields = ""
        Reversed = False
        FieldsTableLoadedOK = LoadAccessDBTable(ControlFieldsTable, DBAccessPath, JustTAG, SearchCriteria, SortFields, Reversed, _
            False, LookupFields, LookupValues, dict_ReturnLookupFields, allLookupFields, LowerTag(TableIDX), UpperTag(TableIDX), ControlLowerTag, False)
        
        If Len(strDeliveryRef) > 0 Then
            SearchText = strDeliveryRef
            SearchCriteria = "[DeliveryReference] = " & Chr(34) & strDeliveryRef & Chr(34)
        Else
            MsgBox ("No Delivery Reference Passed")
            Exit Sub
        End If
        If Len(ASN) > 0 Then
            SearchText = ASN
            SearchCriteria = "[ASNNumber] = " & Chr(34) & ASN & Chr(34)
        End If
        If Len(strDeliveryDate) > 0 Then
            SearchText = strDeliveryDate & "_" & SearchText
            SearchDateCriteria = "[DeliveryDate] = " & "#" & Format(strDeliveryDate, "yyyy/mm/dd") & "#"
        End If
        If Len(SearchDateCriteria) > 0 Then
            SearchCriteria = SearchCriteria & " AND " & SearchDateCriteria
        End If
        JustTAG = True
        Set AllCurrentFields = Nothing
        Set dict_Wholerecs = Nothing
        SortFields = ""
        Reversed = False
        LoadedOK(TableIDX) = LoadAccessDBTable(ControlDBTable, DBAccessPath, JustTAG, SearchCriteria, SortFields, Reversed, _
            False, Fieldnames, FieldValues, dict_Wholerecs, AllCurrentFields, 0, 0, ControlLowerTag - 1, True)
        
        If LoadedOK(TableIDX) Then
            FieldIDX = StartIDX(TableIDX) 'USE THIS !!!!!!!!!!!!!!!!!!!!!
            FieldnameArr = Split(Fieldnames, ",")
            IDX = ControlLowerTag
            'POPULATE all the form controls with the found record - which the user has just selected - either the DeliveryRef or the ASN Number.
            Do While IDX <= UpperTag(TableIDX)
                'Use the function to find the correct database table based on the control name passed ?
                'ctrl.name = txtDeliveryRef for example:
                TAGNumber = IDX
                varControlKey1 = CStr(TAGNumber) 'HERE wE get 41 as the tag number.
                'ReturnValue = dict_Wholerecs.Item(varControlKey1)
                ReturnValue = AllCurrentFields(FieldIDX, 0)
                'ControlFieldname = FLMStartTime - from the AllCurrentFields array. FieldIDX = 4 here.
                ControlValueArr = Split(ReturnValue, "_") '0) VALUE, 1) DBTable, 2) FIELDName
                'FieldNameFromFieldTable
                ControlValue = ControlValueArr(0)
                ControlFieldname = ControlValueArr(2)
                If UCase(ControlFieldname) = UCase("FLMCBStartTime") Then
                    'cheat - this will be fixed for tblLabourHours only: - NEVER entered as FLMCBStartTime is never reached ???
                    TimeStartTAG = TAGNumber + Diff_Between_VTTAG
                    TimeStartValue = AllCurrentFields(FieldIDX, 0)
                    Set TimeStartFormControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", CStr(TimeStartTAG))
                    If Not TimeStartFormControl Is Nothing Then
                        ControlValueArr = Split(TimeStartValue, "_")
                        TimeStartFieldname = ControlValueArr(2)
                        ControlName = TimeStartFormControl.Name
                        ControlLeft = TimeStartFormControl.Left
                        ControlTop = TimeStartFormControl.Top
                        ControlWidth = TimeStartFormControl.Width
                        ControlHeight = TimeStartFormControl.Height
                        ControlType = TypeName(Formctrl)
                        TimeStartFormControl.Text = ControlValueArr(0)
                        Call AddControlInfo(ControlLastSaved, ControlDBTable, TimeStartFieldname, 1, TimeStartTAG, CStr(TimeStartTAG), Formctrl, ControlName, _
                            ControlValueArr(0), ControlType, TagName, Now(), _
                            ControlLeft, ControlTop, ControlWidth, ControlHeight, CDate(strDeliveryDate), strDeliveryRef, ASN, TimeStartTAG, _
                            CStr(ControlLowerTag), CStr(ControlUpperTag), myDic_Collection)
                    End If
                End If
                If UCase(ControlFieldname) = UCase("FLMCBEndTime") Then
                    'cheat - this will be fixed for tblLabourHours only:
                    TimeEndTAG = TAGNumber + Diff_Between_VTTAG
                    TimeEndValue = AllCurrentFields(FieldIDX + 1, 0)
                    Set TimeEndFormControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", CStr(TimeEndTAG))
                    If Not TimeEndFormControl Is Nothing Then
                        ControlValueArr = Split(TimeStartValue, "_")
                        TimeEndFieldname = ControlValueArr(2)
                        ControlName = TimeEndFormControl.Name
                        ControlLeft = TimeEndFormControl.Left
                        ControlTop = TimeEndFormControl.Top
                        ControlWidth = TimeEndFormControl.Width
                        ControlHeight = TimeEndFormControl.Height
                        ControlType = TypeName(TimeEndFormControl)
                        TimeEndFormControl.Text = ControlValueArr(0)
                        Call AddControlInfo(ControlLastSaved, ControlDBTable, TimeEndFieldname, 1, TimeEndTAG, CStr(TimeEndTAG), Formctrl, ControlName, _
                            ControlValueArr(0), ControlType, TagName, Now(), _
                            ControlLeft, ControlTop, ControlWidth, ControlHeight, CDate(strDeliveryDate), strDeliveryRef, ASN, TimeEndTAG, _
                            CStr(ControlLowerTag), CStr(ControlUpperTag), myDic_Collection)
                    End If
                End If
                'DictKey = TheFieldname
                DictItem = ControlValue
                ControlLeft = 0
                ControlTop = 0
                ControlWidth = 0
                ControlHeight = 0
                ControlName = ""
                ControlType = "UNKNOWN"
                'ControlValue = DictItem
                Set Formctrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TextBox", CStr(TAGNumber))
                TagName = CStr(IDX)
                If Formctrl Is Nothing Then
                    Set Formctrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "ComboBox", CStr(IDX))
                    If Formctrl Is Nothing Then 'Could be passing a button or an imaage or a label ?
                        'BUGGER
                        ControlName = "NotFound"
                        ControlLeft = 0
                        ControlTop = 0
                        ControlWidth = 0
                        ControlHeight = 0
                    Else 'ComboBox
                        'grab tablename first:
                        ControlName = Formctrl.Name
                        ControlLeft = Formctrl.Left
                        ControlTop = Formctrl.Top
                        ControlWidth = Formctrl.Width
                        ControlHeight = Formctrl.Height
                        ControlType = TypeName(Formctrl)
                        Formctrl.Text = ControlValue
                    End If
                Else 'TEXTBOX
                    ControlName = Formctrl.Name
                    If UCase(ControlFieldname) = "LASTSAVED" Then
                        If IsDate(ControlValue) Then
                            ControlLastSaved = CDate(ControlValue)
                        Else
                            ControlLastSaved = CDate("01/01/1970")
                        End If
                    End If
                    ControlLeft = Formctrl.Left
                    ControlTop = Formctrl.Top
                    ControlWidth = Formctrl.Width
                    ControlHeight = Formctrl.Height
                    ControlType = TypeName(Formctrl)
                    If UCase(ControlName) = UCase("txtCBReadyLabel") Then
                        If UCase(ControlValue) = "PRE-LABELLED" Then
                            ControlValue = "YES"
                        Else
                            ControlValue = "NO"
                        End If
                    End If
                    Formctrl.Text = ControlValue
                End If
                
                'call AddControlInfo(1, 1, "ID", TheControl As Control, ControlName As String, ControlText As String, _
                'ControlType As String, ControlTag As String, ControlDate As Date, _
                'ControlLeft As Integer, ControlTop As Integer, ControlWidth As Integer, ControlHeight As Integer, _
                'ControlDeliveryDate As Date, ControlDeliveryRef As String, Optional ControlASN As String = "", Optional ControlObjCount As Long, _
                'Optional ControlStartTAG As String = "", Optional ControlEndTag As String = "", Optional ByRef Dic_Collection As Scripting.Dictionary, _
                'Optional ControlRowNumber As Long, Optional ControlTotalRows As Long, Optional MakeVisible As Boolean = True, _
                'Optional ByRef ListArray As Variant = Nothing)
                'REMEMBER - AS PER DELIVERY REF RECORD LOADED:
                Call AddControlInfo(ControlLastSaved, ControlDBTable, ControlFieldname, 1, IDX, CStr(IDX), Formctrl, ControlName, DictItem, ControlType, TagName, Now(), _
                    ControlLeft, ControlTop, ControlWidth, ControlHeight, CDate(strDeliveryDate), strDeliveryRef, ASN, IDX, _
                    CStr(ControlLowerTag), CStr(ControlUpperTag), myDic_Collection)
                'call AddLabourInfo(1,IDX,"Labour" & cstr(IDX),formctrl,controlname,dictitem,controltype,tagname,now(), _

                IDX = IDX + 1
                FieldIDX = FieldIDX + 1
            Loop
        Else 'DID NOT LOAD OK:
            'Could not find Delivery Date and DeliveryReference in table - but still construct the Controls Object:
            ' - as this assigns the AfterChange event to every form control also.
            'Need to loop round each control on the form = or from tblFieldsAndControls - to populate the collections Object in memory.
            'either way - use the global scripting dictionary . Needs the DeliveryDAte and DeliveryRef and TAG for the KEY.
            'Need to save the TABLE name and the FIELDNAME too with each control:
            SearchCriteria = "TableName = " & Chr(34) & ControlDBTable & Chr(34)
            JustTAG = True
            SortFields = ""
            Reversed = False
            If LoadAccessDBTable(ControlFieldsTable, AccessDBpath, JustTAG, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues, _
                dict_ReturnLookupFields, AllFields, 0, 0, ControlLowerTag) Then
                'Now extract relevant info from return script.dictionary:
                myRow = StartIDX(TableIDX) - 1
                LastTAG = EndIDX(TableIDX) - 1
                Do While myRow <= LastTAG
                    ControlFieldname = AllFields(2, myRow)
                    TagName = AllFields(3, myRow)
                    ControlName = AllFields(4, myRow)
                    ControlValue = ""
                    Set Formctrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", TagName)
                    If Not Formctrl Is Nothing Then
                        ControlLeft = Formctrl.Left
                        ControlTop = Formctrl.Top
                        ControlWidth = Formctrl.Width
                        ControlHeight = Formctrl.Height
                        ControlType = UCase(TypeName(Formctrl))
                    Else
                        Set Formctrl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "COMBOBOX", TagName)
                        If Not Formctrl Is Nothing Then
                            ControlLeft = Formctrl.Left
                            ControlTop = Formctrl.Top
                            ControlWidth = Formctrl.Width
                            ControlHeight = Formctrl.Height
                            ControlType = UCase(TypeName(Formctrl))
                        Else
                            'CONTROL NOT a Combo or TextBox.
                            ControlType = "UNKNOWN"
                        End If
                    End If
                    If UCase(ControlType) = "UNKNOWN" Then
                        'CONTROL NOT a COMBO or TEXTBOX - dont save.
                    Else
                        Call AddControlInfo(ControlLastSaved, ControlDBTable, ControlFieldname, 1, IDX, CStr(IDX), Formctrl, ControlName, ControlValue, ControlType, _
                            TagName, Now(), ControlLeft, ControlTop, ControlWidth, ControlHeight, CDate(strDeliveryDate), strDeliveryRef, ASN, IDX, _
                            CStr(ControlLowerTag), CStr(ControlUpperTag), myDic_Collection)
                    End If
                    myRow = myRow + 1
                Loop
            End If
        End If
    Next
    'Now use all of the Control Info based in the CONTROLS COLLECTION to populate ALL of the controls on the userform :
    Set varControlKey1 = Nothing
    TAGNumber = 0
        varControlKey1 = SearchText & "_" & CStr(TAGNumber)
        If InCollection("MISSING", ctrlCollection, varControlKey1) Then
'            Set ctrlProperty = ctrlCollection.Item(varControlKey1)
'            ControlName = ctrlProperty.ControlName
        End If
'    For Each varControlKey1 In ctrlCollection
'        Set ctrlProperty = varControlKey1
'        TagName = ctrlProperty.ControlTAG
'        ControlType = ctrlProperty.ControlType
'        ControlName = ctrlProperty.ControlName
'        ControlLowerTag = ctrlProperty.ControlStartTAG
'        ControlUpperTag = ctrlProperty.ControlEndTAG
'        ControlValue = ctrlProperty.ControlValue
'        ControlDate = ctrlProperty.ControlDate
'        ControlFieldname = ctrlProperty.ControlFieldname
'    Next
    
    
End Sub

Function IsArrayEmpty(arr As Variant, Optional MinLength As Integer = 0) As Boolean
  ' This function returns true if array is empty
  Dim l As Long

  On Error Resume Next
  l = Len(Join(arr))
  If l <= MinLength Then
    IsArrayEmpty = True
  Else
    IsArrayEmpty = False
  End If

  If Err.Number > 0 Then
      IsArrayEmpty = True
  End If

  On Error GoTo 0
End Function

Function Calc_Hours(DateStart As String, DateEnd As String, Optional ByRef strHours As String = "", Optional strMinutes As String = "") As Double
    Dim dtStartDate As Date
    Dim dtEndDate As Date
    Dim intHours As Integer
    Dim dblMins As Double
    Dim dblTime As Double
    Dim FinalHours As Long
    Dim dblHours As Double
    Dim strTime As String
    Dim FullStopPos As Integer
    
    dtStartDate = Convert_strDateToDate(DateStart, True)
    dtEndDate = Convert_strDateToDate(DateEnd, True)
    dblTime = 0#
    dblTime = CDbl(dtEndDate - dtStartDate)
    strTime = CStr(dblTime)
    FullStopPos = InStr(1, strTime, ".")
    strHours = Mid(strTime, 1, FullStopPos - 1)
    dblMins = dblTime - CInt(strHours)
    dblMins = dblMins * 60
    'may want answer in Hours and mins as a string ?
    'multiply the decimal part by 60 to get the minutes.
    'or multiply the whole by 60 to convert all into minutes and then divide by 60 etc.
    strMinutes = CStr(dblMins)
    
    Calc_Hours = dblTime

End Function

Function Convert_Hours_To_strHours(lngTime As Long) As String
    'Convert time into string - Hours and Minutes to display on form:
    Dim dblHours As Double
    Dim dblMins As Double
    Dim strFinalTime As String
    
    'work out hours as a string
    


End Function

Function GetTimes(WB As Workbook, Timesheet As String, RowNumber As Long, TotalCols As Long, ByVal StartTitleCol As Long, ByVal EndTitleCol As Long) As String()
    'Find row. Find Date Start and Date End Columns for particular operative of same Delivery Ref.
    Dim ThisWB As Workbook
    Dim Names() As String
    Dim StartTimes() As String
    Dim EndTimes() As String
    Dim OPIDX As Long
    Dim WholeRow As String
    Dim TitleRow As String
    Dim ColIDX As Long
    Dim Entry As String
    Dim Title As String
    Dim AllTimes() As String
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    WholeRow = ConsolidateRow(Timesheet, RowNumber, TotalCols, ";", " ", " ", ThisWB)
    TitleRow = ConsolidateRow(Timesheet, 1, TotalCols, ";", " ", " ", ThisWB)
    'DateStart = ThisWB.Worksheets(Timesheet).Cells(RowNumber, StartTitleCol+1).value
    'DateEnd = ThisWB.Worksheets(Timesheet).Cells(RowNumber, StartTitleCol+2).value
    ColIDX = StartTitleCol - 1 ' 0-zero based
    OPIDX = 0 'Array Element Index
    ReDim Names(1)
    ReDim StartTimes(1)
    ReDim EndTimes(1)
    ReDim AllTimes(1)
    Do While ColIDX <= EndTitleCol
        Entry = GetFieldValue(WholeRow, ColIDX, ";")
        Title = GetFieldValue(TitleRow, ColIDX, ";")
        If Len(Entry) > 0 And Not Asc(Entry) = 32 Then
            If InStr(1, UCase(Title), "NAME", vbTextCompare) > 0 Then
                ReDim Preserve Names(UBound(Names) + 1)
                ReDim Preserve StartTimes(UBound(StartTimes) + 1)
                ReDim Preserve EndTimes(UBound(EndTimes) + 1)
                ReDim Preserve AllTimes(UBound(AllTimes) + 1)
                Names(OPIDX) = Entry
                Entry = GetFieldValue(WholeRow, ColIDX + 1, ";")
                If Len(Entry) > 0 Then
                    StartTimes(OPIDX) = Entry
                Else
                    StartTimes(OPIDX) = "0"
                End If
                Entry = GetFieldValue(WholeRow, ColIDX + 2, ";")
                If Len(Entry) > 0 Then
                    EndTimes(OPIDX) = Entry
                Else
                    EndTimes(OPIDX) = "0"
                End If
                AllTimes(OPIDX) = Names(OPIDX) & "," & StartTimes(OPIDX) & "," & EndTimes(OPIDX)
                OPIDX = OPIDX + 1
                'ColIDX = ColIDX + 3
            Else
                ColIDX = ColIDX + 1
            End If
            
            
        End If
        
        ColIDX = ColIDX + 1
    Loop
    GetTimes = AllTimes

End Function

Function GetTimes_From_ACCESS(DBTable As String, AccessDBpath As String, SearchCriteria As String) As String()
    Dim Names() As String
    Dim StartTimes() As String
    Dim EndTimes() As String
    Dim OPIDX As Long
    Dim WholeRow As String
    Dim TitleRow As String
    Dim ColIDX As Long
    Dim Entry As String
    Dim Title As String
    Dim AllTimes() As String
    Dim dict_Times As New Scripting.Dictionary
    Dim LoadOK As Boolean
    Dim Fieldnames As String
    Dim FieldValues As String
    Dim strFLMName As String
    Dim strFLMStartTime As String
    Dim strFLMEndTime As String
    Dim OpName As String
    Dim OpStartTime As String
    Dim OpEndTime As String
    Dim SortFields As String
    Dim Reversed As Boolean
    
    'WholeRow = ConsolidateRow(Timesheet, RowNumber, TotalCols, ";", " ", " ", ThisWB)
    'TitleRow = ConsolidateRow(Timesheet, 1, TotalCols, ";", " ", " ", ThisWB)
    'DateStart = ThisWB.Worksheets(Timesheet).Cells(RowNumber, StartTitleCol+1).value
    'DateEnd = ThisWB.Worksheets(Timesheet).Cells(RowNumber, StartTitleCol+2).value
    'ColIDX = StartTitleCol - 1 ' 0-zero based
    OPIDX = 0 'Array Element Index
    ReDim Names(1)
    ReDim StartTimes(1)
    ReDim EndTimes(1)
    ReDim AllTimes(1)
    
    'Set dict_Times = CreateObject("Scripting.Dictionary") 'does not need NEW in front - this is LATE BINDING.
    dict_Times.RemoveAll
    SortFields = ""
    Reversed = False
    LoadOK = LoadAccessDBTable(DBTable, AccessDBpath, False, SearchCriteria, SortFields, Reversed, False, Fieldnames, FieldValues, dict_Times)
    'Now loading from the Database - means that the times MAY NOT be up-to-date - user not saved the data yet.
    'Loading the times from the Controls Collection - from current memory - from the userform as is.
    ' - allows for the user to experiment - put in different times - to work out the TOTAL working Hours and match with the calculated ones !
    ' - getting from the object means quicker access and more organised.
    'SO we need a function to retrieve the precise information from the Controls Collection OBJECT.
    'This function is retrieving from the database - maybe just to test the calculation or to LOAD the OBJECT from the Database ???
    'Not a good idea to load from the DATABASE into userforms or temp memory - bypassing the Control OBJECT - loose information and not recorded. scratchpad.
    'Do While ColIDX <= EndTitleCol
        Entry = GetFieldValue(WholeRow, ColIDX, ";")
        If Len(Entry) > 0 And Not Asc(Entry) = 32 Then
            If InStr(1, UCase(Title), "NAME", vbTextCompare) > 0 Then
                ReDim Preserve Names(UBound(Names) + 1)
                ReDim Preserve StartTimes(UBound(StartTimes) + 1)
                ReDim Preserve EndTimes(UBound(EndTimes) + 1)
                ReDim Preserve AllTimes(UBound(AllTimes) + 1)
                Names(OPIDX) = Entry
                Entry = GetFieldValue(WholeRow, ColIDX + 1, ";")
                If Len(Entry) > 0 Then
                    StartTimes(OPIDX) = Entry
                Else
                    StartTimes(OPIDX) = "0"
                End If
                Entry = GetFieldValue(WholeRow, ColIDX + 2, ";")
                If Len(Entry) > 0 Then
                    EndTimes(OPIDX) = Entry
                Else
                    EndTimes(OPIDX) = "0"
                End If
                AllTimes(OPIDX) = Names(OPIDX) & "," & StartTimes(OPIDX) & "," & EndTimes(OPIDX)
                OPIDX = OPIDX + 1
                'ColIDX = ColIDX + 3
            Else
                ColIDX = ColIDX + 1
            End If
            
            
        End If
        
        ColIDX = ColIDX + 1
    'Loop
    GetTimes_From_ACCESS = AllTimes
End Function



Function SearchRows(WB As Workbook, WorksheetName As String, SearchCol As Long, SearchText As String) As Long
    Dim LastRow As Long
    Dim RowIDX As Long
    Dim Rownum As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    SearchRows = 0
    LastRow = ThisWB.Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).row
    Rownum = 0
    If Len(SearchText) > 0 Then
        RowIDX = 1
        Do While RowIDX <= LastRow
            If UCase(ThisWB.Worksheets(WorksheetName).Cells(RowIDX, SearchCol).value) = UCase(SearchText) Then
                Rownum = RowIDX
                Exit Do
            End If
            RowIDX = RowIDX + 1
        Loop
    End If
    SearchRows = Rownum
End Function

Function ExtractRows(TargetWB As Workbook, DataWB As Workbook, WorksheetName As String, SearchCol As Long, DateSearchText As String, ByRef RowsExtracted As Long, ByRef CountDuplicates As Long, Optional ASNCol As Long = 0, Optional ASNSearchText As String = "", Optional DelRefCol As Long = 0, Optional DelRefText As String = "") As String()
    'EXTRACT all rows matching the SearchText
    'Use an array to extract or a script dictionary ?
    Dim LastRow As Long
    Dim RowIDX As Long
    Dim Rownum As Long
    Dim Temp() As String
    Dim ArrIDX As Long
    Dim WholeRow As String
    Dim LastColumn As Long
    Dim DateEntry As String
    Dim StartRow As Long
    Dim FindRange As Range
    Dim ASNNumEntry As String
    Dim CountInserted As Long
    Dim strProgress As String
    Dim strDate As String
    Dim dtDate As Date
    Dim DeliveryRef As String
    Dim SearchRow As String
   
        
    'TargetWB = WORKBOOK containing DAILY sheet 'now copied to LOCAL workbook.
    'DataWB = WORKBOOK containing GI DATA sheet - maybe needs to be copied locally too ?
        
    LastRow = TargetWB.Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).row + 1
    PercentDone = 0
    'FIND the DATE using the FIND command - to give the loop a head start and improve speed:
    'Set FindRange = Range(Cells(1, SearchCol), Cells(LastRow, SearchCol)).Find(DateSearchText, , xlValues, xlWhole, xlByColumns, xlNext, True)
    'Set FindRange = ActiveWorkbook.Worksheets(WorksheetName).Range("A2:CZ" & CStr(LastRow)).Find(DateSearchText, , xlValues, xlWhole, xlByColumns, xlNext, True)
    'If FindRange Is Nothing Then
        'NOTHING FOUND
        'MsgBox ("NOTHING FOUND")
    'Else
    '    MsgBox ("FOUND: " & FindRange.Address)
    'End If
    'ASSUME WorksheetName is DAILY.
    Rownum = 2
    ArrIDX = 0
    LastColumn = TargetWB.Worksheets(WorksheetName).Cells(2, Columns.Count).End(xlToLeft).Column
    'ReDim Temp(1)
    Application.DisplayAlerts = False
    Do While Rownum <= LastRow
        If Rownum > 0 And SearchCol > 0 Then
            DateEntry = TargetWB.Worksheets(WorksheetName).Cells(Rownum, SearchCol) 'from This source sheet
            ASNNumEntry = TargetWB.Worksheets(WorksheetName).Cells(Rownum, ASNCol) 'from this source sheet
            'WholeRow = ConsolidateRange(ActiveWorkbook.Sheets(WorksheetName).Range(Cells(Rownum, 1), Cells(Rownum, LastColumn)), ";", "", " ", "Err")
            
            WholeRow = ConsolidateRow(WorksheetName, Rownum, LastColumn, ";", " ", " ", TargetWB)
            If IsDate(GetFieldValue(WholeRow, SearchCol - 1, ";")) Then
                'Turn DateEntry (text) into a proper date. Then format back as string again.
                dtDate = CDate(GetFieldValue(WholeRow, SearchCol - 1, ";"))
                DateEntry = CStr(dtDate)
                'Worksheets(WorksheetName).Cells(Rownum, SearchCol).value = DateEntry
                'WholeRow = ConsolidateRow(WorksheetName, Rownum, LastColumn, ";", " ")
            End If
            
            If UCase(DateEntry) = UCase(DateSearchText) Then
                'WHAT IF THE USER IS TRYING TO IMPORT the same data thats already on Timesheet Records ????
                'If RowExistsOnSheet("GI DATA", DateEntry, SearchCol, ASNNumEntry, ASNCol) Then
                If DelRefCol > 0 Then
                    SearchRow = DateSearchText & ";" & ASNSearchText & ";" & DelRefText
                Else
                    
                End If
                
                'If RowExistsOnSheet(DataWB, "GI DATA", WholeRow, 0) Then
                '    CountDuplicates = CountDuplicates + 1
                    'DONT EXTRACT - already in target sheet
                'Else
                    'need to copy WHOLE row into the array:
                    'use consolidation:
                    'WholeRow = ConsolidateRange(ActiveWorkbook.Sheets(WorksheetName).Range(Cells(Rownum, 1), Cells(Rownum, LastColumn)), ";", "", " ", "Err")
                    SaveRow PassArray:=Temp, WholeRow:=WholeRow, ElementIDX:=ArrIDX
                    CountInserted = CountInserted + 1
                    ArrIDX = ArrIDX + 1
                'End If
                'ReDim Preserve Temp(UBound(Temp) + 1)
                
            Else
                'DATE NOT FOUND on current ROW of Daily sheet:
                
            End If
        End If
        DoEvents
        frmGI_TimesheetEntry2_1060x630.lblProgress.Width = CalcPercentage(PercentDone * 0.01, frmGI_TimesheetEntry2_1060x630.FrameProgress.Width, strProgress)
        frmGI_TimesheetEntry2_1060x630.lblProgress.Caption = strProgress
        PercentDone = CLng((Rownum / LastRow) * 100)
        Rownum = Rownum + 1
    Loop
    If SafeArrayGetDim(Temp) = 0 Then
        'TEMP is NOT DEFINED
    Else
        If Not IsEmpty(Temp) Then
            If UBound(Temp) > 0 Then
                RowsExtracted = UBound(Temp)
                ExtractRows = Temp
            Else
                RowsExtracted = 0
                Erase ExtractRows
            End If
        End If
    End If
    Application.DisplayAlerts = True
End Function

Function RowExistsOnSheet(WB As Workbook, SearchWorksheetName As String, SearchText1 As String, SearchColumn1 As Long, _
        Optional SearchText2 As String = "", Optional SearchColumn2 As Long = 0) As Boolean
    'OBJECTIVE to search the given sheet for a DATE And ASN no in the SAME row:
    Dim DicSearch As New Scripting.Dictionary
    Dim objKey1 As Variant
    Dim objKey2 As Variant
    Dim Key1 As String
    Dim Key2 As String
    Dim SearchKey As String
    Dim PassKey As String
    Dim NumItems As Long
    Dim RowIDX As Long
    Dim LastRow As Long
    Dim LastCol As Long
    Dim WholeRow As String
    Dim FoundKey As Boolean
    Dim DicItemCount As Long
    Dim CountRows As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    RowExistsOnSheet = False
    FoundKey = False
    Set DicSearch = CreateObject("Scripting.Dictionary")
    If Not MainGIModule_v1_1.sheetExists(SearchWorksheetName, ThisWB) Then
        MsgBox ("Sheet: " & SearchWorksheetName & " NOT exist in " & WB.Name)
        Exit Function
        
    End If
    LastCol = ThisWB.Worksheets(SearchWorksheetName).Cells(2, Columns.Count).End(xlToLeft).Column
    LastRow = ThisWB.Worksheets(SearchWorksheetName).Cells(Rows.Count, 1).End(xlUp).row
    If LastCol = 1 Then Exit Function
    RowIDX = 2
    DicItemCount = 0
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    'SAVE ALL the ROWS from the target sheet - GI DATE in a script dictionary as two keys combined - Date and ASN. OR pass WHOLEROW from Daily sheet.
    'Then loop through the dictionary keys and compare with the passed made up key from the two fields passed from the same row in the source sheet:
    ' which will be the Daily sheet. If there IS a match - set variable to TRUE and return the value back to the calling program.
    ' This will indicate that row DOES already contain that whole row of information and possibly has already been imported - so dont add it to the
    '   target sheet: GI DATA.
    Do While RowIDX <= LastRow
        '*************************************************
        '***************************************************** consolidate the row into one string:
        WholeRow = ConsolidateRange(ThisWB.Sheets(SearchWorksheetName).Range(Cells(RowIDX, 1), Cells(RowIDX, LastCol)), ";", "", " ", "Err")
        'WholeRow = ConsolidateRow(SearchWorksheetName, RowIDX, LastCol, ";", " ")
        If SearchColumn1 > 0 Then
            Key1 = GetFieldValue(WholeRow, SearchColumn1 - 1, ";")
            DicSearch(Key1) = WholeRow
        Else
            'Wholerow passed as key.
            DicSearch(WholeRow) = WholeRow 'records the whole row from the search spreadsheet - one being copied to.
            'the main sheet - Daily - one single row is passed from that in Searchtext1. Each row in target sheet is compared to the one row.
        
        End If
        If SearchColumn2 > 0 Then
            Key2 = GetFieldValue(WholeRow, SearchColumn2 - 1, ";")
            If Not DicSearch.Exists(Key1 & "_" & Key2) Then
                DicSearch(Key1 & "_" & Key2) = WholeRow
            End If
        End If
        
        RowIDX = RowIDX + 1
    Loop
    CountRows = 0
    For Each objKey1 In DicSearch.Keys
        If Len(objKey1) = 0 Then
            GoTo NEXT_objkey1
        End If
        SearchKey = objKey1
        If Len(SearchText2) > 0 Then
            PassKey = SearchText1 & "_" & SearchText2
        Else
            PassKey = SearchText1
        End If
        'SearchKey is from the Script Dictionary - SearchWorksheetName, Passkey is in the string variable passed in from other sheet:
        If UCase(SearchKey) = UCase(PassKey) Then
            FoundKey = True
            CountRows = CountRows + 1
            Exit For
        End If
        DicItemCount = DicItemCount + 1
NEXT_objkey1:
    Next
    RowExistsOnSheet = FoundKey
    Set DicSearch = Nothing
    
End Function

Sub SortSheet(WB As Workbook, WorksheetName As String, SortColumn1 As Long, Optional SortColumn2 As Long = 0, Optional REverseSort As Boolean = False)
    Dim LastRow As Long
    Dim LastCol As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
       Set ThisWB = ActiveWorkbook
    Else
       Set ThisWB = WB
    End If
    LastCol = ThisWB.Worksheets(WorksheetName).Cells(2, Columns.Count).End(xlToLeft).Column
    LastRow = ThisWB.Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).row

    ThisWB.Worksheets(WorksheetName).sort.SortFields.Clear
    If SortColumn1 > 0 Then
    
        If REverseSort = True Then
            ThisWB.Worksheets(WorksheetName).sort.SortFields.Add key:=Range( _
                Cells(2, SortColumn1), Cells(LastRow, SortColumn1)), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
                xlSortNormal
        Else
            ThisWB.Worksheets(WorksheetName).sort.SortFields.Add key:=Range( _
                Cells(2, SortColumn1), Cells(LastRow, SortColumn1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
        End If
    Else
        
    End If
    If SortColumn2 > 0 Then
        ThisWB.Worksheets(WorksheetName).sort.SortFields.Add key:=Range( _
            Cells(2, SortColumn2), Cells(LastRow, SortColumn2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
    End If
    With ThisWB.Worksheets(WorksheetName).sort
        .SetRange Range(Cells(1, 1), Cells(LastRow, LastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Function PopulateDropdowns(WorksheetName As String, PopCol As Long, Optional StartRow As Long = 0, Optional SORTIt As Boolean = True, Optional WB As Workbook) As String()
    Dim ListArr() As String
    Dim RowIDX As Long
    Dim LastRow As Long
    Dim ItemIDX As Long
    Dim Entry As String
    Dim Dic As New Scripting.Dictionary
    Dim objKey As Variant
    Dim NumItems As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    
    'Declare Scripting Dictionary to eliminate duplicate items:
    
    'Throws automation error if ThisWB = nothing !! - only if SET is not used.
    If Not MainGIModule_v1_1.sheetExists(WorksheetName, ThisWB) Then
        If MainGIModule_v1_1.sheetExists(WorksheetName, ThisWorkbook) Then
            Set ThisWB = ThisWorkbook
        Else
            Exit Function
        End If
        
    End If
    LastRow = ThisWB.Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).row
    If StartRow > 0 Then
        RowIDX = StartRow
    Else
        RowIDX = 2
    End If
    
    Set Dic = CreateObject("Scripting.Dictionary")
    
    ReDim ListArr(1)
    Do While RowIDX <= LastRow
        Entry = ThisWB.Worksheets(WorksheetName).Cells(RowIDX, PopCol).value
        If Not Dic.Exists(Entry) Then
            Dic(Entry) = Entry
            'Listarr(ItemIDX) = Entry
            'ItemIDX = ItemIDX + 1
            'ReDim Preserve Listarr(UBound(Listarr) + 1)
        End If
        
        RowIDX = RowIDX + 1
        
    Loop
    'SortDictionary(Dictionary_Object,SortByKey,Sort_Ascending)
    If SORTIt Then
        NumItems = SortDictionary(Dic, False, True)
    End If
    
    ItemIDX = 0
    For Each objKey In Dic.Keys
        If Len(objKey) = 0 Then
            GoTo NEXT_objkey
        End If
        Entry = CStr(objKey)
        ListArr(ItemIDX) = Entry
        ItemIDX = ItemIDX + 1
        ReDim Preserve ListArr(UBound(ListArr) + 1)
NEXT_objkey:
    Next
    Set Dic = Nothing
    PopulateDropdowns = ListArr
    
End Function

Function Get_DB_Table(ByVal AccessDBpath As String, SearchFieldname As String, Optional DBName As String = "") As String
'Need to get Fieldnames from EACH specified TABLE:
    Dim Fieldnames As String
    Dim Criteria As String
    Dim IDX As Long
    Dim FieldArr() As String
    Dim FieldName As String
    Dim DBTable As String
    
    Get_DB_Table = ""
    DBTable = "tblDeliveryInfo"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    DBTable = "tblLabourHours"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    DBTable = "tblOperatives"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    DBTable = "tblShortsAndExtraParts"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    DBTable = "tblSupplierCompliance"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    DBTable = "tblSecurity"
    If IsFieldIn(DBTable, AccessDBpath, SearchFieldname) Then
        Get_DB_Table = DBTable
        Exit Function
    End If
    
End Function

Function IsFieldIn(ByVal DBTable As String, ByVal DBPathWithDBName As String, SearchFieldname As String, Optional DBName As String = "", _
        Optional Delim As String = ",") As Boolean
    Dim Fieldnames As String
    Dim Criteria As String
    Dim IDX As Long
    Dim FieldArr() As String
    Dim FieldName As String
    
    IsFieldIn = False
    If Len(DBTable) = 0 Then
        MsgBox ("DBTable is empty")
        Exit Function
    End If
    Fieldnames = GetFieldnames_From_ACCESS(DBTable, DBPathWithDBName, Criteria)
    If Len(Fieldnames) > 0 Then
        If InStr(Fieldnames, ",") > 0 Then
            FieldArr = Split(Fieldnames, Delim)
            IDX = 0
            Do While IDX < UBound(FieldArr)
                FieldName = FieldArr(IDX)
                If UCase(FieldName) = UCase(SearchFieldname) Then
                    IsFieldIn = True
                End If
                IDX = IDX + 1
            Loop
        Else
            FieldName = Fieldnames
            If UCase(FieldName) = UCase(SearchFieldname) Then
                IsFieldIn = True
            End If
        End If
    Else
        'Fieldnames returns blank:
        
    End If



End Function

Function PopulateDropdowns_From_ACCESS(DBTable As String, AccessDBpath As String, FieldPos As Long, Optional DBName As String = "", _
    Optional Criteria As String = "", Optional SORTIt As Boolean = True, Optional WB As Workbook) As String()
    
    Dim ListArr() As String
    Dim RowIDX As Long
    Dim LastRow As Long
    Dim ItemIDX As Long
    Dim Entry As String
    Dim Dic As New Scripting.Dictionary
    Dim objKey As Variant
    Dim NumItems As Long
    Dim ThisWB As Workbook
    Dim TotalRecords As Long
    Dim strDBName As String
    Dim strMyPath As String
    Dim strDB As String
    Dim connDB As Object
    Dim ADOSET As Object
    Dim strSQL As String
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    TotalRecords = 0
    'Declare Scripting Dictionary to eliminate duplicate items:
    
    strDBName = DBName
    If Len(AccessDBpath) > 0 Then
        strMyPath = AccessDBpath
    Else
        strMyPath = ThisWorkbook.Path
    End If
    
    If Len(DBName) = 0 Then
        strDB = AccessDBpath
    Else
        strDB = strMyPath & "\" & strDBName
    End If
    If Len(DBTable) = 0 Then
        'MsgBox ("Please specify DB Table to load")
        Application.StatusBar = "No DB Table Specified"
        'usrHeatMapControlPanel.txtOutput.Text = "No DB Table Specified"
        Exit Function
    End If
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    If Len(strDB) > 0 Then
        Set connDB = New ADODB.Connection
        connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    Else
        Exit Function
    End If

'--------------
'OPEN RECORDSET, ACCESS RECORDS AND FIELDS

    'Set the ADO Recordset object:
    Set ADOSET = New ADODB.Recordset
    strSQL = "SELECT * FROM " & DBTable
    If Len(Criteria) > 0 Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
'--------------
    
    ADOSET.Open strSQL, connDB, adOpenStatic, adLockOptimistic, adCmdText

    TotalFields = ADOSET.Fields.Count
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.RemoveAll
    TotalRecords = ADOSET.RecordCount
    If TotalRecords > 0 Then
        ADOSET.MoveFirst
        
        Do While Not ADOSET.EOF
            'got invalid use of null here ??????????
            If Len(ADOSET.Fields(FieldPos).value) > 0 Then
                Entry = ADOSET.Fields(FieldPos).value
                If Not Dic.Exists(Entry) Then
                    Dic(Entry) = Entry
                End If
            End If
            ADOSET.MoveNext
        Loop
    End If
    'SortDictionary(Dictionary_Object,SortByKey,Sort_Ascending)
    If SORTIt Then
        NumItems = SortDictionary(Dic, False, True)
    End If
    
    ItemIDX = 0
    ReDim ListArr(1)
    For Each objKey In Dic.Keys
        If Len(objKey) = 0 Then
            GoTo NEXT_objkey_IN_ACCESS
        End If
        If Len(CStr(objKey)) > 0 Then
            Entry = CStr(objKey)
            ListArr(ItemIDX) = Entry
        End If
        ItemIDX = ItemIDX + 1
        ReDim Preserve ListArr(UBound(ListArr) + 1)
NEXT_objkey_IN_ACCESS:
    Next
    
    ADOSET.Close
    connDB.Close

    'destroy the variables
    Set ADOSET = Nothing
    Set connDB = Nothing
    
    Set Dic = Nothing
    PopulateDropdowns_From_ACCESS = ListArr
    
End Function

Sub SortComboBox(ByRef combo As ComboBox)
    Dim vItems As Variant
    Dim i As Long
    Dim j As Long
    Dim vTemp As Variant
    ' Put the items in a array
    vItems = combo.List
    ' Sort the array
    For i = LBound(vItems, 1) To UBound(vItems, 1) - 1
        For j = i + 1 To UBound(vItems, 1)
            If vItems(i, 0) > vItems(j, 0) Then
                vTemp = vItems(i, 0)
                vItems(i, 0) = vItems(j, 0)
                vItems(j, 0) = vTemp
            End If
        Next j
    Next i
    ' Clear the ComboBox
    combo.Clear
    ' Add the sorted array back to the ComboBox
    For i = LBound(vItems, 1) To UBound(vItems, 1)
        combo.AddItem vItems(i, 0)
    Next i
End Sub


Sub EnableDisableControls(EnableControls As Boolean, Optional TagLowRange As Long = 0, Optional TagUpperRange As Long = 0, Optional ControlType As String = "ALL")
    Dim CTRL As Control
    Dim myRow As Long
    Dim myCol As Long
    Dim myButton As String
    Dim myBtnNum As Long
    Dim Entry As String
    Dim FinalEntry As String
    Dim txtCtrl As TextBox
    Dim btnText As String
    Dim ButtonNumber As Long
    
    'myRow = GetNextAvailablerow("Timesheet Records", 2, 1) 'StartRow = 2 and check col = 1
    For Each CTRL In frmGI_TimesheetEntry2_1060x630.Controls
        myCol = 0
        myBtnNum = 0
        If TypeName(CTRL) = "TextBox" Then
            Entry = CTRL
            'txtCtrl = ctrl
            If UCase(ControlType) = "ALL" Or UCase(ControlType) = "TEXTBOX" Then
                If Len(CTRL.Tag) > 0 Then
                    If IsNumeric(CLng(CTRL.Tag)) Then
                        myCol = CLng(CTRL.Tag)
                        'Entry = txtCtrl.Text
                    
                    End If
                End If
            End If
        End If
        If TypeName(CTRL) = "ComboBox" Then
            Entry = CTRL
            'txtCtrl = ctrl
            If UCase(ControlType) = "ALL" Or UCase(ControlType) = "COMBOBOX" Then
                If Len(CTRL.Tag) > 0 Then
                    If IsNumeric(CLng(CTRL.Tag)) Then
                        myCol = CLng(CTRL.Tag)
                        'Entry = txtCtrl.Text
                        
                    End If
                End If
            End If
        End If
        If TypeName(CTRL) = "CommandButton" Then
            btnText = CTRL 'Button TEXT
            myCol = 0
            If Len(CTRL.Tag) > 0 Then
                If UCase(ControlType) = "ALL" Or UCase(ControlType) = "BUTTONS" Then
                    If UCase(Mid(CTRL.Tag, 1, 3)) = "BTN" Then
                        If IsNumeric(Mid(CTRL.Tag, 4, Len(CTRL.Tag))) Then
                            myBtnNum = CLng(Mid(CTRL.Tag, 4, Len(CTRL.Tag))) 'Prefixed by "btn"
                            If TagLowRange > 0 And myBtnNum <= TagUpperRange And myBtnNum >= TagLowRange Then
                                If EnableControls Then
                                    CTRL.Enabled = True
                                Else
                                    CTRL.Enabled = False
                                End If
                            End If
                            If TagLowRange = 0 And TagUpperRange = 0 Then
                                'DISABLE ALL CONTROLS except txtASNNo and txtDeliveryReference:
                                If EnableControls Then
                                    CTRL.Enabled = True
                                Else
                                    CTRL.Enabled = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If myCol > 0 Then
            If TagLowRange > 0 And myCol <= TagUpperRange And myCol >= TagLowRange Then
                If EnableControls Then
                    CTRL.Enabled = True
                Else
                    CTRL.Enabled = False
                End If
            End If
            If TagLowRange = 0 And TagUpperRange = 0 Then
                'DISABLE ALL CONTROLS except txtASNNo and txtDeliveryReference:
                If EnableControls Then
                    CTRL.Enabled = True
                Else
                    CTRL.Enabled = False
                End If
            End If
        End If
    Next

End Sub

Function SearchRows2(WB As Workbook, WorksheetName As String, ItemToFind As String, SearchColumn As Long) As Long
    Dim RowIDX As Long
    Dim LastRow As Long
    Dim sht As Worksheet
    Dim BlankRow As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
        
    
    Set sht = ThisWB.Sheets(WorksheetName)
    LastRow = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    BlankRow = sht.Cells.Find(" ", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    SearchRows2 = 0
    RowIDX = 2
    Do While RowIDX < LastRow
        If sht.Cells(RowIDX, SearchColumn).value = ItemToFind Then
            SearchRows2 = RowIDX
            Exit Do
        End If
        RowIDX = RowIDX + 1
    Loop
    
    
End Function

Sub DUMP_CONTROL_NAMES_AND_TAG_NUMBER_TO_SHEET(WB As Workbook, UserForm_Name As UserForm, DumpWorksheet As String, Optional RowStart As Long = 2, Optional ColStart As Long = 2)
    Dim CTRL As Control
    Dim ControlName As String
    Dim UserFormName As String
    Dim TAGNumber As String
    Dim RowIDX As Long
    Dim ItemIDX As Long
    Dim ThisWB As Workbook
    
    If WB Is Nothing Then
        Set ThisWB = ActiveWorkbook
    Else
        Set ThisWB = WB
    End If
    If Len(DumpWorksheet) = 0 Then
        MsgBox ("No SHEET SPECIFIED")
    End If
    UserFormName = UserForms(0).Name
    If Not sheetExists(DumpWorksheet) Then
        MsgBox ("Error: Sheet not found: " & DumpWorksheet)
    End If
    ItemIDX = 1
    RowIDX = RowStart
    For Each CTRL In UserForm_Name.Controls
        ControlName = CTRL.Name
        TAGNumber = CTRL.Tag
        ThisWB.Worksheets(DumpWorksheet).Cells(RowIDX, ColStart).value = "Control #" & CStr(ItemIDX)
        ThisWB.Worksheets(DumpWorksheet).Cells(RowIDX, ColStart + 1).value = ControlName
        ThisWB.Worksheets(DumpWorksheet).Cells(RowIDX, ColStart + 2).value = TAGNumber
        ItemIDX = ItemIDX + 1
        RowIDX = RowIDX + 1
    Next

End Sub

Sub Dump_Controls()
Attribute Dump_Controls.VB_Description = "Dump Userform control names to worksheet:\nControl Dump"
Attribute Dump_Controls.VB_ProcData.VB_Invoke_Func = "D\n14"
    Call DUMP_CONTROL_NAMES_AND_TAG_NUMBER_TO_SHEET(ThisWorkbook, frmGI_TimesheetEntry2_1060x630, "Control Dump", 2, 2)
    
    

End Sub

Public Function CalcPercentage(ByVal WhatPercent As Double, ByVal MaxLabelWidth As Double, ByRef NewCaption As String) As Integer
Dim NewWidth As Integer
Dim OutputPercent As Integer

    NewWidth = WhatPercent * MaxLabelWidth
    OutputPercent = WhatPercent * 100
    CalcPercentage = NewWidth
    NewCaption = CStr(OutputPercent) & "%"


End Function

Function ReadXML(XMLFilename As String, Optional RowsToRead As Integer = 12, Optional ColumnsToRead As Integer = 5) As Variant

    Dim FSO As Object
    Dim TS As Object
    Dim xmlArray() As Variant
    Dim xmlData() As Variant
    Dim Data1 As Variant
    Dim Data2 As Variant
    Dim Data3 As Variant
    Dim Data4 As Variant
    Dim Data5 As Variant
    Dim Temp As String
    Dim Names As Long
    Dim Update As Boolean
    Dim Name As String 'EmpNo
    Dim Alias As String
    Dim dtLastPWchange As Date
    Dim strLastPWchange As String
    Dim Description As String
    Dim Superiorgroup As String
    Dim DataIDX As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TS = FSO.OpenTextFile(XMLFilename, 1, False, -2)
    ReDim xmlData(RowsToRead)
    Names = 1
    Do Until TS.AtEndOfStream
        ReDim Preserve xmlArray(1 To ColumnsToRead, 1 To Names)
        Update = False
        Name = ""
        Alias = ""
        strLastPWchange = ""
        Description = ""
        Superiorgroup = ""
        DataIDX = 1
        Do While DataIDX <= RowsToRead
            'Data = TS.ReadLine
            xmlData(DataIDX) = TS.ReadLine
        
            If xmlData(DataIDX) Like "*<name>*" Then
                Temp = Replace(xmlData(DataIDX), "<name>", "")
                Name = Replace(Temp, "</name>", "")
                
                If IsNumeric(Name) Then
                    Name = Val(Name)
                Else
                    Name = Trim(Name)
                End If
                If UCase(Name) = UCase("reporting") Or Mid(Name, 1, 1) = "#" Then
                    Update = False
                    Exit Do
                Else
                    Update = True
                End If
            ElseIf xmlData(DataIDX) Like "*<alias>*" Then
                Temp = Replace(xmlData(DataIDX), "<alias>", "")
                Alias = Replace(Temp, "</alias>", "")
                
                If IsNumeric(Alias) Then
                    Alias = Val(Alias)
                Else
                    Alias = Trim(Alias)
                End If
                If Len(Name) = 0 Or Mid(Alias, 1, 2) = "//" Then
                    Update = False
                    Exit Do
                Else
                    Update = True
                End If
            ElseIf xmlData(DataIDX) Like "*<lastpasswordchange>*" Then
                Temp = Replace(xmlData(DataIDX), "<lastpasswordchange>", "")
                strLastPWchange = Replace(Temp, "</lastpasswordchange>", "")
                
                If IsNumeric(strLastPWchange) Then
                    strLastPWchange = Val(strLastPWchange)
                Else
                    strLastPWchange = Trim(strLastPWchange)
                End If
                If Len(Name) = 0 Then
                    Update = False
                    Exit Do
                Else
                    Update = True
                End If
            ElseIf xmlData(DataIDX) Like "*<description>*" Then
                Temp = Replace(xmlData(DataIDX), "<description>", "")
                Description = Replace(Temp, "</description>", "")
                
                If IsNumeric(Description) Then
                    Description = Val(Description)
                Else
                    Description = Trim(Description)
                End If
                If Len(Name) = 0 Then
                    Update = False
                    Exit Do
                Else
                    Update = True
                End If
            ElseIf xmlData(DataIDX) Like "*<belongsto>*" Then
                Temp = Replace(xmlData(DataIDX), "<superiorgroup>", "")
                Superiorgroup = Replace(Temp, "</superiorgroup>", "")
                
                If IsNumeric(Superiorgroup) Then
                    Superiorgroup = Val(Superiorgroup)
                Else
                    Superiorgroup = Trim(Superiorgroup)
                End If
                If Len(Name) = 0 Then
                    Update = False
                    Exit Do
                Else
                    Update = True
                End If
            End If
            DataIDX = DataIDX + 1
        Loop
        If Update Then
            xmlArray(1, Names) = Name
            xmlArray(2, Names) = Alias
            xmlArray(3, Names) = strLastPWchange
            xmlArray(4, Names) = Description
            xmlArray(5, Names) = Superiorgroup
            Names = Names + 1
        End If
        
    Loop
    
    TS.Close
    
    'Write the array to the active worksheet, starting at A1
    ReadXML = xmlArray
    
End Function

