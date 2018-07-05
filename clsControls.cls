VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cbTimeStartEvent As MSForms.CommandButton
Attribute cbTimeStartEvent.VB_VarHelpID = -1
Public WithEvents cbTimeEndEvent As MSForms.CommandButton
Attribute cbTimeEndEvent.VB_VarHelpID = -1
Public WithEvents comboAfterUpdate As MSForms.ComboBox
Attribute comboAfterUpdate.VB_VarHelpID = -1
Public WithEvents TxtBoxAfterUpdate As MSForms.TextBox
Attribute TxtBoxAfterUpdate.VB_VarHelpID = -1
Private txtboxNameOfChangedValue As String
Private comboNameOfChangedValue As String
Private txtboxTagOfChangedValue As String
Private comboTagOfChangedValue As String
Private txtboxChangedValue As String
Private comboChangedValue As String
Private CtrlID As String 'Delivery Date _ Delivery Ref _ TAG Number
Private CtrlAltID As String 'Delivery Date _ Delivery Ref _ Control Name
Private CtrlASNID As String 'Delivery Date _ ASN _ TAG Number
Private CtrlASNAltID As String 'Delivery Date _ ASN _ Control Name
Private CTRL As Control
Private CtrlName As String
Private CtrlTag As String
Private CtrlDate As Date
Private CtrlLeft As Integer
Private CtrlTop As Integer
Private CtrlWidth As Integer
Private CtrlHeight As Integer
Private CtrlType As String
Private CtrlDeliveryDate As Date
Private CtrlDeliveryRef As String
Private CtrlASN As String
Private CtrlObjNumber As Long 'Count Number of Objects within one FRAME control = Number of CHILD RECORDS in Database
Private CtrlValue As String
Private CtrlList As Variant
Private CtrlStartTag As String
Private CtrlEndTag As String
Private CtrlFrameRowNumber As Long
Private CtrlTotalRows As Long
Private CtrlFLMName As String
Private CtrlFLMStartDateTime As Date
Private CtrlFLMEndDateTime As Date
Private CtrlOpName As String
Private CtrlOpActivity As String
Private CtrlOpStartDateTime As Date
Private CtrlOpEndDateTime As Date
Private CtrlPartNo As String
Private CtrlQty As Long
Private CtrlExtraShort As String
Private CtrlDBTable As String
Private CtrlFieldName As String
Private CtrlLastSaved As Date
Private CtrlBackColour As Long
Private CtrlLeftMargin As Boolean
Private CtrlErrMessage As String



Public Property Get ControlID() As String
    ControlID = CtrlID
End Property


Public Property Let ControlID(value As String)
    CtrlID = value
End Property

Public Property Get ControlAltID() As String
    ControlAltID = CtrlAltID
End Property


Public Property Let ControlAltID(value As String)
    CtrlAltID = value
End Property

Public Property Get ControlASNID() As String
    ControlASNID = CtrlASNID
End Property


Public Property Let ControlASNID(value As String)
    CtrlASNID = value
End Property

Public Property Get ControlASNAltID() As String
    ControlASNAltID = CtrlASNAltID
End Property


Public Property Let ControlASNAltID(value As String)
    CtrlASNAltID = value
End Property

Public Property Get TheControl() As Control
    Set TheControl = CTRL
End Property


Public Property Let TheControl(value As Control)
    Set CTRL = value
End Property

Public Property Get ControlName() As String
    ControlName = CtrlName
End Property


Public Property Let ControlName(value As String)
    CtrlName = value
End Property

Public Property Get ControlTAG() As String
    ControlTAG = CtrlTag
End Property


Public Property Let ControlTAG(value As String)
    CtrlTag = value
End Property

Public Property Get ControlDate() As Date
    ControlDate = CtrlDate
End Property


Public Property Let ControlDate(value As Date)
    CtrlDate = value
End Property
Public Property Get ControlLeftPos() As Integer
    ControlLeftPos = CtrlLeft
End Property


Public Property Let ControlLeftPos(value As Integer)
    CtrlLeft = value
End Property
Public Property Get ControlTopPos() As Integer
    ControlTopPos = CtrlTop
End Property


Public Property Let ControlTopPos(value As Integer)
    CtrlTop = value
End Property
Public Property Get ControlWidthPos() As Integer
    ControlWidthPos = CtrlWidth
End Property


Public Property Let ControlWidthPos(value As Integer)
    CtrlWidth = value
End Property
Public Property Get ControlHeightPos() As Integer
    ControlHeightPos = CtrlHeight
End Property


Public Property Let ControlHeightPos(value As Integer)
    CtrlHeight = value
End Property

Public Property Get ControlType() As String
    ControlType = CtrlType
End Property

Public Property Let ControlType(value As String)
    CtrlType = value
End Property
Public Property Get ControlDeliveryDate() As Date
    ControlDeliveryDate = CtrlDeliveryDate
End Property


Public Property Let ControlDeliveryDate(value As Date)
    CtrlDeliveryDate = value
End Property
Public Property Get ControlDeliveryRef() As String
    ControlDeliveryRef = CtrlDeliveryRef
End Property


Public Property Let ControlDeliveryRef(value As String)
    CtrlDeliveryRef = value
End Property

'CtrlASN

Public Property Get ControlASNNumber() As String
    ControlASNNumber = CtrlASN
End Property


Public Property Let ControlASNNumber(value As String)
    CtrlASN = value
End Property

Public Property Get ControlObjNumber() As Long
    ControlObjNumber = CtrlObjNumber
End Property


Public Property Let ControlObjNumber(value As Long)
    CtrlObjNumber = value
End Property


Public Property Get ControlValue() As String
    ControlValue = CtrlValue
End Property


Public Property Let ControlValue(value As String)
    CtrlValue = value
End Property

Public Property Get ControlList() As Variant
    ControlList = CtrlList
End Property


Public Property Let ControlList(value As Variant)
    CtrlList = value
End Property


Public Property Get ControlStartTAG() As String
    ControlStartTAG = CtrlStartTag
End Property


Public Property Let ControlStartTAG(value As String)
    CtrlStartTag = value
End Property

Public Property Get ControlEndTAG() As String
    ControlEndTAG = CtrlEndTag
End Property


Public Property Let ControlEndTAG(value As String)
    CtrlEndTag = value
End Property

Public Property Get ControlRowNumber() As Long
    ControlRowNumber = CtrlFrameRowNumber
End Property


Public Property Let ControlRowNumber(value As Long)
    CtrlFrameRowNumber = value
End Property

Public Property Get ControlTotalRows() As Long
    ControlTotalRows = CtrlTotalRows
End Property


Public Property Let ControlTotalRows(value As Long)
    CtrlTotalRows = value
End Property

'CtrlBackColour

Public Property Get ControlBACKCOLOUR() As Long
    ControlBACKCOLOUR = CtrlBackColour
End Property


Public Property Let ControlBACKCOLOUR(value As Long)
    CtrlBackColour = value
End Property

'CtrlLeftMargin

Public Property Get ControlLeftMargin() As Boolean
    ControlLeftMargin = CtrlLeftMargin
End Property


Public Property Let ControlLeftMargin(value As Boolean)
    CtrlLeftMargin = value
End Property


'Private CtrlOpName As String
'Private CtrlOpActivity As String
'Private CtrlOpStartDateTime As Date
'Private CtrlOpEndDateTime As Date
'Private CtrlPartNo As String
'Private CtrlQty As Long
'Private CtrlExtraShort As String

Public Property Get ControlFLMName() As String
    ControlFLMName = CtrlFLMName
End Property


Public Property Let ControlFLMName(value As String)
    CtrlFLMName = value
End Property

Public Property Get ControlFLMStartDateTime() As Date
    ControlFLMStartDateTime = CtrlFLMStartDateTime
End Property


Public Property Let ControlFLMStartDateTime(value As Date)
    CtrlFLMStartDateTime = value
End Property

Public Property Get ControlFLMEndDateTime() As Date
    ControlFLMEndDateTime = CtrlFLMEndDateTime
End Property


Public Property Let ControlFLMEndDateTime(value As Date)
    CtrlFLMEndDateTime = value
End Property

Public Property Get ControlOpName() As String
    ControlOpName = CtrlOpName
End Property


Public Property Let ControlOpName(value As String)
    CtrlOpName = value
End Property

Public Property Get ControlOpActivity() As String
    ControlOpActivity = CtrlOpActivity
End Property


Public Property Let ControlOpActivity(value As String)
    CtrlOpActivity = value
End Property

Public Property Get ControlOpStartDateTime() As Date
    ControlOpStartDateTime = CtrlOpStartDateTime
End Property


Public Property Let ControlOpStartDateTime(value As Date)
    CtrlOpStartDateTime = value
End Property

Public Property Get ControlOpEndDateTime() As Date
    ControlOpEndDateTime = CtrlOpEndDateTime
End Property


Public Property Let ControlOpEndDateTime(value As Date)
    CtrlOpEndDateTime = value
End Property

Public Property Get ControlPartNo() As String
    ControlPartNo = CtrlPartNo
End Property


Public Property Let ControlPartNo(value As String)
    CtrlPartNo = value
End Property

Public Property Get ControlQty() As Long
    ControlQty = CtrlQty
End Property


Public Property Let ControlQty(value As Long)
    CtrlQty = value
End Property

Public Property Get ControlExtraShort() As String
    ControlExtraShort = CtrlExtraShort
End Property


Public Property Let ControlExtraShort(value As String)
    CtrlExtraShort = value
End Property

Public Property Get ControlDBTable() As String
    ControlDBTable = CtrlDBTable
End Property


Public Property Let ControlDBTable(value As String)
    CtrlDBTable = value
End Property

Public Property Get ControlFieldname() As String
    ControlFieldname = CtrlFieldName
End Property


Public Property Let ControlFieldname(value As String)
    CtrlFieldName = value
End Property

Public Property Get ControlLastSaved() As Date
    ControlLastSaved = CtrlLastSaved
End Property


Public Property Let ControlLastSaved(value As Date)
    CtrlLastSaved = value
End Property

'Events:

Public Property Get ComboBoxName() As String
    ComboBoxName = comboNameOfChangedValue
End Property


Public Property Let ComboBoxName(value As String)
    comboNameOfChangedValue = value
End Property

Public Property Get TextBoxName() As String
    TextBoxName = txtboxNameOfChangedValue
End Property


Public Property Let TextBoxName(value As String)
    txtboxNameOfChangedValue = value
End Property


Public Property Get ComboTag() As String
    ComboTag = comboTagOfChangedValue
End Property


Public Property Let ComboTag(value As String)
    comboTagOfChangedValue = value
End Property

Public Property Get TextBoxTag() As String
    TextBoxTag = txtboxTagOfChangedValue
End Property


Public Property Let TextBoxTag(value As String)
    txtboxTagOfChangedValue = value
End Property

Public Property Get ComboValue() As String
    ComboValue = comboChangedValue
End Property


Public Property Let ComboValue(value As String)
    comboChangedValue = value
End Property

Public Property Get TextBoxValue() As String
    TextBoxValue = txtboxChangedValue
End Property


Public Property Let TextBoxValue(value As String)
    txtboxChangedValue = value
End Property


Public Property Get ComboBoxChanged() As String
    ComboBoxChanged = comboTagOfChangedValue
End Property


Public Property Let ComboBoxChanged(value As String)
    comboTagOfChangedValue = value
End Property

Public Property Get TextBoxChanged() As String
    TextBoxChanged = txtboxTagOfChangedValue
End Property


Public Property Let TextBoxChanged(value As String)
    txtboxTagOfChangedValue = value
End Property

'CtrlErrMessage

Public Property Get ErrMessage() As String
    ErrMessage = CtrlErrMessage
End Property


Public Property Let ErrMessage(value As String)
    CtrlErrMessage = value
End Property


Private Sub cbTimeStartEvent_click()
    'Insert current Date and time into related TEXT BOX:
    'Need to get tag of CB Text button:
    Dim TAGNumber As String
    Dim TimeControl As Control
    Dim SearchControlName As String
    Dim TimeOut As Date
    'btnOperativeTimeStart1
        
    SearchControlName = Replace(cbTimeStartEvent.Name, "btn", "txt")
    Set TimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", SearchControlName)
    If TimeControl Is Nothing Then
        'PROBLEM !
        MsgBox ("Could not find Control")
    Else
        TAGNumber = TimeControl.Tag
        TimeOut = Now()
        Me.ControlOpStartDateTime = TimeOut
        'Call InsertTimeIntoControl(0, SearchControl.Name, TimeOut, True, 400)
        TimeControl.Text = TimeOut
    End If
    
End Sub

Private Sub cbTimeEndEvent_click()
    'Insert current Date and time into related TEXT BOX:
    'Need to get tag of CB Text button:
    'btnOperativeTimeEnd1
    '17,3 to number. use length. OpNumber = len(cbTimeEndEvent.name)-20
    Dim TAGNumber As String
    Dim TimeControl As Control
    Dim SearchControlName As String
    Dim TimeOut As Date
    
    SearchControlName = Replace(cbTimeEndEvent.Name, "btn", "txt")
    Set TimeControl = FindFormControl(frmGI_TimesheetEntry2_1060x630, "TEXTBOX", "", SearchControlName)
    If TimeControl Is Nothing Then
        'PROBLEM !
        MsgBox ("Could not find Control")
    Else
        TAGNumber = TimeControl.Tag
        TimeOut = Now()
        Me.ControlOpEndDateTime = TimeOut
        'Call InsertTimeIntoControl(0, SearchControl.Name, TimeOut, False, 400)
        TimeControl.Text = TimeOut
    End If
    
End Sub



Private Sub comboAfterUpdate_Change()
    Dim CollectionKey As Variant
    Dim FoundControl As Boolean
    Dim NewValue As Variant
    Dim ErrMessage As String
    Dim clsControl As clsControls
    Dim ReturnValue As Variant
    Dim NewFieldname As String
    
    Me.ComboTag = comboAfterUpdate.Tag
    Me.ComboBoxName = comboAfterUpdate.Name
    Me.ComboValue = comboAfterUpdate.value
    
    NewValue = Me.ComboValue
    CollectionKey = Me.ControlDeliveryDate & "_" & Me.ControlDeliveryRef & "_" & Me.ComboTag
    
    Set ctrlCollection = MainGIModule_v1_1.UpdateCollection(ctrlCollection, CollectionKey, NewValue)
    
    'FoundControl = ReturnControlInfo(CollectionKey, NewValue, Me.ControlDeliveryDate, Me.ControlDeliveryRef, "", Me.ComboTag, "", "TAG", "", _
    '    "ControlValue", ReturnValue, ErrMessage, clsControl)
    'If FoundControl Then
    '    Set ctrlCollection.Item(CollectionKey) = clsControl
        
    'End If
    
    
    
    If Len(comboAfterUpdate.value) > 0 Then
        'txtFurtherComments
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastTagChanged", Me.ComboTag)
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastControlChanged", Me.ComboBoxName)
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastChangedValue", Me.ComboValue)
    End If
End Sub

Private Sub TxtBoxAfterUpdate_Change()
    Dim CollectionKey As Variant
    Dim FoundControl As Boolean
    Dim NewValue As Variant
    Dim ErrMessage As String
    Dim clsControl As clsControls
    Dim ReturnValue As Variant
    Dim NewFieldname As String
    
    Me.TextBoxName = TxtBoxAfterUpdate.Name
    Me.TextBoxTag = TxtBoxAfterUpdate.Tag
    Me.TextBoxValue = TxtBoxAfterUpdate.value
    
    NewValue = Me.TextBoxValue
    CollectionKey = Me.ControlDeliveryDate & "_" & Me.ControlDeliveryRef & "_" & Me.TextBoxTag
    'FoundControl = ReturnControlInfo(CollectionKey, NewValue, Me.ControlDeliveryDate, Me.ControlDeliveryRef, "", Me.ComboTag, "", "TAG", "", _
        "ControlValue", ReturnValue, ErrMessage, clsControl)
    'If FoundControl Then
        'Set ctrlCollection.Item(CollectionKey).ControlValue = clsControl.ControlValue
        'Set ctrlCollection(CollectionKey) = clsControl.ControlValue
    'End If
    
    'Me.UpdateCollection(ctrlCollection, CollectionKey, NewValue)
    Set ctrlCollection = MainGIModule_v1_1.UpdateCollection(ctrlCollection, CollectionKey, NewValue)
    
    If Len(TxtBoxAfterUpdate.value) > 0 Then
        'txtFurtherComments
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastTagChanged", Me.TextBoxTag)
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastControlChanged", Me.TextBoxName)
        Call InsertValueIntoControl("TEXTBOX", 0, "txtLastChangedValue", Me.TextBoxValue)
        'me.ControlOpStartDateTime
        'UPDATE the Collection - but can only do it internally
        'so the value that has changed is in me.txtBoxValue
        ' the control that has the change is in me.TextBoxName
        ' the TAG of the Control that has changed is in me.TextBoxTag
        
    End If
End Sub

Public Function UpdateCollection(coll As Collection, varkey As Variant, ValueToChange As Variant, _
    Optional DeliveryDate As Date, Optional ASNNumber As String = "", Optional DeliveryRef As String, Optional TAGNumber As String) As Collection
    Dim NewKey As Variant
    Dim ProcedureName As String
    Dim NewCollection As Collection
    Dim tempProperty As clsControls
    
    ProcedureName = "cls UpdateCollection"
    On Error GoTo Err_UpdateCollection
    
    If Len(CStr(DeliveryDate)) > 0 Then
        If Len(DeliveryRef) > 0 Then
            NewKey = DeliveryDate & "_" & DeliveryRef & "_" & TAGNumber
        End If
    ElseIf Len(ASNNumber) > 0 Then
        If Len(DeliveryRef) > 0 Then
            NewKey = DeliveryDate & "_" & ASNNumber & "_" & TAGNumber
        End If
    Else
        NewKey = varkey
    End If
    Set NewCollection = coll
    Set tempProperty = NewCollection.Item(NewKey)
    tempProperty.ControlValue = ValueToChange
    NewCollection.Remove NewKey
    NewCollection.Add tempProperty, NewKey
    'Me.TheControl.Text = ValueToChange
    Set UpdateCollection = NewCollection

Exit Function

Err_UpdateCollection:

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
End Function

