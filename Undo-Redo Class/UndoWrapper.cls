VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' **********************************************************************************************
' Class     : UndoWrapper
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Wrapper class for the Undo Class.  This class contains all the methods required
'             for Undo and Redo functionality on a form.
'             The controls that can be undone are: -
'               * Textbox
'               * Listbox
'               * Combobox
'               * Optionbutton
'               * Checkbox
' **********************************************************************************************

' Enum listing the valid types for Undo/Redo
Private Enum ControlTypes
  ctlComboBox
  ctlListBox
  ctlTextBox
  ctlOptionButton
  ctlCheckbox
End Enum

Public colUndoList As New Collection      ' Collection of Undo elements
Public iUndoIndex As Integer              ' Current Undo Index
Private CurControl As Undo                ' Current Undo Element
Private bDoneRedo As Boolean              ' Checks if a Redo has been performed
Private MenuUndo As Menu                  ' A pointer to the Undo Menu element
Private MenuRedo As Menu                  ' A pointer to the Redo Menu element
Private LocalForm As Form                 ' A pointer to the calling form
Private bDisallowClear As Boolean         ' Prevents collection clearing while a undo/redo is in progress

' **********************************************************************************************
' Function  : ClearHistoryTail
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Clear all elements from the history list after the current element
' **********************************************************************************************
Public Sub ClearHistoryTail()
  Dim iIndex As Integer
    If Not bDisallowClear Then
    If iUndoIndex < colUndoList.Count Then
      For iIndex = colUndoList.Count To iUndoIndex + 1 Step -1
        colUndoList.Remove (iIndex)
      Next
      iUndoIndex = colUndoList.Count
    End If
  End If
  If Not MenuUndo Is Nothing Then
    MenuUndo.Caption = "Undo"
    MenuUndo.Enabled = True
  End If
End Sub

' **********************************************************************************************
' Function  : ClearAllHistory
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Clear all elements from the history list
' **********************************************************************************************
Public Sub ClearAllHistory()
  Dim iIndex As Integer
  
  If Not bDisallowClear Then
    For iIndex = colUndoList.Count To 1 Step -1
      colUndoList.Remove (iIndex)
    Next
    iUndoIndex = 0
  End If
  Set CurControl = Nothing
End Sub

' **********************************************************************************************
' Function  : AddUndoHistory
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Insert an item into the history list
' **********************************************************************************************
Public Sub AddUndoHistory(ControlName As String, Optional Index As Integer = -1)
  Dim TempUndo As New Undo
  Dim NewUndo As New Undo
  Dim CheckValue As Variant
  Dim ControlType As ControlTypes
  Dim CurControlType As ControlTypes
  Dim CurControlIndex As Integer
  Dim LocalControl As Control
  
  ' Set the LocalControl pointer.  This points to a single instance of a control
  If Index > -1 Then
    Set LocalControl = LocalForm.Controls(ControlName, Index)
  Else
    Set LocalControl = LocalForm.Controls(ControlName)
  End If
  
  ' Set the ControlType variable depending on the type of the control
  If TypeOf LocalControl Is ComboBox Then
    If LocalControl.Style = 2 Then
      ControlType = ctlListBox
    Else
      ControlType = ctlComboBox
    End If
  ElseIf TypeOf LocalControl Is ListBox Then
    ControlType = ctlListBox
  ElseIf TypeOf LocalControl Is OptionButton Then
    ControlType = ctlOptionButton
  ElseIf TypeOf LocalControl Is CheckBox Then
    ControlType = ctlCheckbox
  Else
    ControlType = ctlTextBox
  End If
  
  ' Set the CurControlIndex.  If a control is not part of a control array
  ' this section will generate an error which is trapped to determine
  ' whether or not the control IS part of a control array
  On Local Error Resume Next
  CurControlIndex = CurControl.Field.Index
  If Err.Number = 343 Then ' Object not an array
    CurControlIndex = -1
  End If
  On Local Error GoTo 0
    
  ' This code is only executed when a control was previously selected.  It determines
  ' whether or not the control's value has changed and if so, inserts it into the
  ' colUndoList Collection
  If Not CurControl Is Nothing Then
    ' set the CurControlType variable to the Type of control
    If TypeOf CurControl.Field Is ComboBox Then
      If CurControl.Field.Style = 2 Then
        CurControlType = ctlListBox
      Else
        CurControlType = ctlComboBox
      End If
    ElseIf TypeOf CurControl.Field Is ListBox Then
      CurControlType = ctlListBox
    ElseIf TypeOf CurControl.Field Is OptionButton Then
      CurControlType = ctlOptionButton
    ElseIf TypeOf CurControl.Field Is CheckBox Then
      CurControlType = ctlCheckbox
    Else
      CurControlType = ctlTextBox
    End If
    
    ' Set the CheckValue variable.  This is the value that will be compared against the
    ' previous value of the control.  It's type is determined by the type of the control
    If CurControl.Field.Name <> ControlName _
    Or (CurControl.Field.Name = ControlName And CurControlIndex <> Index) Then
      If CurControlType = ctlListBox Then
        CheckValue = CurControl.Field.ListIndex
      ElseIf CurControlType = ctlOptionButton Or CurControlType = ctlCheckbox Then
        CheckValue = CurControl.Field.Value
      Else
        CheckValue = CurControl.Field.Text
      End If
      
      ' If the controls value has changed, insert a new undo record into the collection
      If CurControl.PrevVal <> CheckValue And bDoneRedo = False Then
        Set NewUndo.Field = CurControl.Field
        NewUndo.PrevVal = CurControl.PrevVal
        NewUndo.NewVal = CheckValue
        colUndoList.Add NewUndo
        iUndoIndex = iUndoIndex + 1
      End If
    End If
  End If
  
  ' TempUndo holds the values of the currently selected control
  Set TempUndo.Field = LocalControl

  If ControlType = ctlListBox Then
    TempUndo.PrevVal = TempUndo.Field.ListIndex
  ElseIf ControlType = ctlOptionButton Or ControlType = ctlCheckbox Then
    TempUndo.PrevVal = TempUndo.Field.Value
  Else
    TempUndo.PrevVal = TempUndo.Field.Text
  End If
  Set CurControl = TempUndo

  If Not MenuUndo Is Nothing Then
    MenuUndo.Caption = "Undo"
    MenuUndo.Enabled = True
  End If

End Sub

' **********************************************************************************************
' Function  : UndoEdit
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Undoes the current item in the Undo List
' **********************************************************************************************
Public Sub UndoEdit()
  
  On Local Error GoTo ErrorHandler
  
  bDisallowClear = True
  
  ' Only execute this code if elements exist
  If iUndoIndex > 0 Then
    ' If there is an open CurControl (ie. No NewVal or the value has changed), insert it before
    ' we start.  This covers occasions when a field is changed but the user clicks undo before
    ' moving to a new field
    If TypeOf CurControl.Field Is ComboBox Then
      If CurControl.Field.Style = 2 Then
        If IsEmpty(CurControl.NewVal) And CurControl.PrevVal <> CurControl.Field.ListIndex Then
          CurControl.NewVal = CurControl.Field.ListIndex
          colUndoList.Add CurControl
          iUndoIndex = iUndoIndex + 1
        End If
      Else
        If IsEmpty(CurControl.NewVal) And CurControl.PrevVal <> CurControl.Field.Text Then
          CurControl.NewVal = CurControl.Field.Text
          colUndoList.Add CurControl
          iUndoIndex = iUndoIndex + 1
        End If
      End If
    ElseIf TypeOf CurControl.Field Is ListBox Then
      If IsEmpty(CurControl.NewVal) And CurControl.PrevVal <> CurControl.Field.ListIndex Then
        CurControl.NewVal = CurControl.Field.ListIndex
        colUndoList.Add CurControl
        iUndoIndex = iUndoIndex + 1
      End If
    ElseIf TypeOf CurControl.Field Is OptionButton Or TypeOf CurControl.Field Is CheckBox Then
      If IsEmpty(CurControl.NewVal) And CurControl.PrevVal <> CurControl.Field.Value Then
        CurControl.NewVal = CurControl.Field.Value
        colUndoList.Add CurControl
        iUndoIndex = iUndoIndex + 1
      End If
    Else
      If IsEmpty(CurControl.NewVal) And CurControl.PrevVal <> CurControl.Field.Text Then
        CurControl.NewVal = CurControl.Field.Text
        colUndoList.Add CurControl
        iUndoIndex = iUndoIndex + 1
      End If
    End If
    
    ' Restore the field value, depending on the control type
    If TypeOf colUndoList(iUndoIndex).Field Is ComboBox Then
      If colUndoList(iUndoIndex).Field.Style = 2 Then
        colUndoList(iUndoIndex).Field.ListIndex = colUndoList(iUndoIndex).PrevVal
      Else
        colUndoList(iUndoIndex).Field.Text = colUndoList(iUndoIndex).PrevVal
      End If
    ElseIf TypeOf colUndoList(iUndoIndex).Field Is ListBox Then
      colUndoList(iUndoIndex).Field.ListIndex = colUndoList(iUndoIndex).PrevVal
    ElseIf TypeOf colUndoList(iUndoIndex).Field Is OptionButton Or TypeOf colUndoList(iUndoIndex).Field Is CheckBox Then
      colUndoList(iUndoIndex).Field.Value = colUndoList(iUndoIndex).PrevVal
    Else
      colUndoList(iUndoIndex).Field.Text = colUndoList(iUndoIndex).PrevVal
    End If
    
    ' Set the Current control
    Set CurControl = colUndoList(iUndoIndex)
    
    ' Set the Focus to the restored control
    colUndoList(iUndoIndex).Field.SetFocus
    DoEvents

    ' Set the Menu Items
    iUndoIndex = iUndoIndex - 1
    If Not MenuUndo Is Nothing Then
      If iUndoIndex > 0 Then
        MenuUndo.Caption = "Undo"
        MenuUndo.Enabled = True
      Else
        MenuUndo.Caption = "Can't Undo"
        MenuUndo.Enabled = False
      End If
    End If
    If Not MenuRedo Is Nothing Then
      MenuRedo.Caption = "Redo"
      MenuRedo.Enabled = True
    End If
  End If
  bDisallowClear = False
  Exit Sub
ErrorHandler:
  Debug.Print Err.Description
End Sub

' **********************************************************************************************
' Function  : RedoEdit
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Reinserts the next item in the Undo List
' **********************************************************************************************
Public Sub RedoEdit()

  On Local Error GoTo ErrorHandler
  ' If fields are being repopulated by this function, we don't want to clear any of the list
  bDisallowClear = True
  
  ' Only execute this code if we are not at the last list element
  If iUndoIndex < colUndoList.Count Then
    ' If an element exists, Set the Current Control pointer to it
    If iUndoIndex > 0 Then
      Set CurControl = colUndoList(iUndoIndex)
    End If
    
    ' Move to the next list element
    iUndoIndex = iUndoIndex + 1
    
    ' Reset the property, depending on the control type
    If TypeOf colUndoList(iUndoIndex).Field Is ComboBox Then
      If colUndoList(iUndoIndex).Field.Style = 2 Then
        colUndoList(iUndoIndex).Field.ListIndex = colUndoList(iUndoIndex).NewVal
      Else
        colUndoList(iUndoIndex).Field.Text = colUndoList(iUndoIndex).NewVal
      End If
    ElseIf TypeOf colUndoList(iUndoIndex).Field Is ListBox Then
      colUndoList(iUndoIndex).Field.ListIndex = colUndoList(iUndoIndex).NewVal
    ElseIf TypeOf colUndoList(iUndoIndex).Field Is OptionButton Or TypeOf colUndoList(iUndoIndex).Field Is CheckBox Then
      colUndoList(iUndoIndex).Field.Value = colUndoList(iUndoIndex).NewVal
    Else
      colUndoList(iUndoIndex).Field.Text = colUndoList(iUndoIndex).NewVal
    End If
    
    bDoneRedo = True
    
    ' Set the Focus to this control
    colUndoList(iUndoIndex).Field.SetFocus
    DoEvents
    bDoneRedo = False
    
    ' Set the menu options
    If Not MenuRedo Is Nothing Then
      If iUndoIndex < colUndoList.Count Then
        MenuRedo.Caption = "Redo"
        MenuRedo.Enabled = True
      Else
        MenuRedo.Caption = "Can't Redo"
        MenuRedo.Enabled = False
      End If
    End If
    If Not MenuUndo Is Nothing Then
      MenuUndo.Caption = "Undo"
      MenuUndo.Enabled = True
    End If
  End If

  bDisallowClear = False
  Exit Sub
ErrorHandler:
  Debug.Print Err.Description
End Sub

' **********************************************************************************************
' Function  : SetupClass
' Author    : Rob Small
' Date      : 21st October 2004
' Overview  : Initialise the Class.  Sets up the calling form and optionally, the Menu items for
'             Undo and Redo.  If these are set, the menu items will automatically be disabled if
'             the Undo list is empty, at the start or end of the list
' **********************************************************************************************
Public Sub SetupClass(FormObject As Form, Optional UndoMenuItem As Menu, Optional RedoMenuItem As Menu)
  Set LocalForm = FormObject
  Set MenuUndo = UndoMenuItem
  Set MenuRedo = RedoMenuItem

  If Not MenuUndo Is Nothing Then
    MenuUndo.Caption = "Can't Undo"
    MenuUndo.Enabled = False
  End If
  If Not MenuRedo Is Nothing Then
    MenuRedo.Caption = "Can't Redo"
    MenuRedo.Enabled = False
  End If
End Sub
