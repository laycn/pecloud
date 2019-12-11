VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents mvarlist As ListView
Attribute mvarlist.VB_VarHelpID = -1
Private WithEvents mvartext As textbox
Attribute mvartext.VB_VarHelpID = -1

Private edit_item As ListItem

Private active_column As Integer
Private edit_active As Boolean

Private old_cell_value As String

Private Px As Single, Py As Single

Private Const SET_LabelEdit = lvwManual
Private Const SET_HideSelection = False
Private Const SET_FullRowSelect = True
Private Const SET_GridLines = True
Private Const SET_View = lvwReport

Private Column As clsColumnSettings
Private Columns As New Collection

Public Sub removeitem(index As Integer)
'remove an item from the listview (alternative to ListView.ListItems.Remove 1)
    
    mvarlist.ListItems.Remove index

End Sub

Public Sub additem(item_text As String, ParamArray item_subitems() As Variant)

    Dim Item As ListItem, i As Integer
    
    Set Item = mvarlist.ListItems.add(, , item_text)
    
    For i = 1 To UBound(item_subitems) + 1
        If i < mvarlist.ColumnHeaders.Count Then Item.SubItems(i) = item_subitems(i - 1)
    Next

End Sub

Public Sub addcolumn(col_name As String, col_key, col_width As Double, col_grow As Boolean, Optional col_edit As Boolean = False, Optional col_align As Integer = lvwColumnLeft, Optional col_show As Boolean = True)
'adds and defines the column header settings storing them in the clsColumnSettings collection

    Dim Col As ColumnHeader
    
    Set Column = New clsColumnSettings
           
    Column.colwidth = col_width
    Column.align = col_align
    Column.edit = col_edit
    Column.grow = col_grow
    Column.show = col_show
    
    Set Col = mvarlist.ColumnHeaders.add(, col_key, col_name, col_width, col_align)
    
    Column.index = Col.index
    
    Columns.add Column, col_key
    
End Sub

Public Property Set list(ByVal vData As Object)
    
    'define which listview control the class is controlling
    Set mvarlist = vData
    
    'Set the listview to the constant setting defined in the global declerations of the class
    mvarlist.LabelEdit = SET_LabelEdit
    mvarlist.HideSelection = SET_HideSelection
    mvarlist.FullRowSelect = SET_FullRowSelect
    mvarlist.GridLines = SET_GridLines
    mvarlist.View = SET_View
    mvarlist.LabelEdit = SET_LabelEdit

End Property

Public Property Get list() As Object
    
    Set list = mvarlist

End Property

Public Sub Resize()
'resizes the grow columns so that the columns fill the listview while still allowing for the scrollbar
' *** KNOWN BUG: at the moment there is only prevision for one "Grow" column ***

    Dim col_grow_width As Double
    
    For Each Column In Columns
        If Column.grow = False Then 'And Column.show = True Then
            col_grow_width = col_grow_width + Column.colwidth
        End If
    Next
    
    For Each Column In Columns
        If Column.grow Then
            mvarlist.ColumnHeaders(Column.index).Width = mvarlist.Width - col_grow_width - 330
        End If
    Next
    
End Sub

Public Property Set textbox(ByVal vData As Object)
    
    Set mvartext = vData
    
    mvartext.BorderStyle = 0
    mvartext.Appearance = vbFlat
    mvartext.BackColor = vbWhite
    mvartext.Visible = False
    mvartext = ""
    
End Property

Private Sub Class_Initialize()

    edit_active = False
    Set Columns = New Collection

End Sub

Private Sub mvarlist_Click()
'if the list is clicked that is not an item then imitate saving the changes on the list

    If edit_active Then mvartext_KeyPress 13

End Sub

Private Sub mvarlist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Px = x
    Py = y
End Sub

Private Sub mvarlist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Item As ListItem
    
    Set Item = mvarlist.HitTest(x, y)
    
    If Item Is Nothing Then
        Exit Sub
    End If
    
    If edit_active Then
    
        mvartext_KeyPress 13
        
    Else
        
        Dim col_left As Double, col_width As Double, item_top As Double, item_height As Double
           
        For Each Column In Columns
            
            col_left = mvarlist.ColumnHeaders(Column.index).Left
            col_width = mvarlist.ColumnHeaders(Column.index).Width
            
            item_top = Item.Top
            item_height = Item.Height
            
            If x >= col_left And x <= (col_left + col_width) Then
                
                If Column.edit Then
                
                    With mvartext
                        
                        .Left = mvarlist.Left + col_left + 70
                        .Width = col_width - 100
                        .Top = mvarlist.Top + item_top + 60
                        .Height = item_height - 120
                        
                        If Column.align = lvwColumnRight Then
                            .Alignment = vbRightJustify
                        ElseIf Column.align = lvwCenter Then
                            .Alignment = vbCenter
                        Else
                            .Alignment = vbLeftJustify
                        End If
                        
                        If Column.index = 1 Then
                            .Text = Item.Text
                            old_cell_value = Item.Text
                        Else
                            .Text = Item.SubItems(Column.index - 1)
                            old_cell_value = Item.SubItems(Column.index - 1)
                        End If
                        
                        .Visible = True
                        
                        .BackColor = vbHighlight
                        .ForeColor = vbHighlightText
                        
                        .SetFocus
                        .SelStart = 1
                        .SelLength = Len(.Text)
                        
                        Set edit_item = mvarlist.SelectedItem
                        
                        active_column = Column.index
                        
                    End With
                    
                Else
                    active_column = -1
                End If
                    
            End If
        
        Next
    End If

End Sub

Private Sub mvartext_GotFocus()
    
    edit_active = True
    
    With mvartext
        .BackColor = vbButtonFace
        .ForeColor = vbWindowText
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
        SendKeys "{HOME}+{END}", True
    End With
    
End Sub

Private Sub mvartext_KeyPress(KeyAscii As Integer)
    
    If Not edit_item Is Nothing Then
    
        With mvartext
            .BackColor = vbButtonFace
            .ForeColor = vbWindowText
        End With
        
        If KeyAscii = 13 Or KeyAscii = 27 Then
            
            'escape was pressed so restore value to original
            If KeyAscii = 27 Then mvartext = old_cell_value
            
            For Each Column In Columns
                
                If Column.index = active_column Then
                    
                    If Column.Format_Numeric Then
                        
                        mvartext = Val(mvartext)
                        
                        If Column.Format_String <> "" Then
                            mvartext = Format(mvartext, Column.Format_String)
                        End If
                    
                    End If
                    
                    Exit For
                
                End If
            
            Next
                        
            If active_column = 1 Then
                edit_item.Text = mvartext
            ElseIf active_column < 1 Then
                'do nothing
            Else
                edit_item.SubItems(active_column - 1) = mvartext
            End If
            
            mvarlist.SetFocus
            mvartext_LostFocus
        
            mvartext.Visible = False
            edit_active = False
            mvartext = ""
        
        End If
    
    End If

End Sub

Private Sub mvartext_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        mvartext_KeyPress 13
        mvarlist.SetFocus
        SendKeys "{UP}", True
    End If
    
    If KeyCode = vbKeyDown Then
        mvartext_KeyPress 13
        mvarlist.SetFocus
        SendKeys "{DOWN}", True
    End If


End Sub

Private Sub mvartext_LostFocus()
       
    With mvartext
        .BackColor = vbHighlight
        .ForeColor = vbHighlightText
    End With

End Sub

