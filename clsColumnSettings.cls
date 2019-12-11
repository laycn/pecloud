VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumnSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private mvarindex As Integer 'local copy
Private mvargrow As Boolean 'local copy
Private mvaredit As Boolean 'local copy
Private mvarcolwidth As Double 'local copy
Private mvaralign As Integer 'local copy
Private mvarFormat_Numeric As Boolean 'local copy
Private mvarFormat_String As String 'local copy
'local variable(s) to hold property value(s)
Private mvarshow As Boolean 'local copy
Public Property Let show(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.show = 5
    mvarshow = vData
End Property


Public Property Get show() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.show
    show = mvarshow
End Property



Public Property Let Format_String(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Format_String = 5
    mvarFormat_String = vData
End Property


Public Property Get Format_String() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Format_String
    Format_String = mvarFormat_String
End Property



Public Property Let Format_Numeric(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Format_Numeric = 5
    mvarFormat_Numeric = vData
End Property


Public Property Get Format_Numeric() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Format_Numeric
    Format_Numeric = mvarFormat_Numeric
End Property



Public Property Let align(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.align = 5
    mvaralign = vData
End Property


Public Property Get align() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.align
    align = mvaralign
End Property



Public Property Let colwidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.colwidth = 5
    mvarcolwidth = vData
End Property


Public Property Get colwidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colwidth
    colwidth = mvarcolwidth
End Property



Public Property Let edit(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.edit = 5
    mvaredit = vData
End Property


Public Property Get edit() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.edit
    edit = mvaredit
End Property



Public Property Let grow(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.grow = 5
    mvargrow = vData
End Property


Public Property Get grow() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.grow
    grow = mvargrow
End Property



Public Property Let index(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.index = 5
    mvarindex = vData
End Property


Public Property Get index() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.index
    index = mvarindex
End Property



Private Sub Class_Initialize()
    Me.Format_Numeric = False
End Sub

