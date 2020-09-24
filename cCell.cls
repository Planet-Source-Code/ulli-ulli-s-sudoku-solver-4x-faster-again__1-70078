VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#################
'represents a cell
'#################

'my properties
Private myValue             As Long 'number in this cell
Private myRowNumber         As Long 'row this cell is in
Private myColumnNumber      As Long 'column this cell is in
Private myBlockNumber       As Long 'block this cell is in
Private myFixed             As Boolean 'number in this cell is fixed

Friend Property Get BlockNumber() As Long

  'returns the block number this cell is member of

    BlockNumber = myBlockNumber

End Property

Friend Property Let Cellnumber(ByVal nuCellNumber As Long)

  'uses the cellnumber to calculate the row, column and block this cell is member of

    myRowNumber = nuCellNumber \ 9                                          ' 0 to  8 = Row
    myColumnNumber = (nuCellNumber Mod 9) + 9                               ' 9 to 17 = Column
    myBlockNumber = 3 * ((myRowNumber) \ 3) + (myColumnNumber - 9) \ 3 + 18 '18 to 26 = Block

End Property

Friend Property Get ColumnNumber() As Long

  'returns the column number this cell is member of

    ColumnNumber = myColumnNumber

End Property

Friend Property Get Fixed() As Boolean

  'returns true if value in this cell is fixed (is an initial value)

    Fixed = myFixed

End Property

Friend Property Let Fixed(nuFixed As Boolean)

  'sets or unsets cell fixed

    myFixed = nuFixed

End Property

Friend Property Get RowNumber() As Long

  'returns the row number this cell is member of

    RowNumber = myRowNumber

End Property

Friend Property Get Value() As Long

  'returns the number in this cell

    Value = myValue

End Property

Friend Property Let Value(ByVal nuValue As Long)

  'puts a new value in this cell

    If myValue Then 'don't have to add back zero
        Groups(myRowNumber).AddValue myValue 'add the previous value back to available values
        Groups(myColumnNumber).AddValue myValue
        Groups(myBlockNumber).AddValue myValue
    End If
    myValue = nuValue
    If myValue Then 'don't have to remove zero
        Groups(myRowNumber).RemoveValue myValue 'and now remove current value from available values
        Groups(myColumnNumber).RemoveValue myValue
        Groups(myBlockNumber).RemoveValue myValue
    End If

End Property

':) Ulli's VB Code Formatter V2.23.17 (2008-Feb-25 09:14)  Decl: 12  Code: 79  Total: 91 Lines
':) CommentOnly: 14 (15,4%)  Commented: 12 (13,2%)  Empty: 35 (38,5%)  Max Logic Depth: 2
':) Magic Number: 442372841
