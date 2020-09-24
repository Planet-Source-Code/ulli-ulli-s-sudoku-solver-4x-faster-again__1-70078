VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################
'represents a row, a column or a block
'#####################################

Private Bits(1 To 9)                As Long
Private NotBits(1 To 9)             As Long

'my propertie
Private myAvailableValues(0 To 9)   As Boolean
Private myPermitPattern             As Long

Friend Sub AddValue(ByVal Value As Long)

  'adds a value to the available values

    myAvailableValues(Value) = True
    myPermitPattern = myPermitPattern Or Bits(Value)

End Sub

Friend Property Get Agree(ByVal Value As Long) As Boolean

  'returns true when Value is permissible in this group (row, a column or a block)

    Agree = myAvailableValues(Value)

End Property

Private Sub Class_Initialize()

  'sets all available values to true

  Dim i As Long

    myAvailableValues(0) = True
    For i = 1 To 9
        Bits(i) = 2 ^ i
        NotBits(i) = Not Bits(i)
        AddValue i
    Next i

End Sub

Friend Property Get PermitPattern() As Long

  'returns myAvailableValues as binary

    PermitPattern = myPermitPattern

End Property

Friend Sub RemoveValue(ByVal Value As Long)

  'removes a value from the available values

    myAvailableValues(Value) = False
    myPermitPattern = myPermitPattern And NotBits(Value)

End Sub

Friend Sub Reset()

    Class_Initialize

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Feb-25 09:14)  Decl: 12  Code: 58  Total: 70 Lines
':) CommentOnly: 11 (15,7%)  Commented: 0 (0%)  Empty: 28 (40%)  Max Logic Depth: 2
':) Magic Number: 442372841
