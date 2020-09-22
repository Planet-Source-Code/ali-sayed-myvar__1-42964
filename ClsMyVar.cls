VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMyVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ArrTmp() As String
Dim ShiftVal As Long
Dim m_Text As String

Public Sub Clear()
'reset array and text
Erase ArrTmp
ReDim ArrTmp(0)
m_Text = ""
End Sub
Public Property Get Text() As String
'return var

'maybe some text that no assinged in array
'so we add  "& m_Text"
Text = Join(ArrTmp, "") & m_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
Dim Arrlen As Long
'add text to MyVar its add new value to old one
'so you only add the new portion of var
m_Text = m_Text & vNewValue

'MyVar work

'check if big then shiftrate
If Len(m_Text) >= ShiftVal Then
    'if so get array UBound
    Arrlen = UBound(ArrTmp)
    'shift data from text to array
    ArrTmp(Arrlen) = m_Text
    'add array member
    ReDim Preserve ArrTmp(Arrlen + 1)
    'reset text
    m_Text = ""
End If

End Property

'max size for handling
Public Property Get ShiftRate() As Long
ShiftRate = ShiftVal
End Property
Public Property Let ShiftRate(ByVal vNewValue As Long)
ShiftVal = vNewValue
End Property

Private Sub Class_Initialize()
'defualt values
ReDim ArrTmp(0)
ShiftRate = 7000

End Sub
