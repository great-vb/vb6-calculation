VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Stack() As String
Private itemCount As Long
Private Sub Class_Initialize()
    ReDim Stack(0)
    Stack(0) = "#"
End Sub
Public Sub Push(ByVal inString As String)
    ReDim Preserve Stack(itemCount + 1)
    Stack(itemCount + 1) = inString
    itemCount = itemCount + 1
End Sub

Public Function Pop() As String
    If itemCount >= 1 Then
        Pop = Stack(itemCount)
        ReDim Preserve Stack(itemCount - 1)
        itemCount = itemCount - 1
    Else
        Pop = ""
    End If
End Function
Public Function Peek() As String
    If itemCount = 0 Then
        Peek = ""
        Exit Function
    End If
    Peek = Stack(itemCount)
End Function

Sub Clear()
    itemCount = 0
    ReDim Stack(itemCount)
    Stack(itemCount) = "#"
End Sub

Public Function Count()
    Count = itemCount
End Function
Public Function ViewStack() As String
    Dim kOut As String
    Dim i As Long
    If itemCount = 0 Then ViewStack = "": Exit Function
    For i = 1 To itemCount
        kOut = kOut & Format(i, "00") & " " & Stack(i) & vbCrLf
    Next i
    ViewStack = kOut
End Function
