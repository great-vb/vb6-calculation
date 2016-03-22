Attribute VB_Name = "CalcFunc"
Public opNum As New StackClass
Public opChar As New StackClass
Public Function CalcString(ByVal strIn As String) As String
    Dim sTxt As String
    Dim strNumFix As String
    Dim curChar As String
    Dim i As Long
    Dim signCount As Long
    Dim ops1 As String, ops2 As String, opC As String
    '初始化堆栈
        opNum.Clear
        opChar.Clear
    '堆栈初始化结束
    sTxt = strIn
    For i = 1 To Len(sTxt)
        curChar = Mid(sTxt, i, 1)
        If IsSymbol(curChar) = True Then
            '看看数字预备区有没有
            If strNumFix <> "" Then
                opNum.Push strNumFix
                strNumFix = ""
            End If
redo:
            If IsHigh(curChar, opChar.Peek) = 1 Then 'if new come char is higher then push it to stack
                opChar.Push curChar '如果等级高的控制符，则进入
                signCount = signCount + 1
            ElseIf IsHigh(curChar, opChar.Peek) = 0 Then
                If curChar = "#" And opChar.Peek = "#" Then
                    opChar.Pop
                    CalcString = opNum.Pop
                    Exit Function
                End If
            ElseIf IsHigh(curChar, opChar.Peek) = -1 Then 'if low then ready to calculate
                '判断是不是第一个符号
                If signCount = 1 Then '这个符号是刚刚输入#后的那个，无论如何入栈
                    opChar.Push curChar
                    signCount = signCount + 1
                    GoTo nextone
                End If
                ops2 = opNum.Pop
                ops1 = opNum.Pop
                opC = opChar.Pop
                opNum.Push CStr(Calc(ops1, ops2, opC))
                If curChar = ")" And opChar.Peek = "(" Then
                    opChar.Pop  '如果操作数是），就把（弹出来
                    GoTo moveon
                End If
                GoTo redo
moveon:
            End If
        Else '非符号
            strNumFix = strNumFix & curChar
        End If
nextone:
    Next i
End Function

Public Function Calc(ByVal op1 As String, ByVal op2 As String, ByVal options As String) As Double
On Error Resume Next
Calc = 0
Select Case options
    Case "+"
        Calc = CDbl(op1) + CDbl(op2)
    Case "-"
        Calc = CDbl(op1) - CDbl(op2)
    Case "*"
        Calc = CDbl(op1) * CDbl(op2)
    Case "/"
        Calc = CDbl(op1) / CDbl(op2)
End Select
End Function

Public Function IsHigh(ByVal sNew As String, ByVal sOld As String) As Integer
'1大于，-1小于，0等于
Select Case sNew
Case "+"
    Select Case sOld
        Case "("
            IsHigh = 1
            Exit Function
        Case "#"
            IsHigh = 1
            Exit Function
        Case Else
            IsHigh = -1
            Exit Function
    End Select
Case "-"
    Select Case sOld
        Case "("
            IsHigh = 1
            Exit Function
        Case "#"
            IsHigh = 1
            Exit Function
        Case Else
            IsHigh = -1
            Exit Function
    End Select
Case "*"
    Select Case sOld
        Case "("
            IsHigh = 1
            Exit Function
        Case "#"
            IsHigh = 1
            Exit Function
        Case "+"
            IsHigh = 1
            Exit Function
        Case "-"
            IsHigh = 1
            Exit Function
        Case Else
            IsHigh = -1
            Exit Function
    End Select
Case "/"
    Select Case sOld
        Case "("
            IsHigh = 1
            Exit Function
        Case "#"
            IsHigh = 1
            Exit Function
        Case "+"
            IsHigh = 1
            Exit Function
        Case "-"
            IsHigh = 1
            Exit Function
        Case Else
            IsHigh = -1
            Exit Function
    End Select
Case "("
    Select Case sOld
        Case "+"
            IsHigh = 1
            Exit Function
        Case "-"
            IsHigh = 1
            Exit Function
        Case "*"
            IsHigh = 1
            Exit Function
        Case "/"
            IsHigh = 1
            Exit Function
        Case "("
            IsHigh = 1
            Exit Function
        Case Else
            IsHigh = -1
            Exit Function
    End Select
Case ")"
    IsHigh = -1
    Exit Function
Case ""
    IsHigh = -1
    Exit Function
Case "#"
    Select Case sOld
        Case "#"
            IsHigh = 0
            Exit Function
        Case ""
            IsHigh = 1
            Exit Function
        Case "+"
            IsHigh = -1
            Exit Function
        Case "-"
            IsHigh = -1
            Exit Function
        Case "*"
            IsHigh = -1
            Exit Function
        Case "/"
            IsHigh = -1
            Exit Function
        Case ")"
            IsHigh = -1
            Exit Function
    End Select
End Select
End Function

Public Function IsSymbol(ByVal strS As String) As Boolean
    IsSymbol = True
    Select Case strS
        Case "+"
        Case "-"
        Case "*"
        Case "/"
        Case "("
        Case ")"
        Case "#"
        Case Else
            IsSymbol = False
    End Select
End Function
