Sub LogicSymbol()
    'the if-then-else logic
    If 1>0 Then MsgBox "hi" Else MsgBox "no"
    ' the else if
    If 1>0 Then MsgBox "hi" ElseIf 1>0 MsgBox "no"
    ' using End If if you need multiline for the logic
    If 1>0 MsgBox "hi"
    Else MsgBox "no"
    End If

    'Add Logic Symbol, such as And,Or,Not,Xor,Eqv,Imp
    If 1>0 And 2>1 Then MsgBox "yes"

    'Select Case Logic which you can use to replace the multi if-then structure
    Select Case Time()
        Case Is > 0.5
            MsgBox "hello"
        Case Is < 0.5
            MsgBox "hi"
        Case Else 
            MsgBox "你好"
    End Case 
    'for loop, the code in the () can be ignored
    For i = 2 To 10 Step 1
        MsgBox i
        Exit For 'equal to continue
    Next i

    'Do loop, consist of Do While, Do Until
        MsgBox "hi"
        Exit Do 'equal to continue
    Loop

End Sub