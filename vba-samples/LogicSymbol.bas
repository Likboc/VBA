Sub LogicSymbol()
    'the if-then-else logic
    If 1>0 Then MsgBox "hi" Else MsgBox "no"
    ' the else if
    If 1>0 Then MsgBox "hi" ElseIf 1>0 MsgBox "no"
    ' using End If if you need multiline for the logic
    If 1>0 MsgBox "hi"
    Else MsgBox "no"
    End If
    
End Sub