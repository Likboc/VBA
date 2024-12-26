'An WorkSheet_Change Event
Private Sub Worksheet_Change(ByVal Target As Range)
        If Target.Column = 6 And Target.Row > 3 _
        Then Target.Offset(, 2).Value = Target.Offset(, 2).Value & Date & " " & Target.Value & vbNewLine
End Sub