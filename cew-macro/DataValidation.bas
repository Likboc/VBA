Sub DataValidation()
    'Only for DFN & QFN worksheets
      Dim cnt As Integer
        cnt = 4
        If Target.Column = 2 And Target.Row > 3 Then
            Do While Cells(cnt, "B") <> ""
                If (Cells(cnt, "B").Value = Target.Value) Then
                    MsgBox Target.Value & "Device Name duplicated"
                    Target.ClearContents
                    End
                End If
                cnt = cnt + 1
            Loop
        End If
End Sub