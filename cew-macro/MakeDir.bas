Sub MakeDir()
 Dim path As String
 path = ThisWorkbook.path & "\NAME"
 path = Replace(path, "NAME", ActiveCell.Value2)
 If Len(Dir(path, 16)) = 0 Then
    MkDir path
    If ActiveSheet.Name <> "PCB" Then
        MkDir path + "\评估"
        MkDir path + "\封装"
    End If
 End If
 MkDir path + format(Date, "yyyymmdd")
 Selection.Hyperlinks.Add Anchor:=Selection, Address:=path, SubAddress:="", ScreenTip:="", TextToDisplay:=ActiveCell.Value2
End Sub