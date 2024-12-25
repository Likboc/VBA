Attribute VB_Name = "PrintScreen"
Sub PrintScreen()
        ' Regular usage of MsgBox, The Prompt must use less than 1024 character
        MsgBox "Hello, MsgBox"
        ' Add line character symbol
        MsgBox "Hello" & Chr(13) & "MsgBox"
        ' Add Title info
        MsgBox "Hello,MsgBox", Title = "Demo"
        'Change the default style, 20 kinds of styles in total
        'using := to assgin named parameters
        MsgBox "Hello,MsgBox", Buttons:=1
        
        'using debug.print
        Dim Name, Age
        Name = "Mike"
        Age = 12
        Debug.Print Name, Age
End Sub
