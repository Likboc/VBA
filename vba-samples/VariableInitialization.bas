Attribute VB_Name = "VariableDeclaration"
Sub DimVar()
        'Regular declaration
        'Dim First_Name As String, Last_Name As String
        
        'Using symbol, only limited data type support symbol
        'Dim First_Name$, Last_Name$
        
        'Set the String length,if short than the length, i will fill the rest with space, if longer, the abundant part will not count
        ' For Example, Dim abc As String*1, abc = 123, MsgBox abc will output 1;
        'Dim First_Name As String * 10, Last_Name As String * 10
        
        'Initialize var while declaration is not supported by VBA, the code below is illegall
        'Dim First_Name="Mike", Last_Name="Smith"
        
        ' initailize the string,num,date var, you can using let or just ignore the keyword
        ' Let First_Name = "Mike" OR First_Name = "Mike"
        
        'also private & public keyword is supported by VBA, but it's not widely used in my project
        
        'the constant can't be reassigned, you must initialize while declaring
        'Const p As Single = 3.14
        
End Sub


