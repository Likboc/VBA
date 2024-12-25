Attribute VB_Name = "VariableDeclaration"
Sub DimVar()
        'Regular declaration
        Dim First_Name As String, Last_Name As String
        
        'Using symbol, only limited data type support symbol
        Dim First_Name$, Last_Name$
        
        'Set the String length,if short than the length, i will fill the rest with space, if longer, the abundant part will not count
        ' For Example, Dim abc As String*1, abc = 123, MsgBox abc will output 1;
        Dim First_Name As String * 10, Last_Name As String * 10
        
        'Initialize var while declaration is not supported by VBA, the code below is illegall
        Dim First_Name="Mike", Last_Name="Smith"
        
        ' initailize the string,num,date var, you can using let or just ignore the keyword
        Let First_Name = "Mike" OR First_Name = "Mike"
        
        'also private & public keyword is supported by VBA, but it's not widely used in my project
        
        'the constant can't be reassigned, you must initialize while declaring
        Const p As Single = 3.14

        'declare arr
        Dim NameArr(1 To 10) As Integer
        Dim NameArr(10) As Integer
        
        'declare double arr
        Dim NameArr(1 To 2, 1 To 2) As Integer
        Dim NameArr(2,2) As Integer

        'declare a dynamic arr
        Dim NameArr() As String
        Dim Len As Integer
        Len = 10
        ReDim NameArr(Len) As String
        
        'create your array by Array
        Dim arr As Variant
        arr = Array(1,2,3,4,5)
        MsgBox arr(0)

        'create your array by Split
        Dim arr As Variant
        arr = Split("1,22,44",",")

        '---- by Range
        Dim arr As Variant
        arr = Range("A1:B2").Value
        Range("A1:B2").Value = arr
End Sub


