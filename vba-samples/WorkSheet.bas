Sub WorkDemo()
    'use index to return a worksheet obj, and the index starts from 1
    WorkSheets(1).Activate
    'equivalent to
    WorkSheets("Sheet1").Activate
End Sub