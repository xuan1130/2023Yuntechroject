Attribute VB_Name = "Module1"
Option Explicit

Sub 口罩特約藥局排序()
Attribute 口罩特約藥局排序.VB_Description = "目前醫療口罩特約藥局庫存排序"
Attribute 口罩特約藥局排序.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩特約藥局排序 巨集
' 目前醫療口罩特約藥局庫存排序
'
' 快速鍵: Ctrl+q
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 巨集2()
Attribute 巨集2.VB_Description = "醫療口罩庫存排序"
Attribute 巨集2.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' 巨集2 巨集
' 醫療口罩庫存排序
'
' 快速鍵: Ctrl+z
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
