Attribute VB_Name = "Module1"
Option Explicit

Sub �f�n�S���ħ��Ƨ�()
Attribute �f�n�S���ħ��Ƨ�.VB_Description = "�ثe�����f�n�S���ħ��w�s�Ƨ�"
Attribute �f�n�S���ħ��Ƨ�.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�S���ħ��Ƨ� ����
' �ثe�����f�n�S���ħ��w�s�Ƨ�
'
' �ֳt��: Ctrl+q
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ����2()
Attribute ����2.VB_Description = "�����f�n�w�s�Ƨ�"
Attribute ����2.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' ����2 ����
' �����f�n�w�s�Ƨ�
'
' �ֳt��: Ctrl+z
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
