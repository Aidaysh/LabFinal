Attribute VB_Name = "Module1"
Sub Макрос5()
Attribute Макрос5.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Макрос5 Макрос
'
' Сочетание клавиш: Ctrl+p
'
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.Offset(-1, -6).Range("A1:H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:G5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(5, 0).Range("A1:G5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(5, 0).Range("A1:G5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(7, 2).Range("A1").Select
End Sub
