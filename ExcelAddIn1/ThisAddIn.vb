'Public Class ThisAddIn
'
'Private Sub ThisAddIn_Startup() Handles Me.Startup
'
'End Sub
'
'Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
'
'End Sub
'
'End Class

Public m, n As Integer
Sub 清除颜色()
    '
    ' 清除颜色 宏
    '

    '
    Range("C6:AF70").Select
    ActiveWindow.SmallScroll Down:=-69
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub 求成绩()
    '
    ' 求总分
    '

    '
    Dim i As Integer                                'i为纵坐标
    Dim j As Integer                                'j为横坐标
    Dim zongfeng As Integer                         'zongfeng为选择题总成绩
    '    Dim m, n As Integer                             'm为最后一人所在的行号，n为最后一题所在的列号
    i = 6 : j = 3 : n = 3 : m = 6 : zongfeng = 0        '初始化值
    Do                                              '求n的值
        n = n + 1
    Loop Until IsEmpty(Cells(2, n + 1))
    Do                                              '求m的值
        m = m + 1
    Loop Until IsEmpty(Cells(m + 1, 2))
    For i = 6 To m
        zongfeng = 0
        For j = 3 To n
            If Cells(i, j) = Cells(2, j) Then
                zongfeng = zongfeng + 2
            End If
        Next j
        Cells(i, 33) = zongfeng
    Next i
End Sub