Attribute VB_Name = "mark7"
Option Explicit
'   全局变量要求:
'   1、参数需要在不同过程间传递；
'   2、高频过程（step2）用到的参数
    
    Public a$, i%, k%, j%, n%, p%
    Public num(), tech()            '标号数组，名词数组
    Public numT$, techT$            '标号变量，名词变量
    Public numTape$, techTape$      '标号带，名词带，用于检测新输入的是否已经存在,word vba没有针对数组的match函数，日后再优化
    Public numL$                    '同一名词再编号时的旧标号
    Public showtxt$                 '附图标记集合
    Public totalTxtDry$, totalTxtWet$    '干文本纯文字，湿文本带数字，初次标记时两者相等
    Public ckDry%, ckWet%                '干文本校验，湿文本校验
    Public ckPcs%, ckFuz%           '名词校验，分为完全匹配和模糊匹配
    Public ptA$, ptB$
    Public sec%                     '标号状态，判断是否为二次标号
    Public rg As Range
Sub positionTest()
    With mark
        .StartUpPosition = 0
        .Left = 650
        .Top = 200
        .Show 0
        .type1.SetFocus
    End With
End Sub
Sub step1()
    Dim t
    Dim j%, n%
    Dim tube                '旧附图标记单个
    Dim bamboo$             '旧附图标记集合
    Dim oldNo%              '旧附图标记总数
    Dim reg As New RegExp
    
    Set rg = ActiveDocument.Range
    ActiveDocument.TrackRevisions = True

    k = 1                   'k为单次编号的次数，与自动翻页相关；i为总附图标记数量，两者独立
    numTape = ""
    techTape = ""
    rg.WholeStory
    totalTxtWet = rg.Text
    With reg
        .Global = True
        .Pattern = "（\w+）"
        totalTxtDry = .Replace(totalTxtWet, "")
    End With
    '查找附图标记区域
    With rg.Find
        .ClearFormatting
        .Text = "附图标记："            '*************************************
        .Execute
        .Parent.Expand unit:=wdParagraph
    End With
    bamboo = rg.Text
    '读取旧附图标记
    If InStr(bamboo, "-") > 0 Then
        bamboo = Mid(bamboo, 6, Len(bamboo) - 7)
        tube = Split(bamboo, "，")
        oldNo = UBound(tube)
        
        ReDim num(oldNo + 1)
        ReDim tech(oldNo + 1)
        
        For j = 0 To oldNo
            t = Split(tube(j), "-")
            num(j + 1) = t(0)
            tech(j + 1) = t(1)
        Next
        '相关参数初始化
        i = oldNo + 2
        showtxt = ""
        sec = 1
        For n = 1 To i - 1
            showtxt = showtxt & num(n) & "-" & tech(n) & "，"
            numTape = numTape & " " & num(n) & " "
            techTape = techTape & " " & tech(n) & " "
        Next
        showtxt = Left(showtxt, Len(showtxt) - 1) & "。"
    Else
        ReDim tech(1)
        ReDim num(1)
        sec = 0
        i = 1
        showtxt = "输入附图标记,先名称后数字(如:液晶屏10)。" & Chr(10) & "输入'u'页面上翻。" & Chr(10) & "输入'd'页面下翻。"
    End If
    num(0) = 0  '排序少一层首位判断的嵌套
    Selection.HomeKey unit:=wdStory
    
    With mark
        .Show 0
        .display.Caption = showtxt
    End With
    
End Sub
Sub step2()
        If (k Mod 8) = 0 Then ActiveWindow.SmallScroll down:=16 '自动滚动页面

        If a = "u" Then                                            '手动滚动页面
            ActiveWindow.SmallScroll up:=16
        ElseIf a = "d" Then
            ActiveWindow.SmallScroll down:=16
        Else
            For j = Len(a) To 1 Step -1                         '分离名称和标号
                If Mid(a, j, 1) Like "[0-9a-zA-Z]" = False Then     '从倒数第二位开始判断，必定带一位数字，否则会被下一步过滤
                    techT = Left(a, j)
                    numT = Right(a, Len(a) - j)
                    Exit For
                End If
            Next
            
            If j > 0 And j < Len(a) Then               '过滤纯数字字母、纯文字
                If InStr(numTape, numT) = 0 Then      '过滤已有标号，提醒使用者
                
                    ckDry = InStr(totalTxtDry, techT)
                    ckWet = InStr(totalTxtWet, techT)
                    ckPcs = InStr(techTape, " " & techT & " ")
                    ckFuz = InStr(techTape, techT)
                    
'                    ckDry   ckWet   ckPcs   ckFuz   result
'                       1       1       1       /    重新编号
'                       1       1       0       1    先长后短
'                       1       1       0       0      新词
'                       1       0       0       0    先短后长
                    
                    If ckDry > 0 Then                   '过滤打错字
                        '***开始编号***
                        If ckWet > 0 And ckPcs = 0 And ckFuz = 0 Then '新名词
                            Call rp(techT, techT & "（" & numT & "）")
                            Application.ScreenRefresh               '实时显示标号替换
                            
                            numTape = numTape & " " & numT & " "           '*
                            techTape = techTape & " " & techT & " "         '*
                            
                            ReDim Preserve num(i)
                            ReDim Preserve tech(i)
                                                            
                            num(i) = numT
                            tech(i) = techT
                                        
                            Call reorder(num(), tech(), i, numT, techT)
                        ElseIf ckWet > 0 And ckPcs > 0 Then         '旧名词重新标号
                            i = i - 1                               '***名词总数未增加***
                            For n = i To 1 Step -1                  '待优化，不确定是否有官方函数
                                If tech(n) = techT Then Exit For
                            Next
                            
                            numL = num(n)
                            num(n) = numT
                            numTape = Replace(numTape, numL, "")    '在标号带中去除旧标号
                            
                            Call rp(techT & "（" & numL & "）", techT & "（" & numT & "）")
                            Application.ScreenRefresh
                            
                            If i > 1 Then
                                If Val(num(n)) < Val(num(n - 1)) Then
                                    For p = n To 2 Step -1
                                        If Val(num(p)) >= Val(num(p - 1)) Then
                                            Exit For
                                        Else
                                            num(p) = num(p - 1)
                                            tech(p) = tech(p - 1)
                                            num(p - 1) = numT
                                            tech(p - 1) = techT
                                        End If
                                    Next
                                ElseIf n < i Then
                                    If num(n) > num(n + 1) Then
                                        For p = n To i - 1
                                            If Val(num(p)) <= Val(num(p + 1)) Then
                                                Exit For
                                            Else
                                                num(p) = num(p + 1)
                                                tech(p) = tech(p + 1)
                                                num(p + 1) = numT
                                                tech(p + 1) = techT
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        ElseIf ckWet = 0 Then  '先短后长
                            For n = i - 1 To 1 Step -1
                                If InStr(techT, tech(n)) > 0 Then Exit For
                            Next
                            
                            ptA = Left(techT, InStr(techT, tech(n)) + Len(tech(n)) - 1)
                            ptB = Right(techT, Len(techT) - Len(ptA))
                            
                            Call rp(ptA & "（" & num(n) & "）" & ptB, techT & "（" & numT & "）")
                            Application.ScreenRefresh
                            
                            numTape = numTape & " " & numT & " "            '*
                            techTape = techTape & " " & techT & " "         '*
                            
                            ReDim Preserve num(i)
                            ReDim Preserve tech(i)
                                                            
                            num(i) = numT
                            tech(i) = techT
                                        
                            Call reorder(num(), tech(), i, numT, techT)
                        ElseIf ckWet > 0 And ckPcs = 0 And ckFuz > 0 Then '先长后短
                            For n = i - 1 To 1 Step -1
                                If InStr(tech(n), techT) > 0 Then Exit For
                            Next
                            
                            ptA = Left(tech(n), InStr(tech(n), techT) + Len(techT) - 1)
                            ptB = Right(tech(n), Len(tech(n)) - Len(ptA))
                            
                            Call rp(techT, techT & "（" & numT & "）")
                            Call rp(ptA & "（" & numT & "）" & ptB, tech(n))
                            Application.ScreenRefresh
                            
                            numTape = numTape & " " & numT & " "            '*
                            techTape = techTape & " " & techT & " "         '*
                            
                            ReDim Preserve num(i)
                            ReDim Preserve tech(i)
                                                            
                            num(i) = numT
                            tech(i) = techT
                                        
                            Call reorder(num(), tech(), i, numT, techT)
                        End If
                        
                        showtxt = ""
                        For n = 1 To i
                            showtxt = showtxt & num(n) & "-" & tech(n) & "，"
                        Next
                        showtxt = Left(showtxt, Len(showtxt) - 1) & "。"
                        
                        i = i + 1                                   '*******注意计数变量处于嵌套的位置******
                        k = k + 1
                    Else
                        MsgBox ("名称错误")
                    End If
                Else
                    MsgBox ("标号" & numT & "已经有了")
                End If
            End If
        End If
    
        rg.WholeStory
        totalTxtWet = rg.Text
        
        mark.display.Caption = showtxt

End Sub
Function rp(a, b)
    With Selection.Find
        .ClearFormatting
        .Text = a
        With .Replacement
            .ClearFormatting
            .Text = b
        End With
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Function
Function reorder(num(), tech(), i, numT, techT)
    Dim n
    For n = i To 1 Step -1                  '插入排序
        If _
            Val(num(n)) > Val(num(n - 1)) Or _
            ( _
                Val(num(n)) = Val(num(n - 1)) And _
                    (Len(num(n)) > Len(num(n - 1)) Or _
                    Asc(Right(num(n), 1)) > Asc(Right(num(n - 1), 1))) _
            ) Then                          '数字比较，数字相等比末尾字母的ASCII值
            Exit For
        Else
            num(n) = num(n - 1)
            tech(n) = tech(n - 1)
            num(n - 1) = numT
            tech(n - 1) = techT
        End If
    Next
End Function
Sub step3()
    i = i - 1
    k = k - 1
    
    If k = 0 Then
        Exit Sub
    Else
        With Selection.Find
            .ClearFormatting
            .Text = "附图标记："
            .Wrap = wdFindContinue
            .Execute
        End With
        
        Selection.MoveRight
        Selection.TypeText Text:=showtxt
        
        If sec = 1 Then
            Selection.MoveDown unit:=wdParagraph, Extend:=wdExtend
            Selection.End = Selection.End - 1
            Selection.Delete
        End If
    End If
End Sub
Sub step4()
    '清除全局变量
    a = ""
    j = 0
    i = 0
    k = 0
    n = 0
    p = 0
    sec = 0
    ckDry = 0
    ckWet = 0
    ckPcs = 0
    ckFuz = 0
    
    numT = ""
    techT = ""
    numTape = ""
    techTape = ""
    numL = ""
    showtxt = ""
    totalTxtDry = ""
    totalTxtWet = ""
    
    ptA = ""
    ptB = ""
    
    Erase num
    Erase tech
    
    Set rg = Nothing
End Sub

