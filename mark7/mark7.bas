Attribute VB_Name = "mark7"
Option Explicit
'   ȫ�ֱ���Ҫ��:
'   1��������Ҫ�ڲ�ͬ���̼䴫�ݣ�
'   2����Ƶ���̣�step2���õ��Ĳ���
    
    Public a$, i%, k%, j%, n%, p%
    Public num(), tech()            '������飬��������
    Public numT$, techT$            '��ű��������ʱ���
    Public numTape$, techTape$      '��Ŵ������ʴ������ڼ����������Ƿ��Ѿ�����,word vbaû����������match�������պ����Ż�
    Public numL$                    'ͬһ�����ٱ��ʱ�ľɱ��
    Public showtxt$                 '��ͼ��Ǽ���
    Public totalTxtDry$, totalTxtWet$    '���ı������֣�ʪ�ı������֣����α��ʱ�������
    Public ckDry%, ckWet%                '���ı�У�飬ʪ�ı�У��
    Public ckPcs%, ckFuz%           '����У�飬��Ϊ��ȫƥ���ģ��ƥ��
    Public ptA$, ptB$
    Public sec%                     '���״̬���ж��Ƿ�Ϊ���α��
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
    Dim tube                '�ɸ�ͼ��ǵ���
    Dim bamboo$             '�ɸ�ͼ��Ǽ���
    Dim oldNo%              '�ɸ�ͼ�������
    Dim reg As New RegExp
    
    Set rg = ActiveDocument.Range
    ActiveDocument.TrackRevisions = True

    k = 1                   'kΪ���α�ŵĴ��������Զ���ҳ��أ�iΪ�ܸ�ͼ������������߶���
    numTape = ""
    techTape = ""
    rg.WholeStory
    totalTxtWet = rg.Text
    With reg
        .Global = True
        .Pattern = "��\w+��"
        totalTxtDry = .Replace(totalTxtWet, "")
    End With
    '���Ҹ�ͼ�������
    With rg.Find
        .ClearFormatting
        .Text = "��ͼ��ǣ�"            '*************************************
        .Execute
        .Parent.Expand unit:=wdParagraph
    End With
    bamboo = rg.Text
    '��ȡ�ɸ�ͼ���
    If InStr(bamboo, "-") > 0 Then
        bamboo = Mid(bamboo, 6, Len(bamboo) - 7)
        tube = Split(bamboo, "��")
        oldNo = UBound(tube)
        
        ReDim num(oldNo + 1)
        ReDim tech(oldNo + 1)
        
        For j = 0 To oldNo
            t = Split(tube(j), "-")
            num(j + 1) = t(0)
            tech(j + 1) = t(1)
        Next
        '��ز�����ʼ��
        i = oldNo + 2
        showtxt = ""
        sec = 1
        For n = 1 To i - 1
            showtxt = showtxt & num(n) & "-" & tech(n) & "��"
            numTape = numTape & " " & num(n) & " "
            techTape = techTape & " " & tech(n) & " "
        Next
        showtxt = Left(showtxt, Len(showtxt) - 1) & "��"
    Else
        ReDim tech(1)
        ReDim num(1)
        sec = 0
        i = 1
        showtxt = "���븽ͼ���,�����ƺ�����(��:Һ����10)��" & Chr(10) & "����'u'ҳ���Ϸ���" & Chr(10) & "����'d'ҳ���·���"
    End If
    num(0) = 0  '������һ����λ�жϵ�Ƕ��
    Selection.HomeKey unit:=wdStory
    
    With mark
        .Show 0
        .display.Caption = showtxt
    End With
    
End Sub
Sub step2()
        If (k Mod 8) = 0 Then ActiveWindow.SmallScroll down:=16 '�Զ�����ҳ��

        If a = "u" Then                                            '�ֶ�����ҳ��
            ActiveWindow.SmallScroll up:=16
        ElseIf a = "d" Then
            ActiveWindow.SmallScroll down:=16
        Else
            For j = Len(a) To 1 Step -1                         '�������ƺͱ��
                If Mid(a, j, 1) Like "[0-9a-zA-Z]" = False Then     '�ӵ����ڶ�λ��ʼ�жϣ��ض���һλ���֣�����ᱻ��һ������
                    techT = Left(a, j)
                    numT = Right(a, Len(a) - j)
                    Exit For
                End If
            Next
            
            If j > 0 And j < Len(a) Then               '���˴�������ĸ��������
                If InStr(numTape, numT) = 0 Then      '�������б�ţ�����ʹ����
                
                    ckDry = InStr(totalTxtDry, techT)
                    ckWet = InStr(totalTxtWet, techT)
                    ckPcs = InStr(techTape, " " & techT & " ")
                    ckFuz = InStr(techTape, techT)
                    
'                    ckDry   ckWet   ckPcs   ckFuz   result
'                       1       1       1       /    ���±��
'                       1       1       0       1    �ȳ����
'                       1       1       0       0      �´�
'                       1       0       0       0    �ȶ̺�
                    
                    If ckDry > 0 Then                   '���˴����
                        '***��ʼ���***
                        If ckWet > 0 And ckPcs = 0 And ckFuz = 0 Then '������
                            Call rp(techT, techT & "��" & numT & "��")
                            Application.ScreenRefresh               'ʵʱ��ʾ����滻
                            
                            numTape = numTape & " " & numT & " "           '*
                            techTape = techTape & " " & techT & " "         '*
                            
                            ReDim Preserve num(i)
                            ReDim Preserve tech(i)
                                                            
                            num(i) = numT
                            tech(i) = techT
                                        
                            Call reorder(num(), tech(), i, numT, techT)
                        ElseIf ckWet > 0 And ckPcs > 0 Then         '���������±��
                            i = i - 1                               '***��������δ����***
                            For n = i To 1 Step -1                  '���Ż�����ȷ���Ƿ��йٷ�����
                                If tech(n) = techT Then Exit For
                            Next
                            
                            numL = num(n)
                            num(n) = numT
                            numTape = Replace(numTape, numL, "")    '�ڱ�Ŵ���ȥ���ɱ��
                            
                            Call rp(techT & "��" & numL & "��", techT & "��" & numT & "��")
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
                        ElseIf ckWet = 0 Then  '�ȶ̺�
                            For n = i - 1 To 1 Step -1
                                If InStr(techT, tech(n)) > 0 Then Exit For
                            Next
                            
                            ptA = Left(techT, InStr(techT, tech(n)) + Len(tech(n)) - 1)
                            ptB = Right(techT, Len(techT) - Len(ptA))
                            
                            Call rp(ptA & "��" & num(n) & "��" & ptB, techT & "��" & numT & "��")
                            Application.ScreenRefresh
                            
                            numTape = numTape & " " & numT & " "            '*
                            techTape = techTape & " " & techT & " "         '*
                            
                            ReDim Preserve num(i)
                            ReDim Preserve tech(i)
                                                            
                            num(i) = numT
                            tech(i) = techT
                                        
                            Call reorder(num(), tech(), i, numT, techT)
                        ElseIf ckWet > 0 And ckPcs = 0 And ckFuz > 0 Then '�ȳ����
                            For n = i - 1 To 1 Step -1
                                If InStr(tech(n), techT) > 0 Then Exit For
                            Next
                            
                            ptA = Left(tech(n), InStr(tech(n), techT) + Len(techT) - 1)
                            ptB = Right(tech(n), Len(tech(n)) - Len(ptA))
                            
                            Call rp(techT, techT & "��" & numT & "��")
                            Call rp(ptA & "��" & numT & "��" & ptB, tech(n))
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
                            showtxt = showtxt & num(n) & "-" & tech(n) & "��"
                        Next
                        showtxt = Left(showtxt, Len(showtxt) - 1) & "��"
                        
                        i = i + 1                                   '*******ע�������������Ƕ�׵�λ��******
                        k = k + 1
                    Else
                        MsgBox ("���ƴ���")
                    End If
                Else
                    MsgBox ("���" & numT & "�Ѿ�����")
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
    For n = i To 1 Step -1                  '��������
        If _
            Val(num(n)) > Val(num(n - 1)) Or _
            ( _
                Val(num(n)) = Val(num(n - 1)) And _
                    (Len(num(n)) > Len(num(n - 1)) Or _
                    Asc(Right(num(n), 1)) > Asc(Right(num(n - 1), 1))) _
            ) Then                          '���ֱȽϣ�������ȱ�ĩβ��ĸ��ASCIIֵ
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
            .Text = "��ͼ��ǣ�"
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
    '���ȫ�ֱ���
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

