Attribute VB_Name = "Mark64"
Sub mark6_4()
    '***���鿪���޶�ģʽ������Ҫ��ʾ����״̬***
    '��range����selection������ʱҳ�治����ת
    '1��֧�ֱ�Ŵ���ĸ������ֻ���ڵ�����ĸ�����������������
    '2��֧�ֶ��α��
    '3��֧��ͬһ���ʱ�����
    '4��֧�ְ������ʱ�ţ��硰���ӻ������������ӡ�
    '5��Ȩ��Ҫ�����ר�����Ʋ�����ֱ��
    Dim a$, i%, j%, n%, k%, p%
    Dim num(), tech()           '������飬��������
    Dim numT$, techT$           '��ű��������ʱ���
    Dim numTape$, techTape$     '��Ŵ������ʴ������ڼ����������Ƿ��Ѿ�����,word vbaû����������match�������պ����Ż�
    Dim numL$                   'ͬһ�����ٱ��ʱ�ľɱ��
    Dim showTxt$                '��ͼ��Ǽ���
    Dim bamboo$                 '�ɸ�ͼ��Ǽ���
    Dim tube                    '�ɸ�ͼ��ǵ���
    Dim oldNo%                  '�ɸ�ͼ�������
    Dim totalTxtDry$, totalTxtWet$    '���ı������֣�ʪ�ı������֣����α��ʱ�������
    Dim ckDry%, ckWet%                '���ı�У�飬ʪ�ı�У��
    Dim ckPcs%, ckFuz%          '����У�飬��Ϊ��ȫƥ���ģ��ƥ��
    Dim reg As New RegExp
    Dim ptA$, ptB$
    Dim sec%                    '���״̬���ж��Ƿ�Ϊ���α��
    Dim rg As Range
    Dim para As Paragraph
    Dim markPosition$           '��ͼ��Ƕο�ͷ����
    Dim boxX$, boxY$            '�Ի���λ��
    Dim hold$
    
    Set rg = ActiveDocument.Range
    ActiveDocument.TrackRevisions = True
    'ActiveWindow.View.RevisionsFilter.View = wdRevisionsViewFinal '���Ŀ���������±���Ϊ��ʾ����״̬�����������100%��Ч
    
    markPosition = "��ͼ��ǣ�"
    boxX = 12000
    boxY = 4000
    
    k = 1                   'kΪ���α�ŵĴ��������Զ���ҳ��أ�iΪ�ܸ�ͼ������������߶���
    numTape = ""
    techTape = ""
    rg.WholeStory
    totalTxtWet = rg.Text
    With reg
        .Global = True
        .Pattern = "��\w+��"
        '.Pattern = "(��\w+��)|([0-9]+[a-zA-Z]*)"
        totalTxtDry = .Replace(totalTxtWet, "")
    End With
    '���Ҹ�ͼ�������
    With rg.Find
        .ClearFormatting
        .Text = markPosition
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
        showTxt = ""
        sec = 1
        For n = 1 To i - 1
            showTxt = showTxt & num(n) & "-" & tech(n) & "��"
            numTape = numTape & " " & num(n) & " "
            techTape = techTape & " " & tech(n) & " "
        Next
        showTxt = Left(showTxt, Len(showTxt) - 1) & "��"
    Else
        ReDim tech(1)
        ReDim num(1)
        sec = 0
        i = 1
        showTxt = "���븽ͼ���,�����ƺ�����(��:Һ����10)��" & Chr(10) & "����'u'ҳ���Ϸ���" & Chr(10) & "����'d'ҳ���·���"
    End If
    num(0) = 0  '������һ����λ�жϵ�Ƕ��
    Selection.HomeKey unit:=wdStory
    
    '--------------------------------------------Ԥ�������-------------------------------------------------
    a = InputBox(showTxt, "��ͼ���", , boxX, boxY)
    '-------������-------
    Do While a <> ""
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
                        
                        showTxt = ""
                        For n = 1 To i
                            showTxt = showTxt & num(n) & "-" & tech(n) & "��"
                        Next
                        showTxt = Left(showTxt, Len(showTxt) - 1) & "��"
                        
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
        a = InputBox(showTxt, "��ͼ���", , boxX, boxY)
        rg.WholeStory
        totalTxtWet = rg.Text
    Loop

    i = i - 1
    k = k - 1
    
    If k = 0 Then
        Exit Sub
    Else
        With Selection
            With .Find
                .ClearFormatting
                .Text = markPosition
                .Wrap = wdFindContinue
                .Execute
            End With
            .Collapse wdCollapseEnd
            .TypeText showTxt
        End With
        
        If sec = 1 Then
            With Selection
                .MoveDown wdParagraph, 1, wdExtend
                .End = .End - 1
                .Delete
            End With
        End If
    End If
    '************���Ȩ��Ҫ�����У�ר�����⣨ר�����ƣ���ı��**************
    With rg
        With .Find
            .ClearFormatting
            .Text = "˵ �� ��^p"
            .MatchWildcards = False
            .Wrap = wdFindContinue
            .Execute
        End With
        .Start = 0
    End With
    
    For Each para In rg.Paragraphs
        para.Range.Select
        With Selection
            With .Find
                .ClearFormatting
                .Text = "����������"
                .Wrap = wdFindStop
                hold = .Execute
            End With
            If hold = True Then
                .Start = para.Range.Start
                With .Find
                    .ClearFormatting
                    .Text = "��*��"
                    With .Replacement
                        .Text = ""
                    End With
                    .MatchWildcards = True
                    .Wrap = wdFindStop
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End With
    Next
    
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
Sub markclear()
    Application.ScreenUpdating = False
'    ������ĵ���ҳ
    Selection.HomeKey unit:=wdStory
    'ѡ��˵��������
    Call multfind("˵ �� ��*��ͼ���")
    '�滻���
    Call multreplace("��*��", "")
    Selection.MoveRight
    'ѡ��˵����ʵʩ��
    Call multfind("������ʵʩ��*˵ �� �� ժ Ҫ")
    'ɾ������
    Call multreplace("��", "")
    Call multreplace("��", "")
    '�ĵ���ѡ������ĩβ
    Selection.MoveLeft
    Application.ScreenUpdating = True
End Sub
Function multfind(a)
    With Selection.Find
        .ClearFormatting
        .Text = a
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute
    End With
End Function
Function multreplace(a, b)
    With Selection.Find
        .ClearFormatting
        .Text = a
        With .Replacement
            .ClearFormatting
            .Text = b
        End With
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Function

