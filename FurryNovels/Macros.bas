Attribute VB_Name = "Macros"
Sub С˵�Ű�()
Dim filename As String

'
' С˵�Ű� ��
'

    '��һ�к����һ������
    'ActiveDocument.Paragraphs.Add _
    'Range:=ActiveDocument.Paragraphs(2).Range
    
    
    '�ڵ�һ�к�������������
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
    Selection.TypeText Text:="���ߣ�"
    Selection.TypeParagraph
    Selection.TypeText Text:="��ַ��"
    Selection.TypeParagraph
    Selection.TypeText Text:="��ǩ��"
    Selection.TypeParagraph
    Selection.TypeText Text:="������"
    Selection.TypeParagraph
    Selection.TypeParagraph
    
    
    'ȫ�����ó����������ĸ�ʽ
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("��������")
    
    
    'ȫ�ǿո��滻Ϊ��
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchByte = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '�Ʊ���滻Ϊ��
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchByte = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '����һ�����⣺��X��
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 1")
    With Selection.Find
        .Text = "(��[0-9]@��*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '�ڶ������⣨��X�£�ǰ������������
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(��?��*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(��??��*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(��??��*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(��???��*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(��?????��*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '���ö������⣺��X��
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��[0-9]@��*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��?��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��??��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��???��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��????��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��?????��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��??????��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(��???????��*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(����*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("���� 2")
    With Selection.Find
        .Text = "(����*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'ǰ5�����ó�������ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=5, Extend:=wdExtend
    Selection.Style = ActiveDocument.Styles("����")
    
    
    '��һ�����óɱ�����ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Style = ActiveDocument.Styles("����")
    
    
    '��ȡ�ļ����������ļ�
    filename = ActiveDocument.Paragraphs(1).Range.Text
    filename = Left(filename, Len(filename) - 1) & ".docx"
    ChangeFileOpenDirectory "D:\Users\Administrator\Desktop\"
    ActiveDocument.SaveAs2 filename:=filename
    
    
    
End Sub
Sub ���Ϊ()
Dim filename As String

' ���Ϊ ��
'
'
    '��ȡ�ļ����������ļ�
    filename = ActiveDocument.Paragraphs(1).Range.Text
    filename = Left(filename, Len(filename) - 1) & ".docx"
    ChangeFileOpenDirectory "D:\Users\Administrator\Desktop\"
    ActiveDocument.SaveAs2 filename:=filename
    
    
End Sub
