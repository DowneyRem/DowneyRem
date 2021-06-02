Attribute VB_Name = "Macros"
Sub 小说排版()
Dim filename As String

'
' 小说排版 宏
'

    '第一行后添加一个空行
    'ActiveDocument.Paragraphs.Add _
    'Range:=ActiveDocument.Paragraphs(2).Range
    
    
    '在第一行后插入下面的内容
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
    Selection.TypeText Text:="作者："
    Selection.TypeParagraph
    Selection.TypeText Text:="网址："
    Selection.TypeParagraph
    Selection.TypeText Text:="标签："
    Selection.TypeParagraph
    Selection.TypeText Text:="其他："
    Selection.TypeParagraph
    Selection.TypeParagraph
    
    
    '全文设置成正文缩进的格式
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("正文缩进")
    
    
    '全角空格替换为空
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "　"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchByte = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '制表符替换为空
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
    
    
    '设置一级标题：第X卷
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 1")
    With Selection.Find
        .Text = "(第[0-9]@卷*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '在二级标题（第X章）前插入两个空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(第?章*)"
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
        .Text = "(第??章*)"
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
        .Text = "(第??章*)"
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
        .Text = "(第???章*)"
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
        .Text = "(第?????章*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '设置二级标题：第X章
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第[0-9]@章*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第?章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第??章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第???章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第????章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第?????章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第??????章*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第???????章*)"
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
        .Text = "(番外*)"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(番外*)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '前5行设置成正文样式
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=5, Extend:=wdExtend
    Selection.Style = ActiveDocument.Styles("正文")
    
    
    '第一行设置成标题样式
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Style = ActiveDocument.Styles("标题")
    
    
    '获取文件名，保存文件
    filename = ActiveDocument.Paragraphs(1).Range.Text
    filename = Left(filename, Len(filename) - 1) & ".docx"
    ChangeFileOpenDirectory "D:\Users\Administrator\Desktop\"
    ActiveDocument.SaveAs2 filename:=filename
    
    
    
End Sub
Sub 另存为()
Dim filename As String

' 另存为 宏
'
'
    '获取文件名，保存文件
    filename = ActiveDocument.Paragraphs(1).Range.Text
    filename = Left(filename, Len(filename) - 1) & ".docx"
    ChangeFileOpenDirectory "D:\Users\Administrator\Desktop\"
    ActiveDocument.SaveAs2 filename:=filename
    
    
End Sub
