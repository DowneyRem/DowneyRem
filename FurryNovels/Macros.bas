Attribute VB_Name = "Macros"
Sub 小说排版()
Dim filename As String

'
' 小说排版 宏
'

    '第一行后添加一个空行
    'ActiveDocument.Paragraphs.Add _
    'Range:=ActiveDocument.Paragraphs(2).Range
    
    
    '基本文本插入
    '在第一行后插入下面的内容
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="作者："
    Selection.TypeParagraph
    Selection.TypeText Text:="网址："
    Selection.TypeParagraph
    Selection.TypeText Text:="标签："
    Selection.TypeParagraph
    'Selection.TypeText Text:="其他："
    'Selection.TypeParagraph
    Selection.TypeParagraph
    
    
    
    '基本格式设置1
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
    
    
    '2个半角空格替换为空
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchByte = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '制表符转半角空格
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
    
    
    '——————分割线转空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "—{3,20}"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '------------分割线转空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "-{2,20}"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
      
    'Pixiv标记处理
    '[newpage]转空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[newpage\]"
        .Replacement.Text = "^13^13^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '[pixivimage:作品ID]转文字链接
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[pixivimage:(*)\]"
        .Replacement.Text = "^13图片链接：www.pixiv.net/artworks/\1^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
     End With
     Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '[[jumpuri:标题 > 链接目标的URL]] 转文字链接
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[\[jumpuri:(*) \> (*)\]\]"
        .Replacement.Text = "^13【\1】链接：\2^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        
    '[jump:链接目标的页面编号]替换为空
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[jump:(*)\]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    ' [chapter:五]转章节标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[chapter:([序,引,零,〇,一,二,三,四,五,六,七,八,九,十,终])\]"
        .Replacement.Text = "第\1章"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    ' [chapter:第五章][chapter:第5章 1213]转章节标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[chapter:(第*章*)\]"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
        
    '单行中文数字处理
    '十以内数字，转章节标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13([序,引,零,〇,一,二,三,四,五,六,七,八,九,十])^13"
        .Replacement.Text = "第\1章"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '二十以内数字，转章节标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13(十[一,二,三,四,五,六,七,八,九])^13"
        .Replacement.Text = "第\1章"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '一百以内数字，转章节标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13([一,二,三,四,五,六,七,八,九]十[一,二,三,四,五,六,七,八,九])^13"
        .Replacement.Text = "第\1章"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '统一插入空行
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
    
    
    '汉字【第一千章】前，插入空行
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
    
    
    '前言后记等插入空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(前言*^13)"
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
        .Text = "(番外*^13)"
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
        .Text = "(彩蛋*^13)"
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
        .Text = "(后记*^13)"
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
        .Text = "(作者的话*)^13"
        .Replacement.Text = "^13^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     
     
    '统一设置标题
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
    
    
    '第9.5章，设置标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第[0-9]@.[0-9]章*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    '中文序号，设置标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(第?[章,部]*^13)"
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
        .Text = "(第??[章,部]*^13)"
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
        .Text = "(第???章*^13)"
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
        .Text = "(第????章*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    '前言后记等设置标题
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("标题 2")
    With Selection.Find
        .Text = "(前言*^13)"
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
        .Text = "(序^13)"
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
        .Text = "(序章*^13)"
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
        .Text = "(引^13)"
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
        .Text = "(番外*^13)"
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
        .Text = "(彩蛋*^13)"
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
        .Text = "(后记*^13)"
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
        .Text = "(跋^13)"
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
        .Text = "(作者的话*^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
    '基本格式设置2
    '去除多余的空行
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("正文缩进")
    With Selection.Find
        .Text = "^13{4,20}"
        .Replacement.Text = "^13^13^13"
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
