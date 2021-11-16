Sub 另存为TXT()
Dim file As String
Dim filename As String


    '保存一次文件
    ActiveDocument.Save
    file = ActiveDocument.FullName
    
    
    '获取文件名，保存TXT
    filename = ActiveDocument.Paragraphs(1).Range.Text
    filename = Left(filename, Len(filename) - 1) & ".txt"
    ChangeFileOpenDirectory "C:\Users\Administrator\Desktop\"
    ActiveDocument.SaveAs2 filename:=filename, FileFormat:=wdFormatText
    
    
    '关闭文档,关闭软件
    ActiveDocument.Close 0
    Documents.Open filename:=file
    ActiveDocument.Save
    'Application.Quit
    
    
End Sub
