Attribute VB_Name = "NewMacros"

Public Sub exportWordComments_Click()

    FileName = Application.ActiveDocument
    
    varResult = VBA.Split(FileName, ".")
    FileNameStr = varResult(0)
    
    Path = Application.ActiveDocument.Path
    FilePath = Path & "\" & FileName
    LogPath = Path & "\" & FileNameStr & "_comments.txt"
    'LogPath = ".\comments.txt"
    Debug.Print (FilePath)
    If FileName = "False" Then
        Exit Sub
    End If
    
    Rows = ActiveDocument.Comments.Count
    'Debug.Print (Rows)
    
    Open LogPath For Output As #1
    Print #1, "==================================================="
    For i = 1 To Rows
        PageNumber = ActiveDocument.Comments(i).Scope.Information(wdActiveEndPageNumber) '批注在第几页
        CharacterLineNumber = ActiveDocument.Comments(i).Scope.Information(wdFirstCharacterLineNumber) '批注在这页的第几行
        Scope = ActiveDocument.Comments(i).Scope '批注原文
        ScopeComment = ActiveDocument.Comments(i).Range '批注内容
        ScopeDate = ActiveDocument.Comments(i).Date  '批注时间
        ScopeAuthor = ActiveDocument.Comments(i).Contact '批注作者
        ScopeDone = ActiveDocument.Comments(i).Done '批注是否被解决
        
        'Debug.Print ("原文：" & ActiveDocument.Comments(i).Scope) '原文
        'Debug.Print (ActiveDocument.Comments(i).Done)
        'Debug.Print (ActiveDocument.Comments(i).Contact)
        'Debug.Print (ActiveDocument.Comments(i).Creator)
        'Debug.Print (ActiveDocument.Comments(i).Date)
        'Debug.Print (ActiveDocument.Comments(i).Index)
        'Debug.Print (ActiveDocument.Comments(i).Parent)
        'Debug.Print (ActiveDocument.Comments(i).Reference)
        'Debug.Print ("批注内容：" & ActiveDocument.Comments(i).Range) '批注内容
        'Debug.Print (ActiveDocument.Comments(i).IsInk)'是否包含链接
        Print #1, "GET_FILENAME: " & FileNameStr
        Print #1, "GET_FILEPATH: " & FilePath
        Print #1, "GET_PAGE: " & PageNumber
        Print #1, "GET_LINE: " & CharacterLineNumber
        Print #1, "GET_TXT: " & Scope
        Print #1, "GET_COMMENTS: " & ScopeComment
        Print #1, "GET_DATE: " & ScopeDate
        Print #1, "GET_AUTHOR: " & ScopeAuthor
        Print #1, "GET_DONE: " & ScopeDone
        Print #1, "==================================================="
    Next

    Print #1, ""
    Close #1
    
End Sub
