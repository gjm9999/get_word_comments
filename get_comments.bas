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
        PageNumber = ActiveDocument.Comments(i).Scope.Information(wdActiveEndPageNumber) '��ע�ڵڼ�ҳ
        CharacterLineNumber = ActiveDocument.Comments(i).Scope.Information(wdFirstCharacterLineNumber) '��ע����ҳ�ĵڼ���
        Scope = ActiveDocument.Comments(i).Scope '��עԭ��
        ScopeComment = ActiveDocument.Comments(i).Range '��ע����
        ScopeDate = ActiveDocument.Comments(i).Date  '��עʱ��
        ScopeAuthor = ActiveDocument.Comments(i).Contact '��ע����
        ScopeDone = ActiveDocument.Comments(i).Done '��ע�Ƿ񱻽��
        
        'Debug.Print ("ԭ�ģ�" & ActiveDocument.Comments(i).Scope) 'ԭ��
        'Debug.Print (ActiveDocument.Comments(i).Done)
        'Debug.Print (ActiveDocument.Comments(i).Contact)
        'Debug.Print (ActiveDocument.Comments(i).Creator)
        'Debug.Print (ActiveDocument.Comments(i).Date)
        'Debug.Print (ActiveDocument.Comments(i).Index)
        'Debug.Print (ActiveDocument.Comments(i).Parent)
        'Debug.Print (ActiveDocument.Comments(i).Reference)
        'Debug.Print ("��ע���ݣ�" & ActiveDocument.Comments(i).Range) '��ע����
        'Debug.Print (ActiveDocument.Comments(i).IsInk)'�Ƿ��������
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
