

Sub WORD__COLOR_GLOSSARY()

Dim PGH As Word.Paragraph

Dim content As Range
Dim content_text As String

Dim KEYWORD_COLOUR As Long
Dim TEMP_KEYWORD As String
Dim PARAGRAPH_TEXT As String
Dim FOUND_SUBSTRING As Integer

Dim glossary(1 To 10000) As String
Dim intCount As Integer
Dim FILE, GLOSSARY_KW As Variant
Dim GLOSSARY_FOLDER, FILE_NAME, FILE_PATH As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GLOSSARY_FOLDER = "PATH OF YOUR FOLDER HERE"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Application.ScreenUpdating = False
Application.DisplayStatusBar = True


For glossary_idx = 1 To 2              '' NUMBER of Glossaries
    Erase glossary

    Select Case glossary_idx
        Case 1
            FILE_NAME = "rgb_green.txt"
            KEYWORD_COLOUR = RGB(109, 216, 112)
        
        Case 2
            FILE_NAME = "rgb_red.txt"
            KEYWORD_COLOUR = RGB(255, 79, 175)
        
        Case Else
            MsgBox "check code"
    
    End Select

    FILE_PATH = GLOSSARY_FOLDER & "\" & FILE_NAME
    
    
    Set FSO_OBJ = CreateObject("Scripting.FileSystemObject")
    Set FILE = FSO_OBJ.OpenTextFile(FILE_PATH, 1)

    With FILE
        intCount = 1
        Do While .AtEndOfStream = False And intCount < 10001
            glossary(intCount) = .readline
            intCount = intCount + 1
        Loop
    
        .Close
    End With
   
    Set content = ActiveDocument.content
    content_text = content.text
    
    For Each GLOSSARY_KW In glossary
        If GLOSSARY_KW <> "" Then
        
            TEMP_KEYWORD = GLOSSARY_KW
            
            For Each PGH In ActiveDocument.Paragraphs
                PARAGRAPH_TEXT = PGH.Range.text
                FOUND_SUBSTRING = InStr(1, PARAGRAPH_TEXT, GLOSSARY_KW, 1)
                  
                If FOUND_SUBSTRING > 0 Then
                
                    Call CHANGE_KW_COLOUR(PGH.Range, _
                                                                    TEMP_KEYWORD, _
                                                                    KEYWORD_COLOUR)
                                
                End If
            
                FOUND_SUBSTRING = 0
                
            Next
            
            
        End If
    
    Next GLOSSARY_KW
    
    PARAGRAPH_TEXT = ""
    
          
    Application.StatusBar = "Finished glossary_idx = " & glossary_idx
            
    Documents.Save NoPrompt:=True, _
        OriginalFormat:=wdOriginalDocumentFormat
    
Next glossary_idx

Application.ScreenUpdating = True

MsgBox "Finished!"

End Sub



Private Sub CHANGE_KW_COLOUR(pghRange As Range, GLOSSARY_KW As String, KEYWORD_COLOUR As Variant)

    With pghRange.Find
        .text = GLOSSARY_KW
        .Forward = True
        .Format = False
        .MatchCase = False
        
        .MatchWholeWord = True ' Or False
        
        Do While .Execute()
            
            pghRange.Font.Color = KEYWORD_COLOUR
        ''    pghRange.Font.Bold = True
        ''    pghRange.Font.Underline = wdUnderlineSingle
        
        ''    pghRange.HighlightColorIndex = wdDarkYellow
            '' Additional highlight colors here:
            '' https://learn.microsoft.com/en-us/office/vba/api/word.wdcolorindex
                   
        Loop
    
    End With

End Sub
