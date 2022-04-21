Public oWord As Object
Public ResultDoc As Object
Public oDoc As Document
Public Results() As String
Public x As Integer



Sub ukazatel()
'
' ukazatel Ìàêðîñ
'
'
Set oDoc = ActiveDocument
ReDim Families(0) As String
'Dim y As Integer
x = 0
'y = 0
Dim text As String
Dim TargetString As String
Dim TargetLength As Integer

ReDim Results(0) As String

With oDoc.Range.Find
    .text = "([À-ÿ]{1}.) ([À-ÿ]{1}.) (<[À-ÿ]{2;}>)([,.;]{1})"
    .MatchWildcards = True
    While .Execute
'        MsgBox (.Parent.Text)
        
'        WriteText (.Parent.Text)
'        MsgBox .Parent.Paragraphs(1).Range.ListFormat.ListString
        TargetString = .Parent.text
        TargetLength = Len(TargetString)
        TargetString = Left(TargetString, TargetLength - 1)
        
        Control = DuplicateControl(TargetString, Families())
        If Control = True Then
            FindEntry (TargetString)
            Families(x) = TargetString
            x = x + 1
            ReDim Preserve Families(x)
        End If
'        If x = 100 Then
'
'
'
'        End If
               
        
    Wend
    
    Set oWord = CreateObject("Word.Application")
    Set ResultDoc = oWord.Documents.Add
    oWord.Visible = True
    For Each Entry In Results
        text = text & Entry & vbCrLf
    Next
    
    With ResultDoc.Range
        .text = .text & text
    End With
    
    ResultDoc.Save
    MsgBox ("complete")
End With


End Sub

Private Function WriteText(myText As String)
    
    Results(x) = myText
    ReDim Preserve Results(x + 1)

    
End Function

Private Function FindEntry(myText As String)
       Dim resultText As String
       resultText = myText & " -"
       
    With oDoc.Range.Find
        .text = myText
        .MatchWildcards = True
        While .Execute
            resultText = resultText & " " & .Parent.Paragraphs(1).Range.ListFormat.ListString
        Wend
    End With
    WriteText (resultText)

End Function

Function DuplicateControl(myText As String, ByRef Entries() As String) As Boolean
        DuplicateControl = True
    For Each Entry In Entries
        If StrComp(Entry, myText, vbTextCompare) = 0 Then
            DuplicateControl = False
            Exit Function
        End If
    Next
    
End Function
