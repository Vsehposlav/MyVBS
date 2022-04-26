'Option Explicit

    Dim nameForma As String
    Dim nameGroup As String
    Dim nameFacultet As String
    Dim nameFamily As String
    Dim nameFirst As String
    Dim nameOtchestvo As String
    Dim beforeFamily As String
    Dim nameKurs As String
    Dim nameSpec As String
    Dim nameLang As String
    Dim nameLD As String ' личное дело
    Dim TableHeader As Range
    

Private Sub ExitButton_Click()
 End
End Sub

Private Sub GoButton_Click()
    Application.ScreenUpdating = False
    Call ProcessFIO(nameFamily, beforeFamily)
    Call ProcessFIO(nameFirst, beforeFamily)
    Call ProcessFIO(nameOtchestvo, beforeFamily)
    Call ProcessFacultetAndForma(nameFacultet, nameForma)
    Call ProcessSpec(nameSpec, nameForma)
    'Call ProcessGroup(nameGroup)
    Call ProcessKurs(nameKurs, SetOsenniy, SetVesenniy)
    Call ProcessLang(nameLang, beforeFamily)
    Call ProcessLD(nameLD)
    MsgBox "Готово! :)"
    Application.ScreenUpdating = True
    End
End Sub

Private Sub UserForm_Initialize()
'On Error GoTo err_debug

'MsgBox "UserForm_Initialize"
    If SetOsenniy.Value = SetVesenniy.Value Then
        FrameSemestr.ForeColor = &HFF&
        GoButton.Enabled = False
    End If
    
    nameForma = getColNameByVal("Форма")
    
        Call Init_Checkbox(CheckForma, nameForma)
    nameGroup = getColNameByVal("Группа")
        Call Init_Checkbox(CheckGroup, nameGroup)
    nameFacultet = getColNameByVal("Фак.")
        Call Init_Checkbox(CheckFacultet, nameFacultet)
    nameFamily = getColNameByVal("Фамилия")
        Call Init_Checkbox(CheckFamily, nameFamily)
    nameFirst = getColNameByVal("Имя")
        Call Init_Checkbox(CheckFirstName, nameFirst)
    nameOtchestvo = getColNameByVal("Отчество")
        Call Init_Checkbox(CheckOtchestvo, nameOtchestvo)
    beforeFamily = getColNameByVal("Предыдущие ФИО")
        Call Init_Checkbox(CheckBeforeFamily, beforeFamily)
    nameKurs = getColNameByVal("Курс")
        Call Init_Checkbox(CheckKurs, nameKurs)
    nameSpec = getColNameByVal("Спец.")
        Call Init_Checkbox(CheckSpec, nameSpec)
    nameLang = getColNameByVal("Язык")
        Call Init_Checkbox(CheckLang, nameLang)
    nameLD = getColNameByVal("№ л/д")
        Call Init_Checkbox(CheckLD, nameLD)
       
    
'err_debug:
   '  MsgBox Err.Number & ": " & Err.Description & " on line " & Erl, vbError
   
End Sub

Private Sub SetVesenniy_Change()
    GoButton.Enabled = True
End Sub

Private Sub SetOsenniy_Change()
    GoButton.Enabled = True
End Sub


Private Function Init_Checkbox(CheckBox As MSForms.CheckBox, name As String)
    If name = "" Then
            CheckBox.BackColor = &HFF&
            GoButton.Enabled = False
            FrameSemestr.Enabled = False
        Else
            CheckBox.BackColor = &HC000&
            CheckBox.Value = True
            CheckBox.Caption = CheckBox.Caption & " - " & name
        End If
End Function

Private Sub ProcessFacultetAndForma(nameFacultet As String, nameForma As String)
    Dim col As Range
    Dim cell As Range
   ' Dim nameFacultet As String: nameFacultet = getColNameByVal("Факультет")
   ' Dim nameForma As String: nameForma = getColNameByVal("Форма")
   ' Dim nameGroup As String: nameGroup = getColNameByVal("Группа")
   ' Dim nameForma As String: nameFacultet = getColNameByVal("Группа")
    
    
    Set col = Range(nameFacultet + ":" + nameFacultet)
    Set colForma = Range(nameForma + ":" + nameForma)
    Set colGroup = Range(nameGroup + ":" + nameGroup)
    'Set col = Range("G:G")
    
    For Each cell In col.Cells
        
        Select Case cell.Value
            Case "Пед."
                cell.Value = "ПЕД"
            Case "Леч."
                cell.Value = "ЛЕЧ"
                     If colForma.Cells(cell.Row, 1).Value = "7 лет, оч." Then cell.Value = "ЛЕЧВЕЧ"
                     If InStr(1, colGroup.Cells(cell.Row, 1).Value, "-а", vbTextCompare) > 0 Then cell.Value = "ЛЕЧ-ИНОСТР"
            Case "Фарм."
                cell.Value = "ФАРМ"
                    If colForma.Cells(cell.Row, 1).Value = "заоч." Then cell.Value = "ЗАОФАРМ"
            Case "Мед-проф."
                cell.Value = "МЕДПР"
            Case "Стом."
                cell.Value = "СТОМ"
                    If InStr(1, colGroup.Cells(cell.Row, 1).Value, "-а", vbTextCompare) > 0 Then cell.Value = "СТОМ-ИНОСТР"
            Case "Инст. сестр."
                cell.Value = "ВСОД"
        End Select
        
    Next cell
         
    Set col = Range(nameForma + ":" + nameForma)
    'Set col = Range("F:F")
    For Each cell In col.Cells
    
        Select Case cell.Value
            Case "оч."
                cell.Value = "д/о"
            Case "7 лет, оч."
                cell.Value = "в/о"
            Case "заоч."
                cell.Value = "з/о"
        End Select
    Next cell
End Sub

Private Sub ProcessGroup(nameGroup As String)
    Set col = Range(nameGroup + ":" + nameGroup)
    'Set col = Range("J:J")
        
        col.Replace "/ППД", "П"
        col.Replace "/ЛЛД", "Л"
        col.Replace "/ММД", "МП"
        col.Replace "/ФФД", "ФД"
        col.Replace "/ФФЗ", "ФЗ"
        col.Replace "/ФХД", "Б"
        col.Replace "/ЛЛВ", "ЛВ"
        col.Replace "/ССД", "СТ"
        col.Replace "/ИСЗ", "СДЗ"
End Sub

Private Sub ProcessKurs(nameKurs As String, SetOsenniy As MSForms.OptionButton, SetVesenniy As MSForms.OptionButton)
    Set col = Range(nameKurs + ":" + nameKurs)
    'Set col = Range("K:K")
        If SetOsenniy.Value = True Then
            For Each cell In col.Cells
            
                Select Case cell.Value
                    Case "1"
                        cell.Value = "1"
                    Case "2"
                        cell.Value = "3"
                    Case "3"
                        cell.Value = "5"
                    Case "4"
                        cell.Value = "7"
                    Case "5"
                        cell.Value = "9"
                    Case "6"
                        cell.Value = "11"
                    Case "7"
                        cell.Value = "13"
                    Case "Курс"
                        cell.Value = "Семестр"
                End Select
            Next cell
        End If
        
        If SetVesenniy.Value = True Then
            For Each cell In col.Cells
            
                Select Case cell.Value
                    Case "1"
                        cell.Value = "2"
                    Case "2"
                        cell.Value = "4"
                    Case "3"
                        cell.Value = "6"
                    Case "4"
                        cell.Value = "8"
                    Case "5"
                        cell.Value = "10"
                    Case "6"
                        cell.Value = "12"
                    Case "7"
                        cell.Value = "14"
                    Case "Курс"
                        cell.Value = "Семестр"
                End Select
            Next cell
        End If
End Sub

Private Sub ProcessSpec(nameSpec As String, nameForma As String)
    Set col = Range(nameSpec + ":" + nameSpec)
    Set colForma = Range(nameForma + ":" + nameForma)
    'Set col = Range("E:E")
        For Each cell In col.Cells
                If cell.Value <> "" Then cell.NumberFormat = "@"
            Select Case cell.Value
                Case "ЛД"
                     If colForma.Cells(cell.Row, 1).Value = "в/о" Then
                        cell.Value = "060101.01"
                        Else
                        cell.Value = "31.05.01"
                    End If
                Case "Пед"
                    cell.Value = "31.05.02"
                Case "МПД"
                    cell.Value = "32.05.01"
                Case "Стом"
                    cell.Value = "31.05.03"
                Case "Фарм"
                    If colForma.Cells(cell.Row, 1).Value = "з/о" Then
                        cell.Value = "060108.02"
                        Else
                        cell.Value = "33.05.01"
                    End If
                Case "МБХ"
                    cell.Value = "30.05.01"
                Case "Сест.д"
                    If colForma.Cells(cell.Row, 1).Value = "з/о" Then
                        cell.Value = "060109.02"
                        Else
                        cell.Value = "34.03.01"
                    End If
                Case "Фарм.(2в)"
                    If colForma.Cells(cell.Row, 1).Value = "з/о" Then
                        cell.Value = "060108.02"
                        Else
                        cell.Value = "33.05.01"
                    End If
            End Select
        Next cell
End Sub

Private Sub ProcessFIO(nameFIO As String, beforeFIO As String)
   Rem Set col = Range("L:L")
   'Dim name As String: name = getColNameByVal("Фамилия")
   Set col = Range(nameFIO + ":" + nameFIO)
   Set colBeforeFIO = Range(beforeFIO + ":" + beforeFIO)
   col.Replace "Ё", "Е", , , MatchCase:=True
   col.Replace "ё", "е", , , MatchCase:=True
   Dim strPattern As String: strPattern = "(\s*(\(+([A-Za-zА-Яа-яЁё\-\s]+)\s*\)+))"
   Dim strReplace As String: strReplace = ""
   Dim regex As New RegExp
   Dim newval, sourcestr As String
   With regex
        .Pattern = strPattern
   End With
   
        For Each cell In col.Cells
           If regex.Test(cell.Value) Then
                Set result = regex.Execute(cell.Value)
                newval = result(0).SubMatches(2)
                If colBeforeFIO.Cells(cell.Row, 1).Value = "" Then
                    colBeforeFIO.Cells(cell.Row, 1).Value = newval
                Else
                    colBeforeFIO.Cells(cell.Row, 1).Value = colBeforeFIO.Cells(cell.Row, 1).Value & " " & newval
                End If
                cell.Value = regex.Replace(cell.Value, strReplace)
                Rem MsgBox (regex.Replace(cell.Value, strReplace))
            End If
        Next cell
        
Exit Sub
ErrorHandler:
Select Case Err
    Case 65404: MsgBox ("выполнение прервано: требуемый столбец '" + Err.Description + "' не найден")

End Select
    
End Sub

Function getColNameByVal(val As String)
    
    
    'Dim cell As Range
    Dim header As Range
    Set header = Worksheets("Лист1").Rows(1)
    For Each cell In header.Cells
        If cell.Value = val Then
            colname = Split(cell.Address(True, False), "$")
            Rem MsgBox (cell.Value + " - value, " + colname(0) + " - имя столбца")
            getColNameByVal = colname(0)
            Exit Function
        End If
    Next cell
    getColByNameByVal = ""
'Err.Number = 65404
'Err.Description = val
'Err.Raise (65404)

End Function

Private Sub ProcessLang(nameLang As String, beforeFIO As String)

    Set col = Range(nameLang + ":" + nameLang)
    col.Replace "Английский", "eng"
    col.Replace "Немецкий", "ger"
    col.Replace "Французский", "fre"
    col.Replace "Другой", ""
    col.Replace "Не изучал", ""
    Set Anchor = Range(beforeFIO + ":" + beforeFIO)
    Set prev = Range(nameLang + ":" + nameLang)
    col.Cut (Anchor.Offset(0, 1).Cells(1, 1))
    
    Columns(nameLang).Delete

End Sub

Private Sub ProcessLD(nameLD As String)

    Set col = Range(nameLD + ":" + nameLD)
    col.Replace "=-и", ""
    For Each cell In col.Cells
        Select Case cell.Value
            Case "-и"
                cell.Value = ""
            Case "/"
                cell.Value = ""
            Case "/ЛЛД"
                cell.Value = ""
        End Select
    Next cell

End Sub

Sub microtest()
  On Error GoTo ErrorHandler
    Dim name As String: name = getColNameByVal("Страна")
    MsgBox (name)

Exit Sub
ErrorHandler:
Select Case Err
    Case 65404: MsgBox ("выполнение прервано: требуемый столбец '" + Err.Description + "' не найден")

End Select
End Sub
