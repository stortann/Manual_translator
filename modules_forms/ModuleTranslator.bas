'tools-references, then browse for and activate C:\Program Files (x86)\Microsoft Office\root\vfs\SystemX86\FM20.dll
Option Explicit
Private Sub test() 'for testing
'View-ImmediateWindow to see output
    
    Debug.Print "Hi!"
    
    Dim i As Integer
    For i = 0 To 100
        Debug.Print ""
    Next i
    
End Sub
Sub translator(Optional ByVal textGer As Variant)
'find translation to a word input in THIS workbook
'textGer is variant so this sub can be called from excel,
'because vba only allows activation of subs with arguments that are variant type
    
    Application.ScreenUpdating = False
    Load frmTranslator
    
    If Not checkWorkSheetExists("Dictionary") Then
        MsgBox "No Dictionary sheet, create it before running this macro"
        Exit Sub
    End If
    
    'if NOT first call then we have input textGer
    If Not IsMissing(textGer) Then
        textGer = invasiveClean(textGer)
        textGer = nonInvasiveClean(textGer)
        'check if we want Ger-Eng or Eng-Ger translation
        Dim reverse As Boolean
        If Left(textGer, 1) = "-" Then
            reverse = True
            textGer = replace(textGer, "-", "")
        Else
            reverse = False
        End If
        Dim textEng As String
        textEng = ""
        
        Dim word As Variant
        For Each word In Split(textGer, " ")
            
            textEng = textEng + " " + findWordTrans(word, reverse)
            
        Next word
        textEng = invasiveClean(textEng)
        textEng = nonInvasiveClean(textEng)
        
        placeToClipboard (textEng)
        
        Call frmTranslator.passText(CStr(textGer), textEng)

    End If
    
    Application.ScreenUpdating = True
    frmTranslator.Show (VBA.FormShowConstants.vbModeless)
    
End Sub
Sub worksheetCleaner()
'clean worksheet
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    
    Dim lastCol As Integer
    Dim lastRow As Integer
    Dim i As Integer, j As Integer
    Dim cell As String
    
    lastCol = ws.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    lastRow = ws.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    For i = 1 To lastCol
        For j = 1 To lastRow
            ws.Cells(j, i) = deleteCrossOut(ws.Cells(j, i))
            ws.Cells(j, i) = UCase(ws.Cells(j, i))
            ws.Cells(j, i) = invasiveClean(ws.Cells(j, i))
            ws.Cells(j, i) = nonInvasiveClean(ws.Cells(j, i))
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub
Sub updateDictionary()
'update already existing dict or create new

    Dim dictName As String
    dictName = "Dictionary"
    
    If Not checkWorkSheetExists(dictName) Then
        ThisWorkbook.Worksheets.Add Before:=Sheets(1)
        ActiveSheet.Name = dictName
    End If
        
    Dim wsDict As Worksheet
    Set wsDict = ThisWorkbook.Worksheets(dictName)
    
    wsDict.Range("A1:B1").Font.Bold = True
    wsDict.Range("A1").Value = "GERMAN"
    wsDict.Range("B1").Value = "ENGLISH"
    
    'we start filling dictionary from new row
    Dim dictLength As Integer
    dictLength = 1 + wsDict.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Dim wsNumb
    Dim i, j As Integer
    Dim ws As Worksheet
    Dim lastCol, lastRow As Integer
    Dim reverse As Boolean
    reverse = False
    Dim translate As String
    Dim gerText, engText As String
    
    For wsNumb = 2 To ThisWorkbook.Worksheets.Count
    
        Set ws = ThisWorkbook.Worksheets(wsNumb)
        lastCol = ws.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        lastRow = ws.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        For i = 1 To lastCol
            If nonInvasiveClean(ws.Cells(1, i)) = "GERMAN" Then
                
                For j = 2 To lastRow
                    gerText = ws.Cells(j, i)
                    engText = ws.Cells(j, i + 1)
                    
                    gerText = invasiveClean(gerText)
                    gerText = nonInvasiveClean(gerText)
                    
                    engText = invasiveClean(engText)
                    engText = nonInvasiveClean(engText)
                    
                    If Len(gerText) > 0 And Len(engText) > 0 Then
                        
                        Dim word As Variant
                        For Each word In Split(gerText, " ")
                                                        
                            Dim searchResult As Range
                            Set searchResult = wsDict.Range(wsDict.Cells(1, 1), wsDict.Cells(dictLength, 1)).Find(What:=word, LookIn:=xlValues, LookAt:=xlWhole)
            
                            If searchResult Is Nothing And Not IsNumeric(word) Then
                                wsDict.Cells(dictLength, 1).NumberFormat = "@"
                                wsDict.Cells(dictLength, 2).NumberFormat = "@"
                                wsDict.Cells(dictLength, 1).Interior.Color = RGB(255, 0, 0)
                                wsDict.Cells(dictLength, 2).Interior.Color = RGB(255, 0, 0)
                                wsDict.Cells(dictLength, 1) = word
                                wsDict.Cells(dictLength, 2) = engText
                                dictLength = dictLength + 1
                            End If
                        Next word
                    End If
                Next j
            End If
        Next i
    Next wsNumb
    
    wsDict.UsedRange.Columns.AutoFit
    
    Dim cellsWithoutHeader As Range
    Set cellsWithoutHeader = wsDict.Range(wsDict.Cells(1, 1), wsDict.Cells(dictLength, 2))
    cellsWithoutHeader.Sort key1:=Range(wsDict.Cells(2, 1), wsDict.Cells(dictLength, 1)), _
        order1:=xlAscending, Header:=xlYes
    
    wsDict.UsedRange.HorizontalAlignment = xlLeft
End Sub
Sub A_XXX_XXX_XX_XX()
'AXXXXXXXXXX ->A XXX XXX XX XX, uses func a14

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    
    Dim lastCol As Integer
    Dim lastRow As Integer
    Dim i As Integer, j As Integer
    Dim cell As String
    
    lastCol = ws.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    lastRow = ws.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    For i = 1 To lastCol
        For j = 1 To lastRow
            ws.Cells(j, i) = a14(ws.Cells(j, i))
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub
Function a14(ByVal text As String) As String
'AXXXXXXXXXX ->A XXX XXX XX XX
    
    Dim temp, temptext As String
    
    'replace star ChrW(42) with space ChrW(32)
    temptext = replace(text, ChrW(42), ChrW(32))
    
    'replace all spaces ChrW(32) with nothing
    temptext = replace(temptext, ChrW(32), "")
    
    Dim i As Integer
    i = 1
    
    If Len(temptext) = 11 Then
        If Mid(temptext, 1, 1) = "A" And IsNumeric(Mid(temptext, 2)) Then
            text = temptext
            temp = ""
            For i = 1 To Len(text)
                temp = temp + Mid(text, i, 1)
                Select Case i
                Case 1, 4, 7, 9
                    temp = temp + " "
                End Select
            Next i
            text = temp
        End If
    End If
    
    a14 = text
End Function
Sub AXXXXXXXXXX()
'A XXX XXX XX XX -> AXXXXXXXXXX, uses func a10

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.ActiveSheet
    
    Dim lastCol As Integer
    Dim lastRow As Integer
    Dim i As Integer, j As Integer
    Dim cell As String

    lastCol = ws.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    lastRow = ws.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    For i = 1 To lastCol
        For j = 1 To lastRow
            ws.Cells(j, i) = a10(ws.Cells(j, i))
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub
Function a10(ByVal text As String) As String
'A XXX XXX XX XX -> AXXXXXXXXXX
    
    Dim temptext As String
    
    'replace star ChrW(42) with nothing
    temptext = replace(text, ChrW(42), "")
    
    'replace all spaces ChrW(32) with nothing
    temptext = replace(temptext, ChrW(32), "")
    
    If Len(temptext) = 11 Then
        If Mid(temptext, 1, 1) = "A" And IsNumeric(Mid(temptext, 2)) Then
            text = temptext
        End If
    End If
    
    a10 = text
End Function
Function nonInvasiveClean(ByVal text As String) As String
'cleaning without deleting meaningful characters
    
    'uppercase
    text = UCase(text)

    'clear spaces at the start and finish
    text = Trim(text)
    
    'replace non-breaking-space with usual space
    text = replace(text, ChrW(160), ChrW(32))
    
    Dim i As Integer
    'replace from 2 to 5 spaces with 1 space
    For i = 5 To 2 Step -1
        text = replace(text, String(i, ChrW(32)), ChrW(32))
    Next i
    
    'replace As with umlaut chr=196 with AE ChrW(65) + ChrW(69)
    text = replace(text, ChrW(196), ChrW(65) + ChrW(69))
    
    'replace Os with umlaut chr=214 with O ChrW(79)
    text = replace(text, ChrW(214), ChrW(79))
    
    'replace Us with umlaut chr=220 with UE ChrW(85) + ChrW(69)
    text = replace(text, ChrW(220), ChrW(85) + ChrW(69))

    'replace greek doubleSS chr=946 with SS ChrW(83) + ChrW(83)
    text = replace(text, ChrW(946), ChrW(83) + ChrW(83))
    
    'replace german doubleSS chr=223 with SS ChrW(83) + ChrW(83)
    text = replace(text, ChrW(223), ChrW(83) + ChrW(83))
    
    nonInvasiveClean = text
    
End Function
Function invasiveClean(ByVal text As String) As String
'possibly deleting meaningful symbols

    Dim i As Integer
    
    'delete (){}[]
    Dim chars As String
    chars = "(){}[]"
    For i = 1 To Len(chars)
        text = replace(text, Mid(chars, i, 1), "")
    Next i
    
    'replace underscore ChrW(95) with space ChrW(32)
    text = replace(text, ChrW(95), ChrW(32))
    
    'replace star ChrW(42) with space ChrW(32)
    text = replace(text, ChrW(42), ChrW(32))
    
    invasiveClean = text

End Function
Function placeToClipboard(Optional ByVal StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

    Dim x As Variant
    'Store as variant for 64-bit VBA support
    x = StoreText
    'Create HTMLFile Object
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    'Write to the clipboard
                    .setData "text", x
                Case Else
                    'Read from the clipboard (no variable passed through)
                    placeToClipboard = .GetData("text")
            End Select
        End With
    End With
    
End Function
Function findWordTrans(ByVal textInput As String, ByVal reverse As Boolean) As String
'try to find translation of textInput
'reverse=True=Eng-Ger
'if find textInput give the answer from first column to the right to textOutput
'if nothing there leave as is
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dictionary")
    Dim lastCol As Integer, lastRow As Integer
    Dim textOutputRange As Range
    Dim textOutputAddress(0 To 1) As String '(A,1)
    Dim textOutput As String
    textOutput = textInput
    
    lastCol = ws.UsedRange.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    lastRow = ws.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Set textOutputRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)). _
    Find(textInput, SearchDirection:=xlNext, SearchOrder:=xlColumns, _
    LookAt:=xlWhole, LookIn:=xlValues)
    
    If Not textOutputRange Is Nothing Then
        
        textOutputAddress(0) = Split(textOutputRange.Address, "$")(1)
        textOutputAddress(1) = Split(textOutputRange.Address, "$")(2)
        'search to the left and to the right, just in case
        Dim toTheRight As String
        toTheRight = ws.Cells(textOutputAddress(1), _
        Range(textOutputAddress(0) & 1).Column + 1)
        
        If (Range(textOutputAddress(0) & 1).Column - 1) >= 1 And reverse = True Then
            Dim toTheLeft As String
            toTheLeft = ws.Cells(textOutputAddress(1), _
            Range(textOutputAddress(0) & 1).Column - 1)
            If toTheLeft <> "" Then
                textOutput = toTheLeft
            End If
        End If
        
        If toTheRight <> "" And reverse = False Then
            textOutput = toTheRight
        End If
        
    End If
    
    findWordTrans = textOutput
    
End Function
Function deleteCrossOut(ByVal cl As Range) As String
'delete strikethrough text
    
    Dim i As Integer
    Dim newText As String
    newText = ""

    For i = 1 To Len(cl)
        If cl.Characters(i, 1).Font.Strikethrough = False Then
            newText = newText + Mid(cl, i, 1)
        End If
    Next i
    deleteCrossOut = newText
    
End Function
Function checkWorkSheetExists(wsName As String) As Boolean
'does a WS with this name exists?

    checkWorkSheetExists = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Application.Proper(ws.Name) = wsName Then
            checkWorkSheetExists = True
            Exit Function
        End If
    Next ws
    
End Function
