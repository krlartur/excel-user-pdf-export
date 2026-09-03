Option Explicit

Private Const CONFIG_SHEET As String = "Export_Config"

Sub EksportujRaportyUzytkownikow()

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim dataSheetName As String
    Dim userColumnHeader As String
    Dim targetFolder As String
    Dim filePrefix As String
    Dim users() As String
    Dim wantedHeaders() As String
    Dim allColumns As Boolean

    On Error GoTo ErrHandler

    Set wsConfig = ThisWorkbook.Worksheets(CONFIG_SHEET)

    dataSheetName = Trim(wsConfig.Range("B1").Value)
    userColumnHeader = Trim(wsConfig.Range("B2").Value)
    targetFolder = Trim(wsConfig.Range("B3").Value)
    filePrefix = Trim(wsConfig.Range("B4").Value)

    If dataSheetName = "" Or userColumnHeader = "" Then
        MsgBox "Uzupelnij B1 (arkusz danych) i B2 (kolumna z uzytkownikami) w arkuszu " & CONFIG_SHEET, vbExclamation
        Exit Sub
    End If

    If Not SheetExists(dataSheetName) Then
        MsgBox "Nie znaleziono arkusza danych: " & dataSheetName, vbCritical
        Exit Sub
    End If
    Set wsData = ThisWorkbook.Worksheets(dataSheetName)

    ' Folder docelowy - jesli pusty albo nie istnieje, poproz o wybor
    If targetFolder = "" Or Dir(targetFolder, vbDirectory) = "" Then
        targetFolder = WybierzFolder()
        If targetFolder = "" Then
            MsgBox "Nie wybrano folderu docelowego. Anulowano.", vbExclamation
            Exit Sub
        End If
        wsConfig.Range("B3").Value = targetFolder
    End If
    If Right(targetFolder, 1) <> "\" Then targetFolder = targetFolder & "\"

    ' Lista uzytkownikow (kolumna A od A7)
    users = WczytajListe(wsConfig, "A7")
    If Not IsListPopulated(users) Then
        MsgBox "Brak listy uzytkownikow w kolumnie A (od A7) arkusza " & CONFIG_SHEET, vbExclamation
        Exit Sub
    End If

    ' Lista naglowkow (kolumna C od C7) - "*" oznacza wszystkie kolumny
    wantedHeaders = WczytajListe(wsConfig, "C7")
    allColumns = (UBound(wantedHeaders) = LBound(wantedHeaders)) And (Trim(wantedHeaders(LBound(wantedHeaders))) = "*")

    Dim headerRow As Long: headerRow = 1
    Dim lastCol As Long: lastCol = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lastCol
        headerMap(Trim(CStr(wsData.Cells(headerRow, c).Value))) = c
    Next c

    If Not headerMap.Exists(userColumnHeader) Then
        MsgBox "Kolumna '" & userColumnHeader & "' nie istnieje w arkuszu " & dataSheetName, vbCritical
        Exit Sub
    End If
    Dim userColIdx As Long: userColIdx = headerMap(userColumnHeader)

    Dim exportCols() As Long
    Dim exportHeaders() As String
    If allColumns Then
        ReDim exportCols(1 To lastCol)
        ReDim exportHeaders(1 To lastCol)
        For c = 1 To lastCol
            exportCols(c) = c
            exportHeaders(c) = CStr(wsData.Cells(headerRow, c).Value)
        Next c
    Else
        Dim n As Long: n = UBound(wantedHeaders) - LBound(wantedHeaders) + 1
        ReDim exportCols(1 To n)
        ReDim exportHeaders(1 To n)
        Dim k As Long: k = 0
        Dim h As Variant
        For Each h In wantedHeaders
            Dim hh As String: hh = Trim(CStr(h))
            If hh <> "" Then
                If Not headerMap.Exists(hh) Then
                    MsgBox "Naglowek '" & hh & "' nie istnieje w arkuszu " & dataSheetName, vbCritical
                    Exit Sub
                End If
                k = k + 1
                exportCols(k) = headerMap(hh)
                exportHeaders(k) = hh
            End If
        Next h
        ReDim Preserve exportCols(1 To k)
        ReDim Preserve exportHeaders(1 To k)
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim u As Variant
    Dim exportedCount As Long: exportedCount = 0
    Dim skippedUsers As String

    For Each u In users
        Dim userName As String: userName = Trim(CStr(u))
        If userName <> "" Then
            Dim matchCount As Long
            matchCount = EksportujDlaUzytkownika(wsData, headerRow, lastRow, userColIdx, userName, exportCols, exportHeaders, targetFolder, filePrefix)
            If matchCount = 0 Then
                skippedUsers = skippedUsers & userName & vbNewLine
            Else
                exportedCount = exportedCount + 1
            End If
        End If
    Next u

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Wygenerowano PDF dla " & exportedCount & " uzytkownika(ow) w: " & targetFolder
    If skippedUsers <> "" Then
        msg = msg & vbNewLine & vbNewLine & "Brak danych (pominieto):" & vbNewLine & skippedUsers
    End If
    MsgBox msg, vbInformation

    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Blad: " & Err.Description, vbCritical

End Sub

Private Function EksportujDlaUzytkownika(wsData As Worksheet, headerRow As Long, lastRow As Long, _
    userColIdx As Long, userName As String, exportCols() As Long, exportHeaders() As String, _
    targetFolder As String, filePrefix As String) As Long

    Dim wsTemp As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets("__TempExport").Delete
    On Error GoTo 0
    Set wsTemp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsTemp.Name = "__TempExport"

    Dim colCount As Long: colCount = UBound(exportCols) - LBound(exportCols) + 1
    Dim outCol As Long
    For outCol = 1 To colCount
        wsTemp.Cells(1, outCol).Value = exportHeaders(LBound(exportHeaders) + outCol - 1)
    Next outCol
    wsTemp.Rows(1).Font.Bold = True

    Dim outRow As Long: outRow = 1
    Dim r As Long
    For r = headerRow + 1 To lastRow
        If Trim(CStr(wsData.Cells(r, userColIdx).Value)) = userName Then
            outRow = outRow + 1
            For outCol = 1 To colCount
                wsTemp.Cells(outRow, outCol).Value = wsData.Cells(r, exportCols(LBound(exportCols) + outCol - 1)).Value
            Next outCol
        End If
    Next r

    If outRow = 1 Then
        wsTemp.Delete
        EksportujDlaUzytkownika = 0
        Exit Function
    End If

    wsTemp.Columns.AutoFit
    With wsTemp.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = IIf(colCount > 6, xlLandscape, xlPortrait)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .PrintArea = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(outRow, colCount)).Address
    End With

    Dim fileName As String
    Dim prefixPart As String
    If Trim(filePrefix) <> "" Then prefixPart = SanitizeFileName(Trim(filePrefix)) & "_"
    fileName = prefixPart & SanitizeFileName(userName) & "_" & Format(Date, "mm_dd_yyyy") & ".pdf"

    wsTemp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=targetFolder & fileName, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

    wsTemp.Delete

    EksportujDlaUzytkownika = outRow - 1

End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim invalidChars As String: invalidChars = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(invalidChars)
        s = Replace(s, Mid(invalidChars, i, 1), "_")
    Next i
    SanitizeFileName = Trim(s)
End Function

Private Function WczytajListe(ws As Worksheet, startCell As String) As String()
    Dim rng As Range
    Set rng = ws.Range(startCell)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, rng.Column).End(xlUp).Row
    If lastRow < rng.Row Then
        Dim empty1(0) As String
        empty1(0) = ""
        WczytajListe = empty1
        Exit Function
    End If
    Dim n As Long: n = lastRow - rng.Row + 1
    Dim arr() As String
    ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CStr(ws.Cells(rng.Row + i, rng.Column).Value)
    Next i
    WczytajListe = arr
End Function

Private Function IsListPopulated(arr() As String) As Boolean
    Dim v As Variant
    For Each v In arr
        If Trim(CStr(v)) <> "" Then
            IsListPopulated = True
            Exit Function
        End If
    Next v
    IsListPopulated = False
End Function

Private Function SheetExists(ByVal shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function WybierzFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Wybierz folder docelowy dla plikow PDF"
    If fd.Show = -1 Then
        WybierzFolder = fd.SelectedItems(1)
    Else
        WybierzFolder = ""
    End If
End Function
