Option Explicit
' Data Quality Check Macro (Generic Example)
' Original Macro by another author, extended and adapted by Christian S.
' Public Showcase Version (fields anonymized for demonstration purposes)

' Konstanten für Blattnamen
Private Const REPORT_SHEET_NAME As String = "Report"
Private Const DATA_SHEET_NAME As String = "Data"

Sub DataQualityCheck()
    ' Neues Tabellenblatt für den Report anlegen
    Worksheets.Add
    ActiveSheet.Name = REPORT_SHEET_NAME

    ' Kopfzeilen für den Report definieren
    Dim reportHeaders As Variant
    reportHeaders = Array("First Name", "Last Name", _
                          "Field A missing", "Field B missing", "Field C missing", _
                          "Field D missing", "Field E missing", "Field F missing", _
                          "Field G missing", "Field H missing")

    Dim i As Integer
    For i = LBound(reportHeaders) To UBound(reportHeaders)
        Worksheets(REPORT_SHEET_NAME).Cells(1, i + 1).Value = reportHeaders(i)
    Next i

    ' Startzeile in der Datentabelle finden (Marker: "ID")
    Dim z As Integer, y As Integer
    Do
        z = z + 1
    Loop Until Worksheets(DATA_SHEET_NAME).Cells(z, 1).Value = "ID"
    y = z + 1

    ' Spaltenköpfe in der Datentabelle suchen
    Dim columnHeaders As Variant
    columnHeaders = Array("First Name*", "Last Name", "Field A", "Field B", "Field C", _
                          "Field D", "Field E", "Field F", "Field G", "Field H")

    Dim columnIndices() As Integer
    ReDim columnIndices(LBound(columnHeaders) To UBound(columnHeaders))

    Dim Counter As Integer
    Do
        Counter = Counter + 1
        For i = LBound(columnHeaders) To UBound(columnHeaders)
            If Worksheets(DATA_SHEET_NAME).Cells(z, Counter).Value = columnHeaders(i) Then
                columnIndices(i) = Counter
            End If
        Next i
    Loop Until Worksheets(DATA_SHEET_NAME).Cells(z, Counter).Value = ""

    ' Datenzeilen prüfen und Report schreiben
    Dim x As Integer, reportRow As Integer
    x = 0
    reportRow = 2

    Do
        Dim hasIssue As Boolean
        hasIssue = False

        ' Alle Pflichtfelder prüfen
        For i = 2 To UBound(columnHeaders)
            If columnIndices(i) > 0 Then
                If Worksheets(DATA_SHEET_NAME).Cells(y + x, columnIndices(i)).Value = "" Then
                    Worksheets(REPORT_SHEET_NAME).Cells(reportRow, i + 1).Value = "X"
                    hasIssue = True
                End If
            End If
        Next i

        ' Falls Fehler gefunden wurden → Namen in den Report schreiben
        If hasIssue Then
            Worksheets(DATA_SHEET_NAME).Cells(y + x, columnIndices(0)).Copy _
                Destination:=Worksheets(REPORT_SHEET_NAME).Cells(reportRow, 1)
            Worksheets(DATA_SHEET_NAME).Cells(y + x, columnIndices(1)).Copy _
                Destination:=Worksheets(REPORT_SHEET_NAME).Cells(reportRow, 2)
            reportRow = reportRow + 1
        End If

        ' Nächste Zeile prüfen
        x = x + 1
    Loop Until Worksheets(DATA_SHEET_NAME).Cells(y + x, columnIndices(0)).Value = ""

    ' Report formatieren
    With Worksheets(REPORT_SHEET_NAME)
        .Columns("A:" & Chr(64 + UBound(reportHeaders) + 1)).AutoFit
        .Range(.Cells(2, 3), .Cells(reportRow - 1, UBound(reportHeaders) + 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(reportRow - 1, UBound(reportHeaders) + 1)).Borders.LineStyle = xlContinuous
    End With

    ' Export als PDF (optional)
    On Error GoTo ExportError
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\DataQualityReport.pdf"

    With Worksheets(REPORT_SHEET_NAME)
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                             IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End With

    Exit Sub

ExportError:
    MsgBox "Error exporting report as PDF: " & Err.Description, vbExclamation, "Export Error"
End Sub
