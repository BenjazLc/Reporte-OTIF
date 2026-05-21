Attribute VB_Name = "M8_NO_LIBERADO"
Option Explicit

Private Const HOJA_REPORTE As String = "OTIF"
Private Const HOJA_SALIDA As String = "No Liber"

Private Const COL_ID As String = "E"
Private Const COL_IND_LIB As String = "N"
Private Const COL_FECHA_DOC As String = "P"
Private Const COL_GRUPO_COMPRAS As String = "X"

Public Sub OTIF_08_NoLiberados()

    On Error GoTo EH

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsOut As Worksheet

    Dim calcMode As XlCalculation
    Dim scrUpdate As Boolean
    Dim dispAlerts As Boolean
    Dim evtState As Boolean

    Dim rutaReporte As String
Dim fechaHoy As String

fechaHoy = Format(Date, "dd-mm-yyyy")
rutaReporte = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\REPORTE OTIF_" & fechaHoy & ".xlsx"

On Error Resume Next
Set wb = Workbooks("REPORTE OTIF_" & fechaHoy & ".xlsx")
On Error GoTo 0

If wb Is Nothing Then
    Set wb = Workbooks.Open(rutaReporte)
End If

    Set wsData = wb.Worksheets(HOJA_REPORTE)

    calcMode = Application.Calculation
    scrUpdate = Application.ScreenUpdating
    dispAlerts = Application.DisplayAlerts
    evtState = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If HojaExiste(UCase(HOJA_SALIDA), wb) Then wb.Worksheets(UCase(HOJA_SALIDA)).Delete
    If HojaExiste(HOJA_SALIDA, wb) Then wb.Worksheets(HOJA_SALIDA).Delete

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = UCase(HOJA_SALIDA)

' Detalle a la izquierda
ConstruirDetalleNoLiberados wsData, wsOut, 3, 2

' Resumen al costado
ConstruirNoLiberados wsData, wsOut, 3, 11

' Título azul hasta última columna usada
DibujarTitulo wsOut, "NO LIBERADOS"

    wsOut.Columns("B:J").AutoFit
    wsOut.Activate
    wsOut.Range("B3").Select

    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode

    
    Exit Sub

EH:
    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode

    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ConstruirNoLiberados(ByVal wsData As Worksheet, _
                                 ByVal wsOut As Worksheet, _
                                 ByVal filaInicio As Long, _
                                 ByVal colInicio As Long)

    Dim lastRow As Long
    Dim r As Long

    Dim idVal As String
    Dim indLib As String
    Dim grupo As String
    Dim fechaDoc As Variant
    Dim dias As Long
    Dim bucket As String

    Dim dictGrupos As Object
    Dim dictConteos As Object

    Dim key As String
    Dim grupos() As String
    Dim i As Long
    Dim k As Variant

    Set dictGrupos = CreateObject("Scripting.Dictionary")
    dictGrupos.CompareMode = vbTextCompare

    Set dictConteos = CreateObject("Scripting.Dictionary")
    dictConteos.CompareMode = vbTextCompare

    lastRow = UltimaFilaGeneral(wsData)

    For r = 2 To lastRow

        idVal = Trim$(CStr(wsData.Range(COL_ID & r).Value))
        indLib = UCase$(Trim$(CStr(wsData.Range(COL_IND_LIB & r).Value)))
        grupo = Trim$(CStr(wsData.Range(COL_GRUPO_COMPRAS & r).Value))
        fechaDoc = wsData.Range(COL_FECHA_DOC & r).Value

        If indLib = "X" Then
            If Len(idVal) > 0 And Len(grupo) > 0 And IsDate(fechaDoc) Then

                dias = CLng(Date - CDate(fechaDoc))
                bucket = ObtenerSemaforo(dias)

                If Len(bucket) > 0 Then
                    If Not dictGrupos.Exists(grupo) Then dictGrupos.Add grupo, grupo

                    key = bucket & "|" & grupo

                    If dictConteos.Exists(key) Then
                        dictConteos(key) = CLng(dictConteos(key)) + 1
                    Else
                        dictConteos.Add key, 1
                    End If
                End If
            End If
        End If
    Next r

    If dictGrupos.Count = 0 Then
        wsOut.Range(wsOut.Cells(filaInicio, colInicio), wsOut.Cells(filaInicio, colInicio + 3)).Merge
        wsOut.Cells(filaInicio, colInicio).Value = "SIN REGISTROS"
        With wsOut.Range(wsOut.Cells(filaInicio, colInicio), wsOut.Cells(filaInicio, colInicio + 3))
            .Interior.Color = RGB(217, 217, 217)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        Exit Sub
    End If

    ReDim grupos(1 To dictGrupos.Count)
    i = 1
    For Each k In dictGrupos.Keys
        grupos(i) = CStr(k)
        i = i + 1
    Next k

    OrdenarArrayTexto grupos
    EscribirTablaNoLiberados wsOut, dictConteos, grupos, filaInicio, colInicio
End Sub

Private Function ObtenerSemaforo(ByVal dias As Long) As String
    If dias >= 1 And dias <= 3 Then
        ObtenerSemaforo = "1 - 3 días"
    ElseIf dias >= 4 And dias <= 8 Then
        ObtenerSemaforo = "4 - 8 días"
    ElseIf dias >= 9 And dias <= 16 Then
        ObtenerSemaforo = "9 - 16 días"
    ElseIf dias >= 17 And dias <= 30 Then
        ObtenerSemaforo = "17 - 30 días"
    ElseIf dias > 30 Then
        ObtenerSemaforo = "Más de 30 días"
    Else
        ObtenerSemaforo = ""
    End If
End Function

Private Sub EscribirTablaNoLiberados(ByVal ws As Worksheet, _
                                     ByVal dictConteos As Object, _
                                     ByRef grupos() As String, _
                                     ByVal filaInicio As Long, _
                                     ByVal colInicio As Long)

    Dim filaCab As Long
    Dim filaDatos As Long

    Dim semaforos(1 To 5) As String
    Dim r As Long
    Dim c As Long
    Dim totalFila As Long
    Dim totalCol() As Long
    Dim totalGeneral As Long
    Dim key As String
    Dim valor As Long

    semaforos(1) = "1 - 3 días"
    semaforos(2) = "4 - 8 días"
    semaforos(3) = "9 - 16 días"
    semaforos(4) = "17 - 30 días"
    semaforos(5) = "Más de 30 días"

    filaCab = filaInicio + 1
    filaDatos = filaInicio + 2

    ReDim totalCol(1 To UBound(grupos))

    With ws.Range(ws.Cells(filaInicio, colInicio), ws.Cells(filaInicio, colInicio + UBound(grupos) + 1))
        .Merge
        .Value = "Semáforo de No Liberados"
        .Interior.Color = RGB(217, 217, 217)
        .Font.Bold = True
        .Font.Size = 12
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
    End With

    ws.Cells(filaCab, colInicio).Value = "Semáforo"

    For c = 1 To UBound(grupos)
        ws.Cells(filaCab, colInicio + c).Value = grupos(c)
    Next c

    ws.Cells(filaCab, colInicio + UBound(grupos) + 1).Value = "Total general"

    With ws.Range(ws.Cells(filaCab, colInicio), ws.Cells(filaCab, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(189, 215, 238)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    For r = 1 To 5
        ws.Cells(filaDatos + r - 1, colInicio).Value = semaforos(r)
        totalFila = 0

        For c = 1 To UBound(grupos)
            key = semaforos(r) & "|" & grupos(c)

            If dictConteos.Exists(key) Then
                valor = CLng(dictConteos(key))
            Else
                valor = 0
            End If

            ws.Cells(filaDatos + r - 1, colInicio + c).Value = valor
            totalFila = totalFila + valor
            totalCol(c) = totalCol(c) + valor
            totalGeneral = totalGeneral + valor
        Next c

        ws.Cells(filaDatos + r - 1, colInicio + UBound(grupos) + 1).Value = totalFila
    Next r

    ws.Cells(filaDatos + 5, colInicio).Value = "Total general"

    For c = 1 To UBound(grupos)
        ws.Cells(filaDatos + 5, colInicio + c).Value = totalCol(c)
    Next c

    ws.Cells(filaDatos + 5, colInicio + UBound(grupos) + 1).Value = totalGeneral

    With ws.Range(ws.Cells(filaDatos, colInicio), ws.Cells(filaDatos + 5, colInicio + UBound(grupos) + 1))
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Aptos Narrow"
        .Font.Size = 11
    End With

    With ws.Range(ws.Cells(filaDatos, colInicio), ws.Cells(filaDatos, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(255, 242, 204)
    End With

    With ws.Range(ws.Cells(filaDatos + 1, colInicio), ws.Cells(filaDatos + 1, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(255, 230, 153)
    End With

    With ws.Range(ws.Cells(filaDatos + 2, colInicio), ws.Cells(filaDatos + 2, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(244, 176, 132)
    End With

    With ws.Range(ws.Cells(filaDatos + 3, colInicio), ws.Cells(filaDatos + 3, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(255, 153, 0)
    End With

    With ws.Range(ws.Cells(filaDatos + 4, colInicio), ws.Cells(filaDatos + 4, colInicio + UBound(grupos) + 1))
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range(ws.Cells(filaDatos + 5, colInicio), ws.Cells(filaDatos + 5, colInicio + UBound(grupos) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
    End With

' Anchos controlados para tabla Semáforo
ws.Columns(colInicio).ColumnWidth = 14 ' Semáforo

For c = 1 To UBound(grupos)
    ws.Columns(colInicio + c).ColumnWidth = 3.56
Next c

ws.Columns(colInicio + UBound(grupos) + 1).ColumnWidth = 6.78 ' Total general

End Sub

Private Sub ConstruirDetalleNoLiberados(ByVal wsData As Worksheet, _
                                        ByVal wsOut As Worksheet, _
                                        ByVal filaInicio As Long, _
                                        ByVal colInicio As Long)

    Dim lastRow As Long
    Dim r As Long
    Dim filaOut As Long

    Dim idVal As String
    Dim indLib As String
    Dim grupo As String
    Dim fechaDoc As Variant
    Dim dias As Long
    Dim semaforo As String

    With wsOut.Range(wsOut.Cells(filaInicio, colInicio), wsOut.Cells(filaInicio, colInicio + 7))
        .Merge
        .Value = "Detalle No Liberados"
        .Interior.Color = RGB(217, 217, 217)
        .Font.Bold = True
        .Font.Size = 12
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
    End With

    wsOut.Cells(filaInicio + 1, colInicio).Value = "Semáforo"
    wsOut.Cells(filaInicio + 1, colInicio + 1).Value = wsData.Range("C1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 2).Value = wsData.Range("D1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 3).Value = wsData.Range("B1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 4).Value = wsData.Range("F1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 5).Value = wsData.Range("G1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 6).Value = wsData.Range("X1").Value
    wsOut.Cells(filaInicio + 1, colInicio + 7).Value = "Días"

    With wsOut.Range(wsOut.Cells(filaInicio + 1, colInicio), wsOut.Cells(filaInicio + 1, colInicio + 7))
        .Interior.Color = RGB(189, 215, 238)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    lastRow = UltimaFilaGeneral(wsData)
    filaOut = filaInicio + 2

    For r = 2 To lastRow

        idVal = Trim$(CStr(wsData.Range(COL_ID & r).Value))
        indLib = UCase$(Trim$(CStr(wsData.Range(COL_IND_LIB & r).Value)))
        grupo = Trim$(CStr(wsData.Range(COL_GRUPO_COMPRAS & r).Value))
        fechaDoc = wsData.Range(COL_FECHA_DOC & r).Value

        If indLib = "X" Then
            If Len(idVal) > 0 And Len(grupo) > 0 And IsDate(fechaDoc) Then

                dias = CLng(Date - CDate(fechaDoc))
                semaforo = ObtenerSemaforo(dias)

                If Len(semaforo) > 0 Then
                    wsOut.Cells(filaOut, colInicio).Value = semaforo
                    wsOut.Cells(filaOut, colInicio + 1).Value = wsData.Range("C" & r).Value
                    wsOut.Cells(filaOut, colInicio + 2).Value = wsData.Range("D" & r).Value
                    wsOut.Cells(filaOut, colInicio + 3).Value = wsData.Range("B" & r).Value
                    wsOut.Cells(filaOut, colInicio + 4).Value = wsData.Range("F" & r).Value
                    wsOut.Cells(filaOut, colInicio + 5).Value = wsData.Range("G" & r).Value
                    wsOut.Cells(filaOut, colInicio + 6).Value = wsData.Range("X" & r).Value
                    wsOut.Cells(filaOut, colInicio + 7).Value = dias

                    AplicarColorFilaNoLiberadoDetalle wsOut, filaOut, colInicio, semaforo
                    filaOut = filaOut + 1
                End If
            End If
        End If
    Next r

    If filaOut = filaInicio + 2 Then
        wsOut.Range(wsOut.Cells(filaInicio + 2, colInicio), wsOut.Cells(filaInicio + 2, colInicio + 7)).Merge
        wsOut.Cells(filaInicio + 2, colInicio).Value = "SIN REGISTROS"
        With wsOut.Range(wsOut.Cells(filaInicio + 2, colInicio), wsOut.Cells(filaInicio + 2, colInicio + 7))
            .Interior.Color = RGB(217, 217, 217)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    Else
        With wsOut.Range(wsOut.Cells(filaInicio + 2, colInicio), wsOut.Cells(filaOut - 1, colInicio + 7))
            .Borders.LineStyle = xlContinuous
            .Font.Name = "Aptos Narrow"
            .Font.Size = 11
        End With
    End If
End Sub

Private Sub AplicarColorFilaNoLiberadoDetalle(ByVal ws As Worksheet, _
                                              ByVal fila As Long, _
                                              ByVal colInicio As Long, _
                                              ByVal semaforo As String)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(fila, colInicio), ws.Cells(fila, colInicio + 7))

    Select Case semaforo
        Case "1 - 3 días"
            rng.Interior.Color = RGB(255, 242, 204)
            rng.Font.Color = RGB(0, 0, 0)

        Case "4 - 8 días"
            rng.Interior.Color = RGB(255, 230, 153)
            rng.Font.Color = RGB(0, 0, 0)

        Case "9 - 16 días"
            rng.Interior.Color = RGB(244, 176, 132)
            rng.Font.Color = RGB(0, 0, 0)

        Case "17 - 30 días"
            rng.Interior.Color = RGB(255, 153, 0)
            rng.Font.Color = RGB(0, 0, 0)

        Case "Más de 30 días"
            rng.Interior.Color = RGB(255, 0, 0)
            rng.Font.Color = RGB(255, 255, 255)
    End Select
End Sub

Private Sub DibujarTitulo(ByVal ws As Worksheet, ByVal txt As String)

    Dim ultimaCol As Long
    ultimaCol = ws.Cells.Find(What:="*", _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious).Column

    With ws.Range(ws.Cells(1, 2), ws.Cells(1, ultimaCol))
    
        .Merge
        .Value = txt
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 24
End Sub

Private Function HojaExiste(ByVal nombre As String, ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    HojaExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function UltimaFilaGeneral(ByVal ws As Worksheet) As Long
    Dim ultCelda As Range

    On Error Resume Next
    Set ultCelda = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
    On Error GoTo 0

    If ultCelda Is Nothing Then
        UltimaFilaGeneral = 1
    Else
        UltimaFilaGeneral = ultCelda.Row
    End If
End Function

Private Sub OrdenarArrayTexto(ByRef arr() As String)
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If UCase$(arr(j)) < UCase$(arr(i)) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

