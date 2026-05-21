Attribute VB_Name = "M9_RETENIDOS"
Option Explicit

Private Const HOJA_REPORTE As String = "OTIF"
Private Const HOJA_SALIDA As String = "RETENIDOS"

Private Const COL_IND_LIB As String = "N"
Private Const COL_FECHA_DOC As String = "P"

Private Const COL_C As String = "C"
Private Const COL_D As String = "D"
Private Const COL_B As String = "B"
Private Const COL_F As String = "F"
Private Const COL_G As String = "G"
Private Const COL_O As String = "O"

Public Sub OTIF_09_Retenidos()

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

    If HojaExiste(HOJA_SALIDA, wb) Then wb.Worksheets(HOJA_SALIDA).Delete

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = HOJA_SALIDA

    ConstruirRetenidos wsData, wsOut

    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode
    
    
    Dim ws As Worksheet

For Each ws In wb.Worksheets

    Select Case UCase$(ws.Name)

        Case "OTIF"
            ws.Tab.Color = RGB(31, 78, 121)

        Case "RESUM"
            ws.Tab.Color = RGB(0, 176, 80)

        Case "COMPRAD"
            ws.Tab.Color = RGB(91, 155, 213)

        Case "OTIF PROV"
            ws.Tab.Color = RGB(112, 48, 160)

        Case "ATRAS IMP"
            ws.Tab.Color = RGB(255, 0, 0)

        Case "ATRAS NAC"
            ws.Tab.Color = RGB(255, 102, 0)

        Case "NO LIBER"
            ws.Tab.Color = RGB(255, 192, 0)

        Case "RETENIDOS"
            ws.Tab.Color = RGB(192, 0, 0)

    End Select

Next ws

    Exit Sub

EH:
    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode
    
    
    

    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ConstruirRetenidos(ByVal wsData As Worksheet, ByVal wsOut As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim filaOut As Long

    Dim indLib As String
    Dim fechaDoc As Variant
    Dim dias As Long
    Dim semaforo As String

    DibujarTituloRetenidos wsOut, "RETENIDOS"

    ' Cabecera
    wsOut.Cells(3, 2).Value = "Semáforo"
    wsOut.Cells(3, 3).Value = wsData.Range(COL_C & "1").Value
    wsOut.Cells(3, 4).Value = wsData.Range(COL_D & "1").Value
    wsOut.Cells(3, 5).Value = wsData.Range(COL_B & "1").Value
    wsOut.Cells(3, 6).Value = wsData.Range(COL_F & "1").Value
    wsOut.Cells(3, 7).Value = wsData.Range(COL_G & "1").Value
    wsOut.Cells(3, 8).Value = wsData.Range(COL_O & "1").Value
    wsOut.Cells(3, 9).Value = "Días"

    With wsOut.Range(wsOut.Cells(3, 2), wsOut.Cells(3, 9))
        .Interior.Color = RGB(189, 215, 238)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    lastRow = UltimaFilaGeneral(wsData)
    filaOut = 4

    For r = 2 To lastRow

        indLib = Trim$(CStr(wsData.Range(COL_IND_LIB & r).Value))
        fechaDoc = wsData.Range(COL_FECHA_DOC & r).Value

        ' Retenidos = celda vacía
        If Len(indLib) = 0 Then

            If IsDate(fechaDoc) Then
                dias = CLng(Date - CDate(fechaDoc))
                semaforo = ObtenerSemaforoRetenido(dias)

                If Len(semaforo) > 0 Then
                    wsOut.Cells(filaOut, 2).Value = semaforo
                    wsOut.Cells(filaOut, 3).Value = wsData.Range(COL_C & r).Value
                    wsOut.Cells(filaOut, 4).Value = wsData.Range(COL_D & r).Value
                    wsOut.Cells(filaOut, 5).Value = wsData.Range(COL_B & r).Value
                    wsOut.Cells(filaOut, 6).Value = wsData.Range(COL_F & r).Value
                    wsOut.Cells(filaOut, 7).Value = wsData.Range(COL_G & r).Value
                    wsOut.Cells(filaOut, 8).Value = wsData.Range(COL_O & r).Value
                    wsOut.Cells(filaOut, 9).Value = dias

                    AplicarColorFilaRetenido wsOut, filaOut, semaforo

                    filaOut = filaOut + 1
                End If
            End If

        End If
    Next r

    If filaOut = 4 Then
        wsOut.Range("B4:I4").Merge
        wsOut.Range("B4").Value = "SIN REGISTROS"
        With wsOut.Range("B4:I4")
            .Interior.Color = RGB(217, 217, 217)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    Else
        With wsOut.Range(wsOut.Cells(4, 2), wsOut.Cells(filaOut - 1, 9))
            .Borders.LineStyle = xlContinuous
            .Font.Name = "Aptos Narrow"
            .Font.Size = 11
        End With
    End If

    wsOut.Columns("B:I").AutoFit
    wsOut.Activate
    wsOut.Range("B3").Select
End Sub

Private Function ObtenerSemaforoRetenido(ByVal dias As Long) As String
    If dias >= 1 And dias <= 3 Then
        ObtenerSemaforoRetenido = "1 - 3 días"
    ElseIf dias >= 4 And dias <= 7 Then
        ObtenerSemaforoRetenido = "4 - 7 días"
    ElseIf dias >= 8 And dias <= 10 Then
        ObtenerSemaforoRetenido = "8 - 10 días"
    ElseIf dias > 10 Then
        ObtenerSemaforoRetenido = "Más de 10 días"
    Else
        ObtenerSemaforoRetenido = ""
    End If
End Function

Private Sub AplicarColorFilaRetenido(ByVal ws As Worksheet, ByVal fila As Long, ByVal semaforo As String)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 9))

    Select Case semaforo
        Case "1 - 3 días"
            rng.Interior.Color = RGB(255, 242, 204)

        Case "4 - 7 días"
            rng.Interior.Color = RGB(255, 230, 153)

        Case "8 - 10 días"
            rng.Interior.Color = RGB(244, 176, 132)

        Case "Más de 10 días"
            rng.Interior.Color = RGB(255, 0, 0)
            rng.Font.Color = RGB(255, 255, 255)
    End Select
End Sub

Private Sub DibujarTituloRetenidos(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Range("B1:I1")
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

