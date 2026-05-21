Attribute VB_Name = "M5_RESUMEN_COMPRADOR"
Option Explicit

Private Const HOJA_REPORTE As String = "OTIF"
Private Const HOJA_RESUMEN As String = "COMPRAD"

Private Const CAMPO_ESTADO As String = "Estado"
Private Const CAMPO_FECHA As String = "Fecha de entrega"
Private Const CAMPO_COMPRADOR As String = "Creado por"

Private Const EST_ENTREGADO_TIEMPO As String = "ENTREGADO A TIEMPO"
Private Const EST_PENDIENTE_FECHA As String = "A TIEMPO PENDIENTE DE CONFIRMAR"
Private Const EST_PENDIENTE_FECHA_ALT As String = "PENDIENTE DE CONFIRMAR EN FECHA"
Private Const EST_ENTREGADO_ATRASADO As String = "ENTREGADO ATRASADO"
Private Const EST_ATRASADO_PEND As String = "ATRASADO PENDIENTE DE CONFIRMAR"

Public Sub OTIF_05_Resumen_Comprador()

    On Error GoTo EH

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsRes As Worksheet

    Dim colEstado As Long
    Dim colFecha As Long
    Dim colComprador As Long
    Dim lastRow As Long
    Dim r As Long

    Dim anioActual As Long
    Dim mesActual As Long
    Dim fechaEntrega As Variant
    Dim estado As String
    Dim comprador As String
    Dim mesDestino As Long

    Dim compradores As Object
    Dim resumen As Object

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

    anioActual = Year(Date)
    mesActual = Month(Date)

    calcMode = Application.Calculation
    scrUpdate = Application.ScreenUpdating
    dispAlerts = Application.DisplayAlerts
    evtState = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsData = ObtenerHojaPorNombreSeguro(wb, HOJA_REPORTE)
    If wsData Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No se encontró la hoja '" & HOJA_REPORTE & "'."
    End If

    colEstado = BuscarColumna(wsData, CAMPO_ESTADO)
    colFecha = BuscarColumna(wsData, CAMPO_FECHA)
    colComprador = BuscarColumna(wsData, CAMPO_COMPRADOR)

    If colEstado = 0 Then Err.Raise vbObjectError + 1001, , "No se encontró el encabezado: " & CAMPO_ESTADO
    If colFecha = 0 Then Err.Raise vbObjectError + 1002, , "No se encontró el encabezado: " & CAMPO_FECHA
    If colComprador = 0 Then Err.Raise vbObjectError + 1003, , "No se encontró el encabezado: " & CAMPO_COMPRADOR

    lastRow = UltimaFilaGeneral(wsData)
    If lastRow < 2 Then Err.Raise vbObjectError + 1004, , "La hoja REPORTE OTIF no tiene datos."

    Set compradores = CreateObject("Scripting.Dictionary")
    compradores.CompareMode = vbTextCompare

    Set resumen = CreateObject("Scripting.Dictionary")
    resumen.CompareMode = vbTextCompare

    InicializarBloque resumen, "GENERAL"

    For r = 2 To lastRow

        estado = UCase$(Trim$(CStr(wsData.Cells(r, colEstado).Value)))
        comprador = Trim$(CStr(wsData.Cells(r, colComprador).Value))
        fechaEntrega = wsData.Cells(r, colFecha).Value
        mesDestino = 0

        If IsDate(fechaEntrega) Then
            fechaEntrega = CDate(fechaEntrega)
            fechaEntrega = DateSerial(Year(fechaEntrega), Month(fechaEntrega), Day(fechaEntrega))

            If Year(fechaEntrega) = anioActual Then
                mesDestino = Month(fechaEntrega)
            ElseIf Year(fechaEntrega) < anioActual And estado = EST_ATRASADO_PEND Then
                mesDestino = mesActual
            End If
        End If

        If mesDestino >= 1 And mesDestino <= 12 Then

            If Len(comprador) > 0 Then
                If Not compradores.Exists(UCase$(comprador)) Then
                    compradores.Add UCase$(comprador), comprador
                End If

                If Not resumen.Exists(UCase$(comprador)) Then
                    InicializarBloque resumen, UCase$(comprador)
                End If
            End If

            Select Case estado
                Case EST_ENTREGADO_TIEMPO
                    SumarBloque resumen, "GENERAL", 1, mesDestino
                    If Len(comprador) > 0 Then SumarBloque resumen, UCase$(comprador), 1, mesDestino

                Case EST_PENDIENTE_FECHA, EST_PENDIENTE_FECHA_ALT
                    SumarBloque resumen, "GENERAL", 2, mesDestino
                    If Len(comprador) > 0 Then SumarBloque resumen, UCase$(comprador), 2, mesDestino

                Case EST_ENTREGADO_ATRASADO
                    SumarBloque resumen, "GENERAL", 3, mesDestino
                    If Len(comprador) > 0 Then SumarBloque resumen, UCase$(comprador), 3, mesDestino

                Case EST_ATRASADO_PEND
                    SumarBloque resumen, "GENERAL", 4, mesDestino
                    If Len(comprador) > 0 Then SumarBloque resumen, UCase$(comprador), 4, mesDestino
            End Select
        End If
    Next r

    If HojaExiste(HOJA_RESUMEN, wb) Then wb.Worksheets(HOJA_RESUMEN).Delete

    Set wsRes = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsRes.Name = HOJA_RESUMEN

    FormatearTitulo wsRes, anioActual, "RESUMEN OTIF POR COMPRADOR - " & anioActual

    Dim ordenCompradores() As String
    ordenCompradores = OrdenCompradoresDesdeDiccionario(compradores)

    Dim fila As Long
    Dim i As Long

    fila = 3
    

    For i = LBound(ordenCompradores) To UBound(ordenCompradores)
        If Len(ordenCompradores(i)) > 0 Then
            If resumen.Exists(UCase$(ordenCompradores(i))) Then
                fila = EscribirBloqueResumen(wsRes, fila, ordenCompradores(i), resumen(UCase$(ordenCompradores(i))), anioActual)
            End If
        End If
    Next i

    wsRes.Columns("B:B").ColumnWidth = 34
    wsRes.Columns("C:O").AutoFit
    wsRes.Activate
    wsRes.Range("B3").Select

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

Private Sub InicializarBloque(ByRef resumen As Object, ByVal clave As String)
    Dim arr(1 To 4, 1 To 12) As Long
    resumen.Add UCase$(clave), arr
End Sub

Private Sub SumarBloque(ByRef resumen As Object, ByVal clave As String, ByVal filaEstado As Long, ByVal mes As Long)
    Dim arr As Variant
    arr = resumen(UCase$(clave))
    arr(filaEstado, mes) = arr(filaEstado, mes) + 1
    resumen(UCase$(clave)) = arr
End Sub

Private Function EscribirBloqueResumen(ByVal ws As Worksheet, _
                                      ByVal filaInicio As Long, _
                                      ByVal titulo As String, _
                                      ByVal arr As Variant, _
                                      ByVal anioActual As Long) As Long

    Dim r As Long
    Dim c As Long
    Dim filaTabla As Long
    Dim colIni As Long
    Dim totalMes(1 To 12) As Long
    Dim totalGeneral As Long
    Dim totalFila As Long

    colIni = 2 ' B

    With ws.Range(ws.Cells(filaInicio, colIni), ws.Cells(filaInicio, colIni + 13))
        .Merge
        .Value = titulo
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
    End With

    filaTabla = filaInicio + 1

    ws.Cells(filaTabla, colIni).Value = ""
    For c = 1 To 12
        ws.Cells(filaTabla, colIni + c).Value = Format(DateSerial(anioActual, c, 1), "mm.yyyy")
    Next c
    ws.Cells(filaTabla, colIni + 13).Value = "Total"

    With ws.Range(ws.Cells(filaTabla, colIni), ws.Cells(filaTabla, colIni + 13))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    ws.Cells(filaTabla + 1, colIni).Value = EST_ENTREGADO_TIEMPO
    ws.Cells(filaTabla + 2, colIni).Value = EST_PENDIENTE_FECHA_ALT
    ws.Cells(filaTabla + 3, colIni).Value = EST_ENTREGADO_ATRASADO
    ws.Cells(filaTabla + 4, colIni).Value = EST_ATRASADO_PEND
    ws.Cells(filaTabla + 5, colIni).Value = "Total General"
    ws.Cells(filaTabla + 6, colIni).Value = "% OTIF"

    For r = 1 To 4
        totalFila = 0
        For c = 1 To 12
            ws.Cells(filaTabla + r, colIni + c).Value = arr(r, c)
            totalFila = totalFila + arr(r, c)
            totalMes(c) = totalMes(c) + arr(r, c)
        Next c
        ws.Cells(filaTabla + r, colIni + 13).Value = totalFila
        totalGeneral = totalGeneral + totalFila
    Next r

    For c = 1 To 12
        ws.Cells(filaTabla + 5, colIni + c).Value = totalMes(c)
    Next c
    ws.Cells(filaTabla + 5, colIni + 13).Value = totalGeneral

    For c = 1 To 12
        If totalMes(c) > 0 Then
            ws.Cells(filaTabla + 6, colIni + c).Value = arr(1, c) / totalMes(c)
        Else
            ws.Cells(filaTabla + 6, colIni + c).Value = 0
        End If
    Next c

    If totalGeneral > 0 Then
        ws.Cells(filaTabla + 6, colIni + 13).Value = SumaFila(arr, 1) / totalGeneral
    Else
        ws.Cells(filaTabla + 6, colIni + 13).Value = 0
    End If

    ws.Range(ws.Cells(filaTabla + 6, colIni + 1), ws.Cells(filaTabla + 6, colIni + 13)).NumberFormat = "0%"

    AplicarFormatoBloque ws, filaTabla, colIni

    EscribirBloqueResumen = filaTabla + 9
End Function

Private Function SumaFila(ByVal arr As Variant, ByVal filaEstado As Long) As Long
    Dim c As Long
    For c = 1 To 12
        SumaFila = SumaFila + arr(filaEstado, c)
    Next c
End Function

Private Sub AplicarFormatoBloque(ByVal ws As Worksheet, ByVal filaTabla As Long, ByVal colIni As Long)

    Dim rangoBloque As Range
    Set rangoBloque = ws.Range(ws.Cells(filaTabla, colIni), ws.Cells(filaTabla + 6, colIni + 13))

    rangoBloque.Borders.LineStyle = xlContinuous
    rangoBloque.Font.Name = "Aptos Narrow"
    rangoBloque.Font.Size = 11

    With ws.Range(ws.Cells(filaTabla + 1, colIni), ws.Cells(filaTabla + 1, colIni + 13))
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 0, 0)
    End With

    With ws.Range(ws.Cells(filaTabla + 2, colIni), ws.Cells(filaTabla + 2, colIni + 13))
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(0, 0, 0)
    End With

    With ws.Range(ws.Cells(filaTabla + 3, colIni), ws.Cells(filaTabla + 3, colIni + 13))
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range(ws.Cells(filaTabla + 4, colIni), ws.Cells(filaTabla + 4, colIni + 13))
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range(ws.Cells(filaTabla + 5, colIni), ws.Cells(filaTabla + 5, colIni + 13))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
    End With

    With ws.Range(ws.Cells(filaTabla + 6, colIni), ws.Cells(filaTabla + 6, colIni + 13))
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With

    ws.Range(ws.Cells(filaTabla + 1, colIni + 1), ws.Cells(filaTabla + 6, colIni + 13)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(filaTabla + 1, colIni), ws.Cells(filaTabla + 6, colIni)).HorizontalAlignment = xlLeft
End Sub

Private Sub FormatearTitulo(ByVal ws As Worksheet, ByVal anioActual As Long, ByVal textoTitulo As String)
    With ws.Range("B1:O1")
        .Merge
        .Value = textoTitulo
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Rows(1).RowHeight = 24
End Sub

Private Function OrdenCompradoresDesdeDiccionario(ByVal compradores As Object) As String()
    Dim arrBase() As String
    Dim arrOut() As String
    Dim k As Variant
    Dim i As Long
    Dim n As Long

    If compradores.Count = 0 Then
        ReDim arrOut(1 To 1)
        arrOut(1) = ""
        OrdenCompradoresDesdeDiccionario = arrOut
        Exit Function
    End If

    ReDim arrBase(1 To compradores.Count)
    i = 1
    For Each k In compradores.Keys
        arrBase(i) = CStr(compradores(k))
        i = i + 1
    Next k

    OrdenarArray arrBase

    ReDim arrOut(1 To UBound(arrBase))
    n = 0

    n = AgregarSiExiste(arrOut, n, arrBase, "MLEZAMA1")
    n = AgregarSiExiste(arrOut, n, arrBase, "BLAPA")
    n = AgregarSiExiste(arrOut, n, arrBase, "MRIOS")
    
    For i = LBound(arrBase) To UBound(arrBase)
        If UCase$(Trim$(arrBase(i))) <> "MLEZAMA1" _
           And UCase$(Trim$(arrBase(i))) <> "BLAPA" _
           And UCase$(Trim$(arrBase(i))) <> "MRIOS" Then
            n = n + 1
            arrOut(n) = arrBase(i)
        End If
    Next i

    ReDim Preserve arrOut(1 To n)
    OrdenCompradoresDesdeDiccionario = arrOut
End Function

Private Function AgregarSiExiste(ByRef arrOut() As String, _
                                 ByVal n As Long, _
                                 ByRef arrBase() As String, _
                                 ByVal buscado As String) As Long
    Dim i As Long

    For i = LBound(arrBase) To UBound(arrBase)
        If UCase$(Trim$(arrBase(i))) = UCase$(Trim$(buscado)) Then
            n = n + 1
            arrOut(n) = arrBase(i)
            Exit For
        End If
    Next i

    AgregarSiExiste = n
End Function

Private Function ObtenerHojaPorNombreSeguro(ByVal wb As Workbook, ByVal nombreBuscado As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Trim(UCase$(ws.Name)) = Trim(UCase$(nombreBuscado)) Then
            Set ObtenerHojaPorNombreSeguro = ws
            Exit Function
        End If
    Next ws
End Function

Private Function BuscarColumna(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim c As Long
    Dim ultima As Long

    ultima = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To ultima
        If Trim(UCase$(CStr(ws.Cells(1, c).Value))) = Trim(UCase$(nombre)) Then
            BuscarColumna = c
            Exit Function
        End If
    Next c

    BuscarColumna = 0
End Function

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
    Set ultCelda = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If ultCelda Is Nothing Then
        UltimaFilaGeneral = 1
    Else
        UltimaFilaGeneral = ultCelda.Row
    End If
End Function

Private Sub OrdenarArray(ByRef arr() As String)
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

