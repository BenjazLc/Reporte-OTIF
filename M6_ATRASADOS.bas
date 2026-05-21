Attribute VB_Name = "M6_ATRASADOS"
Option Explicit

Private Const HOJA_REPORTE As String = "OTIF"

Private Const COL_ESTADO As String = "AV"
Private Const COL_DOC As String = "M"
Private Const COL_COMPRADOR As String = "O"
Private Const COL_PROVEEDOR As String = "B"
Private Const COL_VALOR As String = "E"

Private Const ESTADO_FILTRO As String = "ATRASADO PENDIENTE DE CONFIRMAR"

Public Sub OTIF_06_ATRASADOS_PROVEEDOR()

    On Error GoTo EH

    Dim wb As Workbook
    Dim wsData As Worksheet

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

    CrearHojaAtrasados wb, wsData, "ZNBM", "ATRAS NAC"
    CrearHojaAtrasados wb, wsData, "ZIMP", "ATRAS IMP"

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

Private Sub CrearHojaAtrasados(ByVal wb As Workbook, _
                               ByVal wsData As Worksheet, _
                               ByVal tipoDoc As String, _
                               ByVal nombreHoja As String)

    Dim wsOut As Worksheet
    Dim dGeneral As Object
    Dim dCompradores As Object
    Dim dTotalesGeneral As Object
    Dim dTotalesCompradores As Object
    Dim comprador As Variant
    Dim colInicio As Long
    Dim ultimaColTitulo As Long


    Set dGeneral = CreateObject("Scripting.Dictionary")
    Set dCompradores = CreateObject("Scripting.Dictionary")
    Set dTotalesGeneral = CreateObject("Scripting.Dictionary")
    Set dTotalesCompradores = CreateObject("Scripting.Dictionary")

    dGeneral.CompareMode = vbTextCompare
    dCompradores.CompareMode = vbTextCompare
    dTotalesGeneral.CompareMode = vbTextCompare
    dTotalesCompradores.CompareMode = vbTextCompare

    CargarConteos wsData, tipoDoc, dGeneral, dCompradores, dTotalesGeneral, dTotalesCompradores

    If HojaExiste(nombreHoja, wb) Then wb.Worksheets(nombreHoja).Delete

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets("Comprad"))
    wsOut.Name = nombreHoja

EscribirTablaProveedor wsOut, dGeneral, dTotalesGeneral, "GENERAL", 3, 2

colInicio = 6

For Each comprador In dCompradores.Keys
    EscribirTablaProveedor wsOut, dCompradores(comprador), dTotalesCompradores(comprador), CStr(comprador), 3, colInicio
    colInicio = colInicio + 4
Next comprador

ultimaColTitulo = colInicio - 2

 DibujarTitulo wsOut, nombreHoja, ultimaColTitulo

    wsOut.Columns.AutoFit
    wsOut.Activate
    wsOut.Range("B3").Select

End Sub

Private Sub CargarConteos(ByVal ws As Worksheet, _
                          ByVal tipoDoc As String, _
                          ByRef dGeneral As Object, _
                          ByRef dCompradores As Object, _
                          ByRef dTotalesGeneral As Object, _
                          ByRef dTotalesCompradores As Object)

    Dim lastRow As Long
    Dim r As Long

    Dim vEstado As String
    Dim vDoc As String
    Dim vComprador As String
    Dim vProveedor As String
    Dim vClave As String

    lastRow = UltimaFilaGeneral(ws)

    For r = 2 To lastRow

        vEstado = Trim$(CStr(ws.Range(COL_ESTADO & r).Value))
        vDoc = Trim$(CStr(ws.Range(COL_DOC & r).Value))
        vComprador = UCase$(Trim$(CStr(ws.Range(COL_COMPRADOR & r).Value)))
        vProveedor = Trim$(CStr(ws.Range(COL_PROVEEDOR & r).Value))
        vClave = Trim$(CStr(ws.Range(COL_VALOR & r).Value))

        If UCase$(vDoc) = UCase$(tipoDoc) Then
            If Len(vProveedor) > 0 And Len(vClave) > 0 Then

                If vComprador = "" Then vComprador = "SIN COMPRADOR"

                AgregarConteo dTotalesGeneral, vProveedor

                If Not dTotalesCompradores.Exists(vComprador) Then
                    Set dTotalesCompradores(vComprador) = CreateObject("Scripting.Dictionary")
                    dTotalesCompradores(vComprador).CompareMode = vbTextCompare
                End If

                AgregarConteo dTotalesCompradores(vComprador), vProveedor

                If UCase$(vEstado) = UCase$(ESTADO_FILTRO) Then

                    AgregarConteo dGeneral, vProveedor

                    If Not dCompradores.Exists(vComprador) Then
                        Set dCompradores(vComprador) = CreateObject("Scripting.Dictionary")
                        dCompradores(vComprador).CompareMode = vbTextCompare
                    End If

                    AgregarConteo dCompradores(vComprador), vProveedor

                End If

            End If
        End If

    Next r

End Sub

Private Sub AgregarConteo(ByRef dict As Object, ByVal proveedor As String)
    If dict.Exists(proveedor) Then
        dict(proveedor) = CLng(dict(proveedor)) + 1
    Else
        dict.Add proveedor, 1
    End If
End Sub

Private Sub EscribirTablaProveedor(ByVal ws As Worksheet, _
                                   ByVal dict As Object, _
                                   ByVal dictTotales As Object, _
                                   ByVal titulo As String, _
                                   ByVal filaInicio As Long, _
                                   ByVal colInicio As Long)

    Dim arrData() As Variant
    Dim n As Long
    Dim i As Long
    Dim k As Variant
    Dim r As Long
    Dim totalProv As Long

    Dim filaTitulo As Long
    Dim filaCab As Long
    Dim filaDatos As Long

    filaTitulo = filaInicio
    filaCab = filaInicio + 1
    filaDatos = filaInicio + 2

    With ws.Range(ws.Cells(filaTitulo, colInicio), ws.Cells(filaTitulo, colInicio + 2))
        .Merge
        .Value = titulo
        .Interior.Color = RGB(217, 217, 217)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
    End With

    ws.Cells(filaCab, colInicio).Value = "Proveedor"
    ws.Cells(filaCab, colInicio + 1).Value = "Cantidad"
    ws.Cells(filaCab, colInicio + 2).Value = "% Incumplimiento"

    With ws.Range(ws.Cells(filaCab, colInicio), ws.Cells(filaCab, colInicio + 2))
        .Interior.Color = RGB(189, 215, 238)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With

    If dict.Count = 0 Then
        ws.Cells(filaDatos, colInicio).Value = "SIN REGISTROS"
        ws.Cells(filaDatos, colInicio + 1).Value = 0
        ws.Cells(filaDatos, colInicio + 2).Value = 0
        ws.Cells(filaDatos, colInicio + 2).NumberFormat = "0.00%"
        ws.Range(ws.Cells(filaDatos, colInicio), ws.Cells(filaDatos, colInicio + 2)).Borders.LineStyle = xlContinuous
        Exit Sub
    End If

    ReDim arrData(1 To dict.Count, 1 To 3)
    n = 0

    For Each k In dict.Keys
        n = n + 1
        arrData(n, 1) = CStr(k)
        arrData(n, 2) = CLng(dict(k))

        If dictTotales.Exists(CStr(k)) Then
            totalProv = CLng(dictTotales(CStr(k)))
        Else
            totalProv = 0
        End If

        If totalProv > 0 Then
            arrData(n, 3) = CLng(dict(k)) / totalProv
        Else
            arrData(n, 3) = 0
        End If
    Next k

    OrdenarMatrizPorCantidad arrData

    r = filaDatos
    For i = 1 To UBound(arrData, 1)
        ws.Cells(r, colInicio).Value = arrData(i, 1)
        ws.Cells(r, colInicio + 1).Value = arrData(i, 2)
        ws.Cells(r, colInicio + 2).Value = arrData(i, 3)
        r = r + 1
    Next i

    With ws.Range(ws.Cells(filaDatos, colInicio), ws.Cells(r - 1, colInicio + 2))
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Aptos Narrow"
        .Font.Size = 11
    End With

    ws.Range(ws.Cells(filaDatos, colInicio + 1), ws.Cells(r - 1, colInicio + 2)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(filaDatos, colInicio + 2), ws.Cells(r - 1, colInicio + 2)).NumberFormat = "0.00%"

End Sub

Private Sub OrdenarMatrizPorCantidad(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp1 As Variant
    Dim tmp2 As Variant

    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If CLng(arr(j, 2)) > CLng(arr(i, 2)) Then
                tmp1 = arr(i, 1)
                tmp2 = arr(i, 2)

                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)

                arr(j, 1) = tmp1
                arr(j, 2) = tmp2
            End If
        Next j
    Next i
End Sub

Private Sub DibujarTitulo(ByVal ws As Worksheet, ByVal txt As String, ByVal ultimaCol As Long)

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

