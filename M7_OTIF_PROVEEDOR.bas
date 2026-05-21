Attribute VB_Name = "M7_OTIF_PROVEEDOR"
Option Explicit

Private Const HOJA_REPORTE As String = "OTIF"
Private Const HOJA_SALIDA As String = "OTIF PROV"

Private Const COL_PROVEEDOR As String = "B"
Private Const COL_DOC As String = "M"
Private Const COL_ESTADO As String = "AV"

Private Const DOC_NAC As String = "ZNBM"
Private Const DOC_IMP As String = "ZIMP"

Private Const EST_1 As String = "A TIEMPO PENDIENTE DE CONFIRMAR"
Private Const EST_1_ALT As String = "PENDIENTE DE CONFIRMAR EN FECHA"
Private Const EST_2 As String = "ATRASADO PENDIENTE DE CONFIRMAR"
Private Const EST_3 As String = "ENTREGADO A TIEMPO"
Private Const EST_4 As String = "ENTREGADO ATRASADO"

Public Sub OTIF_07_Proveedor()

    On Error GoTo EH

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsOut As Worksheet

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

    If HojaExiste(HOJA_SALIDA, wb) Then wb.Worksheets(HOJA_SALIDA).Delete

    Set wsOut = wb.Worksheets.Add
    wsOut.Name = HOJA_SALIDA

    DibujarTitulo wsOut, "OTIF PROVEEDOR"

    
    Call EscribirBloqueProveedor(wsData, wsOut, 3, 2, DOC_NAC, "PROVEEDORES NACIONALES")
    Call EscribirBloqueProveedor(wsData, wsOut, 3, 10, DOC_IMP, "PROVEEDORES EXTRANJEROS")

    wsOut.Columns.AutoFit
'========================
' FORMATO VISUAL
'========================

wsOut.Rows(4).RowHeight = 60

With wsOut.Rows(4)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
End With

' NACIONALES
wsOut.Columns("C").ColumnWidth = 12.7
wsOut.Columns("D").ColumnWidth = 13.1
wsOut.Columns("E").ColumnWidth = 10.7
wsOut.Columns("F").ColumnWidth = 10.7
wsOut.Columns("G").ColumnWidth = 5
wsOut.Columns("H").ColumnWidth = 6.1

' EXTRANJEROS
wsOut.Columns("K").ColumnWidth = 12.7
wsOut.Columns("L").ColumnWidth = 13.1
wsOut.Columns("M").ColumnWidth = 10.7
wsOut.Columns("N").ColumnWidth = 10.7
wsOut.Columns("O").ColumnWidth = 5
wsOut.Columns("P").ColumnWidth = 6.1
    


    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbCritical
    


End Sub

Private Sub EscribirBloqueProveedor(wsData As Worksheet, wsOut As Worksheet, _
                                   filaInicio As Long, colInicio As Long, _
                                   tipoDoc As String, tituloBloque As String)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    Dim proveedor As String, estado As String, doc As String

    lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row

    For r = 2 To lastRow

        proveedor = Trim(wsData.Range(COL_PROVEEDOR & r).Value)
        doc = UCase(Trim(wsData.Range(COL_DOC & r).Value))
        estado = UCase(Trim(wsData.Range(COL_ESTADO & r).Value))

        If doc = tipoDoc And proveedor <> "" Then

            If Not dict.Exists(proveedor) Then
                dict.Add proveedor, Array(0, 0, 0, 0)
            End If

            Dim vals
            vals = dict(proveedor)

            Select Case estado
                Case UCase(EST_1), UCase(EST_1_ALT): vals(0) = vals(0) + 1
                Case UCase(EST_2): vals(1) = vals(1) + 1
                Case UCase(EST_3): vals(2) = vals(2) + 1
                Case UCase(EST_4): vals(3) = vals(3) + 1
            End Select

            dict(proveedor) = vals
        End If
    Next r

    ' TITULO
    With wsOut.Range(wsOut.Cells(filaInicio, colInicio), wsOut.Cells(filaInicio, colInicio + 6))
        .Merge
        .Value = tituloBloque
        .Interior.Color = RGB(217, 217, 217)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    Dim filaCab As Long: filaCab = filaInicio + 1
    Dim filaDatos As Long: filaDatos = filaInicio + 2

    ' CABECERA
    wsOut.Cells(filaCab, colInicio).Value = "Proveedor"
    wsOut.Cells(filaCab, colInicio + 1).Value = EST_1
    wsOut.Cells(filaCab, colInicio + 2).Value = EST_2
    wsOut.Cells(filaCab, colInicio + 3).Value = EST_3
    wsOut.Cells(filaCab, colInicio + 4).Value = EST_4
    wsOut.Cells(filaCab, colInicio + 5).Value = "Total"
    wsOut.Cells(filaCab, colInicio + 6).Value = "%OTIF"

With wsOut.Range(wsOut.Cells(filaCab, colInicio), wsOut.Cells(filaCab, colInicio + 6))
    .Interior.Color = RGB(189, 215, 238)
    .Font.Bold = True
    .Borders.LineStyle = xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

    Dim i As Long, k As Variant, fila As Long
    fila = filaDatos

    For Each k In dict.Keys

        Dim v
        v = dict(k)

        Dim total As Long
        total = v(0) + v(1) + v(2) + v(3)

        wsOut.Cells(fila, colInicio).Value = k
        wsOut.Cells(fila, colInicio + 1).Value = v(0)
        wsOut.Cells(fila, colInicio + 2).Value = v(1)
        wsOut.Cells(fila, colInicio + 3).Value = v(2)
        wsOut.Cells(fila, colInicio + 4).Value = v(3)
        wsOut.Cells(fila, colInicio + 5).Value = total

        If total > 0 Then
            wsOut.Cells(fila, colInicio + 6).Value = v(2) / total
        End If

        wsOut.Cells(fila, colInicio + 6).NumberFormat = "0%"

        fila = fila + 1
    Next k
    
    Dim ultimaFila As Long
ultimaFila = fila - 1

With wsOut.Sort
    .SortFields.Clear
    
    ' Orden 1: Total (col +5)
    .SortFields.Add key:=wsOut.Range(wsOut.Cells(filaDatos, colInicio + 5), wsOut.Cells(ultimaFila, colInicio + 5)), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    ' Orden 2: %OTIF (col +6)
    .SortFields.Add key:=wsOut.Range(wsOut.Cells(filaDatos, colInicio + 6), wsOut.Cells(ultimaFila, colInicio + 6)), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    .SetRange wsOut.Range(wsOut.Cells(filaCab, colInicio), wsOut.Cells(ultimaFila, colInicio + 6))
    .Header = xlYes
    .Apply
End With

    With wsOut.Range(wsOut.Cells(filaDatos, colInicio), wsOut.Cells(fila - 1, colInicio + 6))
        .Borders.LineStyle = xlContinuous
    End With

End Sub

Private Sub DibujarTitulo(ws As Worksheet, txt As String)
    With ws.Range("B1:P1")
        .Merge
        .Value = txt
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Function HojaExiste(nombre As String, wb As Workbook) As Boolean
    On Error Resume Next
    HojaExiste = Not wb.Worksheets(nombre) Is Nothing
    On Error GoTo 0
End Function

