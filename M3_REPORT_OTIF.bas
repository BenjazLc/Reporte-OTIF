Attribute VB_Name = "M3_REPORT_OTIF"
Option Explicit

Private Const RUTA_OTIF_MHTML As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\OTIF_TEMP.XLSX"

Private Const RUTA_EM_MHTML As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\ENTREGA_MERCANCIAS_TEMP.XLSX"

Private Const RUTA_REPORTE_XLSX As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\REPORTE OTIF.xlsx"

Private Const NOMBRE_HOJA_REPORTE As String = "OTIF"
Private Const HEADER_ROW As Long = 1

Public Sub OTIF_03_Crear_Reporte_OTIF()

    On Error GoTo EH

    Dim wbOTIF As Workbook, wsOTIF As Worksheet
    Dim wbEM As Workbook, wsEM As Worksheet
    Dim wbNuevo As Workbook, wsNuevo As Worksheet
    Dim dict As Object

    Dim lastRowOTIF As Long, lastRowEM As Long
    Dim colFechaEntrega As Long
    Dim colResultado As Long
    Dim i As Long

    Dim idOTIF As String
    Dim idEM As String
    Dim valorEM As Variant

    Dim calcMode As XlCalculation
    Dim scrUpdate As Boolean
    Dim dispAlerts As Boolean
    Dim evtState As Boolean

    '-----------------------------------
    ' Guardar estado de Excel
    '-----------------------------------
    calcMode = Application.Calculation
    scrUpdate = Application.ScreenUpdating
    dispAlerts = Application.DisplayAlerts
    evtState = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '-----------------------------------
    ' Validaciones
    '-----------------------------------
    If Dir(RUTA_OTIF_MHTML) = "" Then
        Err.Raise vbObjectError + 1000, , "No se encontró OTIF_TEMP.MHTML"
    End If

    If Dir(RUTA_EM_MHTML) = "" Then
        Err.Raise vbObjectError + 1001, , "No se encontró ENTREGA_MERCANCIAS_TEMP.MHTML"
    End If

    '-----------------------------------
    ' Abrir archivos fuente
    '-----------------------------------
Set wbOTIF = ObtenerLibroAbiertoOAbrir(RUTA_OTIF_MHTML)
Set wsOTIF = wbOTIF.Worksheets(1)

Set wbEM = ObtenerLibroAbiertoOAbrir(RUTA_EM_MHTML)
Set wsEM = wbEM.Worksheets(1)
    '-----------------------------------
    ' Crear diccionario desde ENTREGA_MERCANCIAS
' ID = Pedido & Posición
' Valor a devolver = Fecha de entrada
    '-----------------------------------
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    lastRowEM = ultimaFila(wsEM, "C")

For i = HEADER_ROW + 1 To lastRowEM
    idEM = BuildKey(wsEM.Cells(i, "C"), wsEM.Cells(i, "D"))
    valorEM = wsEM.Cells(i, "A").Value

        If Len(idEM) > 0 Then
            ' Si hay repetidos, conserva el primero no vacío
            If Not dict.Exists(idEM) Then
                dict.Add idEM, valorEM
            ElseIf Len(CStr(dict(idEM))) = 0 And Len(CStr(valorEM)) > 0 Then
                dict(idEM) = valorEM
            End If
        End If
    Next i

    '-----------------------------------
    ' Crear nuevo workbook
    '-----------------------------------
    Set wbNuevo = Workbooks.Add(xlWBATWorksheet)
    Set wsNuevo = wbNuevo.Worksheets(1)
    wsNuevo.Name = NOMBRE_HOJA_REPORTE

    wsOTIF.UsedRange.Copy
    wsNuevo.Range("A1").PasteSpecial xlPasteValues
    wsNuevo.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    lastRowOTIF = ultimaFila(wsNuevo, "C")

    '-----------------------------------
    ' Insertar columna ID después de D
    ' ID = C & D
    '-----------------------------------
    wsNuevo.Columns("E:E").Insert Shift:=xlToRight
    wsNuevo.Cells(HEADER_ROW, "E").Value = "ID"

    For i = HEADER_ROW + 1 To lastRowOTIF
        idOTIF = BuildKey(wsNuevo.Cells(i, "C"), wsNuevo.Cells(i, "D"))
        wsNuevo.Cells(i, "E").Value = idOTIF
    Next i

    ' Formato numérico de la columna ID
    wsNuevo.Columns("E:E").NumberFormat = "0"

    '-----------------------------------
    ' Buscar "Fecha de entrega"
    '-----------------------------------
    colFechaEntrega = BuscarColumnaPorEncabezado(wsNuevo, "Fecha de entrega", HEADER_ROW)

    If colFechaEntrega = 0 Then
        Err.Raise vbObjectError + 1002, , "No se encontró el encabezado 'Fecha de entrega'."
    End If

    '-----------------------------------
    ' Insertar columna resultado después de Fecha de entrega
    '-----------------------------------
    colResultado = colFechaEntrega + 1
    wsNuevo.Columns(colResultado).Insert Shift:=xlToRight
    wsNuevo.Cells(HEADER_ROW, colResultado).Value = "Entrega de mercancías"

    '-----------------------------------
    ' Cruce por ID
    '-----------------------------------
    lastRowOTIF = ultimaFila(wsNuevo, "E")

    For i = HEADER_ROW + 1 To lastRowOTIF
        idOTIF = LimpiarClave(CStr(wsNuevo.Cells(i, "E").Value))

        If Len(idOTIF) > 0 Then
            If dict.Exists(idOTIF) Then
                wsNuevo.Cells(i, colResultado).Value = dict(idOTIF)
            Else
                wsNuevo.Cells(i, colResultado).Value = ""
            End If
        Else
            wsNuevo.Cells(i, colResultado).Value = ""
        End If
    Next i
    
    Dim colEstado As Long
Dim lastCol As Long
Dim fechaK As Variant
Dim fechaL As Variant
Dim valorI As Variant

'-----------------------------------
' Insertar columna Estado al final
'-----------------------------------
lastCol = wsNuevo.Cells(HEADER_ROW, wsNuevo.Columns.Count).End(xlToLeft).Column
colEstado = lastCol + 1

wsNuevo.Cells(HEADER_ROW, colEstado).Value = "Estado"

wsNuevo.Cells(HEADER_ROW, colEstado).Interior.Color = wsNuevo.Cells(HEADER_ROW, colEstado - 1).Interior.Color
wsNuevo.Cells(HEADER_ROW, colEstado).Font.Bold = True
wsNuevo.Cells(HEADER_ROW, colEstado).Font.Color = wsNuevo.Cells(HEADER_ROW, colEstado - 1).Font.Color
wsNuevo.Cells(HEADER_ROW, colEstado).Borders.LineStyle = xlContinuous
wsNuevo.Cells(HEADER_ROW, colEstado).HorizontalAlignment = xlCenter

'-----------------------------------
' Calcular Estado
'-----------------------------------
For i = HEADER_ROW + 1 To lastRowOTIF

    valorI = wsNuevo.Cells(i, "I").Value
    fechaK = wsNuevo.Cells(i, "K").Value
    fechaL = wsNuevo.Cells(i, "L").Value

    If valorI = 0 Then

        If IsDate(fechaL) And IsDate(fechaK) Then
            If fechaL <= fechaK Then
                wsNuevo.Cells(i, colEstado).Value = "ENTREGADO A TIEMPO"
            Else
                wsNuevo.Cells(i, colEstado).Value = "ENTREGADO ATRASADO"
            End If
        Else
            wsNuevo.Cells(i, colEstado).Value = ""
        End If

    Else

        If IsDate(fechaK) Then
            If fechaK >= Date Then
                wsNuevo.Cells(i, colEstado).Value = "A TIEMPO PENDIENTE DE CONFIRMAR"
            Else
                wsNuevo.Cells(i, colEstado).Value = "ATRASADO PENDIENTE DE CONFIRMAR"
            End If
        Else
            wsNuevo.Cells(i, colEstado).Value = ""
        End If

    End If

Next i

    '-----------------------------------
    ' Ajustes visuales
    '-----------------------------------
    wsNuevo.Rows(HEADER_ROW).Font.Bold = True
    With wsNuevo.Rows(HEADER_ROW)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
    wsNuevo.Cells.EntireColumn.AutoFit


    '-----------------------------------
' Guardar con fecha
'-----------------------------------
Dim rutaFinal As String
Dim fechaHoy As String

fechaHoy = Format(Date, "dd-mm-yyyy")

rutaFinal = Replace(RUTA_REPORTE_XLSX, ".xlsx", "_" & fechaHoy & ".xlsx")

wbNuevo.SaveAs Filename:=rutaFinal, FileFormat:=xlOpenXMLWorkbook
    '-----------------------------------
    ' Cerrar fuentes
    '-----------------------------------
    wbOTIF.Close SaveChanges:=False
    wbEM.Close SaveChanges:=False

    '-----------------------------------
    ' Restaurar Excel
    '-----------------------------------
    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode

    
    Exit Sub

EH:
    Dim nErr As Long
    Dim dErr As String

    nErr = Err.Number
    dErr = Err.Description

    On Error Resume Next

    If Not wbOTIF Is Nothing Then wbOTIF.Close SaveChanges:=False
    If Not wbEM Is Nothing Then wbEM.Close SaveChanges:=False

    Application.ScreenUpdating = scrUpdate
    Application.DisplayAlerts = dispAlerts
    Application.EnableEvents = evtState
    Application.Calculation = calcMode

    MsgBox "Error en OTIF_03_Crear_Reporte_OTIF:" & vbCrLf & _
           "Número: " & nErr & vbCrLf & _
           "Descripción: " & dErr, vbCritical
           
End Sub

Private Function ultimaFila(ws As Worksheet, ByVal colRef As Variant) As Long
    ultimaFila = ws.Cells(ws.Rows.Count, colRef).End(xlUp).Row
End Function

Private Function BuscarColumnaPorEncabezado(ws As Worksheet, _
                                            ByVal encabezado As String, _
                                            Optional ByVal filaEncabezado As Long = 1) As Long
    Dim ultCol As Long
    Dim c As Long
    Dim txtCelda As String
    Dim txtBuscado As String

    ultCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    txtBuscado = UCase(Trim(encabezado))

    For c = 1 To ultCol
        txtCelda = UCase(Trim(CStr(ws.Cells(filaEncabezado, c).Value)))
        If txtCelda = txtBuscado Then
            BuscarColumnaPorEncabezado = c
            Exit Function
        End If
    Next c

    BuscarColumnaPorEncabezado = 0
End Function

Private Function BuildKey(c1 As Range, c2 As Range) As String
    Dim p1 As String
    Dim p2 As String

    ' Usar .Text porque en MHTML muchas veces refleja mejor lo visible
    p1 = LimpiarClave(CStr(c1.Text))
    p2 = LimpiarClave(CStr(c2.Text))

    BuildKey = p1 & p2
End Function

Private Function LimpiarClave(ByVal s As String) As String
    s = Replace(s, Chr(160), "")          ' espacio duro / invisible
    s = Replace(s, vbTab, "")
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    s = Trim$(s)

    LimpiarClave = s
End Function

Private Function ObtenerLibroAbiertoOAbrir(ByVal rutaCompleta As String) As Workbook

    Dim wb As Workbook
    Dim nombreBase As String

    nombreBase = UCase$(Replace(Dir(rutaCompleta), ".XLSX", ""))

    For Each wb In Application.Workbooks
        If InStr(1, UCase$(wb.Name), nombreBase, vbTextCompare) > 0 Then
            Set ObtenerLibroAbiertoOAbrir = wb
            Exit Function
        End If
    Next wb

    Set ObtenerLibroAbiertoOAbrir = Workbooks.Open(rutaCompleta, ReadOnly:=True)

End Function
