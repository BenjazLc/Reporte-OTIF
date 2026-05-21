Attribute VB_Name = "Utilidades"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const HOJA_CONTROL As String = "Control"
Public Const CELDA_ESTADO As String = "A2"

Public Const ESPERA_SAVEAS_SEG As Double = 5
Public Const WAIT_KEY_MS As Long = 250
Public Const ESPERA_ABRIR_ARCHIVO_MS As Long = 2000

Public Function GuardarComo_ME2M_7Tabs(ByVal ruta As String, ByVal nombreSinExt As String) As Boolean
    On Error GoTo EH

    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"

    If Not ActivateSaveAsWindow(15000) Then
        GuardarComo_ME2M_7Tabs = False
        Exit Function
    End If

    Sleep 500

    '--- Poner ruta ---
    SendKeys "^{A}", True
    Sleep WAIT_KEY_MS
    SendKeys ruta, True
    Sleep WAIT_KEY_MS
    SendKeys "{ENTER}", True
    Sleep 900

    '--- Poner nombre ---
    SendKeys "%n", True
    Sleep WAIT_KEY_MS
    SendKeys "^{A}", True
    Sleep WAIT_KEY_MS
    SendKeys nombreSinExt, True
    Sleep WAIT_KEY_MS
    SendKeys "{ENTER}", True
    Sleep 500

    '--- Confirmar reemplazo si aparece ---
    SendKeys "s", True
    Sleep 200
    SendKeys "{ENTER}", True

    GuardarComo_ME2M_7Tabs = True
    Exit Function

EH:
    GuardarComo_ME2M_7Tabs = False
End Function

Public Function ActivateSaveAsWindow(ByVal timeoutMs As Long) As Boolean
    Dim t As Double
    t = Timer

    Do
        DoEvents
        On Error Resume Next

        Err.Clear
        AppActivate "Guardar como"
        If Err.Number = 0 Then
            ActivateSaveAsWindow = True
            Exit Function
        End If

        Err.Clear
        AppActivate "Save As"
        If Err.Number = 0 Then
            ActivateSaveAsWindow = True
            Exit Function
        End If

        On Error GoTo 0
        Sleep 200
    Loop While (Timer - t) * 1000 < timeoutMs

    ActivateSaveAsWindow = False
End Function

Public Sub ResolverPopupsSAP_PostGuardar(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double
    t = Timer

    Do
        DoEvents

        PermitirSeguridadSAP ses, 500

        If AceptarCualquierPopupSAP(ses) Then
            Sleep 250
        Else
            Sleep 200
        End If

        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop
End Sub

Public Function AceptarCualquierPopupSAP(ByVal ses As Object) As Boolean
    On Error Resume Next
    AceptarCualquierPopupSAP = False

    If Not ExisteVentana(ses, 1) Then Exit Function

    If TryPress(ses, "wnd[1]/usr/btnSPOP-OPTION1") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/tbar[0]/btn[0]") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnBUTTON_1") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnBTN_YES") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnB_YES") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/tbar[0]/btn[11]") Then AceptarCualquierPopupSAP = True: Exit Function
End Function

Public Sub PermitirSeguridadSAP(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double
    t = Timer

    Do
        DoEvents
        On Error Resume Next

        If ExisteVentana(ses, 1) Then
            Dim titulo As String
            titulo = ses.FindById("wnd[1]").Text

            If InStr(1, titulo, "Seguridad", vbTextCompare) > 0 Or _
               InStr(1, titulo, "Security", vbTextCompare) > 0 Then

                If Not TryPress(ses, "wnd[1]/tbar[0]/btn[0]") Then
                    Call TryPress(ses, "wnd[1]/usr/btnSPOP-OPTION1")
                End If
                Exit Sub
            End If
        End If

        On Error GoTo 0
        Sleep 200
        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop
End Sub

Public Function TryPress(ByVal ses As Object, ByVal id As String) As Boolean
    On Error Resume Next
    Err.Clear
    ses.FindById(id).Press
    TryPress = (Err.Number = 0)
    Err.Clear
End Function

Public Sub EsperarALV(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double
    t = Timer

    Do
        DoEvents
        Sleep 300
        If Not ses.Busy Then Exit Do
        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop

    Application.Wait Now + TimeValue("0:00:06")
End Sub

Public Function ExisteVentana(ByVal ses As Object, ByVal idx As Integer) As Boolean
    On Error Resume Next
    ExisteVentana = Not ses.FindById("wnd[" & idx & "]") Is Nothing
End Function

Public Function EsperarArchivoHistoricoListo(ByVal carpeta As String, ByVal nombreBase As String, ByVal timeoutMs As Long) As Boolean
    Dim t0 As Double
    Dim rutaArchivo As String
    Dim tam1 As Double
    Dim tam2 As Double

    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    t0 = Timer

    Do
        DoEvents
        Sleep 500

        rutaArchivo = ObtenerArchivoPorBase(carpeta, nombreBase)

        If rutaArchivo <> "" Then
            tam1 = TamanoArchivoSeguro(rutaArchivo)
            Sleep 1000
            DoEvents
            tam2 = TamanoArchivoSeguro(rutaArchivo)

            If tam1 > 0 And tam1 = tam2 Then
                EsperarArchivoHistoricoListo = True
                Exit Function
            End If
        End If

        If (Timer - t0) * 1000 > timeoutMs Then Exit Do
    Loop

    EsperarArchivoHistoricoListo = False
End Function

Public Function ObtenerArchivoPorBase(ByVal carpeta As String, ByVal nombreBase As String) As String
    Dim f As String

    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    f = Dir(carpeta & nombreBase & ".*")
    If f <> "" Then
        ObtenerArchivoPorBase = carpeta & f
    Else
        ObtenerArchivoPorBase = ""
    End If
End Function

Public Function TamanoArchivoSeguro(ByVal rutaArchivo As String) As Double
    On Error Resume Next
    TamanoArchivoSeguro = FileLen(rutaArchivo)
    On Error GoTo 0
End Function

Public Sub ConvertirHtmlASimpleXlsx(ByVal archivoOrigen As String, ByVal archivoXLSX As String)
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOrigen = Workbooks.Open(Filename:=archivoOrigen, ReadOnly:=True)
    Sleep ESPERA_ABRIR_ARCHIVO_MS

    Set wsOrigen = wbOrigen.Sheets(1)

    Set wbNuevo = Workbooks.Add(xlWBATWorksheet)
    Set wsDestino = wbNuevo.Sheets(1)
    wsDestino.Name = "DATA"

    wsOrigen.UsedRange.Copy
    wsDestino.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    wsDestino.Range("A1").PasteSpecial xlPasteFormats

    Application.CutCopyMode = False
    AjustarColumnas wsDestino

    wbNuevo.SaveAs Filename:=archivoXLSX, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=True
    wbOrigen.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub EnsureFolderExists(ByVal ruta As String)
    If Len(Dir(ruta, vbDirectory)) = 0 Then MkDir ruta
End Sub

Public Sub BorrarSiExiste(ByVal ruta As String)
    On Error Resume Next
    If Len(Dir(ruta)) > 0 Then Kill ruta
    On Error GoTo 0
End Sub

Public Sub AjustarColumnas(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.EntireColumn.AutoFit
    ws.Rows(1).Font.Bold = True
    On Error GoTo 0
End Sub

