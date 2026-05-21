Attribute VB_Name = "M1_ME2N"
Option Explicit

Private Const RUTA_OTIF As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\"

Private Const NOMBRE_EXPORT As String = "OTIF_TEMP"
Private Const NOMBRE_XLSX As String = "OTIF_TEMP.xlsx"
Private Const VARIANTE_ME2N As String = "BLAPA2"
Private Const ESPERA_CARGA_MS As Long = 2500

Public Sub OTIF_01_Descargar_ME2N()

    On Error GoTo EH

    Dim SapGuiAuto As Object
    Dim gui As Object
    Dim con As Object
    Dim ses As Object

    Dim fechaDesde As Date
    Dim fechaHasta As Date

    Dim archivoHTML As String
    Dim archivoHTML2 As String
    Dim archivoXLS As String
    Dim archivoXLSX As String
    
    Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False
Application.CutCopyMode = False

    ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = "EN PROCESO"

    archivoHTML = RUTA_OTIF & NOMBRE_EXPORT & ".htm"
    archivoHTML2 = RUTA_OTIF & NOMBRE_EXPORT & ".html"
    archivoXLS = RUTA_OTIF & NOMBRE_EXPORT & ".xls"
    archivoXLSX = RUTA_OTIF & NOMBRE_XLSX

    fechaDesde = DateSerial(Year(Date), 1, 1)
    fechaHasta = Date
    


    EnsureFolderExists RUTA_OTIF
    BorrarSiExiste archivoHTML
    BorrarSiExiste archivoHTML2
    BorrarSiExiste archivoXLS
    BorrarSiExiste archivoXLSX

    Set SapGuiAuto = GetObject("SAPGUI")
    Set gui = SapGuiAuto.GetScriptingEngine
    Set con = gui.Children(0)
    Set ses = con.Children(0)

    ses.FindById("wnd[0]").SetFocus
    AppActivate ses.FindById("wnd[0]").Text
    DoEvents
    Sleep 300

    '========================
    ' Ir a ME2N
    '========================
    ses.FindById("wnd[0]/tbar[0]/okcd").Text = "/nME2N"
    ses.FindById("wnd[0]").SendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Aplicar variante
    '========================
    ses.FindById("wnd[0]/tbar[1]/btn[17]").Press
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    ses.FindById("wnd[1]/usr/txtV-LOW").Text = VARIANTE_ME2N
    ses.FindById("wnd[1]/usr/txtV-LOW").CaretPosition = Len(VARIANTE_ME2N)
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Fechas: inicio de año hasta hoy
    '========================
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = Format(fechaDesde, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-HIGH").Text = Format(fechaHasta, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-HIGH").SetFocus
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-HIGH").CaretPosition = 10

    '========================
    ' Ejecutar reporte
    '========================
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press
    EsperarALV ses, 90000
    Sleep ESPERA_CARGA_MS

    '========================
    ' Cambiar layout / vista si aplica
    '========================
    ses.FindById("wnd[0]").SendVKey 23
    EsperarALV ses, 90000
    Sleep ESPERA_CARGA_MS

    '========================
    ' Lanzar exportación y dejar control a PAD
    '========================
    ses.FindById("wnd[0]").SetFocus
    AppActivate ses.FindById("wnd[0]").Text
    DoEvents
    Sleep 500

    SendKeys "^+{F7}", True
    Sleep 1500

    '========================
    ' Validar que apareció Guardar como
    '========================
    'If Not VentanaGuardarComoActiva(10) Then
        'ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = _
         '  "ERROR: No apareció la ventana Guardar como"
     '   Exit Sub
   'End If

    '========================
    ' Entregar control a PAD
    '========================
    'ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = "WAIT_EXPORT_ME2N"
   
    
    Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True
Application.CutCopyMode = False
    Exit Sub

EH:
    ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = "ERROR: " & Err.Description
    
    Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True
Application.CutCopyMode = False

End Sub

Private Function VentanaGuardarComoActiva(Optional ByVal timeoutSeg As Long = 10) As Boolean

    Dim t As Double
    Dim sh As Object

    Set sh = CreateObject("WScript.Shell")
    t = Timer

    Do
        DoEvents

        If sh.AppActivate("Guardar como") Then
            VentanaGuardarComoActiva = True
            Exit Function
        End If

        If sh.AppActivate("Save As") Then
            VentanaGuardarComoActiva = True
            Exit Function
        End If

        Sleep 300

        If Timer - t >= timeoutSeg Then Exit Do
    Loop

    VentanaGuardarComoActiva = False
End Function

