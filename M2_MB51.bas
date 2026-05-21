Attribute VB_Name = "M2_MB51"
Option Explicit

Private Const RUTA_OTIF As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\"

Private Const VARIANTE_MB51 As String = "BLAPA3"
Private Const HOJA_CONTROL As String = "Control"
Private Const CELDA_ESTADO As String = "A2"
Private Const ESPERA_CARGA_MS As Long = 2500

Public Sub OTIF_02()

    On Error GoTo EH

    Dim SapGuiAuto As Object
    Dim gui As Object
    Dim con As Object
    Dim ses As Object

    Dim fechaDesde As Date
    Dim fechaHasta As Date
    
    Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False
Application.CutCopyMode = False

    ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = "EN PROCESO"

fechaDesde = DateSerial(Year(Date), 1, 1)
fechaHasta = Date

    '========================
    ' Conectar a SAP
    '========================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set gui = SapGuiAuto.GetScriptingEngine
    Set con = gui.Children(0)
    Set ses = con.Children(0)

    ses.FindById("wnd[0]").SetFocus
    AppActivate ses.FindById("wnd[0]").Text
    DoEvents
    Sleep 300

    '========================
    ' Ir a MB51
    '========================
    ses.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMB51"
    ses.FindById("wnd[0]").SendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Aplicar variante
    '========================
    ses.FindById("wnd[0]").SendVKey 17
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    ses.FindById("wnd[1]/usr/txtV-LOW").Text = VARIANTE_MB51
    ses.FindById("wnd[1]/usr/txtV-LOW").CaretPosition = Len(VARIANTE_MB51)
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Fechas: última semana
    '========================
    ses.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Format(fechaDesde, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Format(fechaHasta, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus
    ses.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").CaretPosition = 10
    ses.FindById("wnd[0]").SendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    '========================
    ' Ejecutar
    '========================
    ses.FindById("wnd[0]").SendVKey 8
    EsperarALV ses, 90000
    Sleep ESPERA_CARGA_MS

    '========================
    ' Cambiar de vista
    '========================
    ses.FindById("wnd[0]").SendVKey 48
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell 2, "MATNR"
    ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "2"
    Do While ses.Busy: DoEvents: Loop
    Sleep ESPERA_CARGA_MS

    '========================
    ' Entregar control a PAD ANTES de exportar
    '========================
    ses.FindById("wnd[0]").SetFocus
    AppActivate ses.FindById("wnd[0]").Text
    DoEvents
    Sleep 500

    ThisWorkbook.Sheets(HOJA_CONTROL).Range(CELDA_ESTADO).Value = "WAIT_EXPORT_MB51"
    
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

