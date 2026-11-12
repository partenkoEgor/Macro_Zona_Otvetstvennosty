Attribute VB_Name = "ZonaOtvetstvennosti"
Option Explicit

'=====================================================================================
' Макрос «Зоны ответственности»
'
' Берёт выгрузку тикетов (лист "Report"), создаёт в этой же книге 7 листов
' с теми же названиями и столбцами, что в гугл-таблице, и раскладывает
' тикеты по листам. Дальше строки остаётся скопировать в гугл-таблицу.
'
' Запуск: открыть файл выгрузки -> Alt+F8 -> RazlozhitPoListam -> Выполнить.
' Повторный запуск пересоздаёт листы заново.
'
' Все правила отбора собраны в блоке НАСТРОЙКИ ниже.
' Списки пишутся через точку с запятой ";".
' Регистр букв, лишние пробелы, хвост "(M)" у статусов и вид апострофа (' или ’)
' при сравнении не важны.
'=====================================================================================

'------------------------------------ НАСТРОЙКИ ------------------------------------

' Лист с выгрузкой. Если такого листа нет, берётся первый лист книги.
Private Const SRC_SHEET As String = "Report"

' Пороги: сколько часов тикет должен висеть в текущем статусе
' (от Processing date до момента выгрузки), чтобы попасть на лист.
Private Const HOURS_BTM As Double = 100    ' T&A x Operations B | BT M
Private Const HOURS_PSP As Double = 24     ' Operations B | PSP
Private Const HOURS_SMP As Double = 24     ' Operations B | SMP M
Private Const HOURS_API As Double = 24     ' Operations B | API и Operations A | API
Private Const HOURS_QC As Double = 0       ' QC | Mena Leads 1x (0 = без порога)
Private Const HOURS_CRIT As Double = 72    ' с какого часа ставить "Да" в «Критический 72ч+»

' Метка в Internal comment для листа T&A
Private Const TA_MARK As String = "waiting for TA"
' Статус, при котором тикет с меткой уходит на T&A ("In progress (M)" и "In progress")
Private Const TA_STATUS As String = "In progress"

' Внешние статусы, которые берутся для ВСЕХ гео (листы BT M, SMP M, QC)
Private Const ST_ALL_GEO As String = "Received;Approved;Create new transaction;Received (Fraud);" & _
    "Credited to another account;Revision needed"

' Внешние статусы, которые берутся ТОЛЬКО для активных регионов (листы BT M, QC)
Private Const ST_ACTIVE_GEO As String = "Higher quality and resolution file;Not Received;" & _
    "Request a statement for deposit;Refund details needed;Request deposit screenshot;" & _
    "Returned to sender's account;File does not match ticket"

' Активные регионы
Private Const ACTIVE_GEO As String = "Afghanistan;Azerbaijan;Bolivia;Canada;Guatemala;Honduras;Iran;" & _
    "Jamaica;Kyrgyzstan;Nicaragua;Panama;Paraguay;Taiwan"

' Operations B | PSP: внутренние статусы (столбец Status) и гео
Private Const PSP_STATUS As String = "Customer provided additional info;In progress;New"
Private Const PSP_GEO As String = "Azerbaijan;Turkey;Kyrgyzstan;Somalia;Iran"

' Operations A/B | API: внутренние статусы (столбец Status) и гео команд
Private Const API_STATUS As String = "In progress;Revising a pending request;" & _
    "Customer provided additional info;New;Awaiting response from PS"
Private Const API_GEO_A As String = "Algeria;Egypt;Morocco;Tunisia;Turkey"
Private Const API_GEO_B As String = "Azerbaijan;Bahrain;Somalia;Jordan;Oman;Saudi Arabia;" & _
    "Kyrgyzstan;Iran;Iraq"

' Департаменты и типы тикетов
Private Const DEPT_MAIN As String = "MENA 1x"
Private Const DEPT_LEADS As String = "Mena Leads 1x"
Private Const QC_TYPES As String = "API;BT M;SMP M"

'---------------------------------- КОНЕЦ НАСТРОЕК ----------------------------------

' Номера листов (порядок как в гугл-таблице)
Private Const SH_BTM As Long = 1
Private Const SH_TA As Long = 2
Private Const SH_PSP As Long = 3
Private Const SH_SMP As Long = 4
Private Const SH_API_B As Long = 5
Private Const SH_API_A As Long = 6
Private Const SH_QC As Long = 7
Private Const SH_COUNT As Long = 7

' Подготовленные (нормализованные) списки
Private aStAll As Variant, aStActive As Variant, aActiveGeo As Variant
Private aPspSt As Variant, aPspGeo As Variant
Private aApiSt As Variant, aGeoA As Variant, aGeoB As Variant
Private aQcTypes As Variant
Private kTaMark As String, kTaStatus As String, kDeptMain As String, kDeptLeads As String

' Номера столбцов выгрузки
Private cType As Long, cDept As Long, cGeo As Long, cRef As Long, cTid As Long, cUid As Long
Private cExt As Long, cStatus As Long, cCreated As Long, cProc As Long, cComment As Long

'=====================================================================================
' ГЛАВНЫЙ МАКРОС
'=====================================================================================
Public Sub RazlozhitPoListam()
    Dim wb As Workbook, src As Worksheet
    Dim data As Variant, nRows As Long, r As Long, k As Long
    Dim exportTime As Date, timeSource As String
    Dim tgt() As Long, hrs() As Double
    Dim cnt(1 To SH_COUNT) As Long
    Dim t0 As Double, msg As String, missing As String
    Dim prevCalc As Long

    t0 = Timer
    Set wb = ActiveWorkbook
    Set src = FindSourceSheet(wb)
    If src Is Nothing Then
        MsgBox U("\u041d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u043b\u0438\u0441\u0442 \u0441 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u043e\u0439."), vbExclamation
        Exit Sub
    End If

    data = src.UsedRange.Value
    If Not IsArray(data) Then
        MsgBox U("\u041b\u0438\u0441\u0442 \u0441 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u043e\u0439 \u043f\u0443\u0441\u0442\u043e\u0439."), vbExclamation
        Exit Sub
    End If
    nRows = UBound(data, 1)

    missing = MapColumns(data)
    If Len(missing) > 0 Then
        MsgBox U("\u0412 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0435 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b \u0441\u0442\u043e\u043b\u0431\u0446\u044b:") & vbLf & missing, vbExclamation
        Exit Sub
    End If

    InitLists

    ' Момент выгрузки: из имени файла (..._ГГГГММДДччммсс.xlsx),
    ' иначе - самая поздняя дата Processing date в выгрузке.
    If Not ExportTimeFromName(wb.Name, exportTime) Then
        exportTime = MaxDate(data, cProc)
        timeSource = U("\u0441\u0430\u043c\u0430\u044f \u043f\u043e\u0437\u0434\u043d\u044f\u044f Processing date (\u0432 \u0438\u043c\u0435\u043d\u0438 \u0444\u0430\u0439\u043b\u0430 \u043d\u0435\u0442 \u0434\u0430\u0442\u044b)")
    Else
        timeSource = U("\u0438\u0437 \u0438\u043c\u0435\u043d\u0438 \u0444\u0430\u0439\u043b\u0430")
    End If

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    ' 1. Определяем лист для каждой строки
    ReDim tgt(1 To nRows)
    ReDim hrs(1 To nRows)
    For r = 2 To nRows
        hrs(r) = HoursInStatus(data(r, cProc), exportTime)
        tgt(r) = Classify(data, r, hrs(r))
        If tgt(r) > 0 Then cnt(tgt(r)) = cnt(tgt(r)) + 1
    Next r

    ' 2. Пересоздаём листы и выводим данные
    Application.DisplayAlerts = False
    For k = 1 To SH_COUNT
        DeleteSheetIfExists wb, SheetName(k), src
    Next k
    Application.DisplayAlerts = True

    Dim prevWs As Worksheet
    Set prevWs = src
    For k = 1 To SH_COUNT
        Set prevWs = BuildSheet(wb, prevWs, k, data, tgt, hrs, cnt(k), nRows)
    Next k

    src.Activate
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    ' 3. Итог
    Dim total As Long
    msg = U("\u0413\u043e\u0442\u043e\u0432\u043e \u0437\u0430 ") & Format(Timer - t0, "0.0") & U(" \u0441\u0435\u043a.") & vbLf & vbLf
    msg = msg & U("\u041c\u043e\u043c\u0435\u043d\u0442 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438: ") & Format(exportTime, "dd.mm.yyyy hh:mm:ss") & " (" & timeSource & ")" & vbLf & vbLf
    For k = 1 To SH_COUNT
        msg = msg & SheetName(k) & ": " & cnt(k) & vbLf
        total = total + cnt(k)
    Next k
    msg = msg & vbLf & U("\u0412\u0441\u0435\u0433\u043e \u0440\u0430\u0437\u043b\u043e\u0436\u0435\u043d\u043e: ") & total & U(" \u0438\u0437 ") & (nRows - 1)
    MsgBox msg, vbInformation, U("\u0417\u043e\u043d\u044b \u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0441\u0442\u0438")
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    MsgBox U("\u041e\u0448\u0438\u0431\u043a\u0430: ") & Err.Description, vbCritical
End Sub

'=====================================================================================
' ПРАВИЛА ОТБОРА
' Возвращает номер листа (SH_...) или 0, если тикет никуда не идёт.
'=====================================================================================
Private Function Classify(data As Variant, ByVal r As Long, ByVal h As Double) As Long
    Dim tType As String, dept As String, geo As String
    Dim ext As String, st As String, cm As String

    tType = NormKey(data(r, cType))
    dept = NormKey(data(r, cDept))
    geo = NormKey(data(r, cGeo))
    ext = NormStatus(data(r, cExt))
    st = NormStatus(data(r, cStatus))
    cm = LCase$(CStr(data(r, cComment)))

    ' T&A - все типы и департаменты; такие тикеты не попадают на другие листы
    If InStr(1, cm, kTaMark, vbBinaryCompare) > 0 And ext = kTaStatus Then
        Classify = SH_TA
        Exit Function
    End If

    If dept = kDeptMain Then
        Select Case tType
            Case "bt m"
                If h >= HOURS_BTM And ExtStatusRule(ext, geo) Then Classify = SH_BTM
            Case "psp"
                If h >= HOURS_PSP And InList(st, aPspSt) And InList(geo, aPspGeo) Then Classify = SH_PSP
            Case "smp m"
                If h >= HOURS_SMP And InList(ext, aStAll) Then Classify = SH_SMP
            Case "api"
                If h >= HOURS_API And InList(st, aApiSt) Then
                    If InList(geo, aGeoA) Then
                        Classify = SH_API_A
                    ElseIf InList(geo, aGeoB) Then
                        Classify = SH_API_B
                    End If
                End If
        End Select
    ElseIf dept = kDeptLeads Then
        If InList(tType, aQcTypes) And h >= HOURS_QC And ExtStatusRule(ext, geo) Then Classify = SH_QC
    End If
End Function

' Статусы для всех гео + статусы только для активных регионов
Private Function ExtStatusRule(ByVal ext As String, ByVal geo As String) As Boolean
    If InList(ext, aStAll) Then
        ExtStatusRule = True
    ElseIf InList(ext, aStActive) Then
        ExtStatusRule = InList(geo, aActiveGeo)
    End If
End Function

'=====================================================================================
' ОПИСАНИЕ ЛИСТОВ: название и столбцы (как в гугл-таблице)
'=====================================================================================
Private Function SheetName(ByVal k As Long) As String
    Select Case k
        Case SH_BTM:   SheetName = "T&A x Operations B | BT M"
        Case SH_TA:    SheetName = "T&A"
        Case SH_PSP:   SheetName = "Operations B | PSP"
        Case SH_SMP:   SheetName = "Operations B | SMP M"
        Case SH_API_B: SheetName = "Operations B | API"
        Case SH_API_A: SheetName = "Operations A | API"
        Case SH_QC:    SheetName = "QC | Mena Leads 1x"
    End Select
End Function

' Коды полей:
'   TYPE - Тип тикета        DEPT - Департамент       GEO - ГЕО          REF - Реферал
'   TID  - TICKET ID          UID  - USER ID           EXT - Внешний статус
'   INT  - Внутренний статус  CRT  - Дата создания     INS - Во внешнем статусе с
'   HRS  - Время обработки    CRIT - Критический 72ч+  PROC - Статус обработки (пусто)
'   TCOM - Комментарий тикета COM  - Комментарий (пусто)
Private Function SheetFields(ByVal k As Long) As Variant
    Select Case k
        Case SH_TA, SH_QC
            SheetFields = Split("TYPE;DEPT;GEO;REF;TID;UID;EXT;INT;CRT;INS;HRS;CRIT;PROC;TCOM;COM", ";")
        Case SH_PSP
            SheetFields = Split("TYPE;DEPT;GEO;REF;TID;UID;INT;CRT;INS;HRS;CRIT;PROC;TCOM;COM", ";")
        Case Else
            SheetFields = Split("TYPE;DEPT;GEO;REF;TID;UID;EXT;CRT;INS;HRS;CRIT;PROC;TCOM;COM", ";")
    End Select
End Function

' Заголовки - в точности как в гугл-таблице (включая опечатки)
Private Function FieldHeader(ByVal code As String, ByVal k As Long) As String
    Select Case code
        Case "TYPE": FieldHeader = U("\u0422\u0438\u043f \u0442\u0438\u043a\u0435\u0442\u0430")
        Case "DEPT": FieldHeader = U("\u0414\u0435\u043f\u0430\u0440\u0442\u0430\u043c\u0435\u043d\u0442")
        Case "GEO":  FieldHeader = U("\u0413\u0415\u041e")
        Case "REF":  FieldHeader = U("\u0420\u0435\u0444\u0435\u0440\u0430\u043b")
        Case "TID":  FieldHeader = "TICKET ID"
        Case "UID":  FieldHeader = "USER ID"
        Case "EXT":  FieldHeader = U("\u0412\u043d\u0435\u0448\u043d\u0438\u0439 \u0441\u0442\u0430\u0442\u0443\u0441")
        Case "INT":  FieldHeader = U("\u0412\u043d\u0443\u0442\u0440\u0435\u043d\u043d\u0438\u0439 \u0441\u0442\u0430\u0442\u0443\u0441")
        Case "CRT":  FieldHeader = U("\u0414\u0430\u0442\u0430 \u0441\u043e\u0437\u0434\u0430\u043d\u0438\u044f (MSK)")
        Case "INS":  FieldHeader = U("\u0412\u043e \u0432\u043d\u0435\u0448\u0435\u043c \u0441\u0442\u0430\u0442\u0443\u0441\u0435 \u0441 (\u041c\u0421\u041a)")
        Case "HRS":  FieldHeader = U("\u0412\u0440\u0435\u043c\u044f \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438 (\u0427\u0430\u0441\u044b)")
        Case "CRIT": FieldHeader = U("\u041a\u0440\u0438\u0442\u0438\u0447\u0435\u0441\u043a\u0438\u0439") & vbLf & U("72\u0447+")
        Case "PROC": FieldHeader = U("C\u0442\u0430\u0442\u0443\u0441 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438")
        Case "TCOM"
            If k = SH_PSP Then FieldHeader = U("\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439 \u0442\u0438\u043a\u0435\u0442\u0430") Else FieldHeader = U("\u041a\u043e\u043c\u043c\u0435\u043d\u0430\u0440\u0438\u0439 \u0442\u0438\u043a\u0435\u0442\u0430")
        Case "COM":  FieldHeader = U("\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439")
    End Select
End Function

'=====================================================================================
' СОЗДАНИЕ И ЗАПОЛНЕНИЕ ЛИСТА
'=====================================================================================
Private Function BuildSheet(wb As Workbook, afterWs As Worksheet, ByVal k As Long, data As Variant, _
                            tgt() As Long, hrs() As Double, ByVal n As Long, ByVal nRows As Long) As Worksheet
    Dim ws As Worksheet, flds As Variant, nC As Long, c As Long, i As Long, r As Long
    Dim idx() As Long, out() As Variant, code As String

    Set ws = wb.Worksheets.Add(After:=afterWs)
    ws.Name = SheetName(k)

    flds = SheetFields(k)
    nC = UBound(flds) - LBound(flds) + 1

    ' Строки этого листа, отсортированные: дольше всех висящие - сверху
    If n > 0 Then
        ReDim idx(1 To n)
        i = 0
        For r = 2 To nRows
            If tgt(r) = k Then
                i = i + 1
                idx(i) = r
            End If
        Next r
        SortByHoursDesc idx, hrs, 1, n
    End If

    ReDim out(1 To n + 1, 1 To nC)
    For c = 1 To nC
        out(1, c) = FieldHeader(CStr(flds(c - 1)), k)
    Next c

    For i = 1 To n
        r = idx(i)
        For c = 1 To nC
            code = CStr(flds(c - 1))
            Select Case code
                Case "TYPE": out(i + 1, c) = CStr(data(r, cType))
                Case "DEPT": out(i + 1, c) = CStr(data(r, cDept))
                Case "GEO":  out(i + 1, c) = CStr(data(r, cGeo))
                Case "REF":  out(i + 1, c) = CStr(data(r, cRef))
                Case "TID":  out(i + 1, c) = ToNumberOrText(data(r, cTid))
                Case "UID":  out(i + 1, c) = ToNumberOrText(data(r, cUid))
                Case "EXT":  out(i + 1, c) = CStr(data(r, cExt))
                Case "INT":  out(i + 1, c) = CStr(data(r, cStatus))
                Case "CRT":  out(i + 1, c) = DateOrEmpty(data(r, cCreated))
                Case "INS":  out(i + 1, c) = DateOrEmpty(data(r, cProc))
                Case "HRS":  out(i + 1, c) = Int(hrs(r))
                Case "CRIT"
                    If hrs(r) >= HOURS_CRIT Then out(i + 1, c) = U("\u0414\u0430") Else out(i + 1, c) = Empty
                Case "TCOM": out(i + 1, c) = CStr(data(r, cComment))
                Case Else:   out(i + 1, c) = Empty   ' PROC, COM - заполняются вручную
            End Select
        Next c
    Next i

    ' Форматы столбцов ставим ДО записи: текстовые столбцы - "@",
    ' чтобы комментарии вида "=..." или "-..." не превратились в формулы.
    For c = 1 To nC
        With ws.Range(ws.Cells(1, c), ws.Cells(n + 1, c))
            Select Case CStr(flds(c - 1))
                Case "TID", "UID", "HRS": .NumberFormat = "0"
                Case "CRT", "INS":        .NumberFormat = "dd.mm.yyyy hh:mm:ss"
                Case Else:                .NumberFormat = "@"
            End Select
        End With
    Next c

    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, nC)).Value = out

    ' Оформление
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nC))
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    If n > 0 Then ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, nC)).WrapText = False
    ws.Columns.ColumnWidth = 14
    For c = 1 To nC
        Select Case CStr(flds(c - 1))
            Case "CRT", "INS": ws.Columns(c).ColumnWidth = 20
            Case "EXT", "INT": ws.Columns(c).ColumnWidth = 30
            Case "TCOM":       ws.Columns(c).ColumnWidth = 50
            Case "COM":        ws.Columns(c).ColumnWidth = 30
        End Select
    Next c
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Set BuildSheet = ws
End Function

'=====================================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
'=====================================================================================

' Лист выгрузки: по имени, иначе первый лист, у которого в A1 "Number of reactions" или есть "Ticket ID"
Private Function FindSourceSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = LCase$(SRC_SHEET) Then
            Set FindSourceSheet = ws
            Exit Function
        End If
    Next ws
    For Each ws In wb.Worksheets
        If Not IsError(Application.Match("Ticket ID", ws.Rows(1), 0)) Then
            Set FindSourceSheet = ws
            Exit Function
        End If
    Next ws
End Function

' Находит нужные столбцы по заголовкам. Возвращает список ненайденных.
Private Function MapColumns(data As Variant) As String
    Dim m As String
    cType = HeaderCol(data, "Ticket type"):        If cType = 0 Then m = m & "Ticket type" & vbLf
    cDept = HeaderCol(data, "Department"):         If cDept = 0 Then m = m & "Department" & vbLf
    cGeo = HeaderCol(data, "Country"):             If cGeo = 0 Then m = m & "Country" & vbLf
    cRef = HeaderCol(data, "Referal"):             If cRef = 0 Then m = m & "Referal" & vbLf
    cTid = HeaderCol(data, "Ticket ID"):           If cTid = 0 Then m = m & "Ticket ID" & vbLf
    cUid = HeaderCol(data, "User ID"):             If cUid = 0 Then m = m & "User ID" & vbLf
    cExt = HeaderCol(data, "External Status"):     If cExt = 0 Then m = m & "External Status" & vbLf
    cStatus = HeaderCol(data, "Status"):              If cStatus = 0 Then m = m & "Status" & vbLf
    cCreated = HeaderCol(data, "Date created"):    If cCreated = 0 Then m = m & "Date created" & vbLf
    cProc = HeaderCol(data, "Processing date"):    If cProc = 0 Then m = m & "Processing date" & vbLf
    cComment = HeaderCol(data, "Internal comment"): If cComment = 0 Then m = m & "Internal comment" & vbLf
    MapColumns = m
End Function

Private Function HeaderCol(data As Variant, ByVal title As String) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If NormKey(data(1, c)) = NormKey(title) Then
            HeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Sub InitLists()
    aStAll = PrepList(ST_ALL_GEO, True)
    aStActive = PrepList(ST_ACTIVE_GEO, True)
    aActiveGeo = PrepList(ACTIVE_GEO, False)
    aPspSt = PrepList(PSP_STATUS, True)
    aPspGeo = PrepList(PSP_GEO, False)
    aApiSt = PrepList(API_STATUS, True)
    aGeoA = PrepList(API_GEO_A, False)
    aGeoB = PrepList(API_GEO_B, False)
    aQcTypes = PrepList(QC_TYPES, False)
    kTaMark = LCase$(TA_MARK)
    kTaStatus = NormStatus(TA_STATUS)
    kDeptMain = NormKey(DEPT_MAIN)
    kDeptLeads = NormKey(DEPT_LEADS)
End Sub

Private Function PrepList(ByVal s As String, ByVal isStatus As Boolean) As Variant
    Dim a As Variant, i As Long
    a = Split(s, ";")
    For i = LBound(a) To UBound(a)
        If isStatus Then a(i) = NormStatus(a(i)) Else a(i) = NormKey(a(i))
    Next i
    PrepList = a
End Function

Private Function InList(ByVal key As String, lst As Variant) As Boolean
    Dim i As Long
    For i = LBound(lst) To UBound(lst)
        If lst(i) = key Then
            InList = True
            Exit Function
        End If
    Next i
End Function

' Нижний регистр, без лишних пробелов, единый апостроф
Private Function NormKey(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, ChrW(8217), "'")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormKey = LCase$(Trim$(s))
End Function

' То же + без хвоста "(M)": "Received (M)" = "Received"
Private Function NormStatus(ByVal v As Variant) As String
    Dim s As String
    s = NormKey(v)
    If Right$(s, 3) = "(m)" Then s = Trim$(Left$(s, Len(s) - 3))
    NormStatus = s
End Function

' Часы от Processing date до момента выгрузки (если даты нет - 0)
Private Function HoursInStatus(ByVal v As Variant, ByVal exportTime As Date) As Double
    Dim d As Date
    If ParseDate(v, d) Then HoursInStatus = (CDbl(exportTime) - CDbl(d)) * 24#
End Function

' Дата из текста "ГГГГ-ММ-ДД чч:мм:сс" или из настоящей даты Excel
Private Function ParseDate(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim s As String
    On Error GoTo Bad
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = v
        ParseDate = True
        Exit Function
    End If
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        d = CDate(v)
        ParseDate = True
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        d = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
        If Len(s) >= 19 Then
            d = d + TimeSerial(Val(Mid$(s, 12, 2)), Val(Mid$(s, 15, 2)), Val(Mid$(s, 18, 2)))
        ElseIf Len(s) >= 16 Then
            d = d + TimeSerial(Val(Mid$(s, 12, 2)), Val(Mid$(s, 15, 2)), 0)
        End If
        ParseDate = True
    ElseIf Len(s) > 0 And IsDate(s) Then
        d = CDate(s)
        ParseDate = True
    End If
    Exit Function
Bad:
    ParseDate = False
End Function

Private Function DateOrEmpty(ByVal v As Variant) As Variant
    Dim d As Date
    If ParseDate(v, d) Then DateOrEmpty = d Else DateOrEmpty = Empty
End Function

' ID - числом (без ".0"), если это число; иначе как текст
Private Function ToNumberOrText(ByVal v As Variant) As Variant
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        ToNumberOrText = Empty
    ElseIf IsNumeric(v) Then
        ToNumberOrText = CDbl(v)
    Else
        ToNumberOrText = CStr(v)
    End If
End Function

' Ищет в имени файла 14 цифр подряд (ГГГГММДДччммсс), берёт последнее вхождение
Private Function ExportTimeFromName(ByVal nm As String, ByRef d As Date) As Boolean
    Dim i As Long, j As Long, s As String, ok As Boolean
    For i = Len(nm) - 13 To 1 Step -1
        s = Mid$(nm, i, 14)
        ok = True
        For j = 1 To 14
            If Mid$(s, j, 1) < "0" Or Mid$(s, j, 1) > "9" Then
                ok = False
                Exit For
            End If
        Next j
        If ok Then
            If i + 14 <= Len(nm) Then
                If IsNumeric(Mid$(nm, i + 14, 1)) Then ok = False
            End If
        End If
        If ok Then
            If Val(Left$(s, 4)) >= 2000 And Val(Left$(s, 4)) <= 2100 Then
                On Error GoTo Bad
                d = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 5, 2)), Val(Mid$(s, 7, 2))) + _
                    TimeSerial(Val(Mid$(s, 9, 2)), Val(Mid$(s, 11, 2)), Val(Mid$(s, 13, 2)))
                ExportTimeFromName = True
                Exit Function
            End If
        End If
    Next i
    Exit Function
Bad:
    ExportTimeFromName = False
End Function

Private Function MaxDate(data As Variant, ByVal col As Long) As Date
    Dim r As Long, d As Date, mx As Date, found As Boolean
    For r = 2 To UBound(data, 1)
        If ParseDate(data(r, col), d) Then
            If Not found Or d > mx Then mx = d
            found = True
        End If
    Next r
    If found Then MaxDate = mx Else MaxDate = Now
End Function

Private Sub DeleteSheetIfExists(wb As Workbook, ByVal nm As String, src As Worksheet)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = nm Then
            If Not ws Is src Then ws.Delete
            Exit Sub
        End If
    Next ws
End Sub

' Быстрая сортировка индексов по убыванию часов
Private Sub SortByHoursDesc(idx() As Long, hrs() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Long
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    p = hrs(idx((lo + hi) \ 2))
    Do While i <= j
        Do While hrs(idx(i)) > p: i = i + 1: Loop
        Do While hrs(idx(j)) < p: j = j - 1: Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortByHoursDesc idx, hrs, lo, j
    If i < hi Then SortByHoursDesc idx, hrs, i, hi
End Sub

' Раскодирует "\uXXXX" в символы. Нужна, чтобы русский текст в коде
' не зависел от языковых настроек Windows.
Private Function U(ByVal s As String) As String
    Dim i As Long, res As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            res = res & ChrW(CLng(Val("&H" & Mid$(s, i + 2, 4) & "&")))
            i = i + 6
        Else
            res = res & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = res
End Function
