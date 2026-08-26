Option Explicit

' ============================================================================
' GenerateWeeklyReport
'
' Rebuilds the Weekly Team Activity Report (Data, Weekly Report,
' Tickets Created This Week, Author Summary (Anonymized), Dashboard,
' Internal Ref - DO NOT SHARE) from a raw weekly worklog CSV export.
'
' HOW TO RUN: Alt+F8 -> select GenerateWeeklyReport -> Run.
' It will ask you to pick the raw CSV export, then rebuild every sheet
' in THIS workbook.
'
' Author names are anonymized to "User N" using a mapping that is stored
' in a hidden sheet called AuthorMap inside this same workbook, so the
' same person gets the same User N every week. Do not delete that sheet.
' ============================================================================

Sub GenerateWeeklyReport()

    Dim csvPath As Variant
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim thisWb As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, r As Long

    Dim colDate As Long, colTicket As Long, colSummary As Long, colHours As Long
    Dim colAuthor As Long, colStatus As Long, colPriority As Long

    Set thisWb = ThisWorkbook

    csvPath = Application.GetOpenFilename( _
        "Worklog export (*.csv;*.xlsx;*.xls), *.csv;*.xlsx;*.xls", , _
        "Select this week's raw worklog export")
    If csvPath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    Set csvWb = Workbooks.Open(csvPath)
    Set csvWs = csvWb.Sheets(1)

    lastRow = csvWs.Cells(csvWs.Rows.Count, 1).End(xlUp).Row
    lastCol = csvWs.Cells(1, csvWs.Columns.Count).End(xlToLeft).Column

    colDate = FindCol(csvWs, lastCol, Array("date", "worklogdate", "workdate", "started", "workdate"))
    colTicket = FindCol(csvWs, lastCol, Array("ticketkey", "issuekey", "ticket", "key", "workitemkey"))
    colSummary = FindCol(csvWs, lastCol, Array("summary", "issuesummary", "title", "workitemsummary"))
    colHours = FindCol(csvWs, lastCol, Array("hours", "timespent(hours)", "timespent", "hourslogged", "loggedhours"))
    colAuthor = FindCol(csvWs, lastCol, Array("author", "authorid", "authorid", "worklogauthor", "assignee", "user", "name", "fullname"))
    colStatus = FindCol(csvWs, lastCol, Array("status", "issuestatus", "workitemstatus"))
    colPriority = FindCol(csvWs, lastCol, Array("priority"))

    If colDate = 0 Or colTicket = 0 Or colSummary = 0 Or colHours = 0 _
        Or colAuthor = 0 Or colStatus = 0 Or colPriority = 0 Then
        MsgBox "Could not find all required columns in the export." & vbCrLf & _
               "Headers found: " & GetHeaderList(csvWs, lastCol) & vbCrLf & vbCrLf & _
               "Open the VBA editor (Alt+F11), find the FindCol calls near the top " & _
               "of GenerateWeeklyReport, and add your exact header text to the matching Array(...).", _
               vbExclamation, "Column not found"
        csvWb.Close False
        GoTo CleanExit
    End If

    ' ---- Read rows into arrays, dropping weekends -------------------------
    Dim n As Long
    Dim dArr() As Date, tArr() As String, sArr() As String, hArr() As Double
    Dim aArr() As String, stArr() As String, pArr() As String
    ReDim dArr(1 To lastRow): ReDim tArr(1 To lastRow): ReDim sArr(1 To lastRow)
    ReDim hArr(1 To lastRow): ReDim aArr(1 To lastRow): ReDim stArr(1 To lastRow)
    ReDim pArr(1 To lastRow)

    n = 0
    For i = 2 To lastRow
        If Trim(CStr(csvWs.Cells(i, colDate).Value)) <> "" Then
            Dim dt As Date
            dt = CDate(csvWs.Cells(i, colDate).Value)
            If Weekday(dt, vbMonday) <= 5 Then
                n = n + 1
                dArr(n) = DateSerial(Year(dt), Month(dt), Day(dt))
                tArr(n) = Trim(CStr(csvWs.Cells(i, colTicket).Value))
                sArr(n) = Trim(CStr(csvWs.Cells(i, colSummary).Value))
                hArr(n) = Val(csvWs.Cells(i, colHours).Value)
                aArr(n) = Trim(CStr(csvWs.Cells(i, colAuthor).Value))
                stArr(n) = UCase(Trim(CStr(csvWs.Cells(i, colStatus).Value)))
                pArr(n) = StrConv(Trim(CStr(csvWs.Cells(i, colPriority).Value)), vbProperCase)
            End If
        End If
    Next i
    csvWb.Close False

    If n = 0 Then
        MsgBox "No weekday rows found in that export.", vbExclamation
        GoTo CleanExit
    End If

    ' ---- Author mapping (persisted in hidden AuthorMap sheet) -------------
    Dim mapWs As Worksheet
    On Error Resume Next
    Set mapWs = thisWb.Sheets("AuthorMap")
    On Error GoTo CleanFail
    If mapWs Is Nothing Then
        Set mapWs = thisWb.Sheets.Add(After:=thisWb.Sheets(thisWb.Sheets.Count))
        mapWs.Name = "AuthorMap"
        mapWs.Cells(1, 1).Value = "RealName"
        mapWs.Cells(1, 2).Value = "UserID"
        mapWs.Visible = xlSheetVeryHidden
    End If

    Dim mapLast As Long
    mapLast = mapWs.Cells(mapWs.Rows.Count, 1).End(xlUp).Row
    If mapLast < 1 Then mapLast = 1

    Dim authorDict As Object
    Set authorDict = CreateObject("Scripting.Dictionary")
    For i = 2 To mapLast
        If mapWs.Cells(i, 1).Value <> "" Then
            authorDict(CStr(mapWs.Cells(i, 1).Value)) = CStr(mapWs.Cells(i, 2).Value)
        End If
    Next i

    Dim maxUserNum As Long
    maxUserNum = 0
    Dim k As Variant
    For Each k In authorDict.Keys
        Dim uidStr As String
        uidStr = authorDict(k)
        Dim num As Long
        num = Val(Mid(uidStr, InStrRev(uidStr, " ") + 1))
        If num > maxUserNum Then maxUserNum = num
    Next k

    Dim uniqueAuthors As Object
    Set uniqueAuthors = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Not uniqueAuthors.Exists(aArr(i)) Then uniqueAuthors.Add aArr(i), 1
    Next i

    For Each k In uniqueAuthors.Keys
        If Not authorDict.Exists(CStr(k)) Then
            maxUserNum = maxUserNum + 1
            authorDict(CStr(k)) = "User " & maxUserNum
            mapLast = mapLast + 1
            mapWs.Cells(mapLast, 1).Value = k
            mapWs.Cells(mapLast, 2).Value = "User " & maxUserNum
        End If
    Next k

    Dim uArr() As String
    ReDim uArr(1 To n)
    For i = 1 To n
        uArr(i) = authorDict(aArr(i))
    Next i

    Dim numAuthors As Long
    numAuthors = mapLast - 1
    Dim allAuthorIDs() As String
    ReDim allAuthorIDs(1 To numAuthors)
    For i = 1 To numAuthors
        allAuthorIDs(i) = mapWs.Cells(i + 1, 2).Value
    Next i
    Call SortUserIDs(allAuthorIDs)

    ' ---- Distinct sorted dates ---------------------------------------------
    Dim dateDict As Object
    Set dateDict = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        Dim dkey As String
        dkey = CStr(CLng(dArr(i)))
        If Not dateDict.Exists(dkey) Then dateDict.Add dkey, dArr(i)
    Next i
    Dim nd As Long
    nd = dateDict.Count
    Dim distinctDates() As Date
    ReDim distinctDates(1 To nd)
    Dim idx As Long
    idx = 0
    For Each k In dateDict.Keys
        idx = idx + 1
        distinctDates(idx) = dateDict(k)
    Next k
    Call SortDates(distinctDates)

    ' ---- First-seen date per ticket (for "Created" + "New Tickets") -------
    Dim firstSeenDict As Object
    Set firstSeenDict = CreateObject("Scripting.Dictionary")
    Dim ticketSummary As Object, ticketStatus As Object, ticketPriority As Object
    Set ticketSummary = CreateObject("Scripting.Dictionary")
    Set ticketStatus = CreateObject("Scripting.Dictionary")
    Set ticketPriority = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Not firstSeenDict.Exists(tArr(i)) Then
            firstSeenDict.Add tArr(i), dArr(i)
            ticketSummary.Add tArr(i), sArr(i)
            ticketStatus.Add tArr(i), stArr(i)
            ticketPriority.Add tArr(i), pArr(i)
        ElseIf dArr(i) < firstSeenDict(tArr(i)) Then
            firstSeenDict(tArr(i)) = dArr(i)
        End If
    Next i

    ' ---- Rebuild all sheets -------------------------------------------------
    Call ClearOrCreateSheet(thisWb, "Data")
    Call ClearOrCreateSheet(thisWb, "Weekly Report")
    Call ClearOrCreateSheet(thisWb, "Tickets Created This Week")
    Call ClearOrCreateSheet(thisWb, "Author Summary (Anonymized)")
    Call ClearOrCreateSheet(thisWb, "Dashboard")
    Call ClearOrCreateSheet(thisWb, "Internal Ref - DO NOT SHARE")

    ' ===== Data sheet =========================================================
    Dim wsData As Worksheet
    Set wsData = thisWb.Sheets("Data")
    Dim dataHeaders As Variant
    dataHeaders = Array("Date", "Ticket Key", "Summary", "Hours", "AuthorID", _
                         "Status", "Priority", "FirstAuthorTicketDay", "FirstAuthorDay", "FirstTicketWeek")
    For i = 0 To 9
        wsData.Cells(1, i + 1).Value = dataHeaders(i)
    Next i
    For i = 1 To n
        r = i + 1
        wsData.Cells(r, 1).Value = dArr(i)
        wsData.Cells(r, 1).NumberFormat = "yyyy-mm-dd"
        wsData.Cells(r, 2).Value = tArr(i)
        wsData.Cells(r, 3).Value = sArr(i)
        wsData.Cells(r, 4).Value = hArr(i)
        wsData.Cells(r, 5).Value = uArr(i)
        wsData.Cells(r, 6).Value = stArr(i)
        wsData.Cells(r, 7).Value = pArr(i)
        wsData.Cells(r, 8).Formula = "=IF(COUNTIFS($A$2:A" & r & ",A" & r & ",$B$2:B" & r & ",B" & r & ",$E$2:E" & r & ",E" & r & ")=1,1,0)"
        wsData.Cells(r, 9).Formula = "=IF(COUNTIFS($A$2:A" & r & ",A" & r & ",$E$2:E" & r & ",E" & r & ")=1,1,0)"
        wsData.Cells(r, 10).Formula = "=IF(COUNTIFS($B$2:B" & r & ",B" & r & ")=1,1,0)"
    Next i
    Dim lastDataRow As Long
    lastDataRow = n + 1
    wsData.Columns("A:J").AutoFit

    ' ===== Weekly Report sheet ===============================================
    Dim wsWR As Worksheet
    Set wsWR = thisWb.Sheets("Weekly Report")
    wsWR.Cells(1, 1).Value = "Weekly Team Activity Report"
    Call StyleTitle(wsWR.Cells(1, 1))
    wsWR.Cells(2, 1).Value = "Week of " & Format(distinctDates(1), "yyyy-mm-dd") & " to " & _
        Format(distinctDates(nd), "yyyy-mm-dd") & "  |  Author names withheld, shown as User 1-" & _
        numAuthors & "  |  Weekends excluded"
    Call StyleSubtitle(wsWR.Cells(2, 1))

    wsWR.Cells(4, 1).Value = "Week Summary"
    Call StyleDayHeader(wsWR.Cells(4, 1))
    wsWR.Cells(5, 1).Value = "Total Distinct Tickets Worked On This Week"
    wsWR.Cells(5, 5).Formula = "=SUM(Data!J2:J" & lastDataRow & ")"
    wsWR.Cells(6, 1).Value = "Total New Tickets Created This Week"
    wsWR.Cells(6, 5).Formula = "=SUM(Data!J2:J" & lastDataRow & ")"
    wsWR.Cells(6, 6).Value = "(tickets first appearing in this export - see note on 'Tickets Created' tab)"
    Call StyleSubtitle(wsWR.Cells(6, 6))
    wsWR.Cells(7, 1).Value = "Total Hours Logged This Week"
    wsWR.Cells(7, 5).Formula = "=ROUND(SUM(Data!D:D),2)"
    wsWR.Cells(8, 1).Value = "Team Size (Total Available Authors)"
    wsWR.Cells(8, 5).Formula = "=SUMPRODUCT(1/COUNTIF(Data!E2:E" & lastDataRow & ",Data!E2:E" & lastDataRow & "))"
    For i = 5 To 8
        Call StyleBold(wsWR.Cells(i, 5))
    Next i

    r = 10
    Dim di As Long
    For di = 1 To nd
        Dim dToday As Date
        dToday = distinctDates(di)

        Dim dayTickets As Object
        Set dayTickets = CreateObject("Scripting.Dictionary")
        Dim dayAuthors As Object
        Set dayAuthors = CreateObject("Scripting.Dictionary")
        For i = 1 To n
            If dArr(i) = dToday Then
                If Not dayTickets.Exists(tArr(i)) Then dayTickets.Add tArr(i), 1
                If Not dayAuthors.Exists(uArr(i)) Then dayAuthors.Add uArr(i), 1
            End If
        Next i

        Dim newTicketsToday As Long
        newTicketsToday = 0
        For Each k In dayTickets.Keys
            If firstSeenDict(CStr(k)) = dToday Then newTicketsToday = newTicketsToday + 1
        Next k

        wsWR.Cells(r, 1).Value = Format(dToday, "dddd") & ", " & Format(dToday, "yyyy-mm-dd") & _
            "    |    Team Availability: " & dayAuthors.Count & "/" & numAuthors & _
            "    |    New Tickets Created: " & newTicketsToday
        Call StyleDayHeader(wsWR.Cells(r, 1))
        r = r + 1

        Dim hdrs As Variant
        hdrs = Array("Ticket Key", "Summary", "Authors on Ticket", "Hours Logged", "Priority", "Status")
        For i = 0 To 5
            wsWR.Cells(r, i + 1).Value = hdrs(i)
            Call StyleTableHeader(wsWR.Cells(r, i + 1))
        Next i
        r = r + 1
        Dim firstDataRow As Long
        firstDataRow = r

        For Each k In dayTickets.Keys
            Dim tk As String
            tk = CStr(k)
            Dim dExcel As String
            dExcel = "DATE(" & Year(dToday) & "," & Month(dToday) & "," & Day(dToday) & ")"
            wsWR.Cells(r, 1).Value = tk
            wsWR.Cells(r, 2).Value = ticketSummary(tk)
            wsWR.Cells(r, 3).Formula = "=SUMPRODUCT((Data!$A$2:$A$" & lastDataRow & "=" & dExcel & ")*" & _
                "(Data!$B$2:$B$" & lastDataRow & "=""" & tk & """)*Data!$H$2:$H$" & lastDataRow & ")"
            wsWR.Cells(r, 4).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$A:$A," & dExcel & _
                ",Data!$B:$B,""" & tk & """),2)"
            wsWR.Cells(r, 5).Value = ticketPriority(tk)
            wsWR.Cells(r, 6).Value = ticketStatus(tk)
            r = r + 1
        Next k
        Dim lastDataRowForDay As Long
        lastDataRowForDay = r - 1

        wsWR.Cells(r, 1).Formula = "=""Day Total (""&COUNTA(A" & firstDataRow & ":A" & lastDataRowForDay & ")&"" tickets)"""
        Call StyleBold(wsWR.Cells(r, 1))
        wsWR.Cells(r, 4).Formula = "=ROUND(SUM(D" & firstDataRow & ":D" & lastDataRowForDay & "),2)"
        Call StyleBold(wsWR.Cells(r, 4))
        r = r + 2
    Next di

    wsWR.Columns("A").ColumnWidth = 14
    wsWR.Columns("B").ColumnWidth = 50
    wsWR.Columns("C").ColumnWidth = 16
    wsWR.Columns("D").ColumnWidth = 13
    wsWR.Columns("E").ColumnWidth = 10
    wsWR.Columns("F").ColumnWidth = 22

    ' ===== Tickets Created This Week sheet ===================================
    Dim wsTC As Worksheet
    Set wsTC = thisWb.Sheets("Tickets Created This Week")
    wsTC.Cells(1, 1).Value = "New Tickets Created This Week"
    Call StyleTitle(wsTC.Cells(1, 1), 12)
    wsTC.Cells(2, 1).Value = "'Created' = first work-log date found in this export, not the Jira issue creation timestamp"
    Call StyleSubtitle(wsTC.Cells(2, 1))
    Dim tcHdrs As Variant
    tcHdrs = Array("Ticket Key", "Summary", "Status", "Priority", "Date Created")
    For i = 0 To 4
        wsTC.Cells(3, i + 1).Value = tcHdrs(i)
        Call StyleTableHeader(wsTC.Cells(3, i + 1))
    Next i

    ' Sort tickets by first-seen date
    Dim ticketKeys() As String
    Dim nt As Long
    nt = firstSeenDict.Count
    ReDim ticketKeys(1 To nt)
    idx = 0
    For Each k In firstSeenDict.Keys
        idx = idx + 1
        ticketKeys(idx) = CStr(k)
    Next k
    Call SortTicketsByDate(ticketKeys, firstSeenDict)

    r = 4
    For i = 1 To nt
        tk = ticketKeys(i)
        wsTC.Cells(r, 1).Value = tk
        wsTC.Cells(r, 2).Value = ticketSummary(tk)
        wsTC.Cells(r, 3).Value = StrConv(LCase(ticketStatus(tk)), vbProperCase)
        wsTC.Cells(r, 4).Value = ticketPriority(tk)
        wsTC.Cells(r, 5).Value = firstSeenDict(tk)
        wsTC.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
        r = r + 1
    Next i
    wsTC.Columns("A").ColumnWidth = 14
    wsTC.Columns("B").ColumnWidth = 55
    wsTC.Columns("C").ColumnWidth = 22
    wsTC.Columns("D").ColumnWidth = 12
    wsTC.Columns("E").ColumnWidth = 14

    ' ===== Author Summary (Anonymized) sheet =================================
    Dim wsAS As Worksheet
    Set wsAS = thisWb.Sheets("Author Summary (Anonymized)")
    wsAS.Cells(1, 1).Value = "Author Summary (Anonymized)"
    Call StyleTitle(wsAS.Cells(1, 1), 12)
    wsAS.Cells(2, 1).Value = "Names withheld; each team member is identified only by an anonymized ID"
    Call StyleSubtitle(wsAS.Cells(2, 1))

    Dim hdrRow As Long
    hdrRow = 4
    wsAS.Cells(hdrRow, 1).Value = "Author ID"
    Call StyleTableHeader(wsAS.Cells(hdrRow, 1))
    For di = 1 To nd
        wsAS.Cells(hdrRow, di + 1).Value = Format(distinctDates(di), "yyyy-mm-dd")
        Call StyleTableHeader(wsAS.Cells(hdrRow, di + 1))
    Next di
    Dim totalCol As Long
    totalCol = nd + 2
    wsAS.Cells(hdrRow, totalCol).Value = "Total Hours Logged"
    Call StyleTableHeader(wsAS.Cells(hdrRow, totalCol))

    Dim firstAuthorRow As Long
    firstAuthorRow = hdrRow + 1
    For i = 1 To numAuthors
        r = firstAuthorRow + i - 1
        wsAS.Cells(r, 1).Value = allAuthorIDs(i)
        For di = 1 To nd
            dExcel = "DATE(" & Year(distinctDates(di)) & "," & Month(distinctDates(di)) & "," & Day(distinctDates(di)) & ")"
            wsAS.Cells(r, di + 1).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$A:$A," & dExcel & _
                ",Data!$E:$E,""" & allAuthorIDs(i) & """),2)"
        Next di
        Dim colLetterFirst As String, colLetterLast As String
        colLetterFirst = Split(Cells(1, 2).Address, "$")(1)
        colLetterLast = Split(Cells(1, totalCol - 1).Address, "$")(1)
        wsAS.Cells(r, totalCol).Formula = "=ROUND(SUM(" & colLetterFirst & r & ":" & colLetterLast & r & "),2)"
    Next i
    Dim totalRow As Long
    totalRow = firstAuthorRow + numAuthors
    wsAS.Cells(totalRow, 1).Value = "Team Total"
    Call StyleBold(wsAS.Cells(totalRow, 1))
    For i = 2 To totalCol
        Dim cl As String
        cl = Split(Cells(1, i).Address, "$")(1)
        wsAS.Cells(totalRow, i).Formula = "=ROUND(SUM(" & cl & firstAuthorRow & ":" & cl & (totalRow - 1) & "),2)"
        Call StyleBold(wsAS.Cells(totalRow, i))
    Next i
    wsAS.Columns("A").ColumnWidth = 12
    For i = 2 To totalCol
        wsAS.Columns(i).ColumnWidth = 14
    Next i

    ' ===== Dashboard sheet ====================================================
    Dim wsDB As Worksheet
    Set wsDB = thisWb.Sheets("Dashboard")
    wsDB.Cells(1, 1).Value = "Weekly Report - Visual Summary"
    Call StyleTitle(wsDB.Cells(1, 1))
    wsDB.Cells(3, 1).Value = "Hours Logged Per Day": Call StyleBold(wsDB.Cells(3, 1))
    wsDB.Cells(3, 4).Value = "Hours Logged Per Author": Call StyleBold(wsDB.Cells(3, 4))
    wsDB.Cells(3, 7).Value = "Tickets By Priority": Call StyleBold(wsDB.Cells(3, 7))
    wsDB.Cells(4, 1).Value = "Date": Call StyleTableHeader(wsDB.Cells(4, 1))
    wsDB.Cells(4, 2).Value = "Hours": Call StyleTableHeader(wsDB.Cells(4, 2))
    wsDB.Cells(4, 4).Value = "Author": Call StyleTableHeader(wsDB.Cells(4, 4))
    wsDB.Cells(4, 5).Value = "Hours": Call StyleTableHeader(wsDB.Cells(4, 5))
    wsDB.Cells(4, 7).Value = "Priority": Call StyleTableHeader(wsDB.Cells(4, 7))
    wsDB.Cells(4, 8).Value = "Count": Call StyleTableHeader(wsDB.Cells(4, 8))

    For di = 1 To nd
        dExcel = "DATE(" & Year(distinctDates(di)) & "," & Month(distinctDates(di)) & "," & Day(distinctDates(di)) & ")"
        wsDB.Cells(4 + di, 1).Value = Format(distinctDates(di), "yyyy-mm-dd")
        wsDB.Cells(4 + di, 2).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$A:$A," & dExcel & "),2)"
    Next di

    For i = 1 To numAuthors
        wsDB.Cells(4 + i, 4).Value = allAuthorIDs(i)
        wsDB.Cells(4 + i, 5).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$E:$E,""" & allAuthorIDs(i) & """),2)"
    Next i

    Dim priorityDict As Object
    Set priorityDict = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Not priorityDict.Exists(pArr(i)) Then priorityDict.Add pArr(i), 1
    Next i
    Dim pKeys() As String
    Dim npKeys As Long
    npKeys = priorityDict.Count
    ReDim pKeys(1 To npKeys)
    idx = 0
    For Each k In priorityDict.Keys
        idx = idx + 1
        pKeys(idx) = CStr(k)
    Next k
    Call SortStrings(pKeys)
    For i = 1 To npKeys
        wsDB.Cells(4 + i, 7).Value = pKeys(i)
        wsDB.Cells(4 + i, 8).Formula = "=SUMPRODUCT((Data!$G$2:$G$" & lastDataRow & "=""" & pKeys(i) & _
            """)*Data!$J$2:$J$" & lastDataRow & ")"
    Next i

    wsDB.Columns("A").ColumnWidth = 12: wsDB.Columns("B").ColumnWidth = 8
    wsDB.Columns("D").ColumnWidth = 10: wsDB.Columns("E").ColumnWidth = 8
    wsDB.Columns("G").ColumnWidth = 10: wsDB.Columns("H").ColumnWidth = 8

    ' Charts
    Dim cht1 As ChartObject, cht2 As ChartObject
    Set cht1 = wsDB.ChartObjects.Add(Left:=wsDB.Range("J3").Left, Top:=wsDB.Range("J3").Top, Width:=350, Height:=220)
    cht1.Chart.SetSourceData Source:=wsDB.Range(wsDB.Cells(4, 1), wsDB.Cells(4 + nd, 2))
    cht1.Chart.ChartType = xlColumnClustered
    cht1.Chart.HasTitle = True
    cht1.Chart.ChartTitle.Text = "Hours Logged Per Day"

    Set cht2 = wsDB.ChartObjects.Add(Left:=wsDB.Range("J20").Left, Top:=wsDB.Range("J20").Top, Width:=350, Height:=220)
    cht2.Chart.SetSourceData Source:=wsDB.Range(wsDB.Cells(4, 4), wsDB.Cells(4 + numAuthors, 5))
    cht2.Chart.ChartType = xlColumnClustered
    cht2.Chart.HasTitle = True
    cht2.Chart.ChartTitle.Text = "Hours Logged Per Author"

    ' ===== Internal Ref - DO NOT SHARE sheet =================================
    Dim wsIR As Worksheet
    Set wsIR = thisWb.Sheets("Internal Ref - DO NOT SHARE")
    wsIR.Cells(1, 1).Value = "INTERNAL USE ONLY - remove this sheet before presenting to the team"
    Call StyleBold(wsIR.Cells(1, 1))
    wsIR.Cells(2, 1).Value = "Author ID": Call StyleTableHeader(wsIR.Cells(2, 1))
    wsIR.Cells(2, 2).Value = "Real Name": Call StyleTableHeader(wsIR.Cells(2, 2))
    For i = 1 To numAuthors
        wsIR.Cells(2 + i, 1).Value = allAuthorIDs(i)
        For Each k In authorDict.Keys
            If authorDict(k) = allAuthorIDs(i) Then
                wsIR.Cells(2 + i, 2).Value = k
                Exit For
            End If
        Next k
    Next i
    wsIR.Columns("A").ColumnWidth = 12
    wsIR.Columns("B").ColumnWidth = 30

    thisWb.Sheets("Weekly Report").Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Report rebuilt. Review it, then delete the ""Internal Ref - DO NOT SHARE"" sheet before sending it out.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "Something went wrong: " & Err.Description, vbCritical
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close False
    Resume CleanExit

End Sub

' ============================================================================
' Helper functions
' ============================================================================

Function FindCol(ws As Worksheet, lastCol As Long, aliases As Variant) As Long
    Dim j As Long, h As String, a As Variant
    For j = 1 To lastCol
        h = NormalizeHeader(ws.Cells(1, j).Value)
        For Each a In aliases
            If h = NormalizeHeader(CStr(a)) Then
                FindCol = j
                Exit Function
            End If
        Next a
    Next j
    FindCol = 0
End Function

Function NormalizeHeader(s As String) As String
    Dim t As String
    t = LCase(Trim(s))
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    NormalizeHeader = t
End Function

Function GetHeaderList(ws As Worksheet, lastCol As Long) As String
    Dim j As Long, s As String
    For j = 1 To lastCol
        s = s & ws.Cells(1, j).Value & "; "
    Next j
    GetHeaderList = s
End Function

Sub ClearOrCreateSheet(wb As Workbook, name As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = name
    Else
        ws.Cells.Clear
        Dim cObj As ChartObject
        For Each cObj In ws.ChartObjects
            cObj.Delete
        Next cObj
    End If
End Sub

Sub StyleTitle(rng As Range, Optional sz As Long = 14)
    With rng.Font
        .Name = "Arial": .Size = sz: .Bold = True
    End With
End Sub

Sub StyleSubtitle(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 9: .Italic = True
        .Color = RGB(89, 89, 89)
    End With
End Sub

Sub StyleDayHeader(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True
        .Color = RGB(31, 78, 120)
    End With
End Sub

Sub StyleTableHeader(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True
        .Color = RGB(255, 255, 255)
    End With
    rng.Interior.Color = RGB(31, 78, 120)
End Sub

Sub StyleBold(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True
    End With
End Sub

Sub SortUserIDs(arr() As String)
    Dim i As Long, j As Long, tmp As String
    Dim ni As Long, nj As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            ni = Val(Mid(arr(i), InStrRev(arr(i), " ") + 1))
            nj = Val(Mid(arr(j), InStrRev(arr(j), " ") + 1))
            If nj < ni Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Sub SortDates(arr() As Date)
    Dim i As Long, j As Long, tmp As Date
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Sub SortStrings(arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Sub SortTicketsByDate(arr() As String, dateMap As Object)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If dateMap(arr(j)) < dateMap(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub
