Option Explicit

' ============================================================================
' GenerateWeeklyReport (v3 - corrected Sheet3 creation-date counts,
' correct ticket-creation-date logic)
'
' HOW TO RUN: Alt+F8 -> GenerateWeeklyReport -> Run.
' Picks a raw export (.csv/.xlsx/.xls), rebuilds every sheet in THIS workbook.
'
' If the raw file is a multi-sheet .xlsx that also contains an issue-level
' Jira export (a sheet with "Issue key"/"Key" + "Created" columns), that
' sheet is used as the authoritative source of each ticket's real creation
' date. Tickets NOT found in that sheet are treated as pre-existing (not
' "new" this week) rather than guessed at.
'
' Author names are shown in the report. AuthorRoster is used internally to keep
' the team roster stable week to week even if someone logs zero hours.
' ============================================================================

Dim gDayBlue As Long, gHeaderBlue As Long, gGold As Long, gLightBlue As Long, gGreen As Long, gGray As Long

Sub GenerateWeeklyReport()

    ' Colors (defined here so RGB() is evaluated correctly at runtime)
    gDayBlue = RGB(&H44, &H72, &HC4)      ' 4472C4
    gHeaderBlue = RGB(&H2E, &H75, &HB6)   ' 2E75B6
    gGold = RGB(&HFF, &HE6, &H99)         ' FFE699
    gLightBlue = RGB(&HD9, &HE6, &HF5)    ' D9E6F5
    gGreen = RGB(&HC6, &HEF, &HCE)        ' C6EFCE
    gGray = RGB(&H59, &H59, &H59)         ' 595959
    Dim navy As Long
    navy = RGB(&H1F, &H4E, &H78)          ' 1F4E78

    Dim csvPath As Variant
    Dim csvWb As Workbook
    Dim csvWs As Worksheet
    Dim thisWb As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, r As Long

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

    ' ---- Find the worklog sheet (the one with all 7 required columns) -----
    Dim wlSheet As Worksheet
    Dim sh As Worksheet
    For Each sh In csvWb.Worksheets
        Dim testLastCol As Long
        testLastCol = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
        If FindCol(sh, testLastCol, Array("date", "worklogdate", "workdate", "started")) > 0 And _
           FindCol(sh, testLastCol, Array("ticketkey", "issuekey", "ticket", "key", "workitemkey")) > 0 And _
           FindCol(sh, testLastCol, Array("hours", "loggedhours", "hourslogged", "timespent")) > 0 Then
            Set wlSheet = sh
            Exit For
        End If
    Next sh
    If wlSheet Is Nothing Then Set wlSheet = csvWb.Worksheets(1)
    Set csvWs = wlSheet

    lastRow = csvWs.Cells(csvWs.Rows.Count, 1).End(xlUp).Row
    lastCol = csvWs.Cells(1, csvWs.Columns.Count).End(xlToLeft).Column

    colDate = FindCol(csvWs, lastCol, Array("date", "worklogdate", "workdate", "started"))
    colTicket = FindCol(csvWs, lastCol, Array("ticketkey", "issuekey", "ticket", "key", "workitemkey"))
    colSummary = FindCol(csvWs, lastCol, Array("summary", "issuesummary", "title", "workitemsummary"))
    colHours = FindCol(csvWs, lastCol, Array("hours", "timespent(hours)", "timespent", "hourslogged", "loggedhours"))
    colAuthor = FindCol(csvWs, lastCol, Array("author", "authorid", "worklogauthor", "assignee", "user", "name", "fullname"))
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

    ' ---- Look for an issue-level sheet with real per-ticket "Created" -----
    Dim hasIssueSheet As Boolean
    Dim issueCreated As Object, issueSummary As Object, issueStatus As Object, issuePriority As Object
    Set issueCreated = CreateObject("Scripting.Dictionary")
    Set issueSummary = CreateObject("Scripting.Dictionary")
    Set issueStatus = CreateObject("Scripting.Dictionary")
    Set issuePriority = CreateObject("Scripting.Dictionary")
    hasIssueSheet = False
    For Each sh In csvWb.Worksheets
        Dim iLastCol As Long
        iLastCol = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
        Dim colKey As Long, colCreated As Long
        Dim colIssueSummary As Long, colIssueStatus As Long, colIssuePriority As Long
        colKey = FindCol(sh, iLastCol, Array("issuekey", "key"))
        colCreated = FindCol(sh, iLastCol, Array("created"))
        colIssueSummary = FindCol(sh, iLastCol, Array("summary", "issuesummary", "title"))
        colIssueStatus = FindCol(sh, iLastCol, Array("status", "issuestatus"))
        colIssuePriority = FindCol(sh, iLastCol, Array("priority"))
        If colKey > 0 And colCreated > 0 And Not (sh.Name = wlSheet.Name) Then
            Dim iLastRow As Long
            iLastRow = sh.Cells(sh.Rows.Count, colKey).End(xlUp).Row
            Dim ii As Long
            For ii = 2 To iLastRow
                Dim ik As String
                ik = Trim(CStr(sh.Cells(ii, colKey).Value))
                If ik <> "" And IsDate(sh.Cells(ii, colCreated).Value) Then
                    Dim createdDt As Date
                    createdDt = DateSerial(Year(CDate(sh.Cells(ii, colCreated).Value)), _
                                            Month(CDate(sh.Cells(ii, colCreated).Value)), _
                                            Day(CDate(sh.Cells(ii, colCreated).Value)))
                    If issueCreated.Exists(ik) Then
                        If createdDt < issueCreated(ik) Then issueCreated(ik) = createdDt
                    Else
                        issueCreated.Add ik, createdDt
                    End If

                    If colIssueSummary > 0 Then issueSummary(ik) = Trim(CStr(sh.Cells(ii, colIssueSummary).Value))
                    If colIssueStatus > 0 Then issueStatus(ik) = UCase(Trim(CStr(sh.Cells(ii, colIssueStatus).Value)))
                    If colIssuePriority > 0 Then issuePriority(ik) = StrConv(Trim(CStr(sh.Cells(ii, colIssuePriority).Value)), vbProperCase)
                End If
            Next ii
            hasIssueSheet = True
            Exit For
        End If
    Next sh

    ' ---- Read worklog rows into arrays, dropping weekends ------------------
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

    ' ---- Author roster (persisted in hidden AuthorRoster sheet) -----------
    Dim rosterWs As Worksheet
    On Error Resume Next
    Set rosterWs = thisWb.Sheets("AuthorRoster")
    On Error GoTo CleanFail
    If rosterWs Is Nothing Then
        Set rosterWs = thisWb.Sheets.Add(After:=thisWb.Sheets(thisWb.Sheets.Count))
        rosterWs.Name = "AuthorRoster"
        rosterWs.Cells(1, 1).Value = "RealName"
        rosterWs.Visible = xlSheetVeryHidden
    End If
    Dim rosterLast As Long
    rosterLast = rosterWs.Cells(rosterWs.Rows.Count, 1).End(xlUp).Row
    If rosterLast < 1 Then rosterLast = 1

    Dim rosterSet As Object
    Set rosterSet = CreateObject("Scripting.Dictionary")
    For i = 2 To rosterLast
        If rosterWs.Cells(i, 1).Value <> "" Then rosterSet(CStr(rosterWs.Cells(i, 1).Value)) = 1
    Next i
    Dim uniqueAuthorsThisWeek As Object
    Set uniqueAuthorsThisWeek = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Not uniqueAuthorsThisWeek.Exists(aArr(i)) Then uniqueAuthorsThisWeek.Add aArr(i), 1
    Next i
    Dim k As Variant
    For Each k In uniqueAuthorsThisWeek.Keys
        If Not rosterSet.Exists(CStr(k)) Then
            rosterSet.Add CStr(k), 1
            rosterLast = rosterLast + 1
            rosterWs.Cells(rosterLast, 1).Value = k
        End If
    Next k

    Dim numAuthors As Long
    numAuthors = rosterSet.Count
    Dim roster() As String
    ReDim roster(1 To numAuthors)
    Dim idx0 As Long
    idx0 = 0
    For Each k In rosterSet.Keys
        idx0 = idx0 + 1
        roster(idx0) = CStr(k)
    Next k
    Call SortStrings(roster)

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
    Dim weekStart As Date, weekEnd As Date
    weekStart = distinctDates(1)
    weekEnd = distinctDates(nd)

    ' ---- Per-ticket info + real creation date -------------------------------
    Dim ticketSummary As Object, ticketStatus As Object, ticketPriority As Object
    Dim ticketWorkFirstSeen As Object, ticketWorkLastSeen As Object
    Dim tk As String
    Set ticketSummary = CreateObject("Scripting.Dictionary")
    Set ticketStatus = CreateObject("Scripting.Dictionary")
    Set ticketPriority = CreateObject("Scripting.Dictionary")
    Set ticketWorkFirstSeen = CreateObject("Scripting.Dictionary")
    Set ticketWorkLastSeen = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        If Not ticketSummary.Exists(tArr(i)) Then
            ticketSummary.Add tArr(i), sArr(i)
            ticketStatus.Add tArr(i), stArr(i)
            ticketPriority.Add tArr(i), pArr(i)
            ticketWorkFirstSeen.Add tArr(i), dArr(i)
            ticketWorkLastSeen.Add tArr(i), dArr(i)
        Else
            If dArr(i) < ticketWorkFirstSeen(tArr(i)) Then ticketWorkFirstSeen(tArr(i)) = dArr(i)
            If dArr(i) >= ticketWorkLastSeen(tArr(i)) Then
                ticketWorkLastSeen(tArr(i)) = dArr(i)
                ticketSummary(tArr(i)) = sArr(i)
                ticketStatus(tArr(i)) = stArr(i)
                ticketPriority(tArr(i)) = pArr(i)
            End If
        End If
    Next i

    ' Add issue-level metadata, including newly created tickets with no worklogs.
    If hasIssueSheet Then
        For Each k In issueCreated.Keys
            tk = CStr(k)
            If issueSummary.Exists(tk) Then
                If ticketSummary.Exists(tk) Then
                    ticketSummary(tk) = issueSummary(tk)
                Else
                    ticketSummary.Add tk, issueSummary(tk)
                End If
            ElseIf Not ticketSummary.Exists(tk) Then
                ticketSummary.Add tk, ""
            End If

            If issueStatus.Exists(tk) Then
                If ticketStatus.Exists(tk) Then
                    ticketStatus(tk) = issueStatus(tk)
                Else
                    ticketStatus.Add tk, issueStatus(tk)
                End If
            ElseIf Not ticketStatus.Exists(tk) Then
                ticketStatus.Add tk, ""
            End If

            If issuePriority.Exists(tk) Then
                If ticketPriority.Exists(tk) Then
                    ticketPriority(tk) = issuePriority(tk)
                Else
                    ticketPriority.Add tk, issuePriority(tk)
                End If
            ElseIf Not ticketPriority.Exists(tk) Then
                ticketPriority.Add tk, ""
            End If
        Next k
    End If

    ' Creation dates: Sheet3 / issue-level export is authoritative when present.
    Dim ticketCreated As Object
    Set ticketCreated = CreateObject("Scripting.Dictionary")
    If hasIssueSheet Then
        For Each k In issueCreated.Keys
            ticketCreated.Add CStr(k), issueCreated(CStr(k))
        Next k
    Else
        For Each k In ticketWorkFirstSeen.Keys
            ticketCreated.Add CStr(k), ticketWorkFirstSeen(CStr(k))
        Next k
    End If

    Dim totalNewTickets As Long
    totalNewTickets = 0
    For Each k In ticketCreated.Keys
        If ticketCreated(CStr(k)) >= weekStart And ticketCreated(CStr(k)) <= weekEnd Then
            totalNewTickets = totalNewTickets + 1
        End If
    Next k

    ' ---- Rebuild all sheets -------------------------------------------------
    Call ClearOrCreateSheet(thisWb, "Data")
    Call ClearOrCreateSheet(thisWb, "Weekly Report")
    Call ClearOrCreateSheet(thisWb, "Tickets Created This Week")
    Call ClearOrCreateSheet(thisWb, "Author Summary")
    Call ClearOrCreateSheet(thisWb, "Dashboard")
    On Error Resume Next
    thisWb.Sheets("Internal Ref - DO NOT SHARE").Delete
    On Error GoTo CleanFail

    ' ===== Data sheet =========================================================
    Dim wsData As Worksheet
    Set wsData = thisWb.Sheets("Data")
    Dim dataHeaders As Variant
    dataHeaders = Array("Date", "Ticket Key", "Summary", "Hours", "Author", _
                         "Status", "Priority", "FirstAuthorTicketDay", "FirstAuthorDay", _
                         "FirstTicketWeek", "Ticket Created Date")
    For i = 0 To 10
        wsData.Cells(1, i + 1).Value = dataHeaders(i)
    Next i
    For i = 1 To n
        r = i + 1
        wsData.Cells(r, 1).Value = dArr(i)
        wsData.Cells(r, 1).NumberFormat = "yyyy-mm-dd"
        wsData.Cells(r, 2).Value = tArr(i)
        wsData.Cells(r, 3).Value = sArr(i)
        wsData.Cells(r, 4).Value = hArr(i)
        wsData.Cells(r, 5).Value = aArr(i)
        wsData.Cells(r, 6).Value = stArr(i)
        wsData.Cells(r, 7).Value = pArr(i)
        wsData.Cells(r, 8).Formula = "=IF(COUNTIFS($A$2:A" & r & ",A" & r & ",$B$2:B" & r & ",B" & r & ",$E$2:E" & r & ",E" & r & ")=1,1,0)"
        wsData.Cells(r, 9).Formula = "=IF(COUNTIFS($A$2:A" & r & ",A" & r & ",$E$2:E" & r & ",E" & r & ")=1,1,0)"
        wsData.Cells(r, 10).Formula = "=IF(COUNTIFS($B$2:B" & r & ",B" & r & ")=1,1,0)"
        If ticketCreated.Exists(tArr(i)) Then
            wsData.Cells(r, 11).Value = ticketCreated(tArr(i))
            wsData.Cells(r, 11).NumberFormat = "yyyy-mm-dd"
        End If
    Next i
    Dim lastDataRow As Long
    lastDataRow = n + 1
    wsData.Columns("A:K").AutoFit
    wsData.Columns("E").Hidden = False  ' show real author names

    ' ===== Weekly Report sheet ===============================================
    Dim wsWR As Worksheet
    Set wsWR = thisWb.Sheets("Weekly Report")
    wsWR.Columns("A").ColumnWidth = 14
    wsWR.Columns("B").ColumnWidth = 52
    wsWR.Columns("C").ColumnWidth = 16
    wsWR.Columns("D").ColumnWidth = 13
    wsWR.Columns("E").ColumnWidth = 11
    wsWR.Columns("F").ColumnWidth = 24

    wsWR.Range("A1:F1").Merge
    wsWR.Cells(1, 1).Value = "Weekly Team Activity Report"
    Call StyleTitle(wsWR.Cells(1, 1), navy)
    wsWR.Rows(1).RowHeight = 30

    wsWR.Range("A2:F2").Merge
    wsWR.Cells(2, 1).Value = "Week of " & Format(weekStart, "yyyy-mm-dd") & " to " & _
        Format(weekEnd, "yyyy-mm-dd") & "  |  Weekends excluded"
    Call StyleSubtitle(wsWR.Cells(2, 1))
    wsWR.Rows(2).RowHeight = 18

    r = 4
    Dim di As Long
    For di = 1 To nd
        Dim dToday As Date
        dToday = distinctDates(di)

        Dim dayTickets As Object, dayAuthors As Object
        Set dayTickets = CreateObject("Scripting.Dictionary")
        Set dayAuthors = CreateObject("Scripting.Dictionary")
        For i = 1 To n
            If dArr(i) = dToday Then
                If Not dayTickets.Exists(tArr(i)) Then dayTickets.Add tArr(i), 1
                If Not dayAuthors.Exists(aArr(i)) Then dayAuthors.Add aArr(i), 1
            End If
        Next i

        Dim newTicketsToday As Long
        newTicketsToday = 0
        For Each k In ticketCreated.Keys
            If ticketCreated(CStr(k)) = dToday Then newTicketsToday = newTicketsToday + 1
        Next k

        wsWR.Range("A" & r & ":F" & r).Merge
        wsWR.Cells(r, 1).Value = Format(dToday, "dddd") & ", " & Format(dToday, "yyyy-mm-dd") & _
            "    |    Team Availability: " & dayAuthors.Count & "/" & numAuthors & _
            "    |    New Tickets Created: " & newTicketsToday
        Call StyleDayHeader(wsWR.Cells(r, 1))
        wsWR.Rows(r).RowHeight = 19.5
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
            tk = CStr(k)
            Dim dExcel As String
            dExcel = "DATE(" & Year(dToday) & "," & Month(dToday) & "," & Day(dToday) & ")"
            Dim isNew As Boolean
            isNew = False
            If ticketCreated.Exists(tk) Then
                If ticketCreated(tk) = dToday Then isNew = True
            End If
            wsWR.Cells(r, 1).Value = tk
            wsWR.Cells(r, 2).Value = ticketSummary(tk)
            wsWR.Cells(r, 3).Formula = "=SUMPRODUCT((Data!$A$2:$A$" & lastDataRow & "=" & dExcel & ")*" & _
                "(Data!$B$2:$B$" & lastDataRow & "=""" & tk & """)*Data!$H$2:$H$" & lastDataRow & ")"
            wsWR.Cells(r, 4).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$A:$A," & dExcel & _
                ",Data!$B:$B,""" & tk & """),2)"
            wsWR.Cells(r, 5).Value = ticketPriority(tk)
            wsWR.Cells(r, 6).Value = ticketStatus(tk)
            If isNew Then
                For j = 1 To 6
                    wsWR.Cells(r, j).Interior.Color = gGreen
                Next j
            End If
            r = r + 1
        Next k
        Dim lastDataRowForDay As Long
        lastDataRowForDay = r - 1

        wsWR.Range("A" & r & ":B" & r).Merge
        wsWR.Cells(r, 1).Formula = "=""Day Total (""&COUNTA(A" & firstDataRow & ":A" & lastDataRowForDay & ")&"" tickets)"""
        wsWR.Cells(r, 4).Formula = "=ROUND(SUM(D" & firstDataRow & ":D" & lastDataRowForDay & "),2)"
        For j = 1 To 6
            Call StyleDayTotal(wsWR.Cells(r, j), navy)
        Next j
        r = r + 2
    Next di

    ' ---- Week Summary block (at the end) -----------------------------------
    wsWR.Range("A" & r & ":F" & r).Merge
    wsWR.Cells(r, 1).Value = "Week Summary"
    Call StyleSummaryTitle(wsWR.Cells(r, 1), navy)
    wsWR.Rows(r).RowHeight = 19.5
    r = r + 1

    wsWR.Range("A" & r & ":D" & r).Merge
    wsWR.Cells(r, 1).Value = "Total Distinct Tickets Worked On This Week"
    Call StyleSummaryLabel(wsWR.Cells(r, 1))
    wsWR.Range("E" & r & ":F" & r).Merge
    wsWR.Cells(r, 5).Formula = "=SUM(Data!J2:J" & lastDataRow & ")"
    Call StyleSummaryValue(wsWR.Cells(r, 5), navy)
    r = r + 1

    wsWR.Range("A" & r & ":D" & r).Merge
    wsWR.Cells(r, 1).Value = "Total New Tickets Created This Week"
    Call StyleSummaryLabel(wsWR.Cells(r, 1))
    wsWR.Range("E" & r & ":F" & r).Merge
    wsWR.Cells(r, 5).Value = totalNewTickets
    Call StyleSummaryValue(wsWR.Cells(r, 5), navy)
    r = r + 1

    wsWR.Range("A" & r & ":D" & r).Merge
    wsWR.Cells(r, 1).Value = "Completed Tickets (Done + Waiting For Approval)"
    Call StyleSummaryLabel(wsWR.Cells(r, 1))
    wsWR.Range("E" & r & ":F" & r).Merge
    Dim completedTickets As Long
    completedTickets = 0
    For Each k In ticketWorkFirstSeen.Keys
        tk = CStr(k)
        If ticketStatus.Exists(tk) Then
            If UCase(Trim(ticketStatus(tk))) = "DONE" Or _
               InStr(1, UCase(Trim(ticketStatus(tk))), "WAITING FOR APPROVAL", vbTextCompare) > 0 Then
                completedTickets = completedTickets + 1
            End If
        End If
    Next k
    wsWR.Cells(r, 5).Value = completedTickets
    Call StyleSummaryValue(wsWR.Cells(r, 5), navy)
    r = r + 1

    wsWR.Range("A" & r & ":D" & r).Merge
    wsWR.Cells(r, 1).Value = "Total Hours Logged This Week"
    Call StyleSummaryLabel(wsWR.Cells(r, 1))
    wsWR.Range("E" & r & ":F" & r).Merge
    wsWR.Cells(r, 5).Formula = "=ROUND(SUM(Data!D:D),2)"
    Call StyleSummaryValue(wsWR.Cells(r, 5), navy)
    r = r + 1

    wsWR.Range("A" & r & ":D" & r).Merge
    wsWR.Cells(r, 1).Value = "Team Size (Total Available Authors)"
    Call StyleSummaryLabel(wsWR.Cells(r, 1))
    wsWR.Range("E" & r & ":F" & r).Merge
    wsWR.Cells(r, 5).Value = numAuthors
    Call StyleSummaryValue(wsWR.Cells(r, 5), navy)
    r = r + 2

    wsWR.Cells(r, 1).Value = "Legend:"
    Call StyleBold(wsWR.Cells(r, 1))
    r = r + 1
    wsWR.Cells(r, 1).Value = "  "
    wsWR.Cells(r, 1).Interior.Color = gGreen
    wsWR.Cells(r, 2).Value = "Ticket newly created and worked on the same day"

    ' ===== Tickets Created This Week sheet ===================================
    Dim wsTC As Worksheet
    Set wsTC = thisWb.Sheets("Tickets Created This Week")
    wsTC.Range("A1:E1").Merge
    wsTC.Cells(1, 1).Value = "New Tickets Created This Week"
    Call StyleTitle(wsTC.Cells(1, 1), navy)
    wsTC.Rows(1).RowHeight = 22

    Dim tcHdrs As Variant
    tcHdrs = Array("Ticket Key", "Summary", "Status", "Priority", "Date Created")
    For i = 0 To 4
        wsTC.Cells(3, i + 1).Value = tcHdrs(i)
        Call StyleTableHeader(wsTC.Cells(3, i + 1))
    Next i

    Dim newTicketKeys() As String
    Dim ntCount As Long
    ntCount = 0
    ReDim newTicketKeys(1 To ticketCreated.Count)
    For Each k In ticketCreated.Keys
        If ticketCreated(CStr(k)) >= weekStart And ticketCreated(CStr(k)) <= weekEnd Then
            ntCount = ntCount + 1
            newTicketKeys(ntCount) = CStr(k)
        End If
    Next k
    If ntCount > 0 Then
        ReDim Preserve newTicketKeys(1 To ntCount)
        Call SortTicketsByDate(newTicketKeys, ticketCreated)
    End If

    r = 4
    For i = 1 To ntCount
        tk = newTicketKeys(i)
        wsTC.Cells(r, 1).Value = tk
        wsTC.Cells(r, 2).Value = ticketSummary(tk)
        wsTC.Cells(r, 3).Value = StrConv(LCase(ticketStatus(tk)), vbProperCase)
        wsTC.Cells(r, 4).Value = ticketPriority(tk)
        wsTC.Cells(r, 5).Value = ticketCreated(tk)
        wsTC.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
        r = r + 1
    Next i
    r = r + 1
    wsTC.Cells(r, 1).Value = "Total Tickets Created"
    Call StyleBold(wsTC.Cells(r, 1))
    wsTC.Cells(r, 5).Value = ntCount
    Call StyleBold(wsTC.Cells(r, 5))

    wsTC.Columns("A").ColumnWidth = 14
    wsTC.Columns("B").ColumnWidth = 55
    wsTC.Columns("C").ColumnWidth = 22
    wsTC.Columns("D").ColumnWidth = 12
    wsTC.Columns("E").ColumnWidth = 14

    ' ===== Author Summary sheet ===============================================
    Dim wsAS As Worksheet
    Set wsAS = thisWb.Sheets("Author Summary")
    Dim totalCol As Long
    totalCol = nd + 2
    wsAS.Range(wsAS.Cells(1, 1), wsAS.Cells(1, totalCol)).Merge
    wsAS.Cells(1, 1).Value = "Author Summary"
    Call StyleTitle(wsAS.Cells(1, 1), navy)
    wsAS.Rows(1).RowHeight = 22

    Dim hdrRow As Long
    hdrRow = 3
    wsAS.Cells(hdrRow, 1).Value = "Author"
    Call StyleTableHeader(wsAS.Cells(hdrRow, 1))
    For di = 1 To nd
        wsAS.Cells(hdrRow, di + 1).Value = Format(distinctDates(di), "yyyy-mm-dd")
        Call StyleTableHeader(wsAS.Cells(hdrRow, di + 1))
    Next di
    wsAS.Cells(hdrRow, totalCol).Value = "Total Hours Logged"
    Call StyleTableHeader(wsAS.Cells(hdrRow, totalCol))

    Dim firstAuthorRow As Long
    firstAuthorRow = hdrRow + 1
    For i = 1 To numAuthors
        r = firstAuthorRow + i - 1
        wsAS.Cells(r, 1).Value = roster(i)
        For di = 1 To nd
            dExcel = "DATE(" & Year(distinctDates(di)) & "," & Month(distinctDates(di)) & "," & Day(distinctDates(di)) & ")"
            wsAS.Cells(r, di + 1).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$A:$A," & dExcel & _
                ",Data!$E:$E,""" & roster(i) & """),2)"
        Next di
        Dim clFirst As String, clLast As String
        clFirst = Split(Cells(1, 2).Address, "$")(1)
        clLast = Split(Cells(1, totalCol - 1).Address, "$")(1)
        wsAS.Cells(r, totalCol).Formula = "=ROUND(SUM(" & clFirst & r & ":" & clLast & r & "),2)"
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
    wsAS.Columns("A").ColumnWidth = 28
    For i = 2 To totalCol
        wsAS.Columns(i).ColumnWidth = 14
    Next i

    ' ===== Dashboard sheet ====================================================
    Dim wsDB As Worksheet
    Set wsDB = thisWb.Sheets("Dashboard")
    wsDB.Range("A1:H1").Merge
    wsDB.Cells(1, 1).Value = "Weekly Report - Visual Summary"
    Call StyleTitle(wsDB.Cells(1, 1), navy)
    wsDB.Rows(1).RowHeight = 22
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
        wsDB.Cells(4 + i, 4).Value = roster(i)
        wsDB.Cells(4 + i, 5).Formula = "=ROUND(SUMIFS(Data!$D:$D,Data!$E:$E,""" & roster(i) & """),2)"
    Next i

    Dim priorityDict As Object
    Set priorityDict = CreateObject("Scripting.Dictionary")
    For Each k In ticketWorkFirstSeen.Keys
        tk = CStr(k)
        Dim pr As String
        pr = ticketPriority(tk)
        If pr = "" Then pr = "Unspecified"
        If priorityDict.Exists(pr) Then
            priorityDict(pr) = CLng(priorityDict(pr)) + 1
        Else
            priorityDict.Add pr, 1
        End If
    Next k
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
        wsDB.Cells(4 + i, 8).Value = priorityDict(pKeys(i))
    Next i

    wsDB.Columns("A").ColumnWidth = 12: wsDB.Columns("B").ColumnWidth = 8
    wsDB.Columns("D").ColumnWidth = 26: wsDB.Columns("E").ColumnWidth = 8
    wsDB.Columns("G").ColumnWidth = 10: wsDB.Columns("H").ColumnWidth = 8

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

    thisWb.Sheets("Weekly Report").Activate

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Report rebuilt.", vbInformation
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

Sub StyleTitle(rng As Range, navy As Long)
    With rng.Font
        .Name = "Arial": .Size = 16: .Bold = True: .Color = RGB(255, 255, 255)
    End With
    rng.Interior.Color = navy
End Sub

Sub StyleSubtitle(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 10: .Italic = True
        .Color = RGB(89, 89, 89)
    End With
End Sub

Sub StyleDayHeader(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True: .Color = RGB(255, 255, 255)
    End With
    rng.Interior.Color = RGB(&H44, &H72, &HC4)
End Sub

Sub StyleTableHeader(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 10: .Bold = True: .Color = RGB(255, 255, 255)
    End With
    rng.Interior.Color = RGB(&H2E, &H75, &HB6)
End Sub

Sub StyleDayTotal(rng As Range, navy As Long)
    With rng.Font
        .Name = "Arial": .Size = 10: .Bold = True: .Color = navy
    End With
    rng.Interior.Color = RGB(&HFF, &HE6, &H99)
End Sub

Sub StyleSummaryTitle(rng As Range, navy As Long)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True: .Color = RGB(255, 255, 255)
    End With
    rng.Interior.Color = navy
End Sub

Sub StyleSummaryLabel(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 10: .Bold = True
    End With
    rng.Interior.Color = RGB(&HD9, &HE6, &HF5)
End Sub

Sub StyleSummaryValue(rng As Range, navy As Long)
    With rng.Font
        .Name = "Arial": .Size = 10: .Bold = True: .Color = navy
    End With
    rng.Interior.Color = RGB(&HFF, &HE6, &H99)
End Sub

Sub StyleBold(rng As Range)
    With rng.Font
        .Name = "Arial": .Size = 11: .Bold = True
    End With
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
