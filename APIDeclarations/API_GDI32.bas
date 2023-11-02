Option Explicit
Option Private Module

Sub AddNewTask(s As String)
If FreeVersion Then
If Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A")) - 3 >= cFreeVersionTasksCount Then
Call AddNewTaskPlaceholder: sTempStr1 = msg(80) & msg(82)
frmBuyPro.show
bAddMilestone = False: bAddTask = False
Exit Sub
End If
End If
If ActiveSheet.AutoFilterMode = True Then MsgBox "Task cannot be added when the filter mode is on.", vbInformation, "Information": GoTo PreExit
If IsDataCollapsed = True Then Exit Sub
If Selection.Row <= rownine Then MsgBox "You can add a task only under the header row.", vbInformation, "Information": GoTo Last
If Selection.Rows.Count > 1 Then MsgBox "Select the row where you want to add a new task and try adding again", vbInformation, "Information": GoTo Last
cTaskWBSBeingAdded = vbNullString
Dim lrow As Long, tclrow As Long
If s = "AtSelection" Then
clrow = Selection.Row: lIndentLevel = Cells(clrow, cpg.Task).IndentLevel
ElseIf s = "BelowSelection" Then
If IsParentTask(Selection.Row) Then
clrow = Selection.Row + 1: lIndentLevel = Cells(Selection.Row + 1, cpg.Task).IndentLevel
Else
clrow = Selection.Row + 1: lIndentLevel = Cells(Selection.Row, cpg.Task).IndentLevel
End If
End If
lrow = GetLastRow + 1
If lrow < clrow Then clrow = lrow
If IsLicValid Then
If bAddMilestone = False Then bAddTask = True
frmTask.show
End If
PreExit:
bAddTask = False:bAddMilestone = False:clrow = 0:lIndentLevel = 0
Last:
End Sub
Sub EditExistingTask(Optional t As Boolean)
If Selection.Rows.Count > 1 Then MsgBox "Select a single task to edit", vbInformation, "Information": Exit Sub
clrow = Selection.Row
If Cells(clrow, cpg.GEtype) = vbNullString Then MsgBox "Select a task to edit", vbInformation, "Information": Exit Sub
If Cells(clrow, cpg.GEtype) = "T" Then bEditTask = True: frmTask.show: GoTo Last
If Cells(clrow, cpg.GEtype) = "M" Then bEditMilestone = True: frmTask.show: GoTo Last
Last:
clrow = 0:bEditTask = False:bEditMilestone = False
End Sub
Sub LoadFormOnDblClick()
Dim dblTemp As Double
Dim s As Shape, ws As Worksheet
If dbcRow = firsttaskrow And Cells(dbcRow, cpg.Task) = sAddTaskPlaceHolder And Cells(dbcRow, cpg.GEtype) = vbNullString Then
Call AddNewTask("AtSelection")
ElseIf Cells(dbcRow, cpg.GEtype) <> vbNullString And Cells(dbcRow, cpg.Task) <> vbNullString And Cells(dbcRow, cpg.Task) <> sAddTaskPlaceHolder Then
If dbcCol = cpg.TColor Then
dblTemp = PickNewColor(CDbl(Cells(dbcRow, dbcCol).Interior.Color))
If dblTemp <> -4142 Then
Cells(dbcRow, dbcCol).Interior.Color = dblTemp
Set ws = ActiveSheet
On Error Resume Next
Set s = ws.Shapes("S_E_" & Cells(dbcRow, cpg.TID))
If Left(Cells(dbcRow, cpg.WBS), 1) = "M" Then Set s = ws.Shapes("S_M_" & Cells(dbcRow, cpg.TID))
If Not s Is Nothing Then s.Fill.ForeColor.RGB = dblTemp

Set ws = Nothing: Set s = Nothing
On Error GoTo 0
End If
ElseIf dbcCol = cpg.TPColor And Left(Cells(dbcRow, cpg.WBS), 1) <> "M" Then
dblTemp = PickNewColor(CDbl(Cells(dbcRow, dbcCol).Interior.Color))
If dblTemp <> -4142 Then
Cells(dbcRow, dbcCol).Interior.Color = dblTemp
Set ws = ActiveSheet
On Error Resume Next

Set s = ws.Shapes("S_C_" & Cells(dbcRow, cpg.TID))
If Not s Is Nothing Then s.Fill.ForeColor.RGB = dblTemp
Set ws = Nothing: Set s = Nothing
On Error GoTo 0
End If
ElseIf dbcCol = cpg.BLColor And Left(Cells(dbcRow, cpg.WBS), 1) <> "M" Then
dblTemp = PickNewColor(CDbl(Cells(dbcRow, dbcCol).Interior.Color))
If dblTemp <> -4142 Then
Cells(dbcRow, dbcCol).Interior.Color = dblTemp
Set ws = ActiveSheet
On Error Resume Next

Set s = ws.Shapes("S_B_" & Cells(dbcRow, cpg.TID))
If Not s Is Nothing Then s.Fill.ForeColor.RGB = dblTemp
Set ws = Nothing: Set s = Nothing
On Error GoTo 0
End If
ElseIf dbcCol = cpg.ACColor And Left(Cells(dbcRow, cpg.WBS), 1) <> "M" Then
dblTemp = PickNewColor(CDbl(Cells(dbcRow, dbcCol).Interior.Color))
If dblTemp <> -4142 Then
Cells(dbcRow, dbcCol).Interior.Color = dblTemp
Set ws = ActiveSheet
On Error Resume Next

Set s = ws.Shapes("S_A_" & Cells(dbcRow, cpg.TID))
If Not s Is Nothing Then s.Fill.ForeColor.RGB = dblTemp
Set ws = Nothing: Set s = Nothing
On Error GoTo 0
End If
Else
EditExistingTask
End If
ElseIf Cells(dbcRow, cpg.Task) = sAddTaskPlaceHolder And Cells(dbcRow, cpg.WBS) = vbNullString Then
Call AddNewTask("AtSelection")
ElseIf Cells(dbcRow, cpg.Task) = vbNullString And Cells(dbcRow, cpg.WBS) = vbNullString Then
Cells(rowtwo, cpg.Task).End(xlUp).Select
If Selection = sAddTaskPlaceHolder Then
ElseIf Selection = vbNullString Then
Else
Selection.Offset(1, 0).Select
End If
Call AddNewTask("AtSelection")
End If
dbcCol = 0:dbcRow = 0
End Sub
Sub TriggerAddNewSheet(Optional t As Boolean)
If AddNewGC = False And FreeVersion Then MsgBox msg(87) & msg(82): GoTo Last 'temp workaround to allow add new gc in free version
If IsLicValid(0, 1) Then bAddProject = True: frmNewGantt.show: bAddProject = False
Last:
AddNewGC = False
End Sub
Sub DeleteExtrasRowsInFree()
Dim lrow As Long, taskCount As Long, allowedRow As Long
If FreeVersion Then
lrow = GetLastRow: taskCount = Application.WorksheetFunction.CountA(ActiveSheet.Range(Cells(firsttaskrow, cpg.Task), Cells(lrow + 1000, cpg.Task)))
allowedRow = cFreeVersionTasksCount + 9
Do Until taskCount <= cFreeVersionTasksCount
Rows(allowedRow + 1).EntireRow.Delete: taskCount = taskCount - 1
Loop
End If
End Sub
Sub LoadNewGanttFormOnDblClick(Optional t As Boolean)
If IsLicValid(0, 1) Then bEditProject = True: frmNewGantt.show: bEditProject = False
End Sub

Sub TaskgridFormatting(cRow As Long, newtask As Boolean)
tlog "TaskgridFormatting"
Dim df As String, tStr As String, cur As String
'New tasks
If newtask Then'Rows(crow).Font.Color = RGB(0, 0, 0):
Rows(cRow).Font.Color = rgbBlack
With Range(Cells(cRow, cpg.SS), Cells(cRow, cpg.LC))
If Cells(cRow + 1, cpg.TIL) <= Cells(cRow, cpg.TIL) Then .Font.Bold = False Else .Font.Bold = True
.VerticalAlignment = xlCenter:.HorizontalAlignment = xlLeft: .Font.size = taskfontsize
End With
With Cells(cRow, cpg.Task)
.IndentLevel = Cells(cRow, cpg.TIL): .Font.Italic = False:
End With
With Cells(cRow, cpg.TaskIcon)
.value = "u": .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
If Cells(cRow, cpg.GEtype) = "T" Then .Font.Name = "Wingdings 3" Else .Font.Name = "Wingdings": .Font.size = 11
End With
With Cells(cRow, cpg.Priority)
.Font.Bold = True:.Font.size = 10:.Font.Color = vbWhite:.HorizontalAlignment = xlCenter
End With
With Cells(cRow, cpg.Status)
.HorizontalAlignment = xlCenter:.Font.size = 8
End With
With Cells(cRow, cpg.Done)
.value = 0: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
End With
With Cells(cRow, cpg.PercentageCompleted)
.NumberFormat = "0%": .HorizontalAlignment = xlCenter
End With
Cells(cRow, cpg.ECS).HorizontalAlignment = xlRight: Cells(cRow, cpg.ACS).HorizontalAlignment = xlRight
Cells(cRow, cpg.BCS).HorizontalAlignment = xlRight: Cells(cRow, cpg.ResourceCost).HorizontalAlignment = xlRight: Cells(cRow, cpg.Notes).NumberFormat = "General"
Cells(cRow, cpg.ESD).HorizontalAlignment = xlRight: Cells(cRow, cpg.EED).HorizontalAlignment = xlRight: Cells(cRow, cpg.BSD).HorizontalAlignment = xlRight:
Cells(cRow, cpg.BED).HorizontalAlignment = xlRight: Cells(cRow, cpg.ASD).HorizontalAlignment = xlRight: Cells(cRow, cpg.AED).HorizontalAlignment = xlRight:
Cells(cRow, cpg.ED).HorizontalAlignment = xlRight: Cells(cRow, cpg.BD).HorizontalAlignment = xlRight: Cells(cRow, cpg.AD).HorizontalAlignment = xlRight
Cells(cRow, cpg.Work).HorizontalAlignment = xlRight
Cells(cRow, cpg.WBSPredecessors).HorizontalAlignment = xlLeft: Cells(cRow, cpg.WBSSuccessors).HorizontalAlignment = xlLeft
Call DrawTasksBorders(cRow): Call DrawTimelineBorders(cRow)
End If
'OLD tasks
With Cells(cRow, cpg.Priority)
.Font.Bold = True
End With
With Cells(cRow, cpg.TaskIcon) ' needed again for convmile
If Cells(cRow, cpg.GEtype) = "T" Then .Font.Name = "Wingdings 3" Else .Font.Name = "Wingdings": .Font.size = 11
.WrapText = False
End With

df = st.DateFormat
If st.HGC Then df = df & " HH:MM AM/PM"
Cells(cRow, cpg.ESD).NumberFormat = df: Cells(cRow, cpg.EED).NumberFormat = df: Cells(cRow, cpg.BSD).NumberFormat = df:
Cells(cRow, cpg.BED).NumberFormat = df: Cells(cRow, cpg.ASD).NumberFormat = df: Cells(cRow, cpg.AED).NumberFormat = df
tStr = "[$$-409]#,##0.00": cur = st.CurrencyS
If Trim(cur) = "$" Then
Cells(cRow, cpg.ECS).NumberFormat = tStr: Cells(cRow, cpg.ACS).NumberFormat = tStr
Cells(cRow, cpg.BCS).NumberFormat = tStr: Cells(cRow, cpg.ResourceCost).NumberFormat = tStr
Else
Cells(cRow, cpg.ECS).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Cells(cRow, cpg.ACS).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Cells(cRow, cpg.BCS).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Cells(cRow, cpg.ResourceCost).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
End If
tlog "TaskgridFormatting"
End Sub

Sub FormatAllCosts()
tlog "FormatAllCosts"
Dim lt As Long: Dim tStr As String, cur As String
lt = GetLastRow + 2: tStr = "[$$-409]#,##0.00":: cur = st.CurrencyS
If Trim(cur) = "$" Then
Range(Cells(rownine, cpg.ACS), Cells(lt, cpg.ACS)).NumberFormat = tStr
Range(Cells(rownine, cpg.ECS), Cells(lt, cpg.ECS)).NumberFormat = tStr
Range(Cells(rownine, cpg.BCS), Cells(lt, cpg.BCS)).NumberFormat = tStr
Range(Cells(rownine, cpg.ResourceCost), Cells(lt, cpg.ResourceCost)).NumberFormat = tStr
Else
Range(Cells(rownine, cpg.ACS), Cells(lt, cpg.ACS)).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Range(Cells(rownine, cpg.ECS), Cells(lt, cpg.ECS)).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Range(Cells(rownine, cpg.BCS), Cells(lt, cpg.BCS)).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
Range(Cells(rownine, cpg.ResourceCost), Cells(lt, cpg.ResourceCost)).NumberFormat = Chr(34) & cur & Chr(34) & " " & "#,##0.00"
End If
tlog "FormatAllCosts"
End Sub

Sub FormatAllDates()
tlog "FormatAllDates"
Dim df As String: Dim lt As Long: lt = GetLastRow + 2: df = st.DateFormat
If st.HGC Then df = df & " HH:MM AM/PM"
Range(Cells(firsttaskrow, cpg.ESD), Cells(lt, cpg.ESD)).NumberFormat = df
Range(Cells(firsttaskrow, cpg.EED), Cells(lt, cpg.EED)).NumberFormat = df
Range(Cells(firsttaskrow, cpg.BSD), Cells(lt, cpg.BSD)).NumberFormat = df
Range(Cells(firsttaskrow, cpg.BED), Cells(lt, cpg.BED)).NumberFormat = df
Range(Cells(firsttaskrow, cpg.ASD), Cells(lt, cpg.ASD)).NumberFormat = df
Range(Cells(firsttaskrow, cpg.AED), Cells(lt, cpg.AED)).NumberFormat = df
tlog "FormatAllDates"
End Sub

Sub ShowOutofDatesMessage(Optional t As String)
MsgBox "Enter a date between " & Format(csDate, "DD-MMM-YYYY") & " and " & Format(ceDate, "DD-MMM-YYYY")
End Sub

Sub AddNewTaskPlaceholder(Optional t As Boolean)
Dim r As Range: Dim lasttaskno As Long
lasttaskno = GetLastRow + 1:Set r = Cells(lasttaskno, cpg.Task)
If r <> sAddTaskPlaceHolder Then
With r
.value = sAddTaskPlaceHolder: .Font.Italic = True: .Font.Color = rgbBlack:.IndentLevel = 0: .HorizontalAlignment = xlLeft
End With
Else
Exit Sub
End If
End Sub
Option Explicit
Option Private Module

Sub ReCalculateDates(Optional cRowOnly As Long, Optional field As String, Optional familytype As String)
tlog "ReCalculateDates" & field ' check if we can array thiss
Call RememberResArrays: ResArraysReady = True:
Dim cRow As Long, lrow As Long, StartRow As Long, EndRow As Long: lrow = GetLastRow
If field = "" Then field = allFields
If familytype = "" Then familytype = allRows
If CheckForDependency(ActiveSheet) Then Call CalcDepFormulas
StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
For cRow = StartRow To EndRow
If field = estDates Or field = allFields Then
If st.HGC Then
If Cells(cRow, cpg.GEtype) = "T" Then
Cells(cRow, cpg.EED) = CalEEDHrs(Cells(cRow, cpg.Resource), Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
Else
Cells(cRow, cpg.EED) = Cells(cRow, cpg.ESD)
End If
If Cells(cRow, cpg.Dependents) <> "" Or Cells(cRow, cpg.Dependency) <> "" Then Call ReCalcDepFormulas(cRow, True)
Else
If Cells(cRow, cpg.GEtype) = "T" Then
Cells(cRow, cpg.EED) = GetEndDateFromWorkDays(Cells(cRow, cpg.Resource), Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
Else
Cells(cRow, cpg.EED) = Cells(cRow, cpg.ESD)
End If
If Cells(cRow, cpg.Dependents) <> "" Or Cells(cRow, cpg.Dependency) <> "" Then ReCalcDepFormulas cRow
End If
End If
If st.CalBasDates = False Then GoTo basLast
If field = basDates Or field = allFields Then
If IsDate(Cells(cRow, cpg.BSD)) Then
If IsParentTask(cRow) Then GoTo basLast
Cells(cRow, cpg.BSD) = CDate(Cells(cRow, cpg.BSD))
If IsNumeric(Cells(cRow, cpg.BD)) = False Then If Cells(cRow, cpg.GEtype) = "T" Then Cells(cRow, cpg.BD) = 1 Else Cells(cRow, cpg.BD) = 0
Else
Cells(cRow, cpg.BSD) = "": Cells(cRow, cpg.BED) = "": Cells(cRow, cpg.BD) = "":GoTo basLast:
End If
If st.HGC Then
Cells(cRow, cpg.BED) = CalEEDHrs(CStr(Cells(cRow, cpg.Resource)), CDate(Cells(cRow, cpg.BSD)), CInt(Cells(cRow, cpg.BD)))
Else
Cells(cRow, cpg.BED) = GetEndDateFromWorkDays(CStr(Cells(cRow, cpg.Resource)), CDate(Cells(cRow, cpg.BSD)), CInt(Cells(cRow, cpg.BD)))
End If
End If
basLast:
If st.CalActDates = False Then GoTo actLast
If field = actDates Or field = allFields Then
If IsDate(Cells(cRow, cpg.ASD)) Then
If IsParentTask(cRow) Then GoTo actLast
Cells(cRow, cpg.ASD) = CDate(Cells(cRow, cpg.ASD))
If IsNumeric(Cells(cRow, cpg.AD)) = False Then If Cells(cRow, cpg.GEtype) = "T" Then Cells(cRow, cpg.AD) = 1 Else Cells(cRow, cpg.AD) = 0
Else
Cells(cRow, cpg.ASD) = "": Cells(cRow, cpg.AED) = "": Cells(cRow, cpg.AD) = "":GoTo actLast
End If
If st.HGC Then
Cells(cRow, cpg.AED) = CalEEDHrs(CStr(Cells(cRow, cpg.Resource)), CDate(Cells(cRow, cpg.ASD)), CInt(Cells(cRow, cpg.AD)))
Else
Cells(cRow, cpg.AED) = GetEndDateFromWorkDays(CStr(Cells(cRow, cpg.Resource)), CDate(Cells(cRow, cpg.ASD)), CInt(Cells(cRow, cpg.AD)))
End If
End If
actLast:
Next cRow
Call ClearDepFormulas
tlog "ReCalculateDates" & field
End Sub

Sub AutoPopulatePercentages()
Call PopParentTasks(, allFields)
End Sub

Sub CompletedAction()
Dim cRow As Long: cRow = SelTaskRow: Dim c As Range: Set c = Cells(SelTaskRow, cpg.Done)
Dim bNormaltask As Boolean, bParentTask As Boolean, bChildTask As Boolean
If c.value = 100 Then
If IsParentTask(cRow) Then
MsgBox msg(2)
If Cells(cRow, cpg.PercentageCompleted) = 1 Then c.value = 100 Else c.value = 0
Exit Sub
End If
If st.PercAuto Then
MsgBox msg(10)
If Cells(cRow, cpg.PercentageCompleted) = 1 Then c.value = 100 Else c.value = 0
Exit Sub
Else
c.value = 0: Cells(cRow, cpg.PercentageCompleted) = 0
End If
ElseIf c.value = 0 Then
If IsParentTask(cRow) Then
MsgBox msg(2)
If Cells(cRow, cpg.PercentageCompleted) = 1 Then c.value = 100 Else c.value = 0
Exit Sub
End If
If st.PercAuto Then
MsgBox msg(10)
If Cells(cRow, cpg.PercentageCompleted) = 1 Then c.value = 100 Else c.value = 0
Exit Sub
Else
c.value = 100: Cells(cRow, cpg.PercentageCompleted) = 1
End If
End If
If IsNormalTask(cRow) Then 'classify tasks
bNormaltask = True
Else
bNormaltask = False
If IsParentTask(cRow) Then bParentTask = True: bChildTask = False Else bChildTask = True: bParentTask = False
End If
If bChildTask Then
Call PopParentTasks(, fPerc)
Call FormatTasks(cRow, allFields, aboveFamily)
Call DrawGanttBars(cRow, estBars, aboveFamily)
Else
Call FormatTasks(cRow, allFields, rowOnly)
Call DrawGanttBars(cRow, estBars, rowOnly)
End If
End Sub
Option Explicit
Option Private Module
Private arrAllData()

Sub check()
Call DA
Dim ws As Worksheet, curSheet As Worksheet: Set curSheet = ActiveSheet
Call checkReqWorksheets
Call DA
Call CheckDupSettingsSheets
For Each ws In ActiveWorkbook.Sheets
If GanttChart(ws) Then Call checkgc(ws)
Next ws
curSheet.Activate
MsgBox "Checks Done"
Call EA
End Sub

Sub checkgc(gc As Worksheet)
Dim ws As Worksheet: Dim lrow As Long, iAs Long, j As Long: Dim tidRange As Range
Set ws = gc: ws.Activate
Call checkGCColumns(ws):
Call ActivateGanttChart 'Call CalcColPosGCT: Call CalcColPosTimeline: Call CalcColPosGST: Call ReadSettings: sArr.LoadAllArrays:
lrow = GetLastRow(ws): arrAllData = ws.Range(ws.Cells(1, cpg.GEtype), ws.Cells(lrow, cpg.LC)).value
Set tidRange = ws.Range(ws.Cells(firsttaskrow, cpg.TID), ws.Cells(lrow, cpg.TID))
For i = firsttaskrow To UBound(arrAllData())
If arrAllData(i, cpg.GEtype) = "" Then MsgBox ws.Name & " Err: " & "GEType missing on row " & i
If arrAllData(i, cpg.TID) = "" Then MsgBox ws.Name & " Err: " & "TID missing on row " & i
If arrAllData(i, cpg.TIL) = "" Then MsgBox ws.Name & " Err: " & "TIL missing on row " & i
If arrAllData(i, cpg.TIL) <> ws.Cells(i, cpg.Task).IndentLevel Then MsgBox ws.Name & " Err: " & "Indent level mismatch on row " & i
If arrAllData(i, cpg.WBS) = "" Then MsgBox ws.Name & " Err: " & "WBS missing on row " & i
If arrAllData(i, cpg.Task) = "" Then MsgBox ws.Name & " Err: " & "Task missing on row " & i
If arrAllData(i, cpg.Priority) = "" Then MsgBox ws.Name & " Err: " & "Priority missing on row " & i
If arrAllData(i, cpg.Status) = "" Then MsgBox ws.Name & " Err: " & "Status missing on row " & i
If arrAllData(i, cpg.ESD) = "" Then MsgBox ws.Name & " Err: " & "ESD missing on row " & i
If arrAllData(i, cpg.EED) = "" Then MsgBox ws.Name & " Err: " & "EED missing on row " & i
If arrAllData(i, cpg.ED) = "" Then MsgBox ws.Name & " Err: " & "ED missing on row " & i
If arrAllData(i, cpg.PercentageCompleted) = "" Then MsgBox ws.Name & " Err: " & "Percentage Completed missing on row " & i
Next i
Call DuplicateValuesinRange(tidRange)
End Sub

Sub checkReqWorksheets()
Call DA
Dim ws As Worksheet, gsws As Worksheet, rsws As Worksheet: Dim missingWS As Boolean
Dim SS As Long, GRT As Long, ProjectCounter As Long
If WSExists("GDT") = False Then MsgBox "Gantt Dashboard Template " & msg(31): GoTo Last
If WSExists("GDD") = False Then MsgBox "Gantt Dashboard Data " & msg(31): GoTo Last
If WSExists("GST") = False Then MsgBox "Gantt Settings Template " & msg(31): GoTo Last
If WSExists("PVS") = False Then MsgBox "Gantt Dashboard Pivot " & msg(31): GoTo Last
ProjectCounter = 0: missingWS = False
For Each ws In ActiveWorkbook.Sheets
If GanttChart(ws) Then ProjectCounter = ProjectCounter + 1
Next ws
Debug.Print "Total Projects: " & ProjectCounter & " " & "ProjectCount: " & GST.Cells(rowtwo, cps.SSN)

Dim gcName As String, gcGSName As String, gcRSName As String, ptype As String: Dim projNo As Long

For Each ws In ThisWorkbook.Sheets
gcName = "": gcGSName = "": gcRSName = "": projNo = 0: ptype = ""
If GanttChart(ws) Then
gcName = ws.Name: projNo = getPID(ws)
If CheckSheet(getGSname(ws)) = False Then
MsgBox "Settings sheet for " & ws.Name & msg(31)
gcGSName = "Missing - " & getGSname(ws): missingWS = True
Else
Set gsws = setGSws(ws): gcGSName = gsws.Name
ptype = gsws.Cells(rowtwo, 1)
If ptype = "s0n84" Then ptype = "H" Else ptype = "D"
End If
If CheckSheet(getRSname(ws)) = False Then
MsgBox "Resource sheet for " & ws.Name & msg(31)
gcRSName = "Missing - " & getRSname(ws): missingWS = True
Else
Set rsws = setRSws(ws): gcRSName = rsws.Name
End If
Debug.Print "Project:" & ptype & projNo & " Worksheet:" & gcName & " | " & gcGSName & " | " & gcRSName
Set gsws = Nothing:Set rsws = Nothing
End If
Next ws
For Each ws In ThisWorkbook.Sheets
If ws.Name <> GDT.Name And ws.Name <> GDD.Name And ws.Name <> GST.Name And ws.Name <> PVS.Name And Left(ws.Name, 2) <> "GS" And Left(ws.Name, 2) <> "RS" And ws.Cells(1, 1) <> "UserDashSheet" And ws.Name <> "Help" And Not GanttChart(ws) Then
Debug.Print "User WS: " & ws.Name
End If
Next ws
Call CheckOrphanedGS: Call CheckOrphanedRS
If missingWS Then GoTo Last
Call EA
Exit Sub
Last:
ThisWorkbook.Save
MsgBox msg(40)
Call EA
End Sub

Sub checkGCColumns(Optional gc As Worksheet)
Dim ws As Worksheet: Dim r As Range: Dim Res As Variant: Dim i As Long
If gc Is Nothing Then Set ws = ActiveSheet Else Set ws = gc
Set r = ws.Range("1:1")
For i = 1 To UBound(GCcolumns())
Res = Application.Match(GCcolumns(i), r.value, 0)
If IsError(Res) Then MsgBox GCcolumns(i) & " column not found in " & gc.Name
Next i
Set r = Nothing: Set ws = Nothing
End Sub

Sub DuplicateValuesinRange(rg As Range)
Dim i As Integer
Dim j As Integer
Dim MyCell As Range
For Each MyCell In rg
If WorksheetFunction.CountIf(rg, MyCell.value) > 1 Then
MsgBox "Duplicate TID in row " & MyCell.Row
End If
Next
End Sub

Sub CheckOrphanedGS(Optional DeleteGS As Boolean)
Dim ws As Worksheet, gcws As Worksheet: Dim orphanedGS As Boolean
For Each ws In ThisWorkbook.Sheets
If Left(ws.Name, 2) = "GS" Then
For Each gcws In ThisWorkbook.Sheets
If GanttChart(gcws) Then
If ws.Name = getGSname(gcws) Then orphanedGS = False: GoTo nextGS Else orphanedGS = True
End If
Next
If orphanedGS Then
If DeleteGS Then
MsgBox "Deleting " & ws.Name: ws.visible = xlSheetVisible: ws.Delete
Else
Debug.Print "OrphanedGS: " & ws.Name
End If
End If
End If
nextGS:
Next ws
If orphanedGS = False And DeleteGS = True Then MsgBox "No Orphaned Settings Sheets found"
End Sub

Sub CheckOrphanedRS(Optional DeleteRS As Boolean)
Dim ws As Worksheet, gcws As Worksheet: Dim orphanedRS As Boolean
For Each ws In ThisWorkbook.Sheets
If Left(ws.Name, 2) = "RS" Then
For Each gcws In ThisWorkbook.Sheets
If GanttChart(gcws) Then
If ws.Name = getRSname(gcws) Then orphanedRS = False: GoTo nextRS Else orphanedRS = True
End If
Next
If orphanedRS Then
If DeleteRS Then
MsgBox "Deleting " & ws.Name: ws.visible = xlSheetVisible:ws.Delete
Else
Debug.Print "OrphanedRS: " & ws.Name
End If
End If
End If
nextRS:
Next ws
If orphanedRS = False And DeleteRS = True Then MsgBox "No Orphaned Resource Sheets found"
End Sub

Sub CheckDupSettingsSheets()
Call DA 'checks dup settings sheet on dashboard, resource and settings button click
Dim ws As Worksheet, cs As Worksheet: Dim DupFound As Boolean: Dim gsName As String
Dim PID As Long, newPID As Long
Dim dupgs As Worksheet, duprs As Worksheet, curWs As Worksheet
Set curWs = ActiveSheet
For Each ws In ActiveWorkbook.Sheets
If GanttChart(ws) Then
gsName = getGSname(ws): PID = getPID(ws)
For Each cs In ThisWorkbook.Sheets
If cs.Name = ws.Name Then GoTo nexcs
If GanttChart(cs) Then
If PID = getPID(cs) Then
DupFound = True: Set dupgs = setGSws(ws): Set duprs = setRSws(ws)
Debug.Print "dup " & cs.Name
Call ProjectCountPlusOne: newPID = getProjectCounter
dupgs.visible = xlSheetVisible
dupgs.Copy , Worksheets(Worksheets.Count): ActiveSheet.Name = "GS" & newPID
ActiveSheet.Cells(rowtwo, cps.GRT) = "RS" & newPID: ActiveSheet.Range("LL1", "LZ2").Clear
Call hideWorksheet(ActiveSheet):
duprs.visible = xlSheetVisible
duprs.Copy , Worksheets(Worksheets.Count): ActiveSheet.Name = "RS" & newPID
Call hideWorksheet(ActiveSheet): Call hideWorksheet(dupgs): Call hideWorksheet(duprs)
cs.Activate
Call setGSRSname("GS" & newPID, "RS" & newPID, cs, newPID)
End If
End If
nexcs:
Next cs
End If
Next ws
If DupFound Then Call checkReqWorksheets
Call DA: curWs.Activate: Call ActivateGanttChart
End Sub
Private Sub Class_Initialize()
LoadAllArrays
End Sub
Public Sub LoadAllArrays()
tlog "LoadAllArrays"
LoadResourceData
LoadWorkdaysP
LoadHolidaysP
tlog "LoadAllArrays"
End Sub
Public Property Get WorkdaysP() As Variant
WorkdaysP = arrWorkdaysP
End Property
Public Property Get HolidaysP() As Variant
HolidaysP = arrHolidaysP
End Property
Public Property Get HolidaysFull() As Variant
HolidaysFull = arrHolidaysFull
End Property
Public Property Get ResourceP() As Variant
ResourceP = arrResourceData
End Property

Public Sub LoadResourceData()
Dim i, j As Long
Dim noOfResources As Long
Set gs = setGSws(ActiveSheet): Set rs = setRSws(ActiveSheet)
noOfResources = Application.WorksheetFunction.CountA(rs.Range("A:A")) - 2
ReDim arrResourceData(0 To noOfResources, 0 To 11) As Variant
For i = 0 To noOfResources
arrResourceData(i, 0) = LCase(rs.Range("A2").Offset(i, j).value)
Next
For i = 0 To noOfResources
For j = 1 To 11
arrResourceData(i, j) = rs.Range("A2").Offset(i, j).value
Next j
Next i
End Sub
Public Sub LoadWorkdaysP()
Dim noOfResources As Long
Set gs = setGSws(ActiveSheet): Set rs = setRSws(ActiveSheet)
noOfResources = Application.WorksheetFunction.CountA(rs.Range("A:A")) - 2
ReDim arrWorkdaysP(0 To noOfResources, 0 To 7) As Variant
Dim i, j, k, l As Long
j = 0
For i = 0 To noOfResources
arrWorkdaysP(i, j) = LCase(rs.Range("A2").Offset(i, j).value)
Next
For i = 0 To noOfResources
For j = 1 To 6
arrWorkdaysP(i, j) = rs.Range("N2").Offset(i, j - 1).value
Next j
Next i
For i = 0 To noOfResources
j = 7
arrWorkdaysP(i, j) = rs.Range("M2").Offset(i, 0).value
Next i
End Sub
Public Sub LoadHolidaysP()
Dim noOfResources As Long
Set gs = setGSws(ActiveSheet): Set rs = setRSws(ActiveSheet)
noOfResources = Application.WorksheetFunction.CountA(rs.Range("A:A")) - 2
Dim i, j, k, l As Long
Dim lastholcol, counter, maxnoofholidays As Long
Dim holrange As Range
lastholcol = rs.Range("T1").CurrentRegion.Columns.Count - 9
maxnoofholidays = lastholcol / 2
ReDim arrHolidaysP(0 To noOfResources, 0 To maxnoofholidays) As Variant
j = 0
For i = 0 To noOfResources
arrHolidaysP(i, j) = LCase(rs.Range("A2").Offset(i, j).value)
Next
For i = 0 To noOfResources
counter = 0
k = 1
For j = 1 To maxnoofholidays
arrHolidaysP(i, j) = rs.Range("T2").Offset(i, k - 1 + counter).value
counter = counter + 1
k = k + 1
Next j
Next i
End Sub
Public Sub LoadHolidaysFull()
Dim noOfResources As Long
Set gs = setGSws(ActiveSheet): Set rs = setRSws(ActiveSheet)
noOfResources = Application.WorksheetFunction.CountA(rs.Range("A:A")) - 2
Dim i, j, k, l As Long
Dim lastholcol, counter, maxnoofholidays As Long
Dim holrange As Range
lastholcol = rs.Range("T1").CurrentRegion.Columns.Count - 9
maxnoofholidays = lastholcol
ReDim arrHolidaysFull(0 To noOfResources, 0 To maxnoofholidays) As Variant
j = 0
For i = 0 To noOfResources
arrHolidaysFull(i, j) = LCase(rs.Range("A2").Offset(i, j).value)
Next
For i = 0 To noOfResources
For j = 1 To maxnoofholidays
arrHolidaysFull(i, j) = rs.Range("T2").Offset(i, j - 1).value
Next j
Next i
End Sub
Option Explicit
Const cDayLabels As String = "lbl_D_"
Public WithEvents frmLabel As MSForms.label

Private Sub frmLabel_Click()
MarkDate
End Sub

Private Sub frmLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
MarkDate True
End Sub
Sub MarkDate(Optional fClose As Boolean)

If frmLabel.Caption = vbNullString Then Exit Sub

Dim r As Integer, c As Integer

For r = 1 To 6
For c = 1 To 7
With frmDateSelector.Controls(cDayLabels & r & "_" & c)
If .Caption = vbNullString Then
.BackColor = frmDateSelector.BackColor
Else
.BackColor = vbWhite
End If
End With
Next c
Next r

With frmDateSelector
.Controls(frmLabel.Name).BackColor = vbGreen
.lblSelectedDate.Caption = Format(DateSerial(.cmbYears.value, .cmbMonths.value, CInt(frmLabel.Caption)), .lblNumberFormat.Caption)
.lblSelectedDay.Caption = frmLabel.Caption
End With

If fClose = True Then frmDateSelector.ClickOkButton

End Sub
Option Explicit
Public GEtype As Long, TID As Long, Dependency As Long, Dependents As Long, StartConstrain As Long, EndConstrain As Long, TIL As Long
Public SS As Long, TaskIcon As Long, WBS As Long, Task As Long, Priority As Long, Status As Long, Resource As Long, ResourceCost As Long
Public BSD As Long, BED As Long, BD As Long, ESD As Long, EED As Long, ED As Long, WBSPredecessors As Long, WBSSuccessors As Long
Public Work As Long, Done As Long, PercentageCompleted As Long
Public ASD As Long, AED As Long, AD As Long, BCS As Long, ECS As Long, ACS As Long, Notes As Long
Public TColor As Long, TPColor As Long, BLColor As Long, ACColor As Long
Public Custom1 As Long, Custom2 As Long, Custom3 As Long, Custom4 As Long, Custom5 As Long, Custom6 As Long, Custom7 As Long, Custom8 As Long, Custom9 As Long, Custom10 As Long
Public Custom11 As Long, Custom12 As Long, Custom13 As Long, Custom14 As Long, Custom15 As Long, Custom16 As Long, Custom17 As Long, Custom18 As Long, Custom19 As Long, Custom20 As Long
Public ShapeInfoE As Long, ShapeInfoB As Long, ShapeInfoA As Long, LC As Long

Private Sub Class_Initialize()
Dim ws As Worksheet, r As Range, tStr As String, wsname As String
If ActiveSheet.Range("A1") <> "GEType" Then
MsgBox "clsGetColNosGCT called from non gantt chart worksheet": Exit Sub
Else
Set ws = ActiveSheet:wsname = "CalColPos " & ws.Name
End If
Set r = ws.Range("1:1")
GEtype = Application.WorksheetFunction.Match("GEType", r.value, 0)
TID = Application.WorksheetFunction.Match("TID", r.value, 0)
Dependency = Application.WorksheetFunction.Match("Dependency", r.value, 0)
Dependents = Application.WorksheetFunction.Match("Dependents", r.value, 0)
StartConstrain = Application.WorksheetFunction.Match("StartConstrain", r.value, 0)
EndConstrain = Application.WorksheetFunction.Match("EndConstrain", r.value, 0)
TIL = Application.WorksheetFunction.Match("TIL", r.value, 0)
SS = Application.WorksheetFunction.Match("SS", r.value, 0)
TaskIcon = Application.WorksheetFunction.Match("TaskIcon", r.value, 0)
WBS = Application.WorksheetFunction.Match("WBS", r.value, 0)
Task = Application.WorksheetFunction.Match("Task", r.value, 0)
Priority = Application.WorksheetFunction.Match("Priority", r.value, 0)
Status = Application.WorksheetFunction.Match("Status", r.value, 0)
Resource = Application.WorksheetFunction.Match("Resource", r.value, 0)
ResourceCost = Application.WorksheetFunction.Match("ResourceCost", r.value, 0)
BSD = Application.WorksheetFunction.Match("BSD", r.value, 0)
BED = Application.WorksheetFunction.Match("BED", r.value, 0)
BD = Application.WorksheetFunction.Match("BD", r.value, 0)
ESD = Application.WorksheetFunction.Match("ESD", r.value, 0)
EED = Application.WorksheetFunction.Match("EED", r.value, 0)
ED = Application.WorksheetFunction.Match("ED", r.value, 0)
WBSPredecessors = Application.WorksheetFunction.Match("WBSPredecessors", r.value, 0)
WBSSuccessors = Application.WorksheetFunction.Match("WBSSuccessors", r.value, 0)
Work = Application.WorksheetFunction.Match("Work", r.value, 0)
Done = Application.WorksheetFunction.Match("Done", r.value, 0)
PercentageCompleted = Application.WorksheetFunction.Match("PercentageCompleted", r.value, 0)
ASD = Application.WorksheetFunction.Match("ASD", r.value, 0)
AED = Application.WorksheetFunction.Match("AED", r.value, 0)
AD = Application.WorksheetFunction.Match("AD", r.value, 0)
BCS = Application.WorksheetFunction.Match("BCS", r.value, 0)
ECS = Application.WorksheetFunction.Match("ECS", r.value, 0)
ACS = Application.WorksheetFunction.Match("ACS", r.value, 0)
Notes = Application.WorksheetFunction.Match("Notes", r.value, 0)
TColor = Application.WorksheetFunction.Match("TColor", r.value, 0)
TPColor = Application.WorksheetFunction.Match("TPColor", r.value, 0)
BLColor = Application.WorksheetFunction.Match("BLColor", r.value, 0)
ACColor = Application.WorksheetFunction.Match("ACColor", r.value, 0)
Custom1 = Application.WorksheetFunction.Match("Custom 1", r.value, 0)
Custom2 = Application.WorksheetFunction.Match("Custom 2", r.value, 0)
Custom3 = Application.WorksheetFunction.Match("Custom 3", r.value, 0)
Custom4 = Application.WorksheetFunction.Match("Custom 4", r.value, 0)
Custom5 = Application.WorksheetFunction.Match("Custom 5", r.value, 0)
Custom6 = Application.WorksheetFunction.Match("Custom 6", r.value, 0)
Custom7 = Application.WorksheetFunction.Match("Custom 7", r.value, 0)
Custom8 = Application.WorksheetFunction.Match("Custom 8", r.value, 0)
Custom9 = Application.WorksheetFunction.Match("Custom 9", r.value, 0)
Custom10 = Application.WorksheetFunction.Match("Custom 10", r.value, 0)
Custom11 = Application.WorksheetFunction.Match("Custom 11", r.value, 0)
Custom12 = Application.WorksheetFunction.Match("Custom 12", r.value, 0)
Custom13 = Application.WorksheetFunction.Match("Custom 13", r.value, 0)
Custom14 = Application.WorksheetFunction.Match("Custom 14", r.value, 0)
Custom15 = Application.WorksheetFunction.Match("Custom 15", r.value, 0)
Custom16 = Application.WorksheetFunction.Match("Custom 16", r.value, 0)
Custom17 = Application.WorksheetFunction.Match("Custom 17", r.value, 0)
Custom18 = Application.WorksheetFunction.Match("Custom 18", r.value, 0)
Custom19 = Application.WorksheetFunction.Match("Custom 19", r.value, 0)
Custom20 = Application.WorksheetFunction.Match("Custom 20", r.value, 0)
ShapeInfoE = Application.WorksheetFunction.Match("ShapeInfoE", r.value, 0)
ShapeInfoB = Application.WorksheetFunction.Match("ShapeInfoB", r.value, 0)
ShapeInfoA = Application.WorksheetFunction.Match("ShapeInfoA", r.value, 0)
LC = Application.WorksheetFunction.Match("LC", r.value, 0)
Set r = Nothing: Set ws = Nothing
End Sub
Option Explicit
Public ProjectID As Long, dType As Long
Public GEtype As Long, TID As Long, Dependency As Long, Dependents As Long, StartConstrain As Long, EndConstrain As Long, TIL As Long
Public SS As Long, TaskIcon As Long, WBS As Long, Task As Long, Priority As Long, Status As Long, Resource As Long, ResourceCost As Long
Public BSD As Long, BED As Long, BD As Long, ESD As Long, EED As Long, ED As Long, Work As Long, Done As Long, PercentageCompleted As Long
Public ASD As Long, AED As Long, AD As Long, BCS As Long, ECS As Long, ACS As Long, Notes As Long
Public TColor As Long, TPColor As Long, BLColor As Long, ACColor As Long
Public Custom1 As Long, Custom2 As Long, Custom3 As Long, Custom4 As Long, Custom5 As Long, Custom6 As Long, Custom7 As Long, Custom8 As Long, Custom9 As Long, Custom10 As Long
Public Custom11 As Long, Custom12 As Long, Custom13 As Long, Custom14 As Long, Custom15 As Long, Custom16 As Long, Custom17 As Long, Custom18 As Long, Custom19 As Long, Custom20 As Long
Public ShapeInfoE As Long, ShapeInfoB As Long, ShapeInfoA As Long, LC As Long

Private Sub Class_Initialize()
Dim ws As Worksheet, r As Range, tStr As String, wsname As String

Set r = GDD.Range("1:1")
ProjectID = Application.WorksheetFunction.Match("ProjectID", r.value, 0)
dType = Application.WorksheetFunction.Match("dType", r.value, 0)
GEtype = Application.WorksheetFunction.Match("GEType", r.value, 0)
TID = Application.WorksheetFunction.Match("TID", r.value, 0)
Dependency = Application.WorksheetFunction.Match("Dependency", r.value, 0)
Dependents = Application.WorksheetFunction.Match("Dependents", r.value, 0)
StartConstrain = Application.WorksheetFunction.Match("StartConstrain", r.value, 0)
EndConstrain = Application.WorksheetFunction.Match("EndConstrain", r.value, 0)
TIL = Application.WorksheetFunction.Match("TIL", r.value, 0)
SS = Application.WorksheetFunction.Match("SS", r.value, 0)
TaskIcon = Application.WorksheetFunction.Match("TaskIcon", r.value, 0)
WBS = Application.WorksheetFunction.Match("WBS", r.value, 0)
Task = Application.WorksheetFunction.Match("Task", r.value, 0)
Priority = Application.WorksheetFunction.Match("Priority", r.value, 0)
Status = Application.WorksheetFunction.Match("Status", r.value, 0)
Resource = Application.WorksheetFunction.Match("Resource", r.value, 0)
ResourceCost = Application.WorksheetFunction.Match("ResourceCost", r.value, 0)
BSD = Application.WorksheetFunction.Match("BSD", r.value, 0)
BED = Application.WorksheetFunction.Match("BED", r.value, 0)
BD = Application.WorksheetFunction.Match("BD", r.value, 0)
ESD = Application.WorksheetFunction.Match("ESD", r.value, 0)
EED = Application.WorksheetFunction.Match("EED", r.value, 0)
ED = Application.WorksheetFunction.Match("ED", r.value, 0)
Work = Application.WorksheetFunction.Match("Work", r.value, 0)
Done = Application.WorksheetFunction.Match("Done", r.value, 0)
PercentageCompleted = Application.WorksheetFunction.Match("PercentageCompleted", r.value, 0)
ASD = Application.WorksheetFunction.Match("ASD", r.value, 0)
AED = Application.WorksheetFunction.Match("AED", r.value, 0)
AD = Application.WorksheetFunction.Match("AD", r.value, 0)
BCS = Application.WorksheetFunction.Match("BCS", r.value, 0)
ECS = Application.WorksheetFunction.Match("ECS", r.value, 0)
ACS = Application.WorksheetFunction.Match("ACS", r.value, 0)
Notes = Application.WorksheetFunction.Match("Notes", r.value, 0)
Custom1 = Application.WorksheetFunction.Match("Custom 1", r.value, 0)
Custom2 = Application.WorksheetFunction.Match("Custom 2", r.value, 0)
Custom3 = Application.WorksheetFunction.Match("Custom 3", r.value, 0)
Custom4 = Application.WorksheetFunction.Match("Custom 4", r.value, 0)
Custom5 = Application.WorksheetFunction.Match("Custom 5", r.value, 0)
Custom6 = Application.WorksheetFunction.Match("Custom 6", r.value, 0)
Custom7 = Application.WorksheetFunction.Match("Custom 7", r.value, 0)
Custom8 = Application.WorksheetFunction.Match("Custom 8", r.value, 0)
Custom9 = Application.WorksheetFunction.Match("Custom 9", r.value, 0)
Custom10 = Application.WorksheetFunction.Match("Custom 10", r.value, 0)
Custom11 = Application.WorksheetFunction.Match("Custom 11", r.value, 0)
Custom12 = Application.WorksheetFunction.Match("Custom 12", r.value, 0)
Custom13 = Application.WorksheetFunction.Match("Custom 13", r.value, 0)
Custom14 = Application.WorksheetFunction.Match("Custom 14", r.value, 0)
Custom15 = Application.WorksheetFunction.Match("Custom 15", r.value, 0)
Custom16 = Application.WorksheetFunction.Match("Custom 16", r.value, 0)
Custom17 = Application.WorksheetFunction.Match("Custom 17", r.value, 0)
Custom18 = Application.WorksheetFunction.Match("Custom 18", r.value, 0)
Custom19 = Application.WorksheetFunction.Match("Custom 19", r.value, 0)
Custom20 = Application.WorksheetFunction.Match("Custom 20", r.value, 0)
LC = Application.WorksheetFunction.Match("LC", r.value, 0)
Set r = Nothing: Set ws = Nothing
End Sub
Option Explicit

Public ProjectName As Long
Public GRT As Long, SSN As Long
Public BaselineBudget As Long, EstimatedBudget As Long, ActualCosts As Long, TotalACS As Long
Public ShowOverdueBar As Long, ShowEstBaseBar As Long, ShowPercBar As Long, ShowPercDataBar As Long, ShowBaselineBar As Long, ShowActualBar As Long
Public CurrencySymbol As Long, DateFormat As Long, CurrentView As Long
Public ShowCompleted As Long, ShowInProgress As Long, ShowPlanned As Long, ShowLate As Long
Public WeekStartDay As Long, WeekNumType As Long, csDate As Long, ceDate As Long, csYear As Long, ceYear As Long
Public EnableCostsModule As Long
Public STG As Long
Public HiliteHolidays As Long, HiliteWorkOffDays As Long, HiliteHolidaysPR As Long, HiliteWorkOffDaysPR As Long, HideHolidays As Long, HideWorkOffDays As Long, HideNonWorkingHours As Long
Public ShowGrouping As Long, ShowTodayLines As Long
Public tLicenseVal As Long, tUsrName As Long, tUsrEmailID As Long, tUsrActivatedDate As Long, tliky As Long, tLicDuration As Long, tLiType As Long, tFirstSavedDate As Long
Public BarTextEnable As Long, BarTextCharacters As Long, BarTextFontSize As Long, BarTextDataColumnName As Long, BarTextIsBold As Long
Public PercentageEntryMode As Long, PercentageCalculationType As Long
Public ShowDependencyConnector As Long
Public ThemeName As Long, SelectedTheme As Long
Public EBC As Long, EBaseC As Long, EMC As Long, PBC As Long, PDC As Long, BBC As Long, ABC As Long, OBC As Long, DLC As Long, TLC As Long, TBC As Long, GBC As Long
Public PRC As Long, HBC As Long, hc As Long, WC As Long, HCPR As Long, WCPR As Long, CR1C As Long, CR12C As Long, CR2C As Long, CR3C As Long
Public TPCH As Long, TPCN As Long, TPCL As Long, TSCC As Long, TSCI As Long, TSCP As Long
Public TGB As Long, gcType As Long
Public HCOL As Long, dCol As Long, WCOL As Long, MCOL As Long, QCOL As Long, HYCOL As Long, YCOL As Long
Public HWID As Long, DWID As Long, WWID As Long, MWID As Long, QWID As Long, HYWID As Long, YWID As Long
Public RCCal As Long, PTRC As Long, PTDS As Long
Public TSD As Long, TED As Long
Public RefreshTimeline As Long, TimelineVisible As Long, VerticalBorders As Long
Public CalcBED As Long, CalcAED As Long, SavedDate As Long, LockWB As Long, Campaign As Long

Private Sub Class_Initialize()
Dim ws As Worksheet, sr As Range, tStr As String
Set ws = GST: Set sr = GST.Range("1:1") 'GST sheet
SSN = Application.WorksheetFunction.Match("SSN", sr.value, 0)
GRT = Application.WorksheetFunction.Match("GRT", sr.value, 0)
BaselineBudget = Application.WorksheetFunction.Match("BaselineBudget", sr.value, 0)
EstimatedBudget = Application.WorksheetFunction.Match("EstimatedBudget", sr.value, 0)
ActualCosts = Application.WorksheetFunction.Match("ActualCosts", sr.value, 0)
TotalACS = Application.WorksheetFunction.Match("TotalACS", sr.value, 0)
ShowOverdueBar = Application.WorksheetFunction.Match("ShowOverdueBar", sr.value, 0)
ShowEstBaseBar = Application.WorksheetFunction.Match("ShowEstBaseBar", sr.value, 0)
ShowPercBar = Application.WorksheetFunction.Match("ShowPercBar", sr.value, 0)
ShowPercDataBar = Application.WorksheetFunction.Match("ShowPercDataBar", sr.value, 0)
ShowBaselineBar = Application.WorksheetFunction.Match("ShowBaselineBar", sr.value, 0)
ShowActualBar = Application.WorksheetFunction.Match("ShowActualBar", sr.value, 0)
CurrencySymbol = Application.WorksheetFunction.Match("CurrencySymbol", sr.value, 0)
RefreshTimeline = Application.WorksheetFunction.Match("RefreshTimeline", sr.value, 0)
TimelineVisible = Application.WorksheetFunction.Match("TimelineVisible", sr.value, 0)
CurrentView = Application.WorksheetFunction.Match("CurrentView", sr.value, 0)
ShowCompleted = Application.WorksheetFunction.Match("ShowCompleted", sr.value, 0)
ShowInProgress = Application.WorksheetFunction.Match("ShowInProgress", sr.value, 0)
ShowPlanned = Application.WorksheetFunction.Match("ShowPlanned", sr.value, 0)
ShowLate = Application.WorksheetFunction.Match("ShowLate", sr.value, 0)
WeekStartDay = Application.WorksheetFunction.Match("WeekStartDay", sr.value, 0)
WeekNumType = Application.WorksheetFunction.Match("WeekNumType", sr.value, 0)
tLicenseVal = Application.WorksheetFunction.Match("LV", sr.value, 0)
tUsrName = Application.WorksheetFunction.Match("peru", sr.value, 0)
tUsrEmailID = Application.WorksheetFunction.Match("chirun", sr.value, 0)
tUsrActivatedDate = Application.WorksheetFunction.Match("vidud", sr.value, 0)
tliky = Application.WorksheetFunction.Match("liky", sr.value, 0)
tLicDuration = Application.WorksheetFunction.Match("duli", sr.value, 0)
tFirstSavedDate = Application.WorksheetFunction.Match("FiSa", sr.value, 0)
tLiType = Application.WorksheetFunction.Match("LiType", sr.value, 0)
EnableCostsModule = Application.WorksheetFunction.Match("EC", sr.value, 0)
STG = Application.WorksheetFunction.Match("ShowTimelineGrid", sr.value, 0)
VerticalBorders = Application.WorksheetFunction.Match("VerticalBorders", sr.value, 0)
HiliteHolidays = Application.WorksheetFunction.Match("HiliteHolidays", sr.value, 0)
HiliteWorkOffDays = Application.WorksheetFunction.Match("HiliteWorkOffDays", sr.value, 0)
HiliteHolidaysPR = Application.WorksheetFunction.Match("HiliteHolidaysPR", sr.value, 0)
HiliteWorkOffDaysPR = Application.WorksheetFunction.Match("HiliteWorkOffDaysPR", sr.value, 0)
HideHolidays = Application.WorksheetFunction.Match("HideHolidays", sr.value, 0)
HideWorkOffDays = Application.WorksheetFunction.Match("HideWorkOffDays", sr.value, 0)
HideNonWorkingHours = Application.WorksheetFunction.Match("HideNonWorkingHours", sr.value, 0)
ShowGrouping = Application.WorksheetFunction.Match("ShowGrouping", sr.value, 0)
DateFormat = Application.WorksheetFunction.Match("DateFormat", sr.value, 0)
ShowTodayLines = Application.WorksheetFunction.Match("ShowTodayLines", sr.value, 0)
ShowDependencyConnector = Application.WorksheetFunction.Match("ShowDependencyConnector", sr.value, 0)
BarTextEnable = Application.WorksheetFunction.Match("EnableBarText", sr.value, 0)
BarTextCharacters = Application.WorksheetFunction.Match("BarTextCharacters", sr.value, 0)
BarTextFontSize = Application.WorksheetFunction.Match("BarTextFontSize", sr.value, 0)
BarTextDataColumnName = Application.WorksheetFunction.Match("BarTextDataColumnName", sr.value, 0)
BarTextIsBold = Application.WorksheetFunction.Match("BarTextIsBold", sr.value, 0)
PercentageEntryMode = Application.WorksheetFunction.Match("PercentageEntryMode", sr.value, 0)
PercentageCalculationType = Application.WorksheetFunction.Match("PercentageCalculationType", sr.value, 0)
ThemeName = Application.WorksheetFunction.Match("ThemeName", sr.value, 0)
SelectedTheme = Application.WorksheetFunction.Match("SelectedTheme", sr.value, 0)
EBC = Application.WorksheetFunction.Match("EBC", sr.value, 0)
EBaseC = Application.WorksheetFunction.Match("EBaseC", sr.value, 0)
EMC = Application.WorksheetFunction.Match("EMC", sr.value, 0)
PBC = Application.WorksheetFunction.Match("PBC", sr.value, 0)
PDC = Application.WorksheetFunction.Match("PDC", sr.value, 0)
BBC = Application.WorksheetFunction.Match("BBC", sr.value, 0)
ABC = Application.WorksheetFunction.Match("ABC", sr.value, 0)
OBC = Application.WorksheetFunction.Match("OBC", sr.value, 0)
DLC = Application.WorksheetFunction.Match("DLC", sr.value, 0)
TLC = Application.WorksheetFunction.Match("TLC", sr.value, 0)
TBC = Application.WorksheetFunction.Match("TBC", sr.value, 0)
GBC = Application.WorksheetFunction.Match("GBC", sr.value, 0)
PRC = Application.WorksheetFunction.Match("PRC", sr.value, 0)
HBC = Application.WorksheetFunction.Match("HBC", sr.value, 0)
hc = Application.WorksheetFunction.Match("HC", sr.value, 0)
WC = Application.WorksheetFunction.Match("WC", sr.value, 0)
HCPR = Application.WorksheetFunction.Match("HCPR", sr.value, 0)
WCPR = Application.WorksheetFunction.Match("WCPR", sr.value, 0)
CR1C = Application.WorksheetFunction.Match("CR1C", sr.value, 0)
CR12C = Application.WorksheetFunction.Match("CR12C", sr.value, 0)
CR2C = Application.WorksheetFunction.Match("CR2C", sr.value, 0)
CR3C = Application.WorksheetFunction.Match("CR3C", sr.value, 0)
TPCH = Application.WorksheetFunction.Match("TPCH", sr.value, 0)
TPCN = Application.WorksheetFunction.Match("TPCN", sr.value, 0)
TPCL = Application.WorksheetFunction.Match("TPCL", sr.value, 0)
TSCC = Application.WorksheetFunction.Match("TSCC", sr.value, 0)
TSCI = Application.WorksheetFunction.Match("TSCI", sr.value, 0)
TSCP = Application.WorksheetFunction.Match("TSCP", sr.value, 0)
TGB = Application.WorksheetFunction.Match("TGB", sr.value, 0)
gcType = Application.WorksheetFunction.Match("GCTYPE", sr.value, 0)
HCOL = Application.WorksheetFunction.Match("HCOL", sr.value, 0)
dCol = Application.WorksheetFunction.Match("DCOL", sr.value, 0)
WCOL = Application.WorksheetFunction.Match("WCOL", sr.value, 0)
MCOL = Application.WorksheetFunction.Match("MCOL", sr.value, 0)
QCOL = Application.WorksheetFunction.Match("QCOL", sr.value, 0)
HYCOL = Application.WorksheetFunction.Match("HYCOL", sr.value, 0)
YCOL = Application.WorksheetFunction.Match("YCOL", sr.value, 0)
HWID = Application.WorksheetFunction.Match("HWID", sr.value, 0)
DWID = Application.WorksheetFunction.Match("DWID", sr.value, 0)
WWID = Application.WorksheetFunction.Match("WWID", sr.value, 0)
MWID = Application.WorksheetFunction.Match("MWID", sr.value, 0)
QWID = Application.WorksheetFunction.Match("QWID", sr.value, 0)
HYWID = Application.WorksheetFunction.Match("HYWID", sr.value, 0)
YWID = Application.WorksheetFunction.Match("YWID", sr.value, 0)
RCCal = Application.WorksheetFunction.Match("RCCal", sr.value, 0)
PTRC = Application.WorksheetFunction.Match("PTRC", sr.value, 0)
TSD = Application.WorksheetFunction.Match("TSD", sr.value, 0)
TED = Application.WorksheetFunction.Match("TED", sr.value, 0)
csYear = Application.WorksheetFunction.Match("csYear", sr.value, 0)
ceYear = Application.WorksheetFunction.Match("ceYear", sr.value, 0)
csDate = Application.WorksheetFunction.Match("csDate", sr.value, 0)
ceDate = Application.WorksheetFunction.Match("ceDate", sr.value, 0)
CalcBED = Application.WorksheetFunction.Match("CalcBED", sr.value, 0)
CalcAED = Application.WorksheetFunction.Match("CalcAED", sr.value, 0)
ProjectName = Application.WorksheetFunction.Match("ProjectName", sr.value, 0)
PTDS = Application.WorksheetFunction.Match("PTDS", sr.value, 0)
SavedDate = Application.WorksheetFunction.Match("SavedDate", sr.value, 0)
LockWB = Application.WorksheetFunction.Match("LockWB", sr.value, 0)
Campaign = Application.WorksheetFunction.Match("Campaign", sr.value, 0)
Set sr = Nothing: Set ws = Nothing
End Sub
Option Explicit
Public TimelineStart As Long, TimelineEnd As Long, LLC As Long


Private Sub Class_Initialize()
Dim ws As Worksheet, t As Range, wsname As String
If ActiveSheet.Range("A1") <> "GEType" Then
MsgBox "getcoltimeline was called from a non gantt chart worksheet": Exit Sub
Else
Set ws = ActiveSheet: wsname = "CalColPos Timeline " & ws.Name
End If
Set t = ws.Range("1:1") 'Project sheet
TimelineStart = Application.WorksheetFunction.Match("TimelineStart", t.value, 0)
TimelineEnd = Application.WorksheetFunction.Match("TimelineEnd", t.value, 0)
LLC = Application.WorksheetFunction.Match("LLC", t.value, 0)
Set t = Nothing: Set ws = Nothing
End Sub
Option Explicit
Public ShowOverdueBar As Boolean, ShowEstBaseBar As Boolean, ShowPercBar As Boolean, ShowPercDataBar As Boolean, ShowBaselineBar As Boolean
Public ShowActualBar As Boolean, PercManual As Boolean, PercAuto As Boolean
Public ShowDepLines As Boolean, ShowTodayLines As Boolean, ShowTimelineGrid As Boolean, HideHolidays As Boolean, HideWorkOffDays As Boolean, HideNonWorkingHours As Boolean
Public HiliteHolidays As Boolean, HiliteWorkOffDays As Boolean, HiliteHolidaysPR As Boolean, HiliteWorkOffDaysPR As Boolean, EnableCosts As Boolean
Public ShowTextinBars As Boolean, TextBarIsBold As Boolean
Public CalBasDates As Boolean, CalActDates As Boolean, CalParCosts As Boolean, CalResCosts As Boolean
Public CalParPercSimple As Boolean, CalParPercWeight As Boolean, CalParDurSum As Boolean, HGC As Boolean
Public TextBarColumnName As String, CurrentView As String, DateFormat As String, CurrencyS As String
Public TextBarChars As Long, cEBC As Long, cPBC As Long, cPDC As Long, cEMC As Long, cABC As Long, cBBC As Long, cOBC As Long, TGBFS As Long, cTGB As Long
Public cEBase As Long, cDLC As Long, cTLC As Long, cPRC As Long, cTSCC As Long, cTSCI As Long, cTSCP As Long, cTPCH As Long, cTPCN As Long, cTPCL As Long
Public ShowGrouping As Boolean, ShowCompleted As Boolean, ShowPlanned As Boolean, ShowInProgress As Boolean
Public ShowRefreshTimeline As Boolean, ShowTimeline As Boolean, RefreshTimeline As Boolean, VerticalBorders As Boolean, LockWB As Boolean
Public TSD As Date, TED As Date, csDate As Date, ceDate As Date
Public HWID As Long, DWID As Long, WWID As Long, MWID As Long, QWID As Long, HYWID As Long, YWID As Long, csYear As Long, ceYear As Long
Public HCOL As Long, dCol As Long, WCOL As Long, MCOL As Long, QCOL As Long, HYCOL As Long, YCOL As Long
Public cHBC As Long, cHC As Long, cWC As Long, cHCPR As Long, cWCPR As Long, cCR1C As Long, cCR12C As Long, cCR2C As Long, cCR3C As Long, cTBC As Long, cGBC As Long

Private Sub Class_Initialize()
Dim ws As Worksheet
Set ws = setGSws(ActiveSheet)
ShowOverdueBar = ws.Cells(rowtwo, cps.ShowOverdueBar)
ShowEstBaseBar = ws.Cells(rowtwo, cps.ShowEstBaseBar)
ShowPercBar = ws.Cells(rowtwo, cps.ShowPercBar)
ShowPercDataBar = ws.Cells(rowtwo, cps.ShowPercDataBar)
ShowBaselineBar = ws.Cells(rowtwo, cps.ShowBaselineBar)
ShowActualBar = ws.Cells(rowtwo, cps.ShowActualBar)
ShowDepLines = ws.Cells(rowtwo, cps.ShowDependencyConnector)
ShowTodayLines = ws.Cells(rowtwo, cps.ShowTodayLines)
ShowTimelineGrid = ws.Cells(rowtwo, cps.STG)
VerticalBorders = ws.Cells(rowtwo, cps.VerticalBorders)
ShowTextinBars = ws.Cells(rowtwo, cps.BarTextEnable)
HiliteHolidays = ws.Cells(rowtwo, cps.HiliteHolidays)
HiliteWorkOffDays = ws.Cells(rowtwo, cps.HiliteWorkOffDays)
HiliteHolidaysPR = ws.Cells(rowtwo, cps.HiliteHolidaysPR)
HiliteWorkOffDaysPR = ws.Cells(rowtwo, cps.HiliteWorkOffDaysPR)
HideHolidays = ws.Cells(rowtwo, cps.HideHolidays)
HideWorkOffDays = ws.Cells(rowtwo, cps.HideWorkOffDays)
HideNonWorkingHours = ws.Cells(rowtwo, cps.HideNonWorkingHours)
EnableCosts = ws.Cells(rowtwo, cps.EnableCostsModule)
TextBarIsBold = ws.Cells(rowtwo, cps.BarTextIsBold)
CalBasDates = ws.Cells(rowtwo, cps.CalcBED)
CalActDates = ws.Cells(rowtwo, cps.CalcAED)
CalParCosts = ws.Cells(rowtwo, cps.PTRC)
CalResCosts = ws.Cells(rowtwo, cps.RCCal)
If LCase(ws.Cells(rowtwo, cps.PercentageEntryMode)) = "manual" Then PercManual = True: PercAuto = False Else PercManual = False: PercAuto = True
If LCase(ws.Cells(rowtwo, cps.PercentageCalculationType).value) = "simple" Then CalParPercSimple = True Else CalParPercSimple = False
If LCase(ws.Cells(rowtwo, cps.PercentageCalculationType).value) = "weighted" Then CalParPercWeight = True Else CalParPercWeight = False
If ws.Cells(rowtwo, cps.PTDS).value = True Then CalParDurSum = True Else CalParDurSum = False
If ws.Cells(rowtwo, cps.gcType).value = "s0n84" Then HGC = True Else HGC = False
ShowTimeline = ws.Cells(rowtwo, cps.TimelineVisible)
RefreshTimeline = ws.Cells(rowtwo, cps.RefreshTimeline)
If ShowTimeline = False Or RefreshTimeline = False Then ShowRefreshTimeline = False Else ShowRefreshTimeline = True
TextBarColumnName = ws.Cells(rowtwo, cps.BarTextDataColumnName)
CurrentView = ws.Cells(rowtwo, cps.CurrentView)
DateFormat = ws.Cells(rowtwo, cps.DateFormat).value
CurrencyS = ws.Cells(rowtwo, cps.CurrencySymbol).value
TSD = ws.Cells(rowtwo, cps.TSD).value
TED = ws.Cells(rowtwo, cps.TED).value
csYear = ws.Cells(rowtwo, cps.csYear).value
ceYear = ws.Cells(rowtwo, cps.ceYear).value
csDate = ws.Cells(rowtwo, cps.csDate).value
ceDate = ws.Cells(rowtwo, cps.ceDate).value
TextBarChars = ws.Cells(rowtwo, cps.BarTextCharacters)
ShowCompleted = ws.Cells(rowtwo, cps.ShowCompleted)
ShowPlanned = ws.Cells(rowtwo, cps.ShowPlanned)
ShowInProgress = ws.Cells(rowtwo, cps.ShowInProgress)
ShowGrouping = ws.Cells(rowtwo, cps.ShowGrouping)
TGBFS = CLng(ws.Cells(rowtwo, cps.BarTextFontSize))
cEBC = ws.Cells(trc, cps.EBC).Interior.Color
cPBC = ws.Cells(trc, cps.PBC).Interior.Color
cPDC = ws.Cells(trc, cps.PDC).Interior.Color
cEMC = ws.Cells(trc, cps.EMC).Interior.Color
cABC = ws.Cells(trc, cps.ABC).Interior.Color
cBBC = ws.Cells(trc, cps.BBC).Interior.Color
cOBC = ws.Cells(trc, cps.OBC).Interior.Color
cDLC = ws.Cells(trc, cps.DLC).Interior.Color
cTLC = ws.Cells(trc, cps.TLC).Interior.Color
cTGB = ws.Cells(trc, cps.TGB).Interior.Color
cEBase = ws.Cells(trc, cps.EBaseC).Interior.Color
cPRC = ws.Cells(trc, cps.PRC).Interior.Color
cTSCC = ws.Cells(trc, cps.TSCC).Interior.Color
cTSCI = ws.Cells(trc, cps.TSCI).Interior.Color
cTSCP = ws.Cells(trc, cps.TSCP).Interior.Color
cTPCH = ws.Cells(trc, cps.TPCH).Interior.Color
cTPCN = ws.Cells(trc, cps.TPCN).Interior.Color
cTPCL = ws.Cells(trc, cps.TPCL).Interior.Color
cHBC = ws.Cells(trc, cps.HBC).Interior.Color
cHC = ws.Cells(trc, cps.hc).Interior.Color
cWC = ws.Cells(trc, cps.WC).Interior.Color
cHCPR = ws.Cells(trc, cps.HCPR).Interior.Color
cWCPR = ws.Cells(trc, cps.WCPR).Interior.Color
cCR1C = ws.Cells(trc, cps.CR1C).Interior.Color
cCR12C = ws.Cells(trc, cps.CR12C).Interior.Color
cCR2C = ws.Cells(trc, cps.CR2C).Interior.Color
cCR3C = ws.Cells(trc, cps.CR3C).Interior.Color
cTBC = ws.Cells(trc, cps.TBC).Interior.Color
cGBC = ws.Cells(trc, cps.GBC).Interior.Color

HCOL = ws.Cells(rowtwo, cps.HCOL).value
dCol = ws.Cells(rowtwo, cps.dCol).value
WCOL = ws.Cells(rowtwo, cps.WCOL).value
MCOL = ws.Cells(rowtwo, cps.MCOL).value
QCOL = ws.Cells(rowtwo, cps.QCOL).value
HYCOL = ws.Cells(rowtwo, cps.HYCOL).value
YCOL = ws.Cells(rowtwo, cps.YCOL).value

HWID = ws.Cells(rowtwo, cps.HWID).value
DWID = ws.Cells(rowtwo, cps.DWID).value
WWID = ws.Cells(rowtwo, cps.WWID).value
MWID = ws.Cells(rowtwo, cps.MWID).value
QWID = ws.Cells(rowtwo, cps.QWID).value
HYWID = ws.Cells(rowtwo, cps.HYWID).value
YWID = ws.Cells(rowtwo, cps.YWID).value

LockWB = ws.Cells(rowtwo, cps.LockWB).value
Set ws = Nothing
End Sub
Option Explicit
Option Private Module
Private arrAllData()
Private allProjectsData()
Private arrAllDataSorted()
Private arrDashData()
Public Const cUserDashboardName As String = "Project Dashboard"

Sub CreateDashboard(Optional ProjectNo As Long)
Dim lrow As Long, findProjectRow As Long, i As Long, j As Long
Dim ws As Worksheet, UDS As Worksheet: Dim bDashExists As Boolean: Dim ProjectName As String
If Is2007 Then msg (79): Exit Sub
Call DA: ReDim arrProj(1 To GetGCCount, 1 To 2) As String: i = 1
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then arrProj(i, 1) = ws.Cells(rowsix, 10):arrProj(i, 2) = ws.Name: i = i + 1
Next ws
For i = 1 To UBound(arrProj())
For j = 1 To UBound(arrProj())
If i = j Then GoTo nexj
If arrProj(i, 1) = arrProj(j, 1) Then
MsgBox "Duplicate Project Name in worksheet - " & arrProj(i, 2) & ". Please check and correct the project names in all Gantt Charts and then refresh the dashboard again."
GoTo Last
End If
nexj:
Next j
Next i
For Each ws In ThisWorkbook.Sheets
If DashboardSheet(ws) Then bDashExists = True: Set UDS = ws: Exit For Else bDashExists = False
Next ws
If bDashExists = False Then
GDT.visible = True: GDT.Copy , Worksheets(1)
ActiveSheet.Name = cUserDashboardName: ActiveSheet.Range("A1") = "UserDashSheet"
Call hideWorksheet(GDT): Set UDS = ActiveSheet: UDS.Shapes("Notice4").Delete
End If
Call PopulateGDD
With Sheets(UDS.Name)
.visible = True:.Activate
End With
Application.ScreenUpdating = True: frmStatus.show: frmStatus.lblStatusMsg.Caption = "Creating Dashboard...": DoEvents: Call DA
Call PopulateUDSfromGDD: lrow = Application.WorksheetFunction.CountA(UDS.Range("P:P")):
Call DropDownToRange(UDS.Range("B2"), Range(Cells(3, "R"), Cells(lrow, "R"))) 'Call SetPivotFilters(PVS, UDS)
If ProjectNo <> 0 Then
findProjectRow = Application.WorksheetFunction.Match(ProjectNo, Range("P:P"), 0)
ProjectName = Cells(findProjectRow, "R"): Range("B2") = ProjectName
Else
If Range("B2") = "" Then Range("B2") = "Click here to Select a Project"
End If
'Call RefreshDashboard: Call RefreshRibbon
'last:
'Call EA:'Unload frmStatus:'ActiveWorkbook.RefreshAll:'End Sub
Call RefreshDashboard: 'ActiveWorkbook.RefreshAll
Last:
Unload frmStatus: Call RefreshRibbon: Call EA
End Sub

Sub RefreshDashboard(Optional t As Boolean)
Dim UDS As Worksheet: Dim GDDRowCount As Long, GDDColCount As Long
If ActiveSheet.Range("A1") <> "UserDashSheet" Then Exit Sub
If ActiveSheet.Range("A1") = "UserDashSheet" Then Set UDS = ActiveSheet
Dim lrow As Long, i As Long: lrow = Application.WorksheetFunction.CountA(UDS.Range("P:P"))
arrDashData = Range(Cells(1, "P"), Cells(lrow, "AG"))
If Range("B2") = "Click here to Select a Project" Then MsgBox "Please select a Project": Range("B2").Select: GoTo Last
For i = LBound(arrDashData()) To UBound(arrDashData())
If CStr(arrDashData(i, 3)) = CStr(UDS.Range("B2")) Then
UDS.Range("P2") = arrDashData(i, 1)
UDS.Range("S2") = arrDashData(i, 4)
UDS.Range("T2") = arrDashData(i, 5)
UDS.Range("U2") = arrDashData(i, 6)
UDS.Range("V2") = arrDashData(i, 7)
UDS.Range("W2") = arrDashData(i, 8)
UDS.Range("X2") = arrDashData(i, 9)
UDS.Range("Y2") = arrDashData(i, 10)
UDS.Range("Z2") = arrDashData(i, 11)
UDS.Range("AA2") = arrDashData(i, 12)
UDS.Range("AB2") = arrDashData(i, 13)
UDS.Range("AC2") = arrDashData(i, 14)
UDS.Range("AD2") = arrDashData(i, 15)
UDS.Range("AE2") = arrDashData(i, 16)
UDS.Range("AF2") = arrDashData(i, 17)
UDS.Range("AG2") = arrDashData(i, 18)
End If
Next i
GDDRowCount = Application.WorksheetFunction.CountA(GDD.Range("A:A")): GDDColCount = Application.WorksheetFunction.CountA(GDD.Range("1:1"))
If UDS.Range("P2") = 0 Then
GDD.Range(GDD.Cells(2, cpd.LC + 1), GDD.Cells(GDDRowCount, cpd.LC + 1)) = True
Else
GDD.ListObjects("DashboardTable").ListColumns(cpd.LC + 1).DataBodyRange.Formula = "=[@ProjectID]='" & UDS.Name & "'!$P$2" '
End If
Call SetPivotFilters(PVS, UDS): Call CreateChartColorRefreshButton: Call ColorTheCharts:
UDS.Cells(2, "T").NumberFormat = "0%": UDS.Range("AS2") = UDS.Range("Z2") & Format(UDS.Range("U2"), "#,##0.00"):
UDS.Range("AT2") = UDS.Range("Z2") & Format(UDS.Range("Y2"), "#,##0.00"): UDS.Range("S2").HorizontalAlignment = xlLeft
Last:
ActiveWorkbook.RefreshAll
End Sub

Sub refreshdashbutton()
Call CheckDupSettingsSheets: Call CreateDashboard
End Sub

Sub PopulateUDSfromGDD(Optional t As Boolean)
Dim UDS As Worksheet, ws As Worksheet:
Dim lrow As Long, i As Long, cRow As Long, TCCount As Long, TotalTasks As Long, CompletedTasks As Long, InprogressTasks As Long, PlannedTasks As Long, OverdueTasks As Long
Dim PercComp As Double, ACSTotal As Double, ECSTotal As Double, BCSTotal As Double, TotalBasBudget As Double, TotalEstBudget As Double
Dim minESD As Date, maxEED As Date
Set UDS = ActiveSheet: UDS.Range("P2:AG100").ClearContents:
lrow = 4: TotalBasBudget = 0: TotalEstBudget = 0: minESD = Date: maxEED = Date: TotalTasks = 0: CompletedTasks = 0:
InprogressTasks = 0: PlannedTasks = 0: OverdueTasks = 0
allProjectsData = GDD.ListObjects("DashboardTable").DataBodyRange.value

UDS.Cells(2, "P") = "S": UDS.Cells(2, "Q") = "Selected Project": UDS.Cells(2, "R") = "Selected Project":
UDS.Cells(3, "P") = 0: UDS.Cells(3, "Q") = "All Projects": UDS.Cells(3, "R") = "All Projects":
Dim wsgs As Worksheet, wsrs As Worksheet
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then
ws.Activate: Call CalcColPosGCT
Set wsgs = setGSws(ws)
UDS.Cells(lrow, "P") = getPID(ws)
UDS.Cells(lrow, "Q") = ws.Name
UDS.Cells(lrow, "R") = ws.Cells(rowsix, cpg.WBS)
If ws.Cells(rowseven, cpg.WBS) = "Project Lead: " Or ws.Cells(rowseven, cpg.WBS) = "Project Lead: Click to edit" Then
UDS.Cells(lrow, "S") = ""
Else
UDS.Cells(lrow, "S") = Replace(ws.Cells(rowseven, cpg.WBS), "Project Lead: ", vbNullString)
End If
UDS.Cells(lrow, "U") = Format(wsgs.Cells(rowtwo, cps.EstimatedBudget), "#,##0.00")
UDS.Cells(lrow, "V") = Format(wsgs.Cells(rowtwo, cps.BaselineBudget), "#,##0.00")
UDS.Cells(lrow, "Z") = wsgs.Cells(rowtwo, cps.CurrencySymbol)
UDS.Cells(lrow, "AA") = Int(WorksheetFunction.Min(Columns(cpg.ESD)))
UDS.Cells(lrow, "AB") = Int(WorksheetFunction.Max(Columns(cpg.EED)))
lrow = lrow + 1
End If
Next ws
UDS.Activate
lrow = Application.WorksheetFunction.CountA(UDS.Range("P:P"))

For cRow = 4 To lrow
For i = LBound(allProjectsData()) To UBound(allProjectsData())
If UDS.Cells(cRow, "P") = allProjectsData(i, cpd.ProjectID) Then
If allProjectsData(i, cpd.dType) = "T" Or allProjectsData(i, cpd.dType) = "C" Then
If allProjectsData(i, cpd.PercentageCompleted) = 1 Then CompletedTasks = CompletedTasks + 1
If allProjectsData(i, cpd.PercentageCompleted) > 0 And allProjectsData(i, cpd.PercentageCompleted) < 1 Then
If allProjectsData(i, cpd.EED) >= Date Then InprogressTasks = InprogressTasks + 1
End If
If allProjectsData(i, cpd.PercentageCompleted) = 0 And allProjectsData(i, cpd.EED) >= Date Then
PlannedTasks = PlannedTasks + 1
End If
If allProjectsData(i, cpd.PercentageCompleted) < 1 And allProjectsData(i, cpd.EED) < Date Then
OverdueTasks = OverdueTasks + 1
End If
PercComp = PercComp + allProjectsData(i, cpd.PercentageCompleted): TCCount = TCCount + 1
End If
If allProjectsData(i, cpd.dType) = "T" Or allProjectsData(i, cpd.dType) = "MP" Then
BCSTotal = BCSTotal + allProjectsData(i, cpd.BCS): ECSTotal = ECSTotal + allProjectsData(i, cpd.ECS): ACSTotal = ACSTotal + allProjectsData(i, cpd.ACS):
End If
If allProjectsData(i, cpd.ESD) < minESD Then minESD = allProjectsData(i, cpd.ESD)
If allProjectsData(i, cpd.EED) > maxEED Then maxEED = allProjectsData(i, cpd.EED)
End If
Next i

If PercComp > 0 Then UDS.Cells(cRow, "T") = PercComp / TCCount Else UDS.Cells(cRow, "T") = 0
UDS.Cells(cRow, "T").NumberFormat = "0%"
UDS.Cells(cRow, "W") = Format(BCSTotal, "#,##0.00"): UDS.Cells(cRow, "X") = Format(ECSTotal, "#,##0.00"): UDS.Cells(cRow, "Y") = Format(ACSTotal, "#,##0.00"):
UDS.Cells(cRow, "AC") = TCCount: UDS.Cells(cRow, "AD") = CompletedTasks: UDS.Cells(cRow, "AE") = InprogressTasks
UDS.Cells(cRow, "AF") = PlannedTasks: UDS.Cells(cRow, "AG") = OverdueTasks

PercComp = 0: TCCount = 0: BCSTotal = 0: ECSTotal = 0: ACSTotal = 0: minESD = Date: maxEED = Date: TotalTasks = 0: CompletedTasks = 0:
InprogressTasks = 0: PlannedTasks = 0: OverdueTasks = 0:
Next cRow
'Master
For cRow = 4 To lrow
PercComp = PercComp + Cells(cRow, "T")
TotalEstBudget = TotalEstBudget + UDS.Cells(cRow, "U")
TotalBasBudget = TotalBasBudget + UDS.Cells(cRow, "V")
BCSTotal = BCSTotal + UDS.Cells(cRow, "W")
ECSTotal = ECSTotal + UDS.Cells(cRow, "X")
ACSTotal = ACSTotal + UDS.Cells(cRow, "Y")
If cRow = 4 Then minESD = UDS.Cells(4, "AA"): maxEED = UDS.Cells(4, "AB")

If cRow <> 4 And UDS.Cells(cRow, "AA") <= minESD Then minESD = UDS.Cells(cRow, "AA")
If cRow <> 4 And UDS.Cells(cRow, "AB") >= maxEED Then maxEED = UDS.Cells(cRow, "AB")
TotalTasks = TotalTasks + UDS.Cells(cRow, "AC")
CompletedTasks = CompletedTasks + UDS.Cells(cRow, "AD")
InprogressTasks = InprogressTasks + UDS.Cells(cRow, "AE")
PlannedTasks = PlannedTasks + UDS.Cells(cRow, "AF")
OverdueTasks = OverdueTasks + UDS.Cells(cRow, "AG")
Next cRow

If PercComp > 0 Then UDS.Cells(3, "T") = PercComp / (lrow - 3) Else UDS.Cells(3, "T") = 0
UDS.Cells(3, "T").NumberFormat = "0%"
UDS.Cells(3, "U") = TotalEstBudget: UDS.Cells(3, "V") = TotalBasBudget
UDS.Cells(3, "W") = Format(BCSTotal, "#,##0.00"): UDS.Cells(3, "X") = Format(ECSTotal, "#,##0.00"): UDS.Cells(3, "Y") = Format(ACSTotal, "#,##0.00"):
UDS.Cells(3, "Z") = UDS.Cells(4, "Z")
UDS.Cells(3, "AA") = minESD:
UDS.Cells(3, "AB") = maxEED:
UDS.Cells(3, "AC") = TotalTasks: UDS.Cells(3, "AD") = CompletedTasks:
UDS.Cells(3, "AE") = InprogressTasks: UDS.Cells(3, "AF") = PlannedTasks: UDS.Cells(3, "AG") = OverdueTasks
PercComp = 0: TCCount = 0: BCSTotal = 0: ECSTotal = 0: ACSTotal = 0: minESD = Date: maxEED = Date: TotalTasks = 0: CompletedTasks = 0:
End Sub

Sub DropDownToRange(rngTarget As Range, rngSource As Range)
ActiveSheet.Range("B2").Activate'Delete & Add Validation in Target Range
With rngTarget.Validation
.Delete
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
xlBetween, Formula1:="=" & rngSource.Address
.IgnoreBlank = True
.InCellDropdown = True
.InputTitle = ""
.ErrorTitle = ""
.InputMessage = ""
.ErrorMessage = ""
.ShowInput = True
.ShowError = True
End With
End Sub

Sub PopulateGDD()
GDD.Cells.Clear: GDD.Activate
Call SetupGCColumns("GDD"): Call CalcColPosGDD
Dim i As Long, j As Long, lrow As Long, GDDRowCount As Long, GDDColCount As Long, PID As Long, NTR As Long: NTR = 9 ' NonTaskRows
Dim tbl As ListObject: Dim rng As Range, GDDHeaderRange As Range: Dim ws As Worksheet, UDS As Worksheet
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then
ws.Activate: lrow = GetLastRow(ws): Call CalcColPosGCT: PID = getPID(ws)
arrAllData = ws.Range(ws.Cells(1, 1), ws.Cells(lrow, cpg.LC)).value
If Cells(firsttaskrow, cpg.Task) = sAddTaskPlaceHolder Then GoTo nexws
ReDim arrAllDataSorted(1 To lrow - NTR, 1 To cpd.LC):
For i = 10 To lrow
For j = 1 To cpg.LC
arrAllDataSorted(i - NTR, cpd.ProjectID) = PID
arrAllDataSorted(i - NTR, cpd.dType) = "dType"
arrAllDataSorted(i - NTR, cpd.GEtype) = arrAllData(i, cpg.GEtype)
arrAllDataSorted(i - NTR, cpd.TID) = arrAllData(i, cpg.TID)
arrAllDataSorted(i - NTR, cpd.Dependency) = arrAllData(i, cpg.Dependency)
arrAllDataSorted(i - NTR, cpd.Dependents) = arrAllData(i, cpg.Dependents)
arrAllDataSorted(i - NTR, cpd.StartConstrain) = arrAllData(i, cpg.StartConstrain)
arrAllDataSorted(i - NTR, cpd.EndConstrain) = arrAllData(i, cpg.EndConstrain)
arrAllDataSorted(i - NTR, cpd.TIL) = arrAllData(i, cpg.TIL)
arrAllDataSorted(i - NTR, cpd.SS) = arrAllData(i, cpg.SS)
arrAllDataSorted(i - NTR, cpd.TaskIcon) = arrAllData(i, cpg.TaskIcon)
arrAllDataSorted(i - NTR, cpd.WBS) = arrAllData(i, cpg.WBS)
arrAllDataSorted(i - NTR, cpd.Task) = arrAllData(i, cpg.Task)
arrAllDataSorted(i - NTR, cpd.Priority) = arrAllData(i, cpg.Priority)
arrAllDataSorted(i - NTR, cpd.Status) = arrAllData(i, cpg.Status)
arrAllDataSorted(i - NTR, cpd.Resource) = arrAllData(i, cpg.Resource)
arrAllDataSorted(i - NTR, cpd.ResourceCost) = arrAllData(i, cpg.ResourceCost)
arrAllDataSorted(i - NTR, cpd.BSD) = arrAllData(i, cpg.BSD)
arrAllDataSorted(i - NTR, cpd.BED) = arrAllData(i, cpg.BED)
arrAllDataSorted(i - NTR, cpd.BD) = arrAllData(i, cpg.BD)
arrAllDataSorted(i - NTR, cpd.ESD) = arrAllData(i, cpg.ESD)
arrAllDataSorted(i - NTR, cpd.EED) = arrAllData(i, cpg.EED)
arrAllDataSorted(i - NTR, cpd.ED) = arrAllData(i, cpg.ED)
arrAllDataSorted(i - NTR, cpd.Done) = arrAllData(i, cpg.Done)
arrAllDataSorted(i - NTR, cpd.PercentageCompleted) = arrAllData(i, cpg.PercentageCompleted)
arrAllDataSorted(i - NTR, cpd.ASD) = arrAllData(i, cpg.ASD)
arrAllDataSorted(i - NTR, cpd.AED) = arrAllData(i, cpg.AED)
arrAllDataSorted(i - NTR, cpd.AD) = arrAllData(i, cpg.AD)
arrAllDataSorted(i - NTR, cpd.BCS) = arrAllData(i, cpg.BCS)
arrAllDataSorted(i - NTR, cpd.ECS) = arrAllData(i, cpg.ECS)
arrAllDataSorted(i - NTR, cpd.ACS) = arrAllData(i, cpg.ACS)
arrAllDataSorted(i - NTR, cpd.Notes) = arrAllData(i, cpg.Notes)
arrAllDataSorted(i - NTR, cpd.Custom1) = arrAllData(i, cpg.Custom1)
arrAllDataSorted(i - NTR, cpd.Custom2) = arrAllData(i, cpg.Custom2)
arrAllDataSorted(i - NTR, cpd.Custom3) = arrAllData(i, cpg.Custom3)
arrAllDataSorted(i - NTR, cpd.Custom4) = arrAllData(i, cpg.Custom4)
arrAllDataSorted(i - NTR, cpd.Custom5) = arrAllData(i, cpg.Custom5)
arrAllDataSorted(i - NTR, cpd.Custom6) = arrAllData(i, cpg.Custom6)
arrAllDataSorted(i - NTR, cpd.Custom7) = arrAllData(i, cpg.Custom7)
arrAllDataSorted(i - NTR, cpd.Custom8) = arrAllData(i, cpg.Custom8)
arrAllDataSorted(i - NTR, cpd.Custom9) = arrAllData(i, cpg.Custom9)
arrAllDataSorted(i - NTR, cpd.Custom10) = arrAllData(i, cpg.Custom10)
arrAllDataSorted(i - NTR, cpd.Custom11) = arrAllData(i, cpg.Custom11)
arrAllDataSorted(i - NTR, cpd.Custom12) = arrAllData(i, cpg.Custom12)
arrAllDataSorted(i - NTR, cpd.Custom13) = arrAllData(i, cpg.Custom13)
arrAllDataSorted(i - NTR, cpd.Custom14) = arrAllData(i, cpg.Custom14)
arrAllDataSorted(i - NTR, cpd.Custom15) = arrAllData(i, cpg.Custom15)
arrAllDataSorted(i - NTR, cpd.Custom16) = arrAllData(i, cpg.Custom16)
arrAllDataSorted(i - NTR, cpd.Custom17) = arrAllData(i, cpg.Custom17)
arrAllDataSorted(i - NTR, cpd.Custom18) = arrAllData(i, cpg.Custom18)
arrAllDataSorted(i - NTR, cpd.Custom19) = arrAllData(i, cpg.Custom19)
arrAllDataSorted(i - NTR, cpd.Custom20) = arrAllData(i, cpg.Custom20)
arrAllDataSorted(i - NTR, cpd.LC) = arrAllData(i, cpg.LC)
Next j
Next i

For i = 1 To lrow - NTR
If arrAllDataSorted(i, cpd.PercentageCompleted) = 1 Then arrAllDataSorted(i, cpd.Status) = "Completed"
If arrAllDataSorted(i, cpd.PercentageCompleted) > 0 And arrAllDataSorted(i, cpd.PercentageCompleted) < 1 Then
If arrAllDataSorted(i, cpd.EED) >= Date Then arrAllDataSorted(i, cpd.Status) = "InProgress"
End If
If arrAllDataSorted(i, cpd.PercentageCompleted) = 0 And arrAllDataSorted(i, cpd.EED) >= Date Then
arrAllDataSorted(i, cpd.Status) = "Planned"
End If
If arrAllDataSorted(i, cpd.PercentageCompleted) < 1 And arrAllDataSorted(i, cpd.EED) < Date Then
arrAllDataSorted(i, cpd.Status) = "Overdue"
End If
If i = lrow - NTR Then
If arrAllDataSorted(i, cpd.GEtype) = "M" Then arrAllDataSorted(i, cpd.dType) = "M"
If arrAllDataSorted(i, cpd.GEtype) = "T" Then
If arrAllDataSorted(i, cpd.TIL) = 0 Then arrAllDataSorted(i, cpd.dType) = "T" Else arrAllDataSorted(i, cpd.dType) = "C"
End If
Else
If arrAllDataSorted(i, cpd.GEtype) = "M" Then arrAllDataSorted(i, cpd.dType) = "M"
If arrAllDataSorted(i, cpd.GEtype) = "T" Then
If arrAllDataSorted(i, cpd.TIL) = 0 And arrAllDataSorted(i + 1, cpd.TIL) = 1 Then
arrAllDataSorted(i, cpd.dType) = "MP"
ElseIf arrAllDataSorted(i, cpd.TIL) > 0 Then
If arrAllDataSorted(i + 1, cpd.TIL) > arrAllDataSorted(i, cpd.TIL) Then
arrAllDataSorted(i, cpd.dType) = "SP"
Else
arrAllDataSorted(i, cpd.dType) = "C"
End If
ElseIf arrAllDataSorted(i, cpd.TIL) = 0 And arrAllDataSorted(i + 1, cpd.TIL) = 0 Then
arrAllDataSorted(i, cpd.dType) = "T"
End If
End If
End If
Next i
GDDRowCount = Application.WorksheetFunction.CountA(GDD.Range("A:A")): GDDColCount = Application.WorksheetFunction.CountA(GDD.Range("1:1"))
GDD.Range(GDD.Cells(GDDRowCount + 1, 1), GDD.Cells(GDDRowCount + (lrow - NTR), GDDColCount)) = arrAllDataSorted
End If
nexws:
Next ws

GDDRowCount = Application.WorksheetFunction.CountA(GDD.Range("A:A")): GDDColCount = Application.WorksheetFunction.CountA(GDD.Range("1:1"))
Set rng = GDD.Range(GDD.Cells(1, 1), GDD.Cells(GDDRowCount, GDDColCount)): Set tbl = GDD.ListObjects.Add(xlSrcRange, rng, , xlYes)
tbl.Name = "DashboardTable": tbl.TableStyle = "TableStyleMedium15": tbl.ListColumns.Add(cpd.LC + 1).Name = "ProjectFilter"


For Each ws In ThisWorkbook.Sheets
If ws.Range("A1") = "UserDashSheet" Then Set UDS = ws: Exit For Else Exit Sub
Next ws
If UDS.Range("P2") = 0 Then
GDD.Range(GDD.Cells(2, cpd.LC + 1), GDD.Cells(GDDRowCount, cpd.LC + 1)) = True
Else
tbl.ListColumns(cpd.LC + 1).DataBodyRange.Formula = "=[@ProjectID]='" & UDS.Name & "'!$P$2" '
End If
End Sub

Sub SetPivotFilters(ws As Worksheet, UDS As Worksheet)
Dim Pivot As PivotTable
For Each Pivot In PVS.PivotTables
Pivot.RefreshTable
Pivot.Update
Next
On Error Resume Next
Dim pvt As PivotTable
Set pvt = ws.PivotTables("PT1"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With
Set pvt = ws.PivotTables("PT2"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With
Set pvt = ws.PivotTables("PT3"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)":.PivotItems("C").visible = False:.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("MP").visible = True: .PivotItems("T").visible = True
End With
pvt.PivotFields("Actual Cost").NumberFormat = UDS.Range("Z2") & "#,##0.00"
pvt.PivotFields("Estimated Cost").NumberFormat = UDS.Range("Z2") & "#,##0.00"
pvt.PivotFields("Baseline Cost").NumberFormat = UDS.Range("Z2") & "#,##0.00"
Set pvt = ws.PivotTables("PT4"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("GEType")
.CurrentPage = "(All)": .PivotItems("T").visible = False: .PivotItems("M").visible = True
End With
Set pvt = ws.PivotTables("PT5"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With
Set pvt = ws.PivotTables("PT6"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With

Set pvt = ws.PivotTables("PT7"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With
Set pvt = ws.PivotTables("PT8"):pvt.PivotCache.Refresh
With pvt.PivotFields("ProjectFilter")
.PivotItems("True").visible = True: .PivotItems("False").visible = False
End With
With pvt.PivotFields("dType")
.CurrentPage = "(All)": .PivotItems("MP").visible = False
.PivotItems("SP").visible = False: .PivotItems("M").visible = False
.PivotItems("T").visible = True: .PivotItems("C").visible = True
End With
On Error GoTo 0
End Sub

Sub ColorTheCharts()
If Is2007 Then Exit Sub
Dim UDS As Worksheet
If ActiveSheet.Range("A1") <> "UserDashSheet" Then Exit Sub
If ActiveSheet.Range("A1") = "UserDashSheet" Then Set UDS = ActiveSheet
Dim XVals As Variant: Dim i As Long
UDS.ChartObjects("chTaskStatus").Activate
If UDS.ChartObjects("chTaskStatus").Chart.SeriesCollection.Count > 0 Then
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
With ActiveChart.FullSeriesCollection(1).DataLabels
.ShowCategoryName = True: .ShowPercentage = True
With .Format.Fill
.visible = msoTrue
.ForeColor.ObjectThemeColor = msoThemeColorBackground1
.ForeColor.TintAndShade = 0
.ForeColor.Brightness = 0
.Transparency = 0
.Solid
End With
End With
With UDS.ChartObjects("chTaskStatus").Chart.SeriesCollection(1)
XVals = .XValues
For i = LBound(XVals) To UBound(XVals)
Select Case LCase(XVals(i))
Case "completed"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AD1").Interior.Color
Case "inprogress"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AE1").Interior.Color
Case "planned"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AF1").Interior.Color
Case "overdue"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AG1").Interior.Color
End Select
Next i
End With
End If
UDS.ChartObjects("chTaskPriority").Activate
If UDS.ChartObjects("chTaskPriority").Chart.SeriesCollection.Count > 0 Then
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
With ActiveChart.FullSeriesCollection(1).DataLabels
.ShowCategoryName = True: .ShowPercentage = True
With .Format.Fill
.visible = msoTrue
.ForeColor.ObjectThemeColor = msoThemeColorBackground1
.ForeColor.TintAndShade = 0
.ForeColor.Brightness = 0
.Transparency = 0
.Solid
End With
End With
With UDS.ChartObjects("chTaskPriority").Chart.SeriesCollection(1)
XVals = .XValues
For i = LBound(XVals) To UBound(XVals)
Select Case LCase(XVals(i))
Case "low"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AH1").Interior.Color
Case "normal"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AI1").Interior.Color
Case "high"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("AJ1").Interior.Color
End Select
Next i
End With
End If
UDS.ChartObjects("chTaskCosts").Activate

If UDS.ChartObjects("chTaskCosts").Chart.SeriesCollection.Count > 0 Then
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
With ActiveChart.FullSeriesCollection(1).DataLabels
.ShowCategoryName = False: '.ShowPercentage = True
With .Format.Fill
.visible = msoTrue
.ForeColor.ObjectThemeColor = msoThemeColorBackground1
.ForeColor.TintAndShade = 0
.ForeColor.Brightness = 0
.Transparency = 0
.Solid
End With
End With
With UDS.ChartObjects("chTaskCosts").Chart.SeriesCollection(1)
XVals = .XValues
For i = LBound(XVals) To UBound(XVals)
Select Case LCase(XVals(i))
Case "baseline cost"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("W1").Interior.Color
Case "estimated cost"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("X1").Interior.Color
Case "actual cost"
.Points(i).Format.Fill.ForeColor.RGB = UDS.Range("Y1").Interior.Color
End Select
Next i
End With
End If
UDS.Range("B2").Select
End Sub

Sub CreateChartColorRefreshButton()
Dim UDS As Worksheet: Dim s As Shape: Dim leftfrom As Double, topfrom As Double
If ActiveSheet.Range("A1") <> "UserDashSheet" Then Exit Sub
If ActiveSheet.Range("A1") = "UserDashSheet" Then Set UDS = ActiveSheet
Call DeleteShape("ChartColorRefreshButton", 23)
With UDS.Cells(rowtwo, 13)
leftfrom = .Left + 20:topfrom = .Top + 4
End With
Set s = UDS.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=leftfrom, Top:=topfrom, Width:=10, Height:=10)
With s
.Name = "ChartColorRefreshButton": .Line.visible = msoFalse:
.OnAction = "ColorTheChartsClicked"
With s.Fill
.Solid: .ForeColor.RGB = vbRed
End With
End With
End Sub

Sub ColorTheChartsClicked()
Call DA: Call ColorTheCharts: Call EA
End Sub
Option Explicit
Public tIDList As New Collection
Private CheckedList()
Private p As Long

Sub CalcDepFormulas(Optional bReCalcDepFormulas As Boolean)
If CheckForDependency(ActiveSheet) = False Then Exit Sub
Application.EnableEvents = False: ActiveSheet.Calculate: bStopCalculationOfConstraints = True 'dont turn disable all - different
Call UpdateNewDepCols
Dim vDPDs, vDPD
Dim td As Long, d As Long, cRow As Long, lag As Long, i As Long, tidr As Long, lrow As Long
Dim sWorkDaysAddress As String, fStart As String, fEnd As String:
Dim arrResA(): Dim arrESDA(): Dim arrEEDA():
lrow = GetLastRow
Dim arrAllV()
Dim arrTIDV()
arrAllV = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value:
arrTIDV = Range(Cells(firsttaskrow, cpg.TID), Cells(lrow, cpg.TID)).value:
Dim arrConstrains():
ReDim arrConstrains(firsttaskrow To lrow, 1 To 2) 'arrConstrains = Range(Cells(firsttaskrow, cpg.StartConstrain), Cells(lrow, cpg.EndConstrain)).value:
ReDim arrResA(firsttaskrow To lrow): ReDim arrESDA(firsttaskrow To lrow): ReDim arrEEDA(firsttaskrow To lrow)
If ResArraysReady = False Then Call RememberResArrays
For cRow = firsttaskrow To lrow
arrResA(cRow) = Cells(cRow, cpg.Resource).Address: arrESDA(cRow) = Cells(cRow, cpg.ESD).Address: arrEEDA(cRow) = Cells(cRow, cpg.EED).Address
Next
For cRow = firsttaskrow To lrow
If arrAllV(cRow, cpg.Dependency) = "" Then GoTo nexa

Call getResValue(arrAllV(cRow, cpg.Resource), newResourcesArray)
sWorkDaysAddress = "," & sArr.WorkdaysP(resvalue, 7) & "," & _
sArr.WorkdaysP(resvalue, 1) & "," & _
sArr.WorkdaysP(resvalue, 2) & "," & _
sArr.WorkdaysP(resvalue, 3) & "," & _
sArr.WorkdaysP(resvalue, 4) & "," & _
sArr.WorkdaysP(resvalue, 5) & "," & _
sArr.WorkdaysP(resvalue, 6) & ","

vDPDs = Split(arrAllV(cRow, cpg.Dependency), DepSeperator)
For td = 0 To UBound(vDPDs) - 1
vDPD = Split(vDPDs(td), "_"): tidr = getTIDRow(CLng(vDPD(0))): lag = CInt(vDPD(2))
If st.HGC Then
Select Case vDPD(1)
Case Is = "FS"
If lag >= 0 Then
fStart = fStart & "CalEEDHrsDep(" & arrResA(cRow) & "," & arrEEDA(tidr) & "," & lag & ")" & ","
Else
fStart = fStart & "CalESDHrs(" & arrResA(cRow) & "," & arrEEDA(tidr) & "," & -lag & ")" & ","
End If
Case Is = "SS"
If lag >= 0 Then
fStart = fStart & "CalEEDHrsDep(" & arrResA(cRow) & "," & arrESDA(tidr) & "," & lag & ")" & ","
Else
fStart = fStart & "CalESDHrs(" & arrResA(cRow) & "," & arrESDA(tidr) & "," & -lag & ")" & ","
End If
Case Is = "SF"
If lag >= 0 Then
fEnd = fEnd & "CalESDHrs(" & arrResA(cRow) & "," & arrESDA(tidr) & "," & -lag & ")" & ","
Else
fEnd = fEnd & "CalESDHrs(" & arrResA(cRow) & "," & arrESDA(tidr) & "," & -lag & ")" & ","
End If
Case Is = "FF"
If lag >= 0 Then
fEnd = fEnd & "CalESDHrs(" & arrResA(cRow) & "," & arrEEDA(tidr) & "," & -lag & ")" & ","
Else
fEnd = fEnd & "CalESDHrs(" & arrResA(cRow) & "," & arrEEDA(tidr) & "," & -lag & ")" & ","
End If
End Select
Else
Select Case vDPD(1)
Case Is = "FS"
fStart = fStart & "GetDelayedDate(" & arrEEDA(tidr) & "+1," & AddDelayOperator(CStr(vDPD(2))) & sWorkDaysAddress & Chr(34) & "FS" & Chr(34) & "," & resvalue & ")" & ","
Case Is = "SS"
fStart = fStart & "GetDelayedDate(" & arrESDA(tidr) & "," & AddDelayOperator(CStr(vDPD(2))) & sWorkDaysAddress & Chr(34) & "SS" & Chr(34) & "," & resvalue & ")" & ","
Case Is = "SF"
fEnd = fEnd & "GetDelayedDate(" & arrESDA(tidr) & "-1" & "," & AddDelayOperator(CStr(vDPD(2))) & sWorkDaysAddress & Chr(34) & "SF" & Chr(34) & "," & resvalue & ")" & ","
Case Is = "FF"
fEnd = fEnd & "GetDelayedDate(" & arrEEDA(tidr) & "," & AddDelayOperator(CStr(vDPD(2))) & sWorkDaysAddress & Chr(34) & "FF" & Chr(34) & "," & resvalue & ")" & ","
End Select
End If
Next td
If fStart <> vbNullString Then
arrConstrains(cRow, 1) = "=max(" & Left(fStart, Len(fStart) - 1) & ")"
Else
arrConstrains(cRow, 1) = vbNullString
End If
If fEnd <> vbNullString Then
arrConstrains(cRow, 2) = "=max(" & Left(fEnd, Len(fEnd) - 1) & ")"
Else
arrConstrains(cRow, 2) = vbNullString
End If
fStart = vbNullString: fEnd = vbNullString
nexa:
Next cRow
If ActiveSheet.AutoFilterMode Then
Call ReApplyAutoFilter(arrConstrains, lrow, "CalcDepFormulas")
Else
Call ArrayToRange(arrConstrains, lrow, "CalcDepFormulas") 'Range(Cells(1, 1), Cells(lrow, cpg.lc)) = arrAllV
End If
bStopCalculationOfConstraints = False
End Sub

Sub ReCalcDepFormulas(cRow As Long, Optional bCheckDependentTasks As Boolean)
'This will populate ESD and EED by starting from the lowest dependency to the highest
Dim resname As String: Dim bConstrainsChanged As Boolean: Dim vStr As Variant, dStr As Variant: Dim tidr As Long, i As Long 'Dim tRng As Range, fRng As Range:
resname = Cells(cRow, cpg.Resource).value:ActiveSheet.Calculate'Application.CalculateFull
If Cells(cRow, cpg.StartConstrain) <> vbNullString And Cells(cRow, cpg.EndConstrain) <> vbNullString Then 'Both constraints present -
If Cells(cRow, cpg.ESD) <> Cells(cRow, cpg.StartConstrain) Then Cells(cRow, cpg.ESD) = Cells(cRow, cpg.StartConstrain): bConstrainsChanged = True
If Cells(cRow, cpg.EED) <> Cells(cRow, cpg.EndConstrain) Then Cells(cRow, cpg.EED) = Cells(cRow, cpg.EndConstrain): bConstrainsChanged = True
If Cells(cRow, cpg.EED) < Cells(cRow, cpg.ESD) Then
If st.HGC Then
Cells(cRow, cpg.EED) = CalEEDHrs(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
Else
Cells(cRow, cpg.EED) = GetEndDateFromWorkDays(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
End If
bConstrainsChanged = True
Else
If st.HGC Then
Cells(cRow, cpg.ED) = CalEDHrs(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.EED))
Else
Cells(cRow, cpg.ED) = GetWorkDaysFromDate(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.EED))
End If
End If
ElseIf Cells(cRow, cpg.StartConstrain) <> vbNullString And Cells(cRow, cpg.EndConstrain) = vbNullString Then 'START CONSTRAIN PRESENT
If Cells(cRow, cpg.ESD) <> Cells(cRow, cpg.StartConstrain) Then
Cells(cRow, cpg.ESD) = Cells(cRow, cpg.StartConstrain)
If st.HGC Then
Cells(cRow, cpg.EED) = CalEEDHrs(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
Else
Cells(cRow, cpg.EED) = GetEndDateFromWorkDays(resname, Cells(cRow, cpg.ESD), Cells(cRow, cpg.ED))
End If
bConstrainsChanged = True
End If
ElseIf Cells(cRow, cpg.StartConstrain) = vbNullString And Cells(cRow, cpg.EndConstrain) <> vbNullString Then 'END CONSTRAIN PRESENT
If Cells(cRow, cpg.EED) <> Cells(cRow, cpg.EndConstrain) Then
Cells(cRow, cpg.EED) = Cells(cRow, cpg.EndConstrain)
If st.HGC Then
Cells(cRow, cpg.ESD) = CalESDHrs(resname, Cells(cRow, cpg.EED), Cells(cRow, cpg.ED))
Else
Cells(cRow, cpg.ESD) = GetStartFromWorkDays(resname, Cells(cRow, cpg.EED), Cells(cRow, cpg.ED))
End If
bConstrainsChanged = True
End If
End If
If Cells(cRow, cpg.Dependents) = vbNullString Then GoTo Last
If bConstrainsChanged = False And bCheckDependentTasks = False Then GoTo Last
'Set tRng = Range(Cells(rownine + 1, cpg.TID), Cells(Cells.Rows.Count, cpg.TID)):
vStr = Split(Cells(cRow, cpg.Dependents), DepSeperator)
For i = LBound(vStr) To UBound(vStr) - 1
 Deltax
Set fRng = tRng.Find(CLng(vstr(i)), , xlFormulas, xlWhole)
tidr = getTIDRow(CLng(vStr(i)))
If Not fRng Is Nothing Then
Call ReCalcDepFormulas(fRng.Row, False)
dStr = dStr & vstr(i) & DepSeperator
Else
Debug.Print "ReCalcDepFormulas" & CLng(vstr(i))
End If
If tidr > 0 Then
Call ReCalcDepFormulas(tidr, False)
dStr = dStr & vStr(i) & DepSeperator
Else
Debug.Print "ReCalcDepFormulas" & CLng(vstr(i))
End If
Next
If dStr <> vbNullString Then Cells(cRow, cpg.Dependents) = dStr
dStr = vbNullString
Last:
End Sub

Public Sub PopParentTasks(Optional CalOnlyRowNo As Long, Optional ctype As String)
If checkSheetError Then Exit Sub
CalcColPosGCT: CalcColPosTimeline
Dim bECS As Boolean, bACS As Boolean, bBCS As Boolean, bRCS As Boolean, bChildFound As Boolean: 'bbRecalculateConstrainValues As Boolean
Dim ECS As Double, ACS As Double, BCS As Double, RCS As Double, PercentComplete As Double, tPercentComplete As Double:
Dim ESD As Date, EED As Date, BSD As Date, BED As Date, ASD As Date, AED As Date
Dim ED As Long, cTaskCount As Long, ChildTasksDuration As Long, sRow As Long, fRow As Long, cRow As Long, lrow As Long, lastrow As Long
Dim sumOfED As Long, sumOfWork As Long, sumOfBD As Long, sumOfAD As Long, cLevel As Long
Dim arrAllV(): Dim arrAllF(): Dim arrLvlTIL() As Integer
If CalOnlyRowNo > 0 Then fRow = GetLastRowOfFamily(CalOnlyRowNo) Else fRow = GetLastRow
If fRow <= rownine Then Exit Sub
lrow = fRow
ReDim arrLvlTIL(1 To fRow)
ReDim arrAllV(1 To lrow, 1 To cpg.LC): arrAllV = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
ReDim arrAllF(1 To lrow, 1 To cpg.LC): arrAllF = Range(Cells(1, 1), Cells(lrow, cpg.LC)).Formula
lastrow = fRow:
If ctype = "" Then ctype = allFields
'If CheckForDependency(ActiveSheet) Then Call CalcDepFormulas 'is this really required XX 3.76 onwards
tlog "PopParentTasks: " & CalOnlyRowNo & ctype

Do Until fRow <= rownine
sRow = fRow'Take each block and then calculate from max level to min level
If arrAllV(fRow, cpg.TIL) > 0 Then
cLevel = 0
On Error Resume Next
sRow = 0
sRow = WorksheetFunction.Match(Left(arrAllV(fRow, cpg.WBS), InStr(1, arrAllV(fRow, cpg.WBS), ".") - 1), Range(Cells(firsttaskrow, cpg.WBS), Cells(fRow, cpg.WBS)), 0)
On Error GoTo 0
If sRow = 0 Then MsgBox "Task Indent Issue: Please send your file to support@ganttexcel.com to fix the error in task data", vbError, "Error":Exit Sub
sRow = sRow + rownine
For cRow = sRow To fRow'Populates indent levels to an array for this family ' tried moving outside loop didnt work
arrLvlTIL(cRow) = arrAllV(cRow, cpg.TIL) + 1
If cLevel < arrLvlTIL(cRow) Then cLevel = arrLvlTIL(cRow)
Next

Do Until cLevel = 1 'Loop from down to top and the get min and max dates of ESD and EED
EED = DateSerial(1899, 12, 31): ESD = DateSerial(3000, 12, 31): BED = DateSerial(1899, 12, 31): BSD = DateSerial(3000, 12, 31): AED = DateSerial(1899, 12, 31): ASD = DateSerial(3000, 12, 31)
BCS = 0: ECS = 0: ACS = 0: RCS = 0: bBCS = False: bECS = False: bACS = False: bRCS = False: ChildTasksDuration = 0: sumOfED = 0: sumOfWork = 0: bChildFound = False
For cRow = fRow To sRow Step -1
If arrLvlTIL(cRow) = cLevel Then
bChildFound = True
If ctype = "estDates" Or ctype = "allDates" Or ctype = allFields Then
If arrAllV(cRow, cpg.ESD) <> vbNullString And arrAllV(cRow, cpg.EED) <> vbNullString Then
If arrAllV(cRow, cpg.ESD) < ESD Then ESD = arrAllV(cRow, cpg.ESD)
If arrAllV(cRow, cpg.EED) > EED Then EED = arrAllV(cRow, cpg.EED)
If arrAllV(cRow, cpg.ED) <> "" Then sumOfED = sumOfED + arrAllV(cRow, cpg.ED)
End If
End If
If ctype = "work" Or ctype = allFields Then
If arrAllV(cRow, cpg.Work) <> "" Then sumOfWork = sumOfWork + arrAllV(cRow, cpg.Work)
End If
If ctype = "estDates" Or ctype = "allDates" Or ctype = "perc" Or ctype = allFields Then
If st.PercAuto Then GetCalculatedPercentage arrAllV(), cRow
tPercentComplete = Replace(Replace(arrAllV(cRow, cpg.PercentageCompleted), ".", Application.International(xlDecimalSeparator)), ",", Application.International(xlDecimalSeparator))
If st.CalParPercSimple Then
PercentComplete = PercentComplete + tPercentComplete
Else
If arrAllV(cRow, cpg.ED) = 0 Then
ChildTasksDuration = ChildTasksDuration + 1
PercentComplete = PercentComplete + (tPercentComplete * 1)
Else
ChildTasksDuration = ChildTasksDuration + arrAllV(cRow, cpg.ED)
PercentComplete = PercentComplete + (tPercentComplete * arrAllV(cRow, cpg.ED))
End If
End If
End If
If ctype = "basDates" Or ctype = "allDates" Or ctype = allFields Then
If st.CalBasDates Then
If IsDate(arrAllV(cRow, cpg.BSD)) Then If arrAllV(cRow, cpg.BSD) < BSD Then BSD = arrAllV(cRow, cpg.BSD)
If IsDate(arrAllV(cRow, cpg.BED)) Then If arrAllV(cRow, cpg.BED) > BED Then BED = arrAllV(cRow, cpg.BED)
If IsNumeric(arrAllV(cRow, cpg.BD)) Then sumOfBD = sumOfBD + arrAllV(cRow, cpg.BD)
End If
End If
If ctype = "actDates" Or ctype = "allDates" Or ctype = allFields Then
If st.CalActDates Then
If IsDate(arrAllV(cRow, cpg.ASD)) Then If arrAllV(cRow, cpg.ASD) < ASD Then ASD = arrAllV(cRow, cpg.ASD)
If IsDate(arrAllV(cRow, cpg.AED)) Then If arrAllV(cRow, cpg.AED) > AED Then AED = arrAllV(cRow, cpg.AED)
If IsNumeric(arrAllV(cRow, cpg.AD)) Then sumOfAD = sumOfAD + arrAllV(cRow, cpg.AD)
End If
End If
If ctype = "costs" Or ctype = allFields Then
If st.CalParCosts Then
If arrAllV(cRow, cpg.ECS) <> vbNullString And IsNumeric(arrAllV(cRow, cpg.ECS)) Then
ECS = ECS + Replace(Replace(arrAllV(cRow, cpg.ECS), ".", Application.International(xlDecimalSeparator)), ",", Application.International(xlDecimalSeparator)): bECS = True
End If
If arrAllV(cRow, cpg.BCS) <> vbNullString And IsNumeric(arrAllV(cRow, cpg.BCS)) Then
BCS = BCS + Replace(Replace(arrAllV(cRow, cpg.BCS), ".", Application.International(xlDecimalSeparator)), ",", Application.International(xlDecimalSeparator)): bBCS = True
End If
If arrAllV(cRow, cpg.ACS) <> vbNullString And IsNumeric(arrAllV(cRow, cpg.ACS)) Then
ACS = ACS + Replace(Replace(arrAllV(cRow, cpg.ACS), ".", Application.International(xlDecimalSeparator)), ",", Application.International(xlDecimalSeparator)): bACS = True
End If
If arrAllV(cRow, cpg.ResourceCost) <> vbNullString And IsNumeric(arrAllV(cRow, cpg.ResourceCost)) Then
RCS = RCS + Replace(Replace(arrAllV(cRow, cpg.ResourceCost), ".", Application.International(xlDecimalSeparator)), ",", Application.International(xlDecimalSeparator)): bRCS = True
End If
End If
End If
cTaskCount = cTaskCount + 1
ElseIf bChildFound = True And arrLvlTIL(cRow) = cLevel - 1 Then 'Parent Found now update
If ctype = "estDates" Or ctype = "allDates" Or ctype = allFields Then
If ESD <> DateSerial(1899, 12, 31) Then
If arrAllV(crow, cpg.Dependents) <> vbNullString And arrAllV(crow, cpg.ESD) <> ESD And arrLvlTIL(crow) + 1 = arrLvlTIL(crow) Then bbRecalculateConstrainValues = True 'removed in 3.76
arrAllV(cRow, cpg.ESD) = ESD
End If
If EED <> DateSerial(1899, 12, 31) Then
this condition is when a parent task date is getting changed from its child task and this parent has dependents - we recalcualte constrains and re-populate dates
If arrAllV(crow, cpg.Dependents) <> vbNullString And arrAllV(crow, cpg.EED) <> EED And arrLvlTIL(crow) = 1 Then bbRecalculateConstrainValues = True 'removed in 3.76
arrAllV(cRow, cpg.EED) = EED
End If
If st.CalParDurSum Then
arrAllV(cRow, cpg.ED) = sumOfED
Else
If st.HGC Then
arrAllV(cRow, cpg.ED) = CalEDHrs("organization", ESD, EED)
Else
arrAllV(cRow, cpg.ED) = GetWorkDaysFromDate("organization", ESD, EED)
End If
End If
End If
If ctype = "work" Or ctype = allFields Then
If sumOfWork > 0 Then arrAllV(cRow, cpg.Work) = sumOfWork Else arrAllV(cRow, cpg.Work) = vbNullString
End If
If ctype = "basDates" Or ctype = "allDates" Or ctype = allFields Then
If st.CalBasDates Then
If BSD <> DateSerial(1899, 12, 31) And BSD <> DateSerial(3000, 12, 31) Then arrAllV(cRow, cpg.BSD) = BSD Else arrAllV(cRow, cpg.BSD) = vbNullString
If BED <> DateSerial(1899, 12, 31) And BED <> DateSerial(3000, 12, 31) Then arrAllV(cRow, cpg.BED) = BED Else arrAllV(cRow, cpg.BED) = vbNullString
If st.CalParDurSum Then
If sumOfBD > 0 Then arrAllV(cRow, cpg.BD) = sumOfBD Else arrAllV(cRow, cpg.BD) = vbNullString
Else
If st.HGC Then
If arrAllV(cRow, cpg.BSD) <> vbNullString And arrAllV(cRow, cpg.BED) <> vbNullString Then arrAllV(cRow, cpg.BD) = CalEDHrs("organization", BSD, BED) Else arrAllV(cRow, cpg.BD) = vbNullString
Else
If arrAllV(cRow, cpg.BSD) <> vbNullString And arrAllV(cRow, cpg.BED) <> vbNullString Then arrAllV(cRow, cpg.BD) = GetWorkDaysFromDate("organization", BSD, BED) Else arrAllV(cRow, cpg.BD) = vbNullString
End If
End If
End If
End If
If ctype = "actDates" Or ctype = "allDates" Or ctype = allFields Then
If st.CalActDates Then
If ASD <> DateSerial(1899, 12, 31) And ASD <> DateSerial(3000, 12, 31) Then arrAllV(cRow, cpg.ASD) = ASD Else arrAllV(cRow, cpg.ASD) = vbNullString
If AED <> DateSerial(1899, 12, 31) And AED <> DateSerial(3000, 12, 31) Then arrAllV(cRow, cpg.AED) = AED Else arrAllV(cRow, cpg.AED) = vbNullString
If st.CalParDurSum Then
If sumOfAD > 0 Then arrAllV(cRow, cpg.AD) = sumOfAD Else arrAllV(cRow, cpg.AD) = vbNullString
Else
If st.HGC Then
If arrAllV(cRow, cpg.ASD) <> vbNullString And arrAllV(cRow, cpg.AED) <> vbNullString Then arrAllV(cRow, cpg.AD) = CalEDHrs("organization", ASD, AED) Else arrAllV(cRow, cpg.AD) = vbNullString
Else
If arrAllV(cRow, cpg.ASD) <> vbNullString And arrAllV(cRow, cpg.AED) <> vbNullString Then arrAllV(cRow, cpg.AD) = GetWorkDaysFromDate("organization", ASD, AED) Else arrAllV(cRow, cpg.AD) = vbNullString
End If
End If
End If
End If
If ctype = "estDates" Or ctype = "allDates" Or ctype = "perc" Or ctype = allFields Then
If st.CalParPercSimple Then
arrAllV(cRow, cpg.PercentageCompleted) = PercentComplete / cTaskCount
Else
arrAllV(cRow, cpg.PercentageCompleted) = PercentComplete / ChildTasksDuration
End If
End If
If ctype = "costs" Or ctype = allFields Then
If st.CalParCosts Then
If bECS Then arrAllV(cRow, cpg.ECS) = ECS Else arrAllV(cRow, cpg.ECS) = vbNullString
If bBCS Then arrAllV(cRow, cpg.BCS) = BCS Else arrAllV(cRow, cpg.BCS) = vbNullString
If bACS Then arrAllV(cRow, cpg.ACS) = ACS Else arrAllV(cRow, cpg.ACS) = vbNullString
If bRCS Then arrAllV(cRow, cpg.ResourceCost) = RCS Else arrAllV(cRow, cpg.ResourceCost) = vbNullString
End If
End If
bChildFound = False
cTaskCount = 0: PercentComplete = 0: ChildTasksDuration = 0: sumOfED = 0: sumOfWork = 0: sumOfBD = 0: sumOfAD = 0
EED = DateSerial(1899, 12, 31): ESD = DateSerial(3000, 12, 31): BED = DateSerial(1899, 12, 31): BSD = DateSerial(3000, 12, 31): AED = DateSerial(1899, 12, 31): ASD = DateSerial(3000, 12, 31)
BCS = 0: ECS = 0: ACS = 0: RCS = 0: bBCS = False: bECS = False: bACS = False: bRCS = False
If bbRecalculateConstrainValues Then GoTo last' removed 3.76
End If
Next cRow
cLevel = cLevel - 1
Loop
Else
If ctype = "estDates" Or ctype = "allDates" Or ctype = "perc" Or ctype = allFields Then If st.PercAuto Then GetCalculatedPercentage arrAllV(), fRow
If CalOnlyRowNo > 0 Then Exit Do
End If
fRow = sRow - 1
Loop

If ActiveSheet.AutoFilterMode Then
Call ReApplyAutoFilter(arrAllV, lrow, "PopParentTasksV")
Call ReApplyAutoFilter(arrAllF, lrow, "PopParentTasksF")
Else
Call ArrayToRange(arrAllV, lrow, "PopParentTasksV") ''Range(Cells(1, 1), Cells(lrow, cpg.lc)) = arrAllV
Call ArrayToRange(arrAllF, lrow, "PopParentTasksF") '
End If

Last:
'If bbRecalculateConstrainValues = True Then ' removed 3.76
Range(Cells(1, 1), Cells(lrow, cpg.lc)) = arrAllV
bbRecalculateConstrainValues = False
Call ReCalcDepFormulas(crow, True)
PopParentTasks
'End If

If st.EnableCosts Then Call ReCalculateBudgetLineCosts
tlog "PopParentTasks: " & CalOnlyRowNo & ctype

End Sub

Sub ArrayToRange(v, lrow As Long, copyType As String)
tlog copyType
Dim arrAllTempV(): ReDim arrAllTempV(1 To lrow, 1 To cpg.LC): arrAllTempV = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
Dim cRow As Long: Dim bAutoPerc As Boolean
bAutoPerc = st.PercAuto
If copyType = "PopParentTasksV" Then 'v
For cRow = firsttaskrow To lrow
If IsParentTask(cRow) Then
If arrAllTempV(cRow, cpg.ESD) <> v(cRow, cpg.ESD) Then Cells(cRow, cpg.ESD) = v(cRow, cpg.ESD):
If arrAllTempV(cRow, cpg.EED) <> v(cRow, cpg.EED) Then Cells(cRow, cpg.EED) = v(cRow, cpg.EED):
If arrAllTempV(cRow, cpg.ED) <> v(cRow, cpg.ED) Then Cells(cRow, cpg.ED) = v(cRow, cpg.ED):
If arrAllTempV(cRow, cpg.BSD) <> v(cRow, cpg.BSD) Then Cells(cRow, cpg.BSD) = v(cRow, cpg.BSD)
If arrAllTempV(cRow, cpg.ASD) <> v(cRow, cpg.ASD) Then Cells(cRow, cpg.ASD) = v(cRow, cpg.ASD)
If arrAllTempV(cRow, cpg.BED) <> v(cRow, cpg.BED) Then Cells(cRow, cpg.BED) = v(cRow, cpg.BED)
If arrAllTempV(cRow, cpg.AED) <> v(cRow, cpg.AED) Then Cells(cRow, cpg.AED) = v(cRow, cpg.AED)
If arrAllTempV(cRow, cpg.BD) <> v(cRow, cpg.BD) Then Cells(cRow, cpg.BD) = v(cRow, cpg.BD)
If arrAllTempV(cRow, cpg.AD) <> v(cRow, cpg.AD) Then Cells(cRow, cpg.AD) = v(cRow, cpg.AD)
If arrAllTempV(cRow, cpg.ECS) <> v(cRow, cpg.ECS) Then Cells(cRow, cpg.ECS) = v(cRow, cpg.ECS)
If arrAllTempV(cRow, cpg.BCS) <> v(cRow, cpg.BCS) Then Cells(cRow, cpg.BCS) = v(cRow, cpg.BCS)
If arrAllTempV(cRow, cpg.ACS) <> v(cRow, cpg.ACS) Then Cells(cRow, cpg.ACS) = v(cRow, cpg.ACS)
If arrAllTempV(cRow, cpg.ResourceCost) <> v(cRow, cpg.ResourceCost) Then Cells(cRow, cpg.ResourceCost) = v(cRow, cpg.ResourceCost)
If arrAllTempV(cRow, cpg.PercentageCompleted) <> v(cRow, cpg.PercentageCompleted) Then Cells(cRow, cpg.PercentageCompleted) = v(cRow, cpg.PercentageCompleted)
If arrAllTempV(cRow, cpg.Work) <> v(cRow, cpg.Work) Then Cells(cRow, cpg.Work) = v(cRow, cpg.Work)
Else
If bAutoPerc And arrAllTempV(cRow, cpg.PercentageCompleted) <> v(cRow, cpg.PercentageCompleted) Then Cells(cRow, cpg.PercentageCompleted) = v(cRow, cpg.PercentageCompleted)
End If
Next cRow
End If
If copyType = "PopParentTasksF" Then 'f
For cRow = firsttaskrow To lrow
If IsParentTask(cRow) Then
If Left(v(cRow, cpg.ESD), 1) = "=" Then Cells(cRow, cpg.ESD) = v(cRow, cpg.ESD)
If Left(v(cRow, cpg.BSD), 1) = "=" Then Cells(cRow, cpg.BSD) = v(cRow, cpg.BSD)
If Left(v(cRow, cpg.ASD), 1) = "=" Then Cells(cRow, cpg.ASD) = v(cRow, cpg.ASD)
If Left(v(cRow, cpg.EED), 1) = "=" Then Cells(cRow, cpg.EED) = v(cRow, cpg.EED)
If Left(v(cRow, cpg.BED), 1) = "=" Then Cells(cRow, cpg.BED) = v(cRow, cpg.BED)
If Left(v(cRow, cpg.AED), 1) = "=" Then Cells(cRow, cpg.AED) = v(cRow, cpg.AED)
If Left(v(cRow, cpg.ED), 1) = "=" Then Cells(cRow, cpg.ED) = v(cRow, cpg.ED)
If Left(v(cRow, cpg.BD), 1) = "=" Then Cells(cRow, cpg.BD) = v(cRow, cpg.BD)
If Left(v(cRow, cpg.AD), 1) = "=" Then Cells(cRow, cpg.AD) = v(cRow, cpg.AD)
If Left(v(cRow, cpg.ECS), 1) = "=" Then Cells(cRow, cpg.ECS) = v(cRow, cpg.ECS)
If Left(v(cRow, cpg.BCS), 1) = "=" Then Cells(cRow, cpg.BCS) = v(cRow, cpg.BCS)
If Left(v(cRow, cpg.ACS), 1) = "=" Then Cells(cRow, cpg.ACS) = v(cRow, cpg.ACS)
If Left(v(cRow, cpg.ResourceCost), 1) = "=" Then Cells(cRow, cpg.ResourceCost) = v(cRow, cpg.ResourceCost)
If Left(v(cRow, cpg.PercentageCompleted), 1) = "=" Then Cells(cRow, cpg.PercentageCompleted) = v(cRow, cpg.PercentageCompleted)
If Left(v(cRow, cpg.Work), 1) = "=" Then Cells(cRow, cpg.Work) = v(cRow, cpg.Work)
Else
If bAutoPerc = False And Left(v(cRow, cpg.PercentageCompleted), 1) = "=" Then Cells(cRow, cpg.PercentageCompleted) = v(cRow, cpg.PercentageCompleted)
End If
Next cRow
End If
If copyType = "UpdateNewDepCols" Then 'v
Range(Cells(firsttaskrow, cpg.WBSPredecessors), Cells(lrow, cpg.WBSPredecessors)).ClearContents
Range(Cells(firsttaskrow, cpg.WBSSuccessors), Cells(lrow, cpg.WBSSuccessors)).ClearContents
For cRow = firsttaskrow To lrow
If v(cRow, cpg.WBSPredecessors) <> "" Then Cells(cRow, cpg.WBSPredecessors) = v(cRow, cpg.WBSPredecessors):
If v(cRow, cpg.WBSSuccessors) <> "" Then Cells(cRow, cpg.WBSSuccessors) = v(cRow, cpg.WBSSuccessors):
Next cRow
End If
If copyType = "CalcDepFormulas" Then 'f
Range(Cells(firsttaskrow, cpg.StartConstrain), Cells(lrow, cpg.EndConstrain)) = v
End If
If copyType = "ClearDepFormulas" Then 'v
Range(Cells(firsttaskrow, cpg.StartConstrain), Cells(lrow, cpg.EndConstrain)).value = v
End If
If copyType = "ShapeInfoEst" Then 'v
Range(Cells(firsttaskrow, cpg.ShapeInfoE), Cells(lrow, cpg.ShapeInfoE)) = v
End If
If copyType = "ShapeInfoBas" Then 'v
Range(Cells(firsttaskrow, cpg.ShapeInfoB), Cells(lrow, cpg.ShapeInfoB)) = v
End If
If copyType = "ShapeInfoAct" Then 'v
Range(Cells(firsttaskrow, cpg.ShapeInfoA), Cells(lrow, cpg.ShapeInfoA)) = v
End If
tlog copyType
End Sub

Sub getnondepID(cRow As Long)
ReDim CheckedList(1 To GetLastRow) 'GetLastTaskRowNo - 2)
p = 1
Call GetNonDependentIDs(cRow)
End Sub

Sub GetNonDependentIDs(cRow As Long)
Returns task IDs as 1|2|3| etc. which will cause no loop if added as asdependent to task in tRow
'Dim r As Range:
Dim iCount As Long, pRow As Long: Dim vDependents, vdIDs, d As Integer
'Set r = Range(Cells(1, cpg.TID), Cells(Cells.Rows.Count, cpg.TID)) ' TID range
Dim m As Long
For m = LBound(CheckedList()) To UBound(CheckedList())
If CheckedList(m) = cRow Then Exit Sub
Next m
CheckedList(p) = cRow
p = p + 1
'Loop through its dependencies
If Cells(cRow, cpg.Dependents) = vbNullString Then GoTo CheckIndentLevels
vDependents = Split(Cells(cRow, cpg.Dependents), DepSeperator)
For d = 0 To UBound(vDependents) - 1
vdIDs = Split(vDependents(d), "_")
Remove this ID from the tIDlists
On Error Resume Next
tIDList.Remove "K" & vdIDs(0)
On Error GoTo 0
GetNonDependentIDs r.Find(vdIDs(0), , xlFormulas, xlWhole).Row
Call GetNonDependentIDs(getTIDRow(CLng(vdIDs(0))))
Next d
CheckIndentLevels:
Is crow is main parent all others will become nondependents
If Cells(cRow, cpg.Task).IndentLevel = 0 Then
Exit Sub
Else
Remove parents and main parents
pRow = GetTaskParentRowNumber(cRow):iCount = tIDList.Count
Remove the parent task for task list ids
On Error Resume Next
If IsParentTask(pRow) Then tIDList.Remove "K" & Cells(pRow, cpg.TID)
On Error GoTo 0
It is child task, so we need to check this selection parents are driving any tasks
If iCount > tIDList.Count Then Call GetNonDependentIDs(pRow)
End If
End Sub

Function GetTaskParentRowNumber(ByVal cRow As Long) As Long
Dim iLevel As Long:GetTaskParentRowNumber = cRow - 1
iLevel = Cells(cRow, cpg.Task).IndentLevel
Do Until Cells(GetTaskParentRowNumber, cpg.Task).IndentLevel < iLevel
GetTaskParentRowNumber = GetTaskParentRowNumber - 1
Loop
End Function


Sub ReCalculateBudgetLineCosts(Optional t As Boolean)
Set gs = setGSws
Dim fRow As Long, cRow As Long, i As Long: Dim TotalECS As Double, TotalACS As Double, TotalBCS As Double: Dim tStr As String
If st.EnableCosts = False Then Exit Sub
TotalACS = 0: TotalBCS = 0: TotalECS = 0: fRow = GetLastRow
For cRow = firsttaskrow To fRow
If Cells(cRow, cpg.Task).IndentLevel = 0 Then
If IsNumeric(Cells(cRow, cpg.ECS)) Then TotalACS = TotalACS + Cells(cRow, cpg.ACS):
If IsNumeric(Cells(cRow, cpg.BCS)) Then TotalBCS = TotalBCS + Cells(cRow, cpg.BCS):
If IsNumeric(Cells(cRow, cpg.ACS)) Then TotalECS = TotalECS + Cells(cRow, cpg.ECS)
End If
Next
Cells(roweight, cpg.WBS) = _
Project Budget:& _
Estimated:& gs.Cells(rowtwo, cps.CurrencySymbol) & Format(gs.Cells(rowtwo, cps.EstimatedBudget), "#,##0.00") & " | " & _
"Baseline: " & gs.Cells(rowtwo, cps.CurrencySymbol) & Format(gs.Cells(rowtwo, cps.BaselineBudget), "#,##0.00") & " | " & _
"Task Costs: " & _
"Estimated: " & gs.Cells(rowtwo, cps.CurrencySymbol) & Format(TotalECS, "#,##0.00") & " | " & _
"Baseline: " & gs.Cells(rowtwo, cps.CurrencySymbol) & Format(TotalBCS, "#,##0.00") & " | " & _
"Actual: " & gs.Cells(rowtwo, cps.CurrencySymbol) & Format(TotalACS, "#,##0.00")
tStr = Cells(roweight, cpg.WBS): i = InStr(1, tStr, "Task Costs:", vbTextCompare)
With Cells(roweight, cpg.WBS)
.Font.Bold = False: .Characters(1, 14).Font.Bold = True: .Characters(i, 10).Font.Bold = True
End With
TotalACS = 0: TotalBCS = 0: TotalECS = 0
End Sub

Sub ClearDepFormulas(Optional t As Boolean)
Dim arrAllConstrains(): Dim lrow As Long: lrow = GetLastRow:
arrAllConstrains = Range(Cells(firsttaskrow, cpg.StartConstrain), Cells(lrow, cpg.EndConstrain)).value
If ActiveSheet.AutoFilterMode Then
Call ReApplyAutoFilter(arrAllConstrains, lrow, "ClearDepFormulas")
Else
Range(Cells(firsttaskrow, cpg.StartConstrain), Cells(lrow, cpg.EndConstrain)) = arrAllConstrains
Call ArrayToRange(arrAllConstrains, lrow, "ClearDepFormulas")
End If
End Sub
Function AddDelayOperator(tStr As String) As String
If Left(tStr, 1) = "-" Then AddDelayOperator = tStr Else AddDelayOperator = "+" & tStr
End Function

Sub UpdateNewDepCols(Optional t As Boolean)
tlog "UpdateNewDepCols"
Dim arrAllData()
Dim splitDependency, splitDependencyParts: Dim splitDependencyStr As String
Dim splitDependents, splitDependentsParts: Dim splitDependentsStr As String
Dim lrow As Long, i As Long, j As Long, cRow As Long
lrow = GetLastRow
arrAllData = Range(Cells(1, cpg.GEtype), Cells(lrow, cpg.LC)).value
For cRow = firsttaskrow To UBound(arrAllData())
If arrAllData(cRow, cpg.Dependency) <> "" Then
splitDependency = Split(arrAllData(cRow, cpg.Dependency), DepSeperator)
For i = 0 To UBound(splitDependency) - 1
splitDependencyParts = Split(splitDependency(i), "_")
splitDependencyStr = splitDependencyStr & Cells(getTIDRow(CLng(splitDependencyParts(0))), cpg.WBS) & " " & splitDependencyParts(1) & " " & splitDependencyParts(2) & DepSeperator
Next i
arrAllData(cRow, cpg.WBSPredecessors) = splitDependencyStr: splitDependencyStr = vbNullString
Else
arrAllData(cRow, cpg.WBSPredecessors) = ""
End If
If arrAllData(cRow, cpg.Dependents) <> "" Then
splitDependents = Split(arrAllData(cRow, cpg.Dependents), DepSeperator)
For i = 0 To UBound(splitDependents) - 1
splitDependentsParts = Split(splitDependents(i), "_")
splitDependentsStr = splitDependentsStr & Cells(getTIDRow(CLng(splitDependentsParts(0))), cpg.WBS) & DepSeperator
Next i
arrAllData(cRow, cpg.WBSSuccessors) = splitDependentsStr: splitDependentsStr = vbNullString
Else
arrAllData(cRow, cpg.WBSSuccessors) = ""
End If
Next cRow

If ActiveSheet.AutoFilterMode Then
Call ReApplyAutoFilter(arrAllData, lrow, "UpdateNewDepCols")
Else
Range(Cells(1, cpg.GEType), Cells(lrow, cpg.LC)) = arrAllData
Call ArrayToRange(arrAllData, lrow, "UpdateNewDepCols")
End If
tlog "UpdateNewDepCols"
End Sub
Option Explicit
Private bPopParentTasks As Boolean, bEstDatesChanged As Boolean, bBasDatesChanged As Boolean, bActDatesChanged As Boolean, bNewTask As Boolean, bResChanged As Boolean
Private bWBSChanged As Boolean, bPercChanged As Boolean, bCostsChanged As Boolean, bPriorityChanged As Boolean, bStatusChanged As Boolean, bTaskIconChanged As Boolean
Private bDrivingTask As Boolean, bDrivenTask As Boolean, bPlainTask As Boolean, bParentTask As Boolean, bNormaltask As Boolean, bChildTask As Boolean, bDependencies As Boolean
Private bWorkChanged As Boolean, bDoneChanged As Boolean, bChangeTextBar As Boolean, bWBSPredecessor As Boolean, bWBSSuccessor As Boolean, bForceReDraw As Boolean
Private cRow As Long, cCol As Long, lrow As Long
Private arrRowsEdited As Variant

Sub setBooleanDE(b As Boolean)
bPopParentTasks = b: bEstDatesChanged = b: bBasDatesChanged = b: bActDatesChanged = b: bPriorityChanged = b: bStatusChanged = b
bPercChanged = b: bCostsChanged = b: bNewTask = b: bResChanged = b: bDependencies = b: bTaskIconChanged = b: bDoneChanged = b
bWBSChanged = b: bDrivingTask = b: bDrivenTask = b: bPlainTask = b: bParentTask = b: bNormaltask = b: bChildTask = b: bChangeTextBar = b
bWorkChanged = b: bForceReDraw = b
End Sub

Sub TriggerCellValueChanged(sRng As Range)
tlog "TriggerCellValueChanged"
Call DA: Set gs = setGSws
Dim c As Range: Dim tidl As Long, nextTidl As Long, i As Long:
'Public correctColor As Long correctColor = RGB(226, 239, 163) ' mac issues 'Public incorrectColor As Long incorrectColor = RGB(255, 197, 197)
Call RememberResArrays: ResArraysReady = True: setBooleanDE (False): ReDim arrRowsEdited(1 To sRng.Rows.Count)
For Each c In sRng.Cells
cRow = c.Row: cCol = c.column: lrow = GetLastRow
If FreeVersion And c.column = cpg.Task Then
If Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A")) - 3 >= cFreeVersionTasksCount Then
If Cells(cRow, cpg.GEtype) = vbNullString Then Call DeleteExtrasRowsInFree: Call AddNewTaskPlaceholder: sTempStr1 = msg(80) & msg(82): frmBuyPro.show: GoTo Last
End If
End If
If cRow > rownine And (st.ShowCompleted = False Or st.ShowPlanned = False Or st.ShowInProgress = False Or ActiveSheet.AutoFilterMode = True) And Cells(c.Row, cpg.WBS) = vbNullString Then
MsgBox msg(17) 'filtered
If cRow > GetLastRow + 1 Then c = vbNullString
If cCol = cpg.Task Then AddNewTaskPlaceholder
End If
If cRow = rownine Then
If Cells(rownine, cCol) = "" Then
If Cells(rowone, cCol) = "TIL" Or Cells(rowone, cCol) = "SS" Or Cells(rowone, cCol) = "TaskIcon" Then taskTag = "notTask": GoTo nexa
MsgBox msg(64): Cells(rownine, cCol) = resetColHeaderforCol(cCol)
End If
taskTag = "notTask": GoTo nexa
End If
If cRow < rownine Then c.value = vbNullString: taskTag = "notTask": GoTo nexa

If cCol = cpg.Task Then
If cRow > lrow + 1 Then c.value = vbNullString: taskTag = "notTask": GoTo nexa
Else
If Cells(cRow, cpg.GEtype) = vbNullString Then c.value = vbNullString: taskTag = "notTask": GoTo nexa
End If
Call ExecuteCellChanges(c)
If Cells(cRow, cpg.GEtype) = vbNullString Then GoTo nexa:

If IsNormalTask(cRow) Then
bNormaltask = True
Else
bNormaltask = False
If IsParentTask(cRow) Then bParentTask = True: bChildTask = False Else bChildTask = True: bParentTask = False
End If
If IsDrivenTask(cRow) Then bDrivenTask = True
If IsDrivingTask(cRow) Then bDrivingTask = True: bDrivenTask = False
If bDrivingTask = False And bDrivenTask = False Then bPlainTask = True

If bNewTask Then
Call TaskgridFormatting(cRow, True): Call FormatTasks(cRow, allFields, rowOnly) ': Call DrawGanttBars(crow, estBars, rowOnly)
Else
Call TaskgridFormatting(cRow, False)
End If
If bNewTask Then GoTo nexa
nexa:
Next

If checkSheetError Then MsgBox msg(93): GoTo Last

If taskTag = "notTask" Then GoTo Last

If sRng.Rows.Count = 1 Then
If bNormaltask Then 'Single Normal Task
If bEstDatesChanged Then
If bDependencies Then
Call PopParentTasks(, allFields)
Call colorAllSS
Call DrawAllGanttBars: bChangeTextBar = False
Else
If bBasDatesChanged Or bActDatesChanged Then
Call FormatTasks(cRow, fPerc, rowOnly)
Call DrawGanttBars(cRow, allBars, rowOnly): bChangeTextBar = False
Else
Call FormatTasks(cRow, fPerc, rowOnly)
Call DrawGanttBars(cRow, estBars, rowOnly): bChangeTextBar = False
End If
End If
End If
If bBasDatesChanged And bEstDatesChanged = False Then Call DrawGanttBars(cRow, basBars, rowOnly)
If bActDatesChanged And bEstDatesChanged = False Then Call DrawGanttBars(cRow, actBars, rowOnly)
If bPercChanged And bEstDatesChanged = False Then
Call FormatTasks(cRow, fPerc, rowOnly):
If st.ShowPercBar Then Call DrawGanttBars(cRow, estBars, rowOnly): bChangeTextBar = False
End If
If bResChanged Then Call HiliteHNWDperrow(clrow)
If bCostsChanged Then Call PopParentTasks(cRow, fCosts)
If bPriorityChanged Then Call FormatTasks(cRow, fPriority, rowOnly)
If bStatusChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bTaskIconChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bWBSChanged Then Call WBSNumbering
If bDoneChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bChangeTextBar Then Call ChangeTextInGanttBars(clrow, rowOnly)
If bWBSPredecessor Or bWBSSuccessor Then Call UpdateNewDepCols
End If
If bParentTask Then
If bPopParentTasks Then Call PopParentTasks(, allFields)
If bPriorityChanged Then Call FormatTasks(cRow, fPriority, rowOnly)
If bStatusChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bTaskIconChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bWBSChanged Then Call WBSNumbering
If bCostsChanged Then Call PopParentTasks(, fCosts)
If bWorkChanged Then Call PopParentTasks(, fWork)
If bDoneChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If st.CalBasDates = False And bBasDatesChanged Then Call DrawGanttBars(cRow, basBars, rowOnly)
If st.CalActDates = False And bActDatesChanged Then Call DrawGanttBars(cRow, actBars, rowOnly)
If bWBSPredecessor Or bWBSSuccessor Then Call UpdateNewDepCols
End If
If bChildTask Then ' Single Child Task
If bEstDatesChanged Then
If bDependencies Then
Call PopParentTasks(, allFields)
Call colorAllSS
Call DrawAllGanttBars: bChangeTextBar = False
Else
If bBasDatesChanged Or bActDatesChanged Then
Call PopParentTasks(cRow, allFields)
Call FormatTasks(cRow, fPerc, aboveFamily)
Call DrawGanttBars(cRow, allBars, aboveFamily): bChangeTextBar = False
Else
Call PopParentTasks(cRow, allFields)
Call FormatTasks(cRow, fPerc, aboveFamily)
Call DrawGanttBars(cRow, estBars, aboveFamily): bChangeTextBar = False
End If
End If
End If
If bBasDatesChanged And bEstDatesChanged = False Then
If st.CalBasDates Then Call PopParentTasks(cRow, basDates)
Call DrawGanttBars(cRow, basBars, aboveFamily)
End If
If bActDatesChanged And bEstDatesChanged = False Then
If st.CalActDates Then Call PopParentTasks(cRow, actDates)
Call DrawGanttBars(cRow, actBars, aboveFamily)
End If
If bCostsChanged Then Call PopParentTasks(cRow, fCosts)
If bWorkChanged Then Call PopParentTasks(, fWork)
If bPercChanged And bEstDatesChanged = False Then
Call PopParentTasks(cRow, fPerc)
Call FormatTasks(cRow, fPerc, aboveFamily)
If st.ShowPercBar Then Call DrawGanttBars(cRow, estBars, aboveFamily): bChangeTextBar = False
End If
If bResChanged Then Call HiliteHNWDperrow(clrow)
If bPriorityChanged Then Call FormatTasks(cRow, fPriority, rowOnly)
If bStatusChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bTaskIconChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bWBSChanged Then Call WBSNumbering
If bDoneChanged Then Call FormatTasks(cRow, fPerc, rowOnly)
If bChangeTextBar Then Call ChangeTextInGanttBars(clrow, aboveFamily)
If bWBSPredecessor Or bWBSSuccessor Then Call UpdateNewDepCols
End If
Else
If bWBSChanged Then Call WBSNumbering
If bEstDatesChanged Then
Call PopParentTasks(, allFields)
Call colorAllSS
If bBasDatesChanged Or bActDatesChanged Then
Call DrawAllGanttBars: bChangeTextBar = False
Else
Call DrawGanttBars(, estBars, allRows): bChangeTextBar = False
End If
End If
If bBasDatesChanged And bEstDatesChanged = False Then
If st.CalBasDates Then Call PopParentTasks(, allFields)
Call DrawGanttBars(, basBars, allRows)
End If
If bActDatesChanged And bEstDatesChanged = False Then
If st.CalActDates Then Call PopParentTasks(, allFields)
Call DrawGanttBars(, actBars, allRows)
End If
If bPercChanged And bEstDatesChanged = False Then
Call PopParentTasks(, fPerc)
Call colorAllSS
If st.ShowPercBar Then Call DrawGanttBars(, estBars, allRows): bChangeTextBar = False
End If
If bResChanged Then Call HiliteHNWDperrow(, allRows)
If bCostsChanged Then Call PopParentTasks(, allFields)
If bPriorityChanged Then Call colorAllPriority
If bStatusChanged Then Call colorAllSS
If bTaskIconChanged Then Call colorAllSS
If bDoneChanged Then Call colorAllSS
If bChangeTextBar Then Call ChangeTextInGanttBars(clrow, allRows)
If bWBSPredecessor Or bWBSSuccessor Then Call UpdateNewDepCols
End If
If bForceReDraw Then Call DelnDrawAllGanttBars
Last:
taskTag = ""

Call setBooleanDE(False): ResArraysReady = False: tlog "TriggerCellValueChanged"
Call EA
End Sub

Sub ExecuteCellChanges(ByVal crng As Range)
tlog "ExecuteCellChanges"
If crng.Row <= rownine Then Exit Sub 'If Cells(crng.Row, cpg.Task) = sAddTaskPlaceHolder Or Cells(crng.Row, cpg.GEType) = vbNullString And crng.column <> cpg.Task Then crng = vbNullString: Exit Sub
Select Case crng.column
Case cpg.TaskIcon
Cell_TaskIconChanged crng
Case cpg.WBS
Cell_WBSChanged crng
Case cpg.GEtype
Exit Sub
Case cpg.Task
Cell_TaskNameChanged crng
Case cpg.Priority
Cell_PriorityChanged crng
Case cpg.Status
Cell_StatusChanged crng
Case cpg.Resource
Cell_ResourceChanged crng
Case cpg.ResourceCost
Cell_ResourceCost crng
Case cpg.ESD
Cell_ESDChanged crng
Case cpg.EED
Cell_EEDChanged crng
Case cpg.ED
Cell_EDChanged crng
Case cpg.WBSPredecessors
Cell_WBSPredecessor crng
Case cpg.WBSSuccessors
Cell_WBSSuccessor crng
Case cpg.Work
cell_WorkChanged crng
Case cpg.Done
Cell_DoneChanged crng
Case cpg.PercentageCompleted
cell_PercentageCompleted crng
Case cpg.BSD
Cell_BSDChanged crng
Case cpg.BED
Cell_BEDChanged crng
Case cpg.BD
Cell_BDChanged crng
Case cpg.ASD
Cell_ASDChanged crng
Case cpg.AED
Cell_AEDChanged crng
Case cpg.AD
Cell_ADChanged crng
Case cpg.BCS
cell_BCSChanged crng
Case cpg.ECS
Cell_ECSChanged crng
Case cpg.ACS
cell_ACSChanged crng
Case cpg.Notes
Cell_NotesChanged crng
Case cpg.ShapeInfoE
Call bForceDraw
Case cpg.ShapeInfoB
Call bForceDraw
Case cpg.ShapeInfoA
Call bForceDraw
Case cpg.LC
Cell_LCChanged crng
End Select
If Cells(crng.Row, cpg.GEtype) = "S" And crng.column <> cpg.Task Then
crng = vbNullString
End If
tlog "ExecuteCellChanges"
End Sub

Sub Cell_TaskNameChanged(ByVal crng As Range)
crng.NumberFormat = "General"
Dim rngComp As Range: Dim sPos As Long: Dim s As String, df As String: Dim orgStartHrs As Double: Dim startdatehour As Date
lrow = GetLastRow: cRow = crng.Row
If IsError(crng) Then MsgBox msg(76) & crng.Address: Cells(cRow, cpg.Task) = msg(16): GoTo Last
If crng = vbNullString And Cells(cRow, cpg.GEtype) = vbNullString And Cells(cRow - 1, cpg.GEtype) <> vbNullString Then Call AddNewTaskPlaceholder: Exit Sub
If cRow = lrow + 1 Then bNewTask = True Else If cRow > 9 And cRow < lrow + 1 Then bNewTask = False
If bNewTask = False Then
If crng = vbNullString And Cells(cRow, cpg.GEtype) <> vbNullString Then MsgBox msg(15): Cells(cRow, cpg.Task) = msg(16)
crng.IndentLevel = Cells(crng.Row, cpg.TIL)
GoTo Last
End If

If bNewTask Then'not sure of this
If st.ShowGrouping = True Then
Set gs = setGSws: ActiveSheet.Cells.EntireRow.ClearOutline:
gs.Cells(rowtwo, cps.ShowGrouping) = 0: Call ReadSettings: Call RefreshRibbon
End If
Cells(cRow, cpg.GEtype) = "T": Cells(cRow, cpg.TID) = GetNextIDNumber:
Cells(cRow, cpg.Task).IndentLevel = Cells(cRow - 1, cpg.Task).IndentLevel: Cells(cRow, cpg.TIL) = Cells(cRow, cpg.Task).IndentLevel:
Cells(cRow, cpg.Task) = Trim(Cells(cRow, cpg.Task).value): Cells(cRow, cpg.Priority) = "NORMAL":
Cells(cRow, cpg.ED) = 1: Cells(cRow, cpg.PercentageCompleted).value = 0:
If st.HGC Then
orgStartHrs = sArr.ResourceP(0, 10): startdatehour = Date + orgStartHrs: Cells(cRow, cpg.ESD) = startdatehour:
Else
Cells(cRow, cpg.ESD) = Date
End If
Call Cell_EDChanged(Cells(crng.Row, cpg.ED))

If Cells(cRow + 1, cpg.Task) = vbNullString Then Rows(cRow + 1).EntireRow.Insert: Call AddNewTaskPlaceholder

If cRow = firsttaskrow Then
Cells(cRow, cpg.WBS) = 1
Else
s = Cells(cRow - 1, cpg.WBS)
If InStr(1, s, ".") = 0 Then
Cells(cRow, cpg.WBS) = CStr(CLng(s) + 1)
Else
sPos = InStrRev(s, "."): Cells(cRow, cpg.WBS) = Left(s, sPos) & CLng(Right(s, Len(s) - sPos)) + 1
End If
End If
Rows(cRow).RowHeight = taskRowHeight
End If
Last:
bChangeTextBar = True
End Sub

Sub Cell_EDChanged(ByVal crng As Range)
Dim orgStartHrs As Double: Dim startdatehour As Date, ESD As Date, newESD As Date, selESD As Date: Dim answer As Integer: Dim lDur As Long
Call CheckRngFormula(crng, True)
MsgBox msg(20)
If Cells(crng.Row, cpg.GEType) = "M" Then crng = 0 Else crng = 1
'End If

If IsError(crng) Then MsgBox msg(76) & crng.Address: crng = 1

If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(4): bPopParentTasks = True: Exit Sub
If InStr(1, Cells(crng.Row, cpg.Dependency), "S_", vbTextCompare) > 0 And InStr(1, Cells(crng.Row, cpg.Dependency), "F_", vbTextCompare) > 0 Then
MsgBox msg(57):
If st.HGC Then
Cells(crng.Row, cpg.ED) = CalEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ESD), Cells(crng.Row, cpg.EED))
If Cells(crng.Row, cpg.ED) <= 0 Then If Cells(crng.Row, cpg.GEtype) = "M" Then Cells(crng.Row, cpg.ED) = 0 Else Cells(crng.Row, cpg.ED) = 1
Else
Cells(crng.Row, cpg.ED) = GetWorkDaysFromDate(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.ESD)), CDate(Cells(crng.Row, cpg.EED)))
End If
Exit Sub
End If

If Cells(crng.Row, cpg.GEtype) = "M" Then
If Cells(crng.Row, cpg.ED) <> 0 Then MsgBox msg(13)
Cells(crng.Row, cpg.ED) = 0: Cells(crng.Row, cpg.EED) = Cells(crng.Row, cpg.ESD): bEstDatesChanged = True:
GoTo Last
End If

If IsError(crng) = True Or crng <= 1 Or crng >= 10000 Or crng = vbNullString Or crng <= 0 Or IsNumeric(crng) = False Or IsDate(crng) Then
If Cells(crng.Row, cpg.GEtype) = "M" Then crng = 0 Else crng = 1
End If
lDur = RoundUp(crng) 'crng = RoundUp(crng):
crng.NumberFormat = "General"
If IsNumeric(lDur) = False Then MsgBox msg(27): crng = 1
If Cells(crng.Row, cpg.ESD) = vbNullString Then
If st.HGC Then orgStartHrs = sArr.ResourceP(0, 10): startdatehour = Date + orgStartHrs: Cells(crng.Row, cpg.ED).Select
If st.HGC Then Cells(crng.Row, cpg.ESD) = startdatehour Else Cells(crng.Row, cpg.ESD) = Date
End If
If IsDate(Cells(crng.Row, cpg.ESD)) And IsNumeric(lDur) Then
If InStr(1, Cells(crng.Row, cpg.Dependency), "F_", vbTextCompare) > 0 Then
If st.HGC Then
Cells(crng.Row, cpg.ESD) = CalESDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.EED), lDur)
Else
Cells(crng.Row, cpg.ESD) = GetStartFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.EED)), CLng(lDur))
End If
Else
If st.HGC Then
Cells(crng.Row, cpg.EED) = CalEEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ESD), lDur)
Else
Cells(crng.Row, cpg.EED) = GetEndDateFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.ESD)), CLng(lDur))
End If
End If
End If
Last:
Call CalcDepforRow(crng.Row): Call CalcAutoPercforRow(crng): Call CalcResCostforRow(crng): bEstDatesChanged = True
End Sub

Sub CalcDepforRow(cRow As Long)
If Cells(cRow, cpg.Dependency) = "" And Cells(cRow, cpg.Dependents) = "" Then Exit Sub
Call CalcDepFormulas: Call ReCalcDepFormulas(cRow, True): Call ClearDepFormulas: bDependencies = True
End Sub

Sub Cell_ESDChanged(ByVal crng As Range)
Dim d As Date
On Error GoTo whoa
d = CDate(crng)
On Error Resume Next
overflowchecked:

If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(1): bPopParentTasks = True: Exit Sub
If InStr(1, Cells(crng.Row, cpg.Dependency), "S_", vbTextCompare) > 0 Then
MsgBox msg(28): crng = Cells(crng.Row, cpg.StartConstrain): bEstDatesChanged = True:Exit Sub
End If
Dim orgStartHrs As Double: Dim startdatehour As Date, ESD As Date, newESD As Date, selESD As Date: Dim answer As Integer
If st.HGC Then orgStartHrs = sArr.ResourceP(0, 10): startdatehour = Date + orgStartHrs: Cells(crng.Row, cpg.ED).Select
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then
MsgBox msg(20):
'If st.HGC Then Cells(crng.Row, cpg.ESD) = startdatehour Else Cells(crng.Row, cpg.ESD) = Date 'Formula
'End If
If IsError(crng) Then
MsgBox msg(76) & crng.Address:
If st.HGC Then crng = startdatehour Else crng = Date
End If
If IsError(crng) = True Or crng = vbNullString Or IsDate(crng) = False Or crng < csDate Or crng > ceDate Then
If st.HGC Then Cells(crng.Row, cpg.ESD) = startdatehour Else Cells(crng.Row, cpg.ESD) = Date
End If

If StartDateCheck = False Then
selESD = Cells(crng.Row, cpg.ESD): newESD = GetNewESD(Cells(crng.Row, cpg.Resource), selESD)
If newESD <> selESD Then
Application.ScreenUpdating = True
If st.HGC Then
answer = MsgBox(selESD & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & newESD & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & selESD & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
Else
answer = MsgBox(Format(selESD, "ddd dd-mmm-yy") & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & Format(newESD, "ddd dd-mmm-yy") & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & Format(selESD, "ddd dd-mmm-yy") & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
End If
If answer = vbNo Then crng = newESD Else crng = selESD
Application.ScreenUpdating = False
End If
End If

If Cells(crng.Row, cpg.GEtype) = "M" Then
Cells(crng.Row, cpg.ED) = 0: Cells(crng.Row, cpg.EED) = Cells(crng.Row, cpg.ESD): bEstDatesChanged = True:
End If
'crng = CDate(crng) 'Formula

Call Cell_EDChanged(Cells(crng.Row, cpg.ED))
Exit Sub
whoa:
If st.HGC Then Cells(crng.Row, cpg.ESD) = startdatehour Else Cells(crng.Row, cpg.ESD) = Date: GoTo overflowchecked
End Sub

Sub Cell_EEDChanged(ByVal crng As Range)
Dim d As Date
On Error GoTo whoa
d = CDate(crng)
On Error Resume Next
overflowchecked:

If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(1): bPopParentTasks = True: Exit Sub
If InStr(1, Cells(crng.Row, cpg.Dependency), "F_", vbTextCompare) > 0 Then
MsgBox msg(29): crng = Cells(crng.Row, cpg.EndConstrain):bEstDatesChanged = True:Exit Sub
End If
If Cells(crng.Row, cpg.GEtype) = "M" Then MsgBox msg(41): Cells(crng.Row, cpg.EED) = Cells(crng.Row, cpg.ESD): GoTo Last
Dim EED As Date, ESD As Date: Cells(crng.Row, cpg.ED).Select
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20):
'Call Cell_EDChanged(Cells(crng.Row, cpg.ED)): Exit Sub
If crng = vbNullString Or IsDate(crng) = False Or crng < csDate Or crng > ceDate Then
Call Cell_EDChanged(Cells(crng.Row, cpg.ED)): Exit Sub
End If

EED = crng: ESD = Cells(crng.Row, cpg.ESD)
If IsDate(ESD) And IsDate(EED) Then
If EED < ESD Then MsgBox msg(22): Call Cell_EDChanged(Cells(crng.Row, cpg.ED)): Exit Sub
If st.HGC Then
Cells(crng.Row, cpg.ED) = CalEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ESD), Cells(crng.Row, cpg.EED))
If Cells(crng.Row, cpg.ED) <= 0 Then Cells(crng.Row, cpg.ED) = 1: Call Cell_EDChanged(Cells(crng.Row, cpg.ED))
Else
Cells(crng.Row, cpg.ED) = GetWorkDaysFromDate(Cells(crng.Row, cpg.Resource), CDate(ESD), CDate(EED))
End If
End If
Last:
Call CalcDepforRow(crng.Row): Call CalcAutoPercforRow(crng): Call CalcResCostforRow(crng): bEstDatesChanged = True
Exit Sub
whoa:
Cells(crng.Row, cpg.EED) = Date: GoTo overflowchecked
End Sub

Sub CalcAutoPercforRow(ByVal crng As Range)
If st.PercAuto Then Cells(crng.Row, cpg.PercentageCompleted) = GetCP(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ESD), Cells(crng.Row, cpg.EED)) / 100
End Sub
Sub CalcResCostforRow(ByVal crng As Range)
If st.CalResCosts = False Then Exit Sub
If Cells(crng.Row, cpg.Resource) = "" Then Cells(crng.Row, cpg.ResourceCost) = "": Exit Sub
Dim resCost As Double: resCost = GetResourcesCost(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ED))
If resCost > 0 Then Cells(crng.Row, cpg.ResourceCost) = resCost Else Cells(crng.Row, cpg.ResourceCost) = ""
End Sub
Sub Cell_BSDChanged(ByVal crng As Range)
Dim newBSD As Date, selBSD As Date: Dim answer As Integer
If st.CalBasDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(45): bPopParentTasks = True: Exit Sub
bBasDatesChanged = True: Cells(crng.Row, cpg.BD).Select
If checkdates(crng, "b") = False Or crng = "" Then
If Cells(crng.Row, cpg.GEtype) = "M" Then Cells(crng.Row, cpg.BD) = vbNullString: Cells(crng.Row, cpg.BED) = vbNullString: Exit Sub
If IsDate(Cells(crng.Row, cpg.BED)) And IsNumeric(Cells(crng.Row, cpg.BD)) Then
If st.HGC Then
Cells(crng.Row, cpg.BSD) = CalESDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.BED), Cells(crng.Row, cpg.BD))
Else
Cells(crng.Row, cpg.BSD) = GetStartFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.BED)), CLng(Cells(crng.Row, cpg.BD)))
End If
Else
crng = vbNullString: Cells(crng.Row, cpg.BED) = vbNullString: Cells(crng.Row, cpg.BD) = vbNullString: Exit Sub
End If
End If

If StartDateCheck = False Then
selBSD = Cells(crng.Row, cpg.BSD): newBSD = GetNewESD(Cells(crng.Row, cpg.Resource), selBSD)
If newBSD <> selBSD Then
Application.ScreenUpdating = True
If st.HGC Then
answer = MsgBox(selBSD & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & newBSD & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & selBSD & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
Else
answer = MsgBox(Format(selBSD, "ddd dd-mmm-yy") & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & Format(newBSD, "ddd dd-mmm-yy") & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & Format(selBSD, "ddd dd-mmm-yy") & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
End If
If answer = vbNo Then crng = newBSD Else crng = selBSD
Application.ScreenUpdating = False
End If
End If

If Cells(crng.Row, cpg.GEtype) = "M" Then Cells(crng.Row, cpg.BD) = 0: Cells(crng.Row, cpg.BED) = Cells(crng.Row, cpg.BSD):bBasDatesChanged = True: Exit Sub
If IsDate(Cells(crng.Row, cpg.BSD)) Then Cell_BDChanged (Cells(crng.Row, cpg.BD))
If IsDate(Cells(crng.Row, cpg.BED)) Then If CDate(Cells(crng.Row, cpg.BSD)) > CDate(Cells(crng.Row, cpg.BED)) Then MsgBox msg(22): Cells(crng.Row, cpg.BED) = CDate(Cells(crng.Row, cpg.BSD))
Else
bBasDatesChanged = True
End If
End Sub
Sub Cell_ASDChanged(ByVal crng As Range)
Dim newASD As Date, selASD As Date: Dim answer As Integer
If st.CalActDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(47): bPopParentTasks = True: Exit Sub
bActDatesChanged = True
Cells(crng.Row, cpg.AD).Select
If checkdates(crng, "a") = False Or crng = "" Then
If Cells(crng.Row, cpg.GEtype) = "M" Then Cells(crng.Row, cpg.AD) = vbNullString: Cells(crng.Row, cpg.AED) = vbNullString: Exit Sub
If IsDate(Cells(crng.Row, cpg.AED)) And IsNumeric(Cells(crng.Row, cpg.AD)) Then
If st.HGC Then
Cells(crng.Row, cpg.ASD) = CalESDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.AED), Cells(crng.Row, cpg.AD))
Else
Cells(crng.Row, cpg.ASD) = GetStartFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.AED)), CLng(Cells(crng.Row, cpg.AD)))
End If
Else
crng = vbNullString: Cells(crng.Row, cpg.AED) = vbNullString: Cells(crng.Row, cpg.AD) = vbNullString: Exit Sub
End If
End If

If StartDateCheck = False Then
selASD = Cells(crng.Row, cpg.ASD): newASD = GetNewESD(Cells(crng.Row, cpg.Resource), selASD)
If newASD <> selASD Then
Application.ScreenUpdating = True
If st.HGC Then
answer = MsgBox(selASD & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & newASD & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & selASD & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
Else
answer = MsgBox(Format(selASD, "ddd dd-mmm-yy") & " is a holiday/ workoff day for this Resource." & vbNewLine & "Next available date is " & Format(newASD, "ddd dd-mmm-yy") & " ." & vbNewLine & vbNewLine & "Do you still want to set the task start date to" & vbNewLine & Format(selASD, "ddd dd-mmm-yy") & "? ", vbYesNo + vbQuestion, "Confirm Start Date")
End If
If answer = vbNo Then crng = newASD Else crng = selASD
Application.ScreenUpdating = False
End If
End If

If Cells(crng.Row, cpg.GEtype) = "M" Then Cells(crng.Row, cpg.AD) = 0: Cells(crng.Row, cpg.AED) = Cells(crng.Row, cpg.ASD):bActDatesChanged = True: Exit Sub
If IsDate(Cells(crng.Row, cpg.ASD)) Then Call Cell_ADChanged(Cells(crng.Row, cpg.AD))
If IsDate(Cells(crng.Row, cpg.AED)) Then If CDate(Cells(crng.Row, cpg.ASD)) > CDate(Cells(crng.Row, cpg.AED)) Then MsgBox msg(22): Cells(crng.Row, cpg.AED) = CDate(Cells(crng.Row, cpg.ASD))
Else
bActDatesChanged = True
End If
End Sub
Sub Cell_BEDChanged(ByVal crng As Range)
If st.CalBasDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(45): bPopParentTasks = True: Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" Then MsgBox msg(42): Cells(crng.Row, cpg.BED) = Cells(crng.Row, cpg.BSD): Cells(crng.Row, cpg.BD) = 0
bBasDatesChanged = True
Cells(crng.Row, cpg.BD).Select
If Cells(crng.Row, cpg.BSD) = vbNullString Then crng = vbNullString: Cells(crng.Row, cpg.BD) = vbNullString:MsgBox msg(25): Exit Sub
If checkdates(crng, "b") = False Then crng = vbNullString: Call Cell_BDChanged(Cells(crng.Row, cpg.BD)): Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" Then
Cells(crng.Row, cpg.BD) = 0
Else
If IsDate(Cells(crng.Row, cpg.BSD)) And IsDate(Cells(crng.Row, cpg.BED)) Then
If st.HGC Then
Cells(crng.Row, cpg.BD) = CalEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.BSD), Cells(crng.Row, cpg.BED))
If Cells(crng.Row, cpg.BD) = 0 Then Cells(crng.Row, cpg.BD) = 1: Call Cell_BDChanged(Cells(crng.Row, cpg.BD))
Else
Cells(crng.Row, cpg.BD) = GetWorkDaysFromDate(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.BSD)), CDate(Cells(crng.Row, cpg.BED)))
End If
End If
End If
Else
bBasDatesChanged = True
End If
End Sub
Sub Cell_AEDChanged(ByVal crng As Range)
If st.CalActDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(47): bPopParentTasks = True: Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" Then MsgBox msg(43): Cells(crng.Row, cpg.AED) = Cells(crng.Row, cpg.ASD): Cells(crng.Row, cpg.AD) = 0
bActDatesChanged = True
Cells(crng.Row, cpg.AD).Select
If Cells(crng.Row, cpg.ASD) = vbNullString Then crng = vbNullString: Cells(crng.Row, cpg.AD) = vbNullString:MsgBox msg(26): Exit Sub
If checkdates(crng, "a") = False Then crng = vbNullString: Call Cell_ADChanged(Cells(crng.Row, cpg.AD)): Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" Then
Cells(crng.Row, cpg.AD) = 0
Else
If IsDate(Cells(crng.Row, cpg.ASD)) And IsDate(Cells(crng.Row, cpg.AED)) Then
If st.HGC Then
Cells(crng.Row, cpg.AD) = CalEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ASD), Cells(crng.Row, cpg.AED))
If Cells(crng.Row, cpg.AD) = 0 Then Cells(crng.Row, cpg.AD) = 1: Call Cell_ADChanged(Cells(crng.Row, cpg.AD))
Else
Cells(crng.Row, cpg.AD) = GetWorkDaysFromDate(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.ASD)), CDate(Cells(crng.Row, cpg.AED)))
End If
End If
End If
Else
bActDatesChanged = True
End If
End Sub
Sub Cell_BDChanged(ByVal crng As Range)
Dim lDur As Long: lDur = RoundUp(crng)
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20)
If Cells(crng.Row, cpg.GEType) = "M" Then crng = 0 Else crng = 1
'End If
If st.CalBasDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(46): bPopParentTasks = True: Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" And crng <> 0 Then MsgBox msg(13): crng = 0
bBasDatesChanged = True
If Cells(crng.Row, cpg.BSD) = vbNullString Then Cells(crng.Row, cpg.BED) = vbNullString: crng = vbNullString: MsgBox msg(25): Exit Sub
If checkDuration(crng, "b") = False Then If Cells(crng.Row, cpg.GEtype) = "M" Then crng = 0 Else crng = 1
crng.NumberFormat = "General"'crng = RoundUp(crng):
If IsDate(Cells(crng.Row, cpg.BSD)) And IsNumeric(lDur) Then
If st.HGC Then
Cells(crng.Row, cpg.BED) = CalEEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.BSD), lDur)
Else
Cells(crng.Row, cpg.BED) = GetEndDateFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.BSD)), CLng(lDur))
End If
End If
Else
bBasDatesChanged = True
End If
End Sub

Sub Cell_ADChanged(ByVal crng As Range)
Dim lDur As Long: lDur = RoundUp(crng)
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20)
If Cells(crng.Row, cpg.GEType) = "M" Then crng = 0 Else crng = 1
'End If
If st.CalActDates Then
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(48): bPopParentTasks = True: Exit Sub
If Cells(crng.Row, cpg.GEtype) = "M" And crng <> 0 Then MsgBox msg(13): crng = 0
bActDatesChanged = True
If Cells(crng.Row, cpg.ASD) = vbNullString Then Cells(crng.Row, cpg.AED) = vbNullString: crng = vbNullString: MsgBox msg(26): Exit Sub
If checkDuration(crng, "a") = False Then If Cells(crng.Row, cpg.GEtype) = "M" Then crng = 0 Else crng = 1
crng.NumberFormat = "General": 'crng = RoundUp(crng):
If IsDate(Cells(crng.Row, cpg.ASD)) And IsNumeric(lDur) Then
If st.HGC Then
Cells(crng.Row, cpg.AED) = CalEEDHrs(Cells(crng.Row, cpg.Resource), Cells(crng.Row, cpg.ASD), lDur)
Else
Cells(crng.Row, cpg.AED) = GetEndDateFromWorkDays(Cells(crng.Row, cpg.Resource), CDate(Cells(crng.Row, cpg.ASD)), CLng(lDur))
End If
End If
Else
bActDatesChanged = True
End If
End Sub
Function checkdates(ByVal crng As Range, dateType As String) As Boolean
If IsError(crng) = True Then crng = vbNullString: GoTo Last
Call CheckRngFormula(crng, True) 'CheckRngFormula(crng) Then MsgBox msg(20): crng = vbNullString: GoTo last
If crng = vbNullString Or IsDate(crng) = False Then crng = vbNullString: GoTo Last
If CDate(crng) < csDate Or CDate(crng) > ceDate Then Call ShowOutofDatesMessage: crng = vbNullString: GoTo Last
If dateType = "b" And st.CalBasDates = False Then
If IsDate(Cells(crng.Row, cpg.BED)) Then If CDate(Cells(crng.Row, cpg.BSD)) > CDate(Cells(crng.Row, cpg.BED)) Then MsgBox msg(22): Cells(crng.Row, cpg.BED) = CDate(Cells(crng.Row, cpg.BSD)) ': GoTo last
End If
If dateType = "a" And st.CalActDates = False Then
If IsDate(Cells(crng.Row, cpg.AED)) Then If CDate(Cells(crng.Row, cpg.ASD)) > CDate(Cells(crng.Row, cpg.AED)) Then MsgBox msg(22): Cells(crng.Row, cpg.AED) = CDate(Cells(crng.Row, cpg.ASD)) ': GoTo last
End If
checkdates = True: Exit Function
Last:
checkdates = False
End Function

Function checkDuration(ByVal crng As Range, dateType As String) As Boolean
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20)
If IsError(crng) = True Or crng = "" Or crng = 0 Or IsNumeric(crng) = False Or crng <= 1 Or crng >= 10000 Then
If dateType = "b" Then If st.CalBasDates Then GoTo Last
If dateType = "a" Then If st.CalActDates Then GoTo Last
End If
'crng = RoundUp(crng):
crng.NumberFormat = "General": checkDuration = True: Exit Function
Last:
checkDuration = False
End Function
Sub Cell_ResourceChanged(ByVal crng As Range)
If IsParentTask(crng.Row) Then MsgBox msg(3)
If CheckRngFormula(crng, False) = True Then Cells(crng.Row, cpg.Resource) = vbNullString
crng.Select:
Dim vStr, i As Integer, sresources As String: Dim resfound As Boolean: Dim j As Long
crng.value = Trim(crng.value):: resfound = False
If Cells(crng.Row, cpg.Resource) = vbNullString Then
If st.CalResCosts Then Cells(crng.Row, cpg.ResourceCost) = vbNullString
Else
vStr = Split(crng & sResourceSeperator, sResourceSeperator)
For i = 0 To UBound(vStr) - 1
For j = LBound(sArr.ResourceP) To UBound(sArr.ResourceP)
If LCase(vStr(i)) = sArr.ResourceP(j, 0) Then resfound = True: Exit For
Next j
If resfound = False Then
MsgBox msg(18)
Else
If InStr(1, sResourceSeperator & sresources, sResourceSeperator & vStr(i) & sResourceSeperator, vbTextCompare) = 0 Then
sresources = sresources & vStr(i) & sResourceSeperator: resfound = False 'prevent same resource name being added twice
End If
End If
Next i
End If
If InStr(1, crng, ",") > 0 And IsParentTask(crng.Row) = False Then MsgBox msg(11)
If sresources = vbNullString Then crng = vbNullString Else crng = Left(sresources, Len(sresources) - 2)
If IsParentTask(crng.Row) Then Exit Sub

Last:
Call Cell_EDChanged(Cells(crng.Row, cpg.ED)): bResChanged = True
If Cells(crng.Row, cpg.BSD) <> vbNullString Then Call Cell_BDChanged(Cells(crng.Row, cpg.BD))
If Cells(crng.Row, cpg.ASD) <> vbNullString Then Call Cell_ADChanged(Cells(crng.Row, cpg.AD))
End Sub

Sub cell_PercentageCompleted(ByVal crng As Range)
If IsError(crng) Then MsgBox msg(76) & crng.Address: crng = 0
If st.PercAuto Then
MsgBox msg(10): Call CalcAutoPercforRow(crng)
Else
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then MsgBox msg(2): bPopParentTasks = True: Exit Sub
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20) ': crng = vbNullString
If crng = vbNullString Or IsNumeric(crng) = False Then crng = 0
If crng < 0 Then crng = 0
If crng > 1 Then crng = 1
End If
bPercChanged = True
End Sub

Sub cell_WorkChanged(ByVal crng As Range)
If IsParentTask(crng.Row) Then MsgBox msg(62): bPopParentTasks = True: Exit Sub
If IsError(crng) = True Or IsNumeric(crng) = False Or crng = vbNullString Or crng <= 0 Then crng = vbNullString
Call CheckRngFormula(crng, True) 'If CheckRngFormula(crng) Then MsgBox msg(20): 'crng = vbNullString
bWorkChanged = True
End Sub

Sub cell_BCSChanged(ByVal crng As Range)
Call checkCostCell(crng)
End Sub
Sub Cell_ECSChanged(ByVal crng As Range)
Call checkCostCell(crng)
End Sub
Sub cell_ACSChanged(ByVal crng As Range)
Call checkCostCell(crng)
End Sub
Sub checkCostCell(ByVal crng As Range)
If st.CalParCosts Then
If IsParentTask(crng.Row) Then MsgBox msg(5): bPopParentTasks = True: Exit Sub
If IsError(crng) = True Or IsNumeric(crng) = False Or crng = vbNullString Or crng <= 0 Then crng = vbNullString
Call CheckRngFormula(crng, True) ' crng = vbNullString
End If
bCostsChanged = True
End Sub
Sub Cell_ResourceCost(ByVal crng As Range)
If st.CalParCosts Then If IsParentTask(crng.Row) Then MsgBox msg(6): bPopParentTasks = True: Exit Sub
If st.CalResCosts Then
MsgBox msg(8):
Call CalcResCostforRow(crng):
End If
bCostsChanged = True
End Sub
Sub Cell_NotesChanged(ByVal crng As Range)
Cells(crng.Row, cpg.Notes).NumberFormat = "General"
End Sub
Sub Cell_LCChanged(ByVal crng As Range)
MsgBox msg(49): crng = ""
End Sub

Sub Cell_PriorityChanged(ByVal crng As Range)
Call CheckRngFormula(crng, True)
If IsError(crng) Then MsgBox msg(76) & crng.Address: crng = "NORMAL"
If crng = vbNullString Then crng = "NORMAL"
If Left(crng.Formula, 1) <> "=" Then
If LCase(crng) <> "high" And LCase(crng) <> "normal" And LCase(crng) <> "low" Then MsgBox msg(19): crng = "Normal" ' priority can be set High, Normal or Low
crng = UCase(crng)
End If
bPriorityChanged = True
End Sub
Sub Cell_StatusChanged(ByVal crng As Range)
If IsError(crng) Then MsgBox msg(76) & crng.Address: crng = ""
MsgBox msg(12): bStatusChanged = True
End Sub
Sub Cell_WBSChanged(ByVal crng As Range)
MsgBox msg(7)
bWBSChanged = True
End Sub
Sub Cell_TaskIconChanged(ByVal crng As Range)
MsgBox msg(30): bTaskIconChanged = True
With Cells(cRow, cpg.TaskIcon)
.value = "u": .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
If Cells(cRow, cpg.GEtype) = "T" Then .Font.Name = "Wingdings 3" Else .Font.Name = "Wingdings": .Font.size = 11
End With
End Sub
Sub Cell_DoneChanged(ByVal crng As Range)
MsgBox msg(59): bDoneChanged = True
With Cells(cRow, cpg.Done)
.value = 0: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
End With
End Sub

Sub Cell_WBSPredecessor(ByVal crng As Range)
MsgBox msg(63): bWBSPredecessor = True
End Sub

Sub Cell_WBSSuccessor(ByVal crng As Range)
MsgBox msg(63): bWBSSuccessor = True
End Sub

Sub bForceDraw(Optional t As String)
MsgBox msg(77): bForceReDraw = True
End Sub
95349HC4T1789B0FH3AF045
About
Buy Pro Version
Select Date
Delect Gantt Chart
Active License
Gantt Excel
Add Gantt Chart
Priority
Resources
Select Resources
UserForm1
Gantt Excel
Settings
Status
Add Edit Task
Timeline
Welcome to Gantt Excel 
Option Explicit
#If Mac Then
#If VBA7 Then
Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
#Else
Private Declare Function popen Lib "/usr/lib/libc.dylib" (ByVal command As String, ByVal mode As String) As Long
Private Declare Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As Long) As Long
Private Declare Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal items As Long, ByVal stream As Long) As Long
Private Declare Function feof Lib "/usr/lib/libc.dylib" (ByVal file As Long) As Long
#End If
#End If
Public Const sVersionNo As String = "v 4.28"
Public Const min_ASC As Integer = 32
Public Const Max_ASC As Integer = 126
Public Const No_of_Chars As Integer = Max_ASC - min_ASC + 1
Public dFormDate As Date
Public sResourcesNamesFromForm As String, formResname As String, cTaskWBSBeingAdded As String, taskTag As String, lLicDur As String
Public sfUsrName As String, sfUsrEmailID As String, sfLicenseCode As String, pstrLicType As String, sTempStr As String, sTempStr1 As String
Public sRowOrWidth As String, bStopCalculationOfConstraints As Boolean, bAllowFormulas As Boolean, bOcee As Boolean
Public bAddTask As Boolean, bEditTask As Boolean, bAddMilestone As Boolean, bEditMilestone As Boolean
Public bTriggerFromForm As Boolean, bNoDateSelectedForForm As Boolean, bClosing As Boolean, ResSelectorDB As Boolean
Public disabledStatus As Boolean, enabledStatus As Boolean, bDeleteAllAndDrawGB As Boolean, bHideTime As Boolean, bAddResTaskForm As Boolean
Public StartDateCheck As Boolean, bImportGC As Boolean, bImportLic As Boolean, AddNewGC As Boolean
Public tlogg As Boolean, dlogg As Boolean, bESD As Boolean, bEED As Boolean, bBSD As Boolean, bBED As Boolean, bASD As Boolean, bAED As Boolean
Public bShowAll As Boolean, bAddProject As Boolean, bEditProject As Boolean, ResArraysReady As Boolean
Public clrow As Long, barTextColNo As Long, dbcRow As Long, dbcCol As Long, resvalue As Long, lIndentLevel As Long, SelectedRow As Long, SelTaskRow As Long
Public currentSheet As Worksheet, rs As Worksheet, gs As Worksheet
Public Const wbPass As String = "GanttExcel"
Public Const productName As String = "Gantt Excel 2023 Pro - "
Public Const productTypeFree As String = "Free Version "
Public Const productTypeDP As String = "Daily Planner "
Public Const productTypeDPM As String = "Daily Planner - Mac Version "
Public Const productTypeHP As String = "Hourly Planner "
Public Const productTypeHPM As String = "Hourly Planner - Mac Version "
Public Const productTypeHD As String = "Hourly Daily Planner "
Public Const productTypeHDM As String = "Hourly Daily Planner - Mac Version "
Public Const sAddTaskPlaceHolder As String = "Type here to add a new task"
Public Const sLFreeKey As String = "N4fL3sgyYf-Gm2xds9E-6dG7b72mevs" '.ForeColor.RGB = RGB(180, 180, 180)
Public Const Ocee As String = "psc222914"
Public Const pstrFree As String = "T42VP"
Public Const pstrDP As String = "10D6W"
Public Const pstrDPM As String = "10D2M"
Public Const pstrHP As String = "10H8W"
Public Const pstrHPM As String = "10H3M"
Public Const pstrHD As String = "10DH2W"
Public Const pstrHDM As String = "10DH4M"
Public Const sCProD As String = "D68GTU"
Public Const sCProDMac As String = "DMT7HJ"
Public Const sCProH As String = "HYT2KL"
Public Const sCProHMac As String = "HM3R5T"
Public Const sCProHD As String = "HDU7OP"
Public Const sCProHDMac As String = "HDM8U3"
Public Const DepSeperator As String = "|"
Public Const sResourceSeperator As String = ", "
Public Const rowOnly As String = "rowonly": Public Const aboveFamily As String = "abovefamily": Public Const family As String = "family"
Public Const allRows As String = "allrows"
Public Const estBars As String = "estbars": Public Const basBars As String = "basbars": Public Const actBars As String = "actbars": Public Const allBars As String = "allbars"
Public Const allFields As String = "allfields"
Public Const fCosts As String = "costs"
Public Const fWork As String = "work"
Public Const fPerc As String = "perc"
Public Const fPriority As String = "priority"
Public Const fStatus As String = "status"
Public Const allDates As String = "allDates"
Public Const estDates As String = "estDates"
Public Const basDates As String = "basDates"
Public Const actDates As String = "actDates"
Public Const CheckSymbol As String = "Ÿ"
Public Const UnCheckSymbol As String = "¬"
Public Const Org As String = "Organization"
Public Const cLicConst As String = "a999z"
Public Const cFreeVersionTasksCount As Long = 15
Public Const cHeaderName As String = "WBS"
Public Const trc As Long = 7
Public cps As New clsGetColNosGST
Public st As New clsSettings
Public cpg As New clsGetColNosGCT
Public cpd As New clsGetColNosGDD
Public cpt As New clsGetColNosTimeline
Public sArr As New clsArrProp
Public Const rowone As Long = 1
Public Const rowtwo As Long = 2
Public Const rowsix As Long = 6
Public Const rowseven As Long = 7
Public Const roweight As Long = 8
Public Const rownine As Long = 9
Public Const taskfontsize As Long = 10
Public Const firsttaskrow As Long = 10
Public Const taskRowHeight As Long = 18
Public Const maxHCol As Long = 480
Public Const maxDCol As Long = 720
Public Const sPassD = vbNullString
Public Const csYear As Long = 1950
Public Const ceYear As Long = 2100
Public Const csDate As Date = #1/1/1950#
Public Const ceDate As Date = #12/1/2100#
Public Const siteURL As String = "https://www.ganttexcel.com/"
Public Const freeURL1 As String = "free-version-o1/"
Public Const freeURL2 As String = "free-version-n23/"
Public Const freeURL3 As String = "free-version-p34/"
Public Const freeURL4 As String = "free-version-q45/"
Public Const macroURL As String = "enable-macros/"
Public Const docURL As String = "documentation/"
Public Const howtoURL As String = "how-to-create-a-gantt-chart-in-excel/"
Public Const buyURL As String = "buy/"
Public Const contactURL As String = "contact-us/"
Public Const sourceURL As String = "?utm_source=XL"
Public Const mediumURL As String = "&utm_medium="
Public tidArr()
Public newHolidaysArray()
Public newResourcesArray()
Public newWorkdaysArray()

Function GCcolumns()
GCcolumns = Array("zero", "GEType", "TID", "Dependency", "Dependents", "StartConstrain", "EndConstrain", _
TIL, "SS", "TaskIcon", "WBS", "Task", "Priority", "Status", "Resource", "ResourceCost", "BSD", "BED", "BD", "ESD", "EED", "ED", _
WBSPredecessors, "WBSSuccessors", _
Work, "Done", "PercentageCompleted", "ASD", "AED", "AD", "BCS", "ECS", "ACS", "Notes", _
TColor, "TPColor", "BLColor", "ACColor", _
Custom 1, "Custom 2", "Custom 3", "Custom 4", "Custom 5", "Custom 6", "Custom 7", "Custom 8", "Custom 9", "Custom 10", _
Custom 11, "Custom 12", "Custom 13", "Custom 14", "Custom 15", "Custom 16", "Custom 17", "Custom 18", "Custom 19", "Custom 20", "ShapeInfoE", "ShapeInfoB", "ShapeInfoA", "LC")
End Function

Function GCcolumnsEngName()
GCcolumnsEngName = Array("zero", "GEType", "TID", "Dependency", "Dependents", "StartConstrain", "EndConstrain", _
, "", "", "WBS", "Task", "Priority", "Status", "Resource", "Resource Cost", "Baseline Start", "Baseline End", "Baseline Duration", "Start", "Finish", "Duration", _
WBS Predecessors, "WBS Successors", _
Work, "Done", "% Complete", "Actual Start", "Actual End", "Actual Duration", "Baseline Cost", "Est. Cost", "Actual Cost", "Notes", _
Bar Color, "% Color", "Baseline Color", "Actual Color", _
Custom 1, "Custom 2", "Custom 3", "Custom 4", "Custom 5", "Custom 6", "Custom 7", "Custom 8", "Custom 9", "Custom 10", _
Custom 11, "Custom 12", "Custom 13", "Custom 14", "Custom 15", "Custom 16", "Custom 17", "Custom 18", "Custom 19", "Custom 20", "ShapeInfoE", "ShapeInfoB", "ShapeInfoA", "")
End Function

Function parCheck(cRow As Long, cCol As Long, ctype As String) As Boolean
If IsParentTask(cRow) = False Then Exit Function
If cCol = cpg.ESD Or cCol = cpg.EED Or cCol = cpg.BSD Or cCol = cpg.BED Or cCol = cpg.ASD Or cCol = cpg.AED Then Call myMsgBox(cRow, cCol, 1): parCheck = True: Exit Function
If cCol = cpg.PercentageCompleted Then Call myMsgBox(cRow, cCol, 2): parCheck = True: Exit Function
If cCol = cpg.Resource Then MsgBox msg(3): parCheck = False: Exit Function
If cCol = cpg.ED Or cCol = cpg.BD Or cCol = cpg.AD Then Call myMsgBox(cRow, cCol, 4): parCheck = True: Exit Function
If st.CalParCosts Then
If cCol = cpg.ECS Or cCol = cpg.BCS Or cCol = cpg.ACS Then Call myMsgBox(cRow, cCol, 5): parCheck = True: Exit Function
If cCol = cpg.ResourceCost Then Call myMsgBox(cRow, cCol, 6): parCheck = True: Exit Function
Else
Exit Function
End If
End Function

Function CheckForDependency(ws As Worksheet) As Boolean
CheckForDependency = False
If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(rownine + 1, 3), ws.Cells(GetLastRow + 2, 6))) > 0 Then CheckForDependency = True
End Function
Function GetFolder() As String
Dim fldr As FileDialog:Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
.Title = "Select a Folder"
.AllowMultiSelect = False
.InitialFileName = Application.DefaultFilePath & "\"
If .show <> -1 Then GoTo NextCode
sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem:Set fldr = Nothing
End Function
Function GetFolderMac() As String
Dim folderPath As String, RootFolder As String, scriptstr As String
On Error Resume Next
RootFolder = MacScript("return (path to desktop folder) as String")
If val(Application.Version) < 15 Then
scriptstr = "(choose folder with prompt ""Select the folder""" & _
" default location alias """ & RootFolder & """) as string"
Else
scriptstr = "return posix path of (choose folder with prompt ""Select the folder""" & _
" default location alias """ & RootFolder & """) as string"
End If
GetFolderMac = MacScript(scriptstr)
On Error GoTo 0
End Function

Function CheckSheet(ByVal sSheetName As String) As Boolean ' uses active wb
Dim oSheet As Excel.Worksheet
Dim bReturn As Boolean
For Each oSheet In ActiveWorkbook.Sheets
If oSheet.Name = sSheetName Then bReturn = True: Exit For
Next oSheet
CheckSheet = bReturn
If CheckSheet = False Then Exit Function
End Function

Function WSExists(sCodeName As String) As Boolean ' uses this wb
Dim ws As Object
For Each ws In ThisWorkbook.Sheets
If ws.CodeName = sCodeName Then WSExists = True: Exit Function
Next
End Function
Function RoundUp(ByVal value As Double)
If Int(value) = value Then RoundUp = value Else RoundUp = Int(value) + 1
End Function
Function findSettingRow(ResourceName As String) As Integer
Set gs = setGSws: Set rs = setRSws
Dim resnumber As Double
If IsNumeric(ResourceName) Then
resnumber = CDbl(ResourceName)
On Error Resume Next
findSettingRow = Application.WorksheetFunction.Match(resnumber, rs.Range("A1:A10000"), 0)
On Error GoTo 0
Else
On Error Resume Next
findSettingRow = Application.WorksheetFunction.Match(ResourceName, rs.Range("A1:A10000"), 0)
On Error GoTo 0
End If
If findSettingRow <= 0 Then findSettingRow = 2 'Set organizational
End Function
Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
Const BGColor As Long = 13160660
Const ColorIndexLast As Long = 32
Dim myOrgColor As Double, myNewColor As Double:Dim myRGB_R As Integer, myRGB_G As Integer, myRGB_B As Integer
myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
If i_OldColor = xlNone Then
Color2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
Else
Color2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
End If
If Application.Dialogs(xlDialogEditColor).show(ColorIndexLast, myRGB_R, myRGB_G, myRGB_B) = True Then
PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
Else
PickNewColor = i_OldColor
End If
End Function
Sub Color2RGB(ByVal i_Color As Long, _
o_R As Integer, o_G As Integer, o_B As Integer)
o_R = i_Color Mod 256
i_Color = i_Color \ 256
o_G = i_Color Mod 256
i_Color = i_Color \ 256
o_B = i_Color Mod 256
End Sub

Function DashboardSheet(ws As Worksheet) As Boolean
If ws.Range("A1") = "UserDashSheet" Then DashboardSheet = True
End Function
Function IsAnyOpenWorkbookProtected(Optional t As Boolean) As Boolean
Dim oWin As Object
On Error GoTo Last
Set oWin = CallByName(Application, "ProtectedViewWindows", VbGet)
If Not oWin Is Nothing Then
IsAnyOpenWorkbookProtected = oWin.Count > 0
End If
Last:
On Error GoTo 0
Set oWin = Nothing
End Function
Function DA()
If tlogg Then
If disabledStatus = True Then
Debug.Print ("-Already Disabled - Disabling Again-")
Else
Debug.Print ("-Disabling All-")
End If
End If
With Application
.ScreenUpdating = False:.EnableEvents = False:.DisplayAlerts = False:.Calculation = xlCalculationManual
End With
disabledStatus = True: enabledStatus = False
End Function
Function EA()
If tlogg Then
If enabledStatus = True Then
Debug.Print ("-Already Enabled - Enabling again-")
Else
Debug.Print ("-Enabling All-")
End If
End If
With Application
.EnableEvents = True:.DisplayAlerts = True:.Calculation = xlCalculationAutomatic:.ScreenUpdating = True:
End With
disabledStatus = False: enabledStatus = True
End Function
Public Function GetFirstDateOfWeek(ByVal curDate As Date, Optional lWeekStartDayNum As Long) As Date
If lWeekStartDayNum = 0 Then lWeekStartDayNum = 1 'Set gs = setGSws gs.Cells(rowtwo, cps.WeekStartDay).value
GetFirstDateOfWeek = (curDate - (WorksheetFunction.Weekday(curDate, 2) - lWeekStartDayNum))
If GetFirstDateOfWeek > curDate Then GetFirstDateOfWeek = GetFirstDateOfWeek - 7
End Function
Public Function GetLastDateOfWeek(ByVal curDate As Date, Optional lWeekStartDayNum As Long) As Date
GetLastDateOfWeek = GetFirstDateOfWeek(curDate) + 6
End Function
Function GetFirstDateOfMonth(Optional d As Date) As Date
If d = 0 Then d = Date
GetFirstDateOfMonth = DateSerial(Year(d), Int(Month(d)), 1)
End Function
Function GetLastDateOfMonth(Optional d As Date) As Date
If d = 0 Then d = Date
GetLastDateOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Function GetFirstDateOfQuarter(Optional d As Date) As Date
If d = 0 Then d = Date
GetFirstDateOfQuarter = DateSerial(Year(d), Int((Month(d) - 1) / 3) * 3 + 1, 1)
End Function
Function GetLastDateOfQuarter(Optional d As Date) As Date
If d = 0 Then d = Date
GetLastDateOfQuarter = DateSerial(Year(d), ((Int((Month(d) - 1) / 3) + 1) * 3) + 1, 1) - 1
End Function
Function GetFirstDateOfHalfYearly(Optional d As Date) As Date
If d = 0 Then d = Date
If Month(d) > 6 Then
GetFirstDateOfHalfYearly = DateSerial(Year(d), 7, 1)
Else
GetFirstDateOfHalfYearly = DateSerial(Year(d), 1, 1)
End If
End Function
Function GetLastDateOfHalfYearly(Optional d As Date) As Date
If d = 0 Then d = Date
If Month(d) > 6 Then
GetLastDateOfHalfYearly = DateSerial(Year(d), 12, 31)
Else
GetLastDateOfHalfYearly = DateSerial(Year(d), 6, 30)
End If
End Function
Function GetFirstDateOfYear(Optional d As Date) As Date
If d = 0 Then d = Date
GetFirstDateOfYear = DateSerial(Year(d), 1, 1)
End Function
Function GetLastDateOfYear(Optional d As Date) As Date
If d = 0 Then d = Date
GetLastDateOfYear = DateSerial(Year(d), 12, 31)
End Function
Function GetResourcesCost(sresources As String, Optional lDur) As Double
If sresources = vbNullString Then Exit Function
If st.CalResCosts = False Then Exit Function
If ResArraysReady = False Then Call RememberResArrays
Dim vR, i As Integer: Dim r As Range: Dim n As Long: Dim resnewcost As Double
vR = Split(sresources & sResourceSeperator, sResourceSeperator)
For i = 0 To UBound(vR) - 1
On Error Resume Next
err.Clear
n = LBound(newResourcesArray())
If err.Number = 0 Then
Else
End If
Call getResValue(vR(i), newResourcesArray): resnewcost = resnewcost + newResourcesArray(resvalue, 1)
On Error GoTo 0
Next
On Error Resume Next
If lDur = vbNullString Then lDur = 0
If IsNumeric(lDur) = False Then lDur = 0
If lDur > 0 Then GetResourcesCost = resnewcost * lDur
On Error GoTo 0
End Function

Function IsTask(cRow As Long) As Boolean
If Cells(cRow, cpg.GEtype) = "T" Then IsTask = True
End Function

Function IsOverdue(cRow As Long) As Boolean
If st.HGC Then
If CDate(Cells(cRow, cpg.EED)) < Now And Cells(cRow, cpg.PercentageCompleted) < 1 Then IsOverdue = True Else IsOverdue = False
Else
If CDate(Cells(cRow, cpg.EED)) < Date And Cells(cRow, cpg.PercentageCompleted) < 1 Then IsOverdue = True Else IsOverdue = False
End If
End Function



Function IsNormalTask(cRow As Long) As Boolean
Dim tidl As Long, nextTidl As Long: tidl = Cells(cRow, cpg.TIL): nextTidl = Cells(cRow + 1, cpg.TIL)
If tidl = 0 And nextTidl = 0 Then IsNormalTask = True Else IsNormalTask = False
End Function
Function IsParentTask(ByVal cRow As Long, Optional ws As Worksheet) As Boolean
If ws Is Nothing Then Set ws = ActiveSheet
Dim tidl As Long, nextTidl As Long: tidl = ws.Cells(cRow, cpg.TIL): nextTidl = ws.Cells(cRow + 1, cpg.TIL)
If nextTidl > tidl Then IsParentTask = True Else IsParentTask = False
End Function
Function IsDrivingTask(cRow As Long) As Boolean
If Cells(cRow, cpg.Dependents) <> "" Then IsDrivingTask = True Else IsDrivingTask = False
End Function
Function IsDrivenTask(cRow As Long) As Boolean
If Cells(cRow, cpg.Dependency) <> "" Then IsDrivenTask = True Else IsDrivenTask = False
End Function
Function GetWorkDaysFromDate(resname As String, sdate As Date, eDate As Date) As Long
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim c As Date, wd As Long, i As Long, b As Boolean
For c = sdate To eDate
wd = Weekday(c, 2)
If IsDateAWorkday(resvalue, wd) Then
If IsDateAHoliday(resvalue, c) Then GoTo NextDate
GetWorkDaysFromDate = GetWorkDaysFromDate + 1
End If
NextDate:
Next c
If GetWorkDaysFromDate = 0 Then GetWorkDaysFromDate = 1
End Function

Public Function GanttChart(Optional ws As Worksheet) As Boolean
If ws Is Nothing Then Set ws = ActiveSheet
If ws.Cells(1, 1) = "GEType" Or ws.Cells(1, 1) = "tType" Then GanttChart = True
End Function

Public Sub CalcColPosGST(Optional t As Boolean)
Set cps = Nothing
Set cps = New clsGetColNosGST
End Sub
Public Sub ReadSettings(Optional t As Boolean)
Set st = Nothing
Set st = New clsSettings
End Sub
Public Sub CalcColPosGCT(Optional t As Boolean)
Set cpg = Nothing
Set cpg = New clsGetColNosGCT
End Sub
Public Sub CalcColPosGDD(Optional t As Boolean)
Set cpd = Nothing
Set cpd = New clsGetColNosGDD
End Sub
Public Sub CalcColPosTimeline(Optional t As Boolean)
Set cpt = Nothing
Set cpt = New clsGetColNosTimeline
End Sub

Function GetLicType() As String
If GST.Cells(rowtwo, cps.tUsrName) <> vbNullString Then GetLicType = GST.Cells(rowtwo, cps.tLiType)
End Function

Function IsLicValid(Optional t As Boolean, Optional bBypass As Boolean) As Boolean
Dim tBool As Boolean, IsGSheet As Boolean: Dim sl As String
IsGSheet = False
If GST.Cells(rowtwo, cps.tUsrName) <> vbNullString Then tBool = GST.Cells(rowtwo, cps.tLicenseVal) Else tBool = False
If tBool = False And bBypass = True Then IsLicValid = False:Exit Function
If tBool = False Then
If sTempStr <> "OnStartUp" Then frmLicEntry.show
If sfLicenseCode = vbNullString Then GoTo Last
If GST.Cells(rowtwo, cps.tLicenseVal) Then GoTo Last
Application.EnableEvents = False
With GST
.Cells(rowtwo, cps.tLicenseVal) = 1: .Cells(rowtwo, cps.tLiType) = pstrLicType: .Cells(rowtwo, cps.tUsrActivatedDate).NumberFormat = "@"
.Cells(rowtwo, cps.tUsrActivatedDate) = Decipher(Format(Date, "YYYYMMDD")): .Cells(rowtwo, cps.tUsrEmailID).NumberFormat = "@"
.Cells(rowtwo, cps.tUsrEmailID) = Decipher(sfUsrEmailID): .Cells(rowtwo, cps.tUsrName).NumberFormat = "@": .Cells(rowtwo, cps.tUsrName) = (sfUsrName)
.Cells(rowtwo, cps.tliky) = sfLicenseCode
If IsDate(lLicDur) Then
.Cells(rowtwo, cps.tLicDuration) = Decipher(CStr(Format(lLicDur, "dd-mmm-yyyy hh:mm:ss")) & cLicConst)
ElseIf lLicDur = "lifetime" Then
.Cells(rowtwo, cps.tLicDuration) = Decipher(CStr(Format(Now() + 36500, "dd-mmm-yyyy hh:mm:ss")) & cLicConst)
End If
End With
Application.EnableEvents = True: Call ThankYouMsg

If sTempStr = vbNullString Then RefreshRibbon
sTempStr = vbNullString: IsLicValid = True: GoTo Last
End If
If IsGSheet = False Then IsLicValid = True: GoTo Last
pstrLicType = GST.Cells(rowtwo, cps.tLiType)
If FreeVersion Then
If Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A")) - 3 >= cFreeVersionTasksCount Then
sTempStr1 = msg(80) & msg(82)
frmBuyPro.show
IsLicValid = 0
GoTo Last
Else
IsLicValid = True
End If
ElseIf pstrLicType = vbNullString Then
GoTo Last
Else
IsLicValid = True
End If
Last:
End Function
Function FreeVersion(Optional t As Boolean) As Boolean
If GetLicType = pstrFree Then FreeVersion = True
End Function
Sub ShowLimitation(Optional t As Boolean)
sTempStr1 = msg(81) & msg(82)
frmBuyPro.show
End Sub
Sub AddNewLicense(Optional t As String)
sTempStr = "Upgrade"
If GetLicType = vbNullString And GetGCCount = 1 Then Call IsLicValid Else Call UpgradeLicenseToPro
sTempStr = vbNullString
End Sub
Sub UpgradeLicenseToPro(Optional t As Boolean, Optional bBypass As Boolean)
Dim licstring As String, strCampaign As String
Dim curWs As Worksheet: Set curWs = ActiveSheet
licstring = GetLicType
If licstring = pstrDP Or licstring = pstrDPM Or licstring = pstrHP Or licstring = pstrHPM Or licstring = pstrHD Or licstring = pstrHDM Then
MsgBox msg(73)
GoTo Last
Else
If bBypass = 0 Then frmLicEntry.show
If sfLicenseCode = vbNullString Then GoTo Last
If GST.Cells(rowtwo, cps.tLicenseVal) Then GoTo Last
Application.EnableEvents = False
On Error Resume Next
strCampaign = Worksheets("Help").Range("D5")
On Error GoTo 0
If strCampaign = "cpc" Then strCampaign = "Cpc"
If strCampaign = "org" Then strCampaign = "Org"
If strCampaign = "cus" Then strCampaign = "Cus"
If Not bOcee Then
If strCampaign = "Cpc" Or strCampaign = "Org" Then strCampaign = strCampaign & "Cus"
End If
If strCampaign = "" Then strCampaign = "NotSure"
With GST
.Cells(rowtwo, cps.tLicenseVal) = 1: .Cells(rowtwo, cps.tLiType) = pstrLicType: .Cells(rowtwo, cps.tUsrActivatedDate).NumberFormat = "@"
.Cells(rowtwo, cps.tUsrActivatedDate) = Decipher(Format(Date, "YYYYMMDD"))
.Cells(rowtwo, cps.Campaign) = strCampaign
With GST.Cells(rowtwo, cps.tUsrEmailID)
.NumberFormat = "@": .value = Decipher(sfUsrEmailID)
End With
.Cells(rowtwo, cps.tUsrName).NumberFormat = "@": .Cells(rowtwo, cps.tUsrName) = sfUsrName: .Cells(rowtwo, cps.tliky) = sfLicenseCode

If IsDate(lLicDur) Then
.Cells(rowtwo, cps.tLicDuration) = Decipher(CStr(Format(lLicDur, "dd-mmm-yyyy hh:mm:ss")) & cLicConst)
ElseIf lLicDur = "lifetime" Then
End If
End With
If CheckSheet("Help") Then Call resetHelp
curWs.Activate
ThisWorkbook.Save: Application.EnableEvents = True: Call ThankYouMsg
End If
Last:
Call RefreshRibbon
If bOcee Then AddNewGC = True: Call TriggerAddNewSheet
End Sub

Sub ThankYouMsg(Optional t As String)
If pstrLicType = pstrFree Then
MsgBox msg(70) & vbLf & msg(71), vbInformation, "Information": Exit Sub
ElseIf pstrLicType = pstrDP Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeDP, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
ElseIf pstrLicType = pstrDPM Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeDPM, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
ElseIf pstrLicType = pstrHP Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeHP, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
ElseIf pstrLicType = pstrHPM Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeHPM, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
ElseIf pstrLicType = pstrHD Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeHD, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
ElseIf pstrLicType = pstrHDM Then
MsgBox msg(70) & vbLf & msg(72) & productName & productTypeHDM, vbInformation, "Information": MsgBox msg(74), vbInformation, "Information"
End If
End Sub

Sub RemoveLicense(Optional t As Boolean)
Call DA
With GST
.Cells(rowtwo, cps.tLicenseVal) = 0:.Cells(rowtwo, cps.tLiType) = vbNullString:.Cells(rowtwo, cps.tUsrName) = vbNullString:.Cells(rowtwo, cps.tUsrEmailID) = vbNullString
.Cells(rowtwo, cps.tUsrActivatedDate) = vbNullString:.Cells(rowtwo, cps.tliky) = "-":.Cells(rowtwo, cps.tLicDuration) = vbNullString
End With
Call EA
End Sub
Function IsDataCollapsed(Optional t As Boolean) As Boolean
Dim cRow As Long, r As Range, lrow As Long
cRow = firsttaskrow
lrow = GetLastRow: If lrow = 9 Then Exit Function
Set r = Range(Cells(cRow, cpg.WBS), Cells(lrow, cpg.WBS))
If Cells(rownine, cpg.WBS).ColumnWidth = 0 Then
r.Columns.AutoFit
End If
If r.Cells.Count = 1 Then Exit Function
If r.SpecialCells(xlCellTypeVisible).Cells.Count <> lrow - rownine Then
If MsgBox("You cannot perform this action when task groups are collapsed or filtered." & vbLf & _
"Do you want to continue by expanding or unfiltering them?", vbQuestion + vbYesNo, "Information") = vbYes Then
ExpandAlLGroups
Exit Function
Else
IsDataCollapsed = True
Exit Function
End If
End If
End Function
Function GetDateFormatOverride(sdf As String) As String
GetDateFormatOverride = Replace(Replace(sdf, "/", "\/"), "-", "\-")
End Function
Function GetDateFormatForDisplay(sdf As String) As String
GetDateFormatForDisplay = Replace(Replace(sdf, "\/", "/"), "\-", "-")
End Function
Sub DisableKeysDummy(Optional t As Boolean)
'has been blank forever
End Sub
Function GetNextIDNumber(Optional t As Boolean) As Long
GetNextIDNumber = Application.WorksheetFunction.Max(Range(Cells(firsttaskrow, cpg.TID), Cells(Cells.Rows.Count, cpg.TID))) + 1
End Function
Function MoveAsc(ByVal a, ByVal mLvl)
mLvl = mLvl Mod No_of_Chars
a = a + mLvl
If a < min_ASC Then
a = a + No_of_Chars
ElseIf a > Max_ASC Then
a = a - No_of_Chars
End If
MoveAsc = a
End Function
Function Decipher(ByVal s As String, Optional ByVal key As String) As String
Dim p, keyPos, c, E, k, chkSum
key = "GET"
If key = vbNullString Then
Decipher = s
Exit Function
End If
For p = 1 To Len(s)
If Asc(Mid(s, p, 1)) < min_ASC Or Asc(Mid(s, p, 1)) > Max_ASC Then Exit Function
Next p
For keyPos = 1 To Len(key)
chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
Next keyPos
keyPos = 0
For p = 1 To Len(s)
c = Asc(Mid(s, p, 1))
keyPos = keyPos + 1
If keyPos > Len(key) Then keyPos = 1
k = Asc(Mid(key, keyPos, 1))
c = MoveAsc(c, k)
c = MoveAsc(c, k * Len(key))
c = MoveAsc(c, chkSum * k)
c = MoveAsc(c, p * k)
c = MoveAsc(c, Len(s) * p)
E = E & Chr(c)
Next p
Decipher = E
End Function
Function UnDecipher(ByVal s As String, Optional ByVal key As String) As String
Dim p, keyPos, c, d, k, chkSum
key = "GET"
If key = vbNullString Then
UnDecipher = s
Exit Function
End If
For keyPos = 1 To Len(key)
chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
Next keyPos
keyPos = 0
For p = 1 To Len(s)
c = Asc(Mid(s, p, 1))
keyPos = keyPos + 1
If keyPos > Len(key) Then keyPos = 1
k = Asc(Mid(key, keyPos, 1))

c = MoveAsc(c, -(Len(s) * p))
c = MoveAsc(c, -(p * k))
c = MoveAsc(c, -(chkSum * k))
c = MoveAsc(c, -(k * Len(key)))
c = MoveAsc(c, -k)
d = d & Chr(c)
Next p
UnDecipher = d
End Function
Function GetGCCount() As Long
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then GetGCCount = GetGCCount + 1
Next ws
End Function
Function GetLastRow(Optional ws As Worksheet) As Long ' latest
If ws Is Nothing Then GetLastRow = Evaluate("counta(A:A)+8") - 2 Else GetLastRow = Evaluate("counta(" & "'" & ws.Name & "'!" & "A:A)+8") - 2
End Function
Sub OpenHyperlink(sURL As String)
On Error Resume Next
ThisWorkbook.FollowHyperlink Address:=sURL, NewWindow:=True
If err.Number <> 0 Then
MsgBox "Please goto to " & "www.ganttexcel.com", , "Gantt Excel"
End If
On Error GoTo 0
End Sub

#If Mac Then



#If VBA7 Then

Function execShell(command As String, Optional ByRef exitCode As Long) As String
Dim file As LongPtr
file = popen(command, "r")

If file = 0 Then
Exit Function
End If

While feof(file) = 0
Dim sGarb As String
Dim read As Long
sGarb = Space(50)
read = fread(sGarb, 1, Len(sGarb) - 1, file)
If read > 0 Then
sGarb = Left$(sGarb, read)
execShell = execShell & sGarb
End If
Wend

exitCode = pclose(file)
End Function

Function GetWebResponseMac(sURL As String, squery As String) As String

Dim sCmd As String
Dim sResult As String
Dim lExitCode As Long

sCmd = "curl --get -d """ & squery & """" & " " & sURL
sResult = execShell(sCmd, lExitCode)
GetWebResponseMac = sResult

End Function

#Else

Function execShell(command As String, Optional ByRef exitCode As Long) As String

Dim file As Long
file = popen(command, "r")

If file = 0 Then
Exit Function
End If

While feof(file) = 0
Dim sGarb As String
Dim read As Long
sGarb = Space(50)
read = fread(sGarb, 1, Len(sGarb) - 1, file)
If read > 0 Then
sGarb = Left$(sGarb, read)
execShell = execShell & sGarb
End If
Wend

exitCode = pclose(file)

End Function

Function GetWebResponseMac(sURL As String, squery As String) As String
''
Dim sCmd As String
Dim sResult As String
Dim lExitCode As Long


sCmd = "curl --get -d """ & squery & """" & " " & sURL
sResult = execShell(sCmd, lExitCode)
GetWebResponseMac = sResult

End Function

#End If

#End If

#If Mac Then
#Else
Function GetWebResponse(sURL As String) As String
If sURL = vbNullString Then Exit Function
Dim wR As Object'WinHttp.WinHttpRequest
Set wR = CreateObject("WinHttp.WinHttpRequest.5.1")
On Error Resume Next
wR.Open "GET", sURL
wR.Send
If err.Number <> 0 Then
GetWebResponse = "Error - " & err.Description
err.Clear
Else
GetWebResponse = wR.ResponseText
End If
On Error GoTo 0
Set wR = Nothing
End Function

#End If


Function GetDataFromURL(strURL, strMethod, strPostData, lngTimeout)
Dim strUserAgentString
Dim intSslErrorIgnoreFlags
Dim blnEnableRedirects
Dim blnEnableHttpsToHttpRedirects
Dim strHostOverride
Dim strLogin
Dim strPassword
Dim strResponseText
Dim objWinHttp
strUserAgentString = "http_requester/0.1"
intSslErrorIgnoreFlags = 13056
blnEnableRedirects = True
blnEnableHttpsToHttpRedirects = True
strHostOverride = vbNullString
strLogin = vbNullString
strPassword = vbNullString
Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
objWinHttp.Open strMethod, strURL
If strMethod = "POST" Then
objWinHttp.SetRequestHeader "Content-type", _
"application/x-www-form-urlencoded"
End If
If strHostOverride <> vbNullString Then
objWinHttp.SetRequestHeader "Host", strHostOverride
End If
objWinHttp.Option(0) = strUserAgentString
objWinHttp.Option(4) = intSslErrorIgnoreFlags
objWinHttp.Option(6) = blnEnableRedirects
objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
If (strLogin <> vbNullString) And (strPassword <> vbNullString) Then
objWinHttp.SetCredentials strLogin, strPassword, 0
End If
On Error Resume Next
objWinHttp.Send (strPostData)
If err.Number = 0 Then
If objWinHttp.Status = "200" Then
GetDataFromURL = objWinHttp.ResponseText
Else
GetDataFromURL = "HTTP " & objWinHttp.Status & " " & _
objWinHttp.StatusText
End If
Else
GetDataFromURL = "Error " & err.Number & " " & err.Source & " " & _
err.Description
End If
On Error GoTo 0
Set objWinHttp = Nothing
End Function

Public Function WeekNumVBA(ByVal dDate As Date, sWeekType) As Integer
If sWeekType = "ISO" Then
WeekNumVBA = DatePart("ww", dDate - Weekday(dDate, 2) + 4, 2, 2)
Else
WeekNumVBA = WorksheetFunction.weeknum(dDate, 2)
End If
End Function

Function GetLastRowOfFamily(cRow As Long) As Long
Dim vArr()
Dim lrow As Long, i As Long, curRowTIL As Long, nextRowTil As Long: lrow = GetLastRow
If cRow = lrow Then GetLastRowOfFamily = cRow: Exit Function
curRowTIL = Cells(cRow, cpg.TIL): nextRowTil = Cells(cRow + 1, cpg.TIL)
If curRowTIL = 0 And nextRowTil = 0 Then GetLastRowOfFamily = cRow: Exit Function

vArr = Range(Cells(1, cpg.TIL), Cells(lrow, cpg.TIL))
If curRowTIL = 0 And nextRowTil = 1 Then i = cRow + 1 Else i = cRow
For i = i To UBound(vArr)
If vArr(i, 1) = 0 Then GetLastRowOfFamily = i - 1: Exit Function
If i = lrow Then GetLastRowOfFamily = i: Exit Function
Next
Last:
End Function

Function GetFirstRowOfFamily(cRow As Long) As Long
Dim vArr()
Dim lrow As Long, i As Long, curRowTIL As Long
curRowTIL = Cells(cRow, cpg.TIL): lrow = GetLastRow:
If curRowTIL = 0 Then GetFirstRowOfFamily = cRow: Exit Function
vArr = Range(Cells(1, cpg.TIL), Cells(lrow, cpg.TIL))
i = cRow - 1
For i = cRow To LBound(vArr) Step -1
If vArr(i, 1) = 0 Then GetFirstRowOfFamily = i: Exit Function
Next
End Function

Sub DoSpellCheck(Optional t As Boolean)
Dim ws As Worksheet, r As Range, sRng As Range:Set ws = ActiveSheet
If Not GanttChart(ws) Then Exit Sub
Set sRng = Union(ws.Range(ws.Cells(1, cpg.TID), ws.Cells(ws.Cells.Rows.Count, cpg.WBS - 1)), _
ws.Range(ws.Cells(1, 1), ws.Cells(rownine, cpg.LC)))
Set r = Selection
If Intersect(r, sRng) Is Nothing Then
Else
MsgBox "Please select only task data cells", vbInformation, "Invalid Selection"
Exit Sub
End If
If r.Cells.Count = 1 Then
Set r = ws.Range(ws.Cells(firsttaskrow, cpg.WBS), ws.Cells(GetLastRow, cpg.LC))
End If
r.Cells.CheckSpelling
MsgBox "Spell Check Completed", vbInformation, "Gantt Excel"
Set ws = Nothing
End Sub
Sub AddFilterToTasksTrigger(Optional t As Boolean)
If Not GanttChart Then Exit Sub
Call DA:
If Not ActiveSheet.AutoFilterMode Then MsgBox msg(61):
Call AddFilteringToTasks(, True): Call EA
End Sub
Sub AddFilteringToTasks(Optional t As Boolean, Optional btimeline As Boolean)
If GanttChart = False Then Exit Sub
Set gs = setGSws
Dim ws As Worksheet: Dim cCol As Long: Dim tRng As Range: Dim bEnableEvents As Boolean
Set ws = ActiveSheet: Set tRng = Selection
If Not ws.AutoFilterMode Then
ws.Range(ws.Cells(rownine, cpg.TID), ws.Cells(rownine, cpg.LC)).Select
Selection.AutoFilter
On Error Resume Next
bEnableEvents = Application.EnableEvents
Application.EnableEvents = False
gs.Cells(rowtwo, cps.ShowPlanned) = 1
gs.Cells(rowtwo, cps.ShowInProgress) = 1
gs.Cells(rowtwo, cps.ShowCompleted) = 1
gs.Cells(rowtwo, cps.ShowGrouping) = 0
Application.EnableEvents = bEnableEvents
For cCol = cpt.TimelineEnd To cpg.LC Step -1
ws.Cells(rownine, 1).AutoFilter field:=cCol, VisibleDropDown:=False
Next cCol
Else
ws.AutoFilterMode = False
End If
tRng.Select
On Error GoTo 0
Call RefreshRibbon:
If btimeline Then Call CreateTimeline
End Sub
Sub ClearFilters(Optional t As Boolean)
If GanttChart = False Then Exit Sub
Dim ws As Worksheet:Set ws = ActiveSheet
If ws.FilterMode Then ws.ShowAllData
RefreshRibbon
End Sub

Sub ReApplyAutoFilter(v, lrow As Long, copyType As String)
Dim ws As Worksheet:Dim filterArray():Dim currentFiltRange As String:Dim col As Integer:Set ws = ActiveSheet
If ws.AutoFilterMode = True Then 'Capture AutoFilter settings
With ws.AutoFilter
currentFiltRange = .Range.Address
With .Filters
ReDim filterArray(1 To .Count, 1 To 3)
For col = 1 To .Count
With .Item(col)
If .On Then
filterArray(col, 1) = .Criteria1
If .Operator Then
filterArray(col, 2) = .Operator
If .Operator = xlAnd Or .Operator = xlOr Then
filterArray(col, 3) = .Criteria2
End If
End If
End If
End With
Next col
End With
End With
End If
ws.AutoFilterMode = False 'Remove AutoFilter
'Your code here - Range(Cells(1, 1), Cells(lrow, cpg.lc)).value = v
Call ArrayToRange(v, lrow, copyType)
If Not currentFiltRange = "" Then 'Restore Filter settings
Call AddFilteringToTasks(, False)
For col = 1 To UBound(filterArray(), 1)
If Not IsEmpty(filterArray(col, 1)) Then
If filterArray(col, 2) Then
check if Criteria2 exists and needs to be populated
If filterArray(col, 2) = xlAnd Or filterArray(col, 2) = xlOr Then
ws.Range(currentFiltRange).AutoFilter field:=col, _
Criteria1:=filterArray(col, 1), _
Operator:=filterArray(col, 2), _
Criteria2:=filterArray(col, 3)
Else
ws.Range(currentFiltRange).AutoFilter field:=col, _
Criteria1:=filterArray(col, 1), _
Operator:=filterArray(col, 2)
End If
Else
ws.Range(currentFiltRange).AutoFilter field:=col, _
Criteria1:=filterArray(col, 1)
End If
End If
Next col
End If
End Sub

Sub RSDS(Optional t As Boolean)
Call DA: GST.Cells(rowtwo, cps.tFirstSavedDate) = vbNullString:
Call PrepAllGanttCharts(False): ThisWorkbook.Save:
Application.DisplayAlerts = False: ThisWorkbook.Close: Application.DisplayAlerts = True
Call EA
End Sub
Function OpenBuyHyperlink(Optional t As Boolean)
OpenHyperlink siteURL & buyURL & strSource & "BuyBtn" & strMedium
End Function
Function OpenURLOnStartup()
Dim sWkNum As Long: Dim url_path As String: Dim sdate As Date:
If GST.Cells(rowtwo, cps.tFirstSavedDate) = vbNullString Then Exit Function Else sdate = GST.Cells(rowtwo, cps.tFirstSavedDate)
If sdate = Date Then Exit Function
sWkNum = Int((Date - sdate) / 7)
Select Case sWkNum
Case Is = 0
OpenHyperlink siteURL & freeURL1 & strSource & "W" & sWkNum & strMedium
Case Is = 1
OpenHyperlink siteURL & freeURL2 & strSource & "W" & sWkNum & strMedium
Case Is = 2
OpenHyperlink siteURL & freeURL3 & strSource & "W" & sWkNum & strMedium
Case Is >= 3
OpenHyperlink siteURL & freeURL4 & strSource & "W" & sWkNum & strMedium
End Select
End Function
Function EncodeEmail(s As String)
Const cStringAs String = "abcdefghijklmnopqrstuvwxyz1234567890_.()-@!#$%^&*[]{}|?/,<>"
Const nString As String = "1338745223488631774272104097418232296590923454576885693984899550751962469959834398945112339188278764806730201436531511"
Dim f As String
Dim i As Long, t As String
Dim p As Long
For i = 1 To Len(s)
t = LCase(Mid(s, i, 1))
p = InStr(1, cString, t, vbTextCompare)
If p = 0 Then
f = f & t
Else
f = f & Mid(nString, p * 2 - 1, 2)
End If
Next i
EncodeEmail = f
End Function
Function CheckRngFormula(ByVal Target As Range, Optional AllowFormula As Boolean) As Boolean
If Target.HasFormula Then
CheckRngFormula = True:
If AllowFormula Then
If Not bAllowFormulas Then MsgBox msg(20)
Else
MsgBox msg(21)
End If
Else
CheckRngFormula = False
End If
End Function
Function containsFormula(ByVal Target As Range) As Boolean
If Target.HasFormula Then containsFormula = True Else containsFormula = False
End Function
Function NoOfDaysInMonth(dt As Date)
NoOfDaysInMonth = Day(DateSerial(Year(dt), Month(dt) + 1, 1) - 1)
End Function

Function CalEDHrs(resname As String, ESD As Date, EED As Date)
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim dTSD As Double, dTED As Double: Dim tempenddate As Date
Dim i As Long, wkday As Long, noofdays As Long, durationhours As Long: dTSD = ESD: dTED = EED
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer
resStartHour = Format(newResourcesArray(resvalue, 10), "h"): resEndHour = Format(newResourcesArray(resvalue, 11), "h")
If resEndHour = 0 Then resEndHour = 24
daystartHour = Format(dTSD, "h"): dayendHour = Format(dTED, "h"): durationhours = 0: tempenddate = Int(CDbl(EED)): noofdays = Int(CDbl(EED)) - Int(CDbl(ESD)) + 1
For i = 1 To noofdays
If tempenddate < Int(CDbl(ESD)) Then Exit For
wkday = Weekday(Int(CDbl(tempenddate)), 2)
If IsDateAHoliday(resvalue, Int(CDbl(tempenddate))) = True Or IsDateAWorkday(resvalue, wkday) = False Then
GoTo Last
Else
If Int(CDbl(EED)) = Int(CDbl(ESD)) Then
If daystartHour <= resStartHour Then daystartHour = resStartHour
If daystartHour > resEndHour Then GoTo Last
If dayendHour > resEndHour Then dayendHour = resEndHour
durationhours = durationhours + (dayendHour - daystartHour)
GoTo Last
End If
If tempenddate = Int(CDbl(ESD)) Then
If daystartHour <= resStartHour Then daystartHour = resStartHour
If daystartHour > resEndHour Then GoTo Last
durationhours = durationhours + (resEndHour - daystartHour)
GoTo Last
End If
If tempenddate = Int(CDbl(EED)) Then
If dayendHour <= resStartHour Then GoTo Last
If dayendHour > resEndHour Then dayendHour = resEndHour
durationhours = durationhours + (dayendHour - resStartHour)
GoTo Last
End If
durationhours = durationhours + (resEndHour - resStartHour)
End If
Last:
tempenddate = tempenddate - 1
Next
CalEDHrs = durationhours
End Function

Function CalESDHrs(resname As String, EED As Date, ED As Long)
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim dTED As Double: Dim tempenddate As Date: Dim noofworkhours As Long, i As Long, wkday As Long, durationhours As Long, workhrs As Long: dTED = EED
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer
resStartHour = Format(newResourcesArray(resvalue, 10), "h"): resEndHour = Format(newResourcesArray(resvalue, 11), "h")
If resEndHour = 0 Then resEndHour = 24:
dayendHour = Format(dTED, "h"): durationhours = ED: workhrs = resEndHour - resStartHour: tempenddate = Int(CDbl(EED)): i = 0
Do While i < ED
wkday = Weekday(Int(CDbl(tempenddate)), 2)
If IsDateAHoliday(resvalue, Int(CDbl(tempenddate))) = True Or IsDateAWorkday(resvalue, wkday) = False Then
GoTo Last
Else
If tempenddate = Int(CDbl(EED)) Then
If dayendHour <= resStartHour Then GoTo Last
If dayendHour < resEndHour Then durationhours = dayendHour - resStartHour Else durationhours = resEndHour - resStartHour
Else
durationhours = resEndHour - resStartHour
End If
If durationhours >= ED - i Then durationhours = ED - i: Exit Do
i = i + durationhours
If i >= ED Then Exit Do
End If
Last:
tempenddate = tempenddate - 1
Loop
Dim n As Long:
If tempenddate = Int(CDbl(EED)) Then
If dayendHour > resEndHour Then n = resEndHour - durationhours Else n = dayendHour - durationhours
Else
n = resEndHour - durationhours
End If
CalESDHrs = DateAdd("h", n, tempenddate)
End Function

Function CalEEDHrs(resname As String, ESD As Date, ED As Long)
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim dTSD As Double: Dim tempenddate As Date: Dim noofworkhours As Long, i As Long, wkday As Long, durationhours As Long, workhrs As Long: dTSD = ESD
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer
resStartHour = Format(newResourcesArray(resvalue, 10), "h"): resEndHour = Format(newResourcesArray(resvalue, 11), "h")
If resEndHour = 0 Then resEndHour = 24:
daystartHour = Format(dTSD, "h"): durationhours = ED: workhrs = resEndHour - resStartHour: tempenddate = Int(CDbl(ESD)): i = 0
Do While i <= ED
wkday = Weekday(Int(CDbl(tempenddate)), 2)
If IsDateAHoliday(resvalue, Int(CDbl(tempenddate))) = True Or IsDateAWorkday(resvalue, wkday) = False Then
GoTo Last
Else
If tempenddate = Int(CDbl(ESD)) Then
If daystartHour >= resEndHour Then GoTo Last
If daystartHour <= resStartHour Then durationhours = resEndHour - resStartHour Else durationhours = resEndHour - daystartHour
Else
durationhours = resEndHour - resStartHour
End If

If ED - i <= durationhours Then Exit Do
i = i + durationhours
End If
Last:
tempenddate = tempenddate + 1
Loop
Dim n As Long:
If i = 0 And tempenddate = Int(CDbl(ESD)) Then
If daystartHour >= resStartHour Then n = daystartHour + (ED - i) Else n = resStartHour + (ED - i)
Else
n = resStartHour + (ED - i)
End If
CalEEDHrs = DateAdd("h", n, tempenddate)
End Function
Function CalEEDHrsDep(resname As String, ESD As Date, ED As Long)
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim dTSD As Double: Dim tempenddate As Date, calcdate As Date: Dim noofworkhours As Long, i As Long, wkday As Long, durationhours As Long, workhrs As Long, restofday As Long: dTSD = ESD
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer
resStartHour = Format(newResourcesArray(resvalue, 10), "h"): resEndHour = Format(newResourcesArray(resvalue, 11), "h")
If resEndHour = 0 Then resEndHour = 24:
daystartHour = Format(dTSD, "h"): durationhours = ED: workhrs = resEndHour - resStartHour
tempenddate = Int(CDbl(ESD)): i = 0
Do While i <= ED
wkday = Weekday(Int(CDbl(tempenddate)), 2)
If IsDateAHoliday(resvalue, Int(CDbl(tempenddate))) = True Or IsDateAWorkday(resvalue, wkday) = False Then
GoTo Last
Else
If tempenddate = Int(CDbl(ESD)) Then
If daystartHour >= resEndHour Then GoTo Last
If daystartHour <= resStartHour Then durationhours = resEndHour - resStartHour Else durationhours = resEndHour - daystartHour
Else
durationhours = resEndHour - resStartHour
End If
If ED - i <= durationhours Then Exit Do
i = i + durationhours
End If
Last:
tempenddate = tempenddate + 1
Loop
Dim n As Long:
If i = 0 And tempenddate = Int(CDbl(ESD)) Then
If daystartHour >= resStartHour Then n = daystartHour + (ED - i) Else n = resStartHour + (ED - i)
Else
n = resStartHour + (ED - i)
End If
CalEEDHrsDep = DateAdd("h", n, tempenddate):
calcdate = CalEEDHrsDep: wkday = Weekday(Int(CDbl(calcdate)), 2)
If IsDateAHoliday(resvalue, Int(CDbl(calcdate))) = True Or IsDateAWorkday(resvalue, wkday) = False Then
CalEEDHrsDep = GetNewESD(resname, calcdate + 1)
End If
If Format(CalEEDHrsDep, "h") >= resEndHour Then
CalEEDHrsDep = GetNewESD(resname, calcdate + 1)
End If
If Format(CalEEDHrsDep, "h") < resStartHour Then
CalEEDHrsDep = GetNewESD(resname, calcdate)
End If
End Function

Function sBrowseForFileMac() As String
#If Mac Then
Dim sPath As String
Dim sScript As String
Dim sFileFormat As String
sFileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"

On Error Resume Next
sPath = MacScript("return (path to desktop folder) as String")
If val(Application.Version) < 15 Then
sScript = _
"set theFile to (choose file of type" & _
" " & sFileFormat & " " & _
"with prompt ""Please select a file"" default location alias """ & _
sPath & """ without multiple selections allowed) as string" & vbNewLine & _
"return theFile"
Else
sScript = _
"set theFile to (choose file of type" & _
" " & sFileFormat & " " & _
"with prompt ""Please select a file"" default location alias """ & _
sPath & """ without multiple selections allowed) as string" & vbNewLine & _
"return posix path of theFile"
End If
sBrowseForFileMac = MacScript(sScript)
On Error GoTo 0
#End If
End Function

Function GetResourcesFilePath() As String
Dim fPath
#If Mac Then
#Else
fPath = Application.GetOpenFilename( _
FileFilter:="XLSX Files (*.xlsx),*.xlsx", _
Title:="Select a file", _
MultiSelect:=False)
If fPath = False Then
GetResourcesFilePath = vbNullString
Else
GetResourcesFilePath = fPath
End If
#End If
End Function

Function GetNewESD(resname As String, ESD As Date) As Date
Dim newESD As Date: newESD = Int(CDbl(ESD))
Dim wkday As Long, holiday As Long, weekend As Long: Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
If st.HGC Then resStartHour = Format(newResourcesArray(resvalue, 10), "h"): resEndHour = Format(newResourcesArray(resvalue, 11), "h"):daystartHour = Format(ESD, "h"):
If resEndHour = 0 Then resEndHour = 24
If daystartHour > resEndHour Then newESD = newESD + 1
wkday = Weekday(newESD, 2)
Do While IsDateAWorkday(resvalue, wkday) = False Or IsDateAHoliday(resvalue, newESD) = True
newESD = newESD + 1: wkday = Weekday(newESD, 2)
Loop
If st.HGC Then
If Int(CDbl(ESD)) = Int(CDbl(newESD)) Then
If daystartHour < resStartHour Then
GetNewESD = DateAdd("h", resStartHour, newESD)
Else
GetNewESD = DateAdd("h", daystartHour, newESD)
End If
Else
GetNewESD = DateAdd("h", resStartHour, newESD)
End If
Else
GetNewESD = newESD
End If
End Function

Function GetEndDateFromWorkDays(resname As String, ByVal sdate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim wkday As Long
Do While 1 = 1
wkday = Weekday(sdate, 2)
If IsDateAHoliday(resvalue, sdate) = False Then
If IsDateAWorkday(resvalue, wkday) = True Then
wd = wd - 1
If wd <= 0 Then GetEndDateFromWorkDays = sdate: Exit Do
End If
End If
sdate = sdate + 1
Loop
End Function

Function GetStartFromWorkDays(resname As String, ByVal eDate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim wkday As Long
Do While 1 = 1
wkday = Weekday(eDate, 2)
If IsDateAHoliday(resvalue, eDate) = False Then
If IsDateAWorkday(resvalue, wkday) = False Then
Else
wd = wd - 1
If wd <= 0 Then
GetStartFromWorkDays = eDate
Exit Do
End If
End If
End If
eDate = eDate - 1
Loop
End Function
Function IsDateAHoliday(resvalue As Long, d As Date) As Boolean
Dim i As Long: Dim includeOrgHol As Boolean: includeOrgHol = newResourcesArray(resvalue, 2)
For i = 1 To UBound(newHolidaysArray, 2)
If newHolidaysArray(resvalue, i) = d Then IsDateAHoliday = True: Exit For
If includeOrgHol Then If newHolidaysArray(0, i) = d Then IsDateAHoliday = True: Exit For
Next
End Function

Function IsDateAWorkday(resvalue As Long, index As Long)
IsDateAWorkday = CBool(newWorkdaysArray(resvalue, index))
End Function

Function getResValue(ByVal resname As String, ByVal arr As Variant) As Long
Dim i As Long
If resname = "" Or InStr(1, resname, ",") > 0 Or resname = "organization" Then resname = "organization": resvalue = 0: Exit Function
For i = LBound(arr) To UBound(arr)
If LCase(resname) = arr(i, 0) Then resvalue = i: Exit For
Next i
End Function

Function GetCP(resname As String, startDate As Date, endDate As Date) As Long
If ResArraysReady = False Then Call RememberResArrays
Dim wkday As Long, holiday As Long, weekend As Long: Dim cpc As Double
Call getResValue(resname, newResourcesArray): wkday = Weekday(startDate, 2)
Do While IsDateAWorkday(resvalue, wkday) = False
startDate = startDate + 1:wkday = Weekday(startDate, 2)
Loop
Do While IsDateAHoliday(resvalue, startDate) = True
startDate = startDate + 1
Loop
If st.HGC Then
If startDate > Now Then cpc = 0: GoTo Last
If endDate < Now Then cpc = 100: GoTo Last
If startDate < Now And endDate >= Now Then cpc = CInt(Round(CalEDHrs(resname, startDate, Now) / CalEDHrs(resname, startDate, endDate) * 100)): GoTo Last
Else
If startDate > Date Then cpc = 0: GoTo Last
If endDate < Date Then cpc = 100: GoTo Last
If startDate = Date Then cpc = 0: GoTo Last
If startDate < Date And endDate >= Date Then cpc = CInt(Round((GetWorkDaysFromDate(resname, startDate, Date - 1) / GetWorkDaysFromDate(resname, startDate, endDate)) * 100)): GoTo Last
End If
Last:
GetCP = cpc
End Function

Function getSetTaskBarColumnNo() As Long
Dim r As Range: 'Set gs = setGSws
Set r = ActiveSheet.Range("1:1"):
'If barTextColNo = 0 Then gs.Cells(rowtwo, cps.BarTextDataColumnName) = "Task": Call ReadSettings
On Error Resume Next
getSetTaskBarColumnNo = Application.WorksheetFunction.Match(st.TextBarColumnName, r.value, 0)
On Error GoTo 0
Set r = Nothing
End Function

Function getStartRow(Optional cRowOnly As Long, Optional familytype As String) As Long
Dim lrow As Long: lrow = GetLastRow
If cRowOnly = 0 Or familytype = "" Then familytype = allRows
If familytype = rowOnly Then
getStartRow = cRowOnly
ElseIf familytype = family Then
getStartRow = GetFirstRowOfFamily(cRowOnly)
ElseIf familytype = aboveFamily Then
getStartRow = GetFirstRowOfFamily(cRowOnly)
ElseIf familytype = allRows Then
getStartRow = firsttaskrow
End If
End Function

Function getEndRow(Optional cRowOnly As Long, Optional familytype As String) As Long
Dim lrow As Long: lrow = GetLastRow
If cRowOnly = 0 Or familytype = "" Then familytype = allRows
If familytype = rowOnly Then
getEndRow = cRowOnly:
ElseIf familytype = family Then
getEndRow = GetLastRowOfFamily(cRowOnly)
ElseIf familytype = aboveFamily Then
getEndRow = cRowOnly
ElseIf familytype = allRows Then
getEndRow = lrow
End If
End Function

Public Function GetDelayedDateNew(ByVal sdate As Date, ByVal dDelay As Long, _
Optional Su As Boolean, Optional Mo As Boolean, Optional Tu As Boolean, _
Optional We As Boolean, Optional Th As Boolean, Optional Fr As Boolean, Optional Sa As Boolean, _
Optional sDependencyType As String, Optional resvalue As Long) As Date

Dim bReverse As Boolean, i As Long
GetDelayedDateNew = sdate
Set the given date a to working day
Do Until IsWorkingDayForDependencies(Weekday(GetDelayedDateNew, vbSunday), Su, Mo, Tu, We, Th, Fr, Sa) = True _
And IsDateAHoliday(resvalue, GetDelayedDateNew) = False
If sDependencyType = "FS" Then
GetDelayedDateNew = GetDelayedDateNew + 1
ElseIf sDependencyType = "SF" Then
GetDelayedDateNew = GetDelayedDateNew - 1
ElseIf sDependencyType = "SS" Then
GetDelayedDateNew = GetDelayedDateNew + 1
ElseIf sDependencyType = "FF" Then
GetDelayedDateNew = GetDelayedDateNew + 1
End If
Loop
After the above loop executed the date will a working date. We now add the delay value to get the final date
and ensure that the final date is also a working day
i = dDelay
If i > 0 Then
Do Until i = 0
GetDelayedDateNew = GetDelayedDateNew + 1
Do Until IsWorkingDayForDependencies(Weekday(GetDelayedDateNew, vbSunday), Su, Mo, Tu, We, Th, Fr, Sa) = True _
And IsDateAHoliday(resvalue, GetDelayedDateNew) = False
GetDelayedDateNew = GetDelayedDateNew + 1
Loop
i = i - 1
Loop
ElseIf i < 0 Then
Do Until i = 0
GetDelayedDateNew = GetDelayedDateNew - 1
Do Until IsWorkingDayForDependencies(Weekday(GetDelayedDateNew, vbSunday), Su, Mo, Tu, We, Th, Fr, Sa) = True _
And IsDateAHoliday(resvalue, GetDelayedDateNew) = False
GetDelayedDateNew = GetDelayedDateNew - 1
Loop
i = i + 1
Loop
ElseIf i = 0 Then
End If
End Function
Public Function GetDelayedDate(ByVal sdate As Date, ByVal dDelay As Long, _
Optional Su As Boolean, Optional Mo As Boolean, Optional Tu As Boolean, _
Optional We As Boolean, Optional Th As Boolean, Optional Fr As Boolean, Optional Sa As Boolean, _
Optional sDependencyType As String, Optional resvalue As Long) As Date

If bStopCalculationOfConstraints Then
GetDelayedDate = Date
Exit Function
End If
GetDelayedDate = GetDelayedDateNew(sdate, dDelay, Su, Mo, Tu, We, Th, Fr, Sa, sDependencyType, resvalue)
End Function

Function IsWorkingDayForDependencies(wDayNum As Long, Su As Boolean, Mo As Boolean, _
Tu As Boolean, We As Boolean, Th As Boolean, Fr As Boolean, Sa As Boolean) As Boolean
Select Case wDayNum
Case 1
If Su Then IsWorkingDayForDependencies = True
Case 2
If Mo Then IsWorkingDayForDependencies = True
Case 3
If Tu Then IsWorkingDayForDependencies = True
Case 4
If We Then IsWorkingDayForDependencies = True
Case 5
If Th Then IsWorkingDayForDependencies = True
Case 6
If Fr Then IsWorkingDayForDependencies = True
Case 7
If Sa Then IsWorkingDayForDependencies = True
End Select
End Function

Function GetCountOfChar(ByRef ar_sText As String, ByVal a_sChar As String) As Integer
Dim l_iIndex As Integer, l_iMax As Integer, l_iLen As Integer
GetCountOfChar = 0: l_iMax = Len(ar_sText): l_iLen = Len(a_sChar)
For l_iIndex = 1 To l_iMax
If (Mid(ar_sText, l_iIndex, l_iLen) = a_sChar) Then
GetCountOfChar = GetCountOfChar + 1
If (l_iLen > 1) Then l_iIndex = l_iIndex + (l_iLen - 1)
End If
Next l_iIndex
End Function

Function resetColHeaderforCol(colNo As Long) As String
Dim colName As String: colName = Cells(rowone, colNo)
Dim i As Long
For i = LBound(GCcolumns()) To UBound(GCcolumns())
If colName = GCcolumns(i) Then
resetColHeaderforCol = GCcolumnsEngName(i)
Exit Function
End If
Next i
End Function

Function getPID(Optional ws As Worksheet) As Long
Dim splitSS
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
splitSS = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
getPID = splitSS(0)
Else
MsgBox "getPID function " & msg(65)
End If
End Function

Function setGSws(Optional ws As Worksheet) As Worksheet
Dim splitSS
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
splitSS = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
Set setGSws = Worksheets(splitSS(1))
Else
MsgBox "setGSws function " & msg(65)
End If
End Function

Function setRSws(Optional ws As Worksheet) As Worksheet
Dim splitSS
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
splitSS = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
Set setRSws = Worksheets(splitSS(2))
Else
MsgBox "setRSws function " & msg(65)
End If
End Function

Function getGSname(Optional ws As Worksheet) As String
Dim splitSS
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
splitSS = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
getGSname = splitSS(1)
Else
MsgBox "getGSname function " & msg(65)
End If
End Function

Function getRSname(Optional ws As Worksheet) As String
Dim splitSS
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
splitSS = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
getRSname = splitSS(2)
Else
MsgBox "getRSname function " & msg(65)
End If
End Function

Function getProjectCounter() As Long
getProjectCounter = GST.Cells(rowtwo, cps.SSN).value
End Function

Sub ProjectCountPlusOne(Optional t As String)
GST.Cells(rowtwo, cps.SSN).value = GST.Cells(rowtwo, cps.SSN).value + 1
End Sub

Sub setGSRSname(gsName As String, rsname As String, ws As Worksheet, Optional ProjectNo As Long)
Dim SSvalue As String
Dim oldSplit
Dim SSCol As Long
If ws Is Nothing Then Set ws = ActiveSheet
If GanttChart(ws) Then
SSCol = Application.WorksheetFunction.Match("SS", ws.Range("1:1"), 0)
oldSplit = Split(ws.Cells(rowtwo, SSCol).value, DepSeperator)
If ProjectNo = 0 Then
SSvalue = oldSplit(0) & DepSeperator & gsName & DepSeperator & rsname
Else
SSvalue = ProjectNo & DepSeperator & gsName & DepSeperator & rsname
End If
ws.Cells(rowtwo, SSCol).value = SSvalue
End If
End Sub

Function getTIDfromShape(s As Shape) As Long
Dim vStr As Variant: vStr = Split(s.Name, "_"): getTIDfromShape = CLng(vStr(2))
End Function
Function getTIDRowfromShape(s As Shape) As Long
Dim vStr As Variant: vStr = Split(s.Name, "_"): getTIDRowfromShape = getTIDRow(CLng(vStr(2))):
End Function
Function getGETypefromShape(s As Shape) As String
Dim vStr As Variant: vStr = Split(s.Name, "_"): getGETypefromShape = Cells(getTIDRow(CLng(vStr(2))), 1)
End Function
Function getShapeType(s As Shape) As String
If InStr(1, s.Name, "TE") Then getShapeType = "TE"
If InStr(1, s.Name, "TP") Then getShapeType = "TP"
If InStr(1, s.Name, "ME") Then getShapeType = "ME"
If InStr(1, s.Name, "TB") Then getShapeType = "TB"
If InStr(1, s.Name, "MB") Then getShapeType = "MB"
If InStr(1, s.Name, "TA") Then getShapeType = "TA"
If InStr(1, s.Name, "MA") Then getShapeType = "MA"
If InStr(1, s.Name, "TGB") Then getShapeType = "TGB"
End Function
Function getSavedShapeL(shInfo As String) As Double
Dim vstr1 As Variant: vstr1 = Split(shInfo, DepSeperator)
getSavedShapeL = vstr1(2):
End Function
Function getSavedShapeT(shInfo As String) As Double
Dim vstr1 As Variant: vstr1 = Split(shInfo, DepSeperator)
getSavedShapeT = vstr1(3)
End Function
Function getSavedShapeR(shInfo As String) As Double
Dim vstr1 As Variant: vstr1 = Split(shInfo, DepSeperator)
getSavedShapeR = vstr1(4)
End Function
Function getSavedShapeB(shInfo As String) As Double
Dim vstr1 As Variant: vstr1 = Split(shInfo, DepSeperator)
getSavedShapeB = vstr1(5)
End Function
Function shapeExists(shapeName As String) As Boolean
Dim Sh As Shape
For Each Sh In ActiveSheet.Shapes
If Sh.Name = shapeName Then shapeExists = True
Next Sh
End Function

Function getTIDRow(TID As Long) As Long
getTIDRow = Range(Cells(rowone, cpg.TID), Cells(10000, cpg.TID)).Find(TID, , xlFormulas, xlWhole).Row
End Function

Function GetCalculatedPercentage(ByRef arrData(), ByRef cRow As Long)
If ResArraysReady = False Then Call RememberResArrays
Dim newESD As Date: newESD = arrData(cRow, cpg.ESD)
Dim resname As String: Dim wkday As Long, holiday As Long, weekend As Long
resname = arrData(cRow, cpg.Resource): Call getResValue(resname, newResourcesArray): wkday = Weekday(newESD, 2)
Do While IsDateAWorkday(resvalue, wkday) = False
newESD = newESD + 1:wkday = Weekday(newESD, 2)
Loop
Do While IsDateAHoliday(resvalue, newESD) = True
newESD = newESD + 1
Loop
If st.HGC Then
If newESD > Now Then
arrData(cRow, cpg.PercentageCompleted) = 0
ElseIf CDate(arrData(cRow, cpg.EED)) < Now Then
arrData(cRow, cpg.PercentageCompleted) = 1
Else
arrData(cRow, cpg.PercentageCompleted) = CalEDHrs(Cells(cRow, cpg.Resource), newESD, Now) / CalEDHrs(Cells(cRow, cpg.Resource), newESD, CDate(arrData(cRow, cpg.EED)))
End If
Else
If newESD > Date Then arrData(cRow, cpg.PercentageCompleted) = 0: GoTo Last
If CDate(arrData(cRow, cpg.EED)) < Date Then arrData(cRow, cpg.PercentageCompleted) = 1: GoTo Last
If newESD = Date Then arrData(cRow, cpg.PercentageCompleted) = 0: GoTo Last
If newESD < Date And CDate(arrData(cRow, cpg.EED)) >= Date Then
arrData(cRow, cpg.PercentageCompleted) = GetWorkDaysFromDate(Cells(cRow, cpg.Resource), newESD, Date - 1) / GetWorkDaysFromDate(Cells(cRow, cpg.Resource), newESD, CDate(arrData(cRow, cpg.EED)))
GoTo Last
End If
End If
Last:
End Function

Function GetDependencyTypeCodeFromName(d As String) As String
Select Case d
Case Is = "Finish to Start"
GetDependencyTypeCodeFromName = "FS"
Case Is = "Start to Start"
GetDependencyTypeCodeFromName = "SS"
Case Is = "Start to Finish"
GetDependencyTypeCodeFromName = "SF"
Case Is = "Finish to Finish"
GetDependencyTypeCodeFromName = "FF"
End Select
End Function
Function GetDependencyTypeNameFromCode(d As String) As String
Select Case d
Case Is = "FS"
GetDependencyTypeNameFromCode = "Finish to Start"
Case Is = "SS"
GetDependencyTypeNameFromCode = "Start to Start"
Case Is = "SF"
GetDependencyTypeNameFromCode = "Start to Finish"
Case Is = "FF"
GetDependencyTypeNameFromCode = "Finish to Finish"
End Select
End Function

Function Is2007() As Boolean
#If VBA7 Then
Is2007 = False
#Else
Is2007 = True
#End If
End Function

Function checkSheetError(Optional ws As Worksheet) As Boolean
Dim errorCount As Long
If ws Is Nothing Then Set ws = ActiveSheet
On Error Resume Next
errorCount = ws.UsedRange.Cells.SpecialCells(xlCellTypeFormulas, xlErrors).Count
On Error GoTo 0
If errorCount > 0 Then MsgBox msg(92): checkSheetError = True Else checkSheetError = False
End Function
Function getGEtype(cRow As Long) As String
If Cells(cRow, cpg.GEtype) <> "" Then getGEtype = Cells(cRow, cpg.GEtype)
End Function
Function GetNxtWDAfHol(resname As String, ByVal sdate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Do Until IsDateAHoliday(resvalue, sdate) = False
sdate = sdate + 1
Loop
GetNxtWDAfHol = sdate
End Function
Function GetNxtWDAfWO(resname As String, ByVal sdate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim wkday As Long: wkday = Weekday(sdate, 2)
Do Until IsDateAWorkday(resvalue, wkday)
sdate = sdate + 1: wkday = Weekday(sdate, 2)
Loop
GetNxtWDAfWO = sdate:
End Function
Function GetPreWDAfHol(resname As String, ByVal sdate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Do Until IsDateAHoliday(resvalue, sdate) = False
sdate = sdate - 1
Loop
GetPreWDAfHol = sdate
End Function
Function GetPreWDAfWO(resname As String, ByVal sdate As Date, ByVal wd As Long) As Date
If ResArraysReady = False Then Call RememberResArrays
Call getResValue(resname, newResourcesArray)
Dim wkday As Long: wkday = Weekday(sdate, 2)
Do Until IsDateAWorkday(resvalue, wkday)
sdate = sdate - 1: wkday = Weekday(sdate, 2)
Loop
GetPreWDAfWO = sdate:
End Function
Function Filled(MyCell As Range) As Boolean
If MyCell.Interior.ColorIndex > 0 Then
Filled = True
Else
Filled = False
End If
End Function

Function strSource() As String
strSource = sourceURL & GST.Cells(rowtwo, cps.Campaign)
End Function

Function strMedium() As String
If FreeVersion Then
strMedium = mediumURL & GST.Cells(rowtwo, cps.tUsrEmailID)
Else
strMedium = mediumURL & EncodeEmail(UnDecipher(GST.Cells(rowtwo, cps.tUsrEmailID)))
End If
End Function
Option Explicit
Option Private Module
Private Const BHeightRatio As Double = 0.17
Private Const MHeightRatio As Double = 0.15
Private Const lheightFactorOthers As Double = 0.35
Private vArrAllData()
Private vArrShapes()
Private vArrDatesValues()
Dim CountNum As Long
Dim sColor As Long
Dim bFreeVersion As Boolean, bShowPBar As Boolean, bShowOverdueBars As Boolean

Sub ShapeCount()
Call ShapeArray: MsgBox "ShapeCount = " & UBound(vArrShapes)
End Sub

Sub ShapeArray()
Dim Sh As Shape: Dim x As Long: x = 1
If ActiveSheet.Shapes.Count = 0 Then Exit Sub
ReDim vArrShapes(1 To ActiveSheet.Shapes.Count)
For Each Sh In ActiveSheet.Shapes
vArrShapes(x) = Sh.Name: x = x + 1
Next Sh
End Sub

Sub ChangeTextInGanttBars(Optional cRowOnly As Long, Optional familytype As String) 'used from settings to find tgb and change them
If st.ShowRefreshTimeline = False Then Exit Sub
bFreeVersion = FreeVersion
If Not bFreeVersion Then If st.ShowTextinBars = False Then Exit Sub 'force draw in free
Dim lrow As Long, cRow As Long, lzoomLevel As Long, StartRow As Long, EndRow As Long: Dim s As Shape
barTextColNo = getSetTaskBarColumnNo: lrow = GetLastRow: lzoomLevel = ActiveWindow.Zoom: ActiveWindow.Zoom = 100:
StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
For cRow = StartRow To EndRow
On Error Resume Next
Set s = ActiveSheet.Shapes("S_TGB_" & Cells(cRow, cpg.TID))
On Error GoTo 0
If s Is Nothing Then GoTo nexa:
Call setTGB(s, cRow): Set s = Nothing
nexa:
Next cRow
ActiveWindow.Zoom = lzoomLevel
End Sub

Sub setTGB(s As Shape, shapeRow As Long) ' sets color and text
With s
.Line.visible = msoFalse: .Fill.visible = msoFalse: .ZOrder msoBringToFront
If st.CurrentView = "D" Then .OnAction = "ShapeClicked"
With .TextFrame.Characters.Font
.Color = st.cTGB: .size = st.TGBFS: .Bold = st.TextBarIsBold:
End With
With .TextFrame
.VerticalAlignment = xlVAlignCenter: .HorizontalAlignment = xlHAlignLeft
End With
If bFreeVersion And shapeRow >= 20 Then
.TextFrame.Characters.Text = msg(60)
Else
If Cells(shapeRow, barTextColNo).NumberFormat <> "General" Then
.TextFrame.Characters.Text = Format(Cells(shapeRow, barTextColNo).Value2, Cells(shapeRow, barTextColNo).NumberFormat)
Else
.TextFrame.Characters.Text = Left(Cells(shapeRow, barTextColNo).Value2, st.TextBarChars)
End If
End If
End With
End Sub

Sub DrawTGBAll(Optional es As Shape) ' draws all tgb based onposition of est bars
If st.ShowRefreshTimeline = False Then Exit Sub
Dim s As Shape, Sh As Shape: Dim pleft As Double: Dim shapeRow As Long, lrow As Long, x As Long, y As Long, tidr As Long: Dim cView As String: Dim vArrShapesTGB
bFreeVersion = FreeVersion: cView = st.CurrentView
If Not bFreeVersion Then If st.ShowTextinBars = False Then Exit Sub 'force draw in free
lrow = GetLastRow: barTextColNo = getSetTaskBarColumnNo: x = 1
If ActiveSheet.Shapes.Count = 0 Then Exit Sub
ReDim vArrShapesTGB(1 To ActiveSheet.Shapes.Count, 1 To 5):
If es Is Nothing Then
For Each Sh In ActiveSheet.Shapes
If cView = "HH" Then
If Left(Sh.Name, 5) = "S_TH_" Or Left(Sh.Name, 4) = "S_ME" Then
pleft = Sh.Left
If Left(Sh.Name, 4) = "S_ME" Then pleft = pleft + 5 Else pleft = pleft - 10
vArrShapesTGB(x, 1) = Sh.Name: vArrShapesTGB(x, 2) = pleft: vArrShapesTGB(x, 3) = Sh.Top
vArrShapesTGB(x, 4) = Sh.Height: vArrShapesTGB(x, 5) = Sh.TopLeftCell.Row: x = x + 1
End If
Else
If Left(Sh.Name, 5) = "S_TE_" Or Left(Sh.Name, 5) = "S_ME_" Then
vArrShapesTGB(x, 1) = Sh.Name: vArrShapesTGB(x, 2) = Sh.Left: vArrShapesTGB(x, 3) = Sh.Top
vArrShapesTGB(x, 4) = Sh.Height: vArrShapesTGB(x, 5) = Sh.TopLeftCell.Row: x = x + 1
End If
End If
Next Sh
Else
ReDim vArrShapesTGB(1 To 1, 1 To 5):
tidr = getTIDfromShape(es):
On Error Resume Next 'draws tgb on demand for single row
Set Sh = ActiveSheet.Shapes("S_TGB_" & tidr): Sh.Delete
On Error GoTo 0
If cView = "HH" Then
If Left(es.Name, 5) = "S_TH_" Or Left(es.Name, 4) = "S_ME" Then
pleft = es.Left
If Left(es.Name, 4) = "S_ME" Then pleft = pleft + 5 Else pleft = pleft - 10
vArrShapesTGB(x, 1) = es.Name: vArrShapesTGB(x, 2) = pleft: vArrShapesTGB(x, 3) = es.Top
vArrShapesTGB(x, 4) = es.Height: vArrShapesTGB(x, 5) = es.TopLeftCell.Row
End If
Else
If Left(es.Name, 5) = "S_TE_" Or Left(es.Name, 5) = "S_ME_" Then
vArrShapesTGB(x, 1) = es.Name: vArrShapesTGB(x, 2) = es.Left: vArrShapesTGB(x, 3) = es.Top
vArrShapesTGB(x, 4) = es.Height: vArrShapesTGB(x, 5) = es.TopLeftCell.Row
End If
End If
End If
For y = 1 To UBound(vArrShapesTGB, 1)
If vArrShapesTGB(y, 1) <> "" Then
If Left(vArrShapesTGB(y, 1), 4) = "S_TE" Then pleft = vArrShapesTGB(y, 2) + 1 Else pleft = vArrShapesTGB(y, 2) + 10
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=pleft, Top:=vArrShapesTGB(y, 3), Width:=300, Height:=vArrShapesTGB(y, 4))
shapeRow = vArrShapesTGB(y, 5): s.Name = "S_TGB_" & Mid(vArrShapesTGB(y, 1), 6): Call setTGB(s, shapeRow)
Else
GoTo Last
End If
Next y
Last:
Set s = Nothing: Set Sh = Nothing: ReDim vArrShapesTGB(1, 1)
End Sub

Sub DelnDrawAllGanttBars()
Call checkSheetError
bDeleteAllAndDrawGB = True: Call DrawGanttBars(, allBars, allRows)
End Sub
Sub DrawAllGanttBars()
Call checkSheetError
Call DrawGanttBars(, allBars, allRows)
End Sub

Sub DrawGanttBars(Optional cRowOnly As Long, Optional typeofbar As String, Optional familytype As String)
If st.ShowRefreshTimeline = False Then Exit Sub
Dim pcompleted As Double, taskpercentmark As Double: Dim resname As String, cView As String, strGEtype As String
Dim sRng As Range: Dim sno As Long, StartRow As Long, EndRow As Long, i As Long: Dim bNoBar As Boolean
Dim ESCol As Long, EECol As Long, BSCol As Long, BECol As Long, ASCol As Long, AECol As Long, PSColAs Long, lrow As Long, cRow As Long, lzoomLevel As Long, taskduration As Long
Dim ESD As Date, EED As Date, BSD As Date, BED As Date, ASD As Date, AED As Date, TSD As Date, TED As Date, TPD As Date
Call LoadCalDateValuesArray: barTextColNo = getSetTaskBarColumnNo: bShowPBar = st.ShowPercBar: bFreeVersion = FreeVersion
If typeofbar = "" Then typeofbar = allBars
lrow = GetLastRow: lzoomLevel = ActiveWindow.Zoom: ActiveWindow.Zoom = 100: StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
tlog "DrawGanttBars: " & typeofbar & "-" & familytype
ReDim vArrAllData(1 To lrow, 1 To cpg.LC): vArrAllData = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value: cView = st.CurrentView
If ResArraysReady = False Then Call RememberResArrays
If st.ShowOverdueBar Then bShowOverdueBars = True Else bShowOverdueBars = False
If cView = "HH" Then 'HOURLY View
If st.HideNonWorkingHours Or st.HideHolidays Or st.HideWorkOffDays Then Call BuildHourlyView
tlog "Add gantt bars for hours"
Call DeleteAllGanttShapes
For i = firsttaskrow To lrow
strGEtype = getGEtype(i)
If Not IsDate(CDate(vArrAllData(i, cpg.ESD))) And Not IsDate(CDate(vArrAllData(i, cpg.EED))) Then GoTo Last
ESD = vArrAllData(i, cpg.ESD): EED = vArrAllData(i, cpg.EED): resname = vArrAllData(i, cpg.Resource)
If EED = ESD And (EED = Cells(rowsix, cpt.TimelineStart).value Or ESD = Cells(rowsix, cpt.TimelineEnd).value) Then GoTo okToDraw
If EED <= Cells(rowsix, cpt.TimelineStart).value Or ESD >= Cells(rowsix, cpt.TimelineEnd).value Then bNoBar = True: GoTo Last
okToDraw:
Call DrawGanttBarsHours("E", resname, i, ESD, EED)
Last:
If bNoBar Then bNoBar = False: Call setColors(i, strGEtype, "E")
Next i
tlog "Add gantt bars for hours":
GoTo final
End If
If cView = "D" Then If st.HideHolidays Or st.HideWorkOffDays Then Call BuildDailyView 'DAILY and Other Views
Set sRng = Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)): TSD = Cells(rowsix, cpt.TimelineStart): TED = Cells(rowsix, cpt.TimelineEnd)
Select Case cView
Case Is = "D"
TED = Cells(rowsix, cpt.TimelineEnd)
Case Is = "W"
TED = Cells(rowsix, cpt.TimelineEnd) + 6
Case Is = "M"
TED = DateAdd("M", 1, Cells(rowsix, cpt.TimelineEnd)) - 1
Case Is = "Q"
TED = DateAdd("M", 3, Cells(rowsix, cpt.TimelineEnd)) - 1
Case Is = "HY"
TED = DateAdd("M", 6, Cells(rowsix, cpt.TimelineEnd)) - 1
Case Is = "Y"
TED = DateAdd("Y", 1, Cells(rowsix, cpt.TimelineEnd)) - 1
End Select
If familytype <> allRows Then If (EndRow - StartRow) > 100 Then StartRow = firsttaskrow: EndRow = lrow:
If bDeleteAllAndDrawGB = True Then
Call DeleteAllGanttShapes
Else
If ActiveSheet.Shapes.Count > 0 Then Call ShapeArray
End If
For cRow = StartRow To EndRow
If vArrAllData(cRow, cpg.GEtype) = vbNullString Then GoTo nexa
ESCol = 0: EECol = 0: BSCol = 0: BECol = 0: ASCol = 0: AECol = 0: strGEtype = getGEtype(cRow):
If typeofbar = estBars Or typeofbar = allBars Then
If IsDate(vArrAllData(cRow, cpg.ESD)) = False Or IsDate(vArrAllData(cRow, cpg.EED)) = False Then GoTo EstPercBarsDrawn
If vArrAllData(cRow, cpg.ESD) = "" Or vArrAllData(cRow, cpg.EED) = "" Then GoTo EstPercBarsDrawn
ESD = Int(CDbl(vArrAllData(cRow, cpg.ESD))): EED = Int(CDbl(vArrAllData(cRow, cpg.EED)))
Select Case cView
Case Is = "W"
ESD = GetFirstDateOfWeek(ESD): EED = GetFirstDateOfWeek(EED)
Case Is = "M"
ESD = DateSerial(Year(ESD), Month(ESD), 1): EED = DateSerial(Year(EED), Month(EED), 1)
Case Is = "Q"
ESD = GetFirstDateOfQuarter(ESD): EED = GetFirstDateOfQuarter(EED)
Case Is = "HY"
ESD = GetFirstDateOfHalfYearly(ESD): EED = GetFirstDateOfHalfYearly(EED)
Case Is = "Y"
ESD = GetFirstDateOfYear(ESD): EED = GetFirstDateOfYear(EED)
End Select
If EED < TSD Or ESD > TED Then ' completely OUTSIDE timeline
If bDeleteAllAndDrawGB = False Then
On Error Resume Next
ActiveSheet.Shapes("S_TE_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
On Error Resume Next
ActiveSheet.Shapes("S_TP_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
On Error Resume Next
ActiveSheet.Shapes("S_TGB_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
On Error Resume Next
ActiveSheet.Shapes("S_ME_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
End If
bNoBar = True: GoTo EstPercBarsDrawn
End If
If ESD >= TSD And EED <= TED Then 'task dates are within the timeline
ESCol = GetTimelineDateColumnNo(sRng, ESD, cpt.TimelineStart): EECol = GetTimelineDateColumnNo(sRng, EED, cpt.TimelineEnd): GoTo drawEst
End If
If ESD >= TSD And ESD <= TED And EED >= TED Then 'Start date is within the timeline and End Date is outside the timeline
ESCol = GetTimelineDateColumnNo(sRng, ESD, cpt.TimelineStart): EECol = cpt.TimelineEnd: GoTo drawEst
End If
If ESD <= TSD And EED >= TSD And EED <= EED Then 'Start date is outside the timeline and End Date is inside the timeline
ESCol = cpt.TimelineStart: EECol = GetTimelineDateColumnNo(sRng, EED, cpt.TimelineEnd): GoTo drawEst
End If
drawEst:
Call DrawShapeForRowD(cRow, ESCol, EECol, "E")
DrawComplexPercBars: 'perc bars
If st.ShowPercBar = False Or vArrAllData(cRow, cpg.GEtype) = "M" Then GoTo EstPercBarsDrawn
If vArrAllData(cRow, cpg.PercentageCompleted) <> vbNullString Then pcompleted = vArrAllData(cRow, cpg.PercentageCompleted) Else pcompleted = 0
If bDeleteAllAndDrawGB = False Then
If pcompleted = 1 Or pcompleted = 0 Then
For sno = LBound(vArrShapes) To UBound(vArrShapes)
If vArrShapes(sno) = "S_TP_" & Cells(cRow, cpg.TID) Then ActiveSheet.Shapes(vArrShapes(sno)).Delete
Next sno
GoTo EstPercBarsDrawn
End If
End If
If pcompleted > 0 And pcompleted < 1 Then
taskduration = EED - ESD + 1: taskpercentmark = taskduration * pcompleted
If taskpercentmark < 1 Then taskpercentmark = 1
TPD = Int(CDbl(ESD + taskpercentmark)) - 1
Select Case cView
Case Is = "W"
TPD = GetFirstDateOfWeek(TPD)
Case Is = "M"
TPD = DateSerial(Year(TPD), Month(TPD), 1)
Case Is = "Q"
TPD = GetFirstDateOfQuarter(TPD)
Case Is = "HY"
TPD = GetFirstDateOfHalfYearly(TPD)
Case Is = "Y"
TPD = GetFirstDateOfYear(TPD)
End Select
If bDeleteAllAndDrawGB = False And TPD < TSD Then
For sno = LBound(vArrShapes) To UBound(vArrShapes)
If vArrShapes(sno) = "S_TP_" & Cells(cRow, cpg.TID) Then ActiveSheet.Shapes(vArrShapes(sno)).Delete
Next sno
GoTo EstPercBarsDrawn
End If
If TPD < TSD Then GoTo EstPercBarsDrawn
If TPD >= TSD And TPD <= TED Then 'TPD inside timeline
If ESD <= TSD Then
ESCol = cpt.TimelineStart: EECol = GetTimelineDateColumnNo(sRng, TPD, cpt.TimelineEnd): GoTo drawPerc
End If
If ESD > TSD Then
ESCol = GetTimelineDateColumnNo(sRng, ESD, cpt.TimelineStart): EECol = GetTimelineDateColumnNo(sRng, TPD, cpt.TimelineEnd): GoTo drawPerc
End If
End If
If TPD > TED Then 'TPD after timeline
If ESD <= TSD Then
ESCol = cpt.TimelineStart: EECol = cpt.TimelineEnd: GoTo drawPerc
End If
If ESD > TSD Then
ESCol = GetTimelineDateColumnNo(sRng, ESD, cpt.TimelineStart): EECol = cpt.TimelineEnd: GoTo drawPerc
End If
End If
drawPerc:
Call DrawShapeForRowD(cRow, ESCol, EECol, "P")
End If
End If
EstPercBarsDrawn:
If typeofbar = basBars Or typeofbar = allBars Then
If st.ShowBaselineBar Then
If IsDate(vArrAllData(cRow, cpg.BSD)) And IsDate(vArrAllData(cRow, cpg.BED)) Then
BSD = Int(CDbl(CDate(vArrAllData(cRow, cpg.BSD)))): BED = Int(CDbl(CDate(vArrAllData(cRow, cpg.BED))))
Select Case cView
Case Is = "W"
BSD = GetFirstDateOfWeek(BSD): BED = GetFirstDateOfWeek(BED)
Case Is = "M"
BSD = DateSerial(Year(BSD), Month(BSD), 1): BED = DateSerial(Year(BED), Month(BED), 1)
Case Is = "Q"
BSD = GetFirstDateOfQuarter(BSD): BED = GetFirstDateOfQuarter(BED)
Case Is = "HY"
BSD = GetFirstDateOfHalfYearly(BSD): BED = GetFirstDateOfHalfYearly(BED)
Case Is = "Y"
BSD = GetFirstDateOfYear(BSD): BED = GetFirstDateOfYear(BED)
End Select
If BED < TSD Or BSD > TED Then
If bDeleteAllAndDrawGB = False Then
On Error Resume Next
ActiveSheet.Shapes("S_TB_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
End If
GoTo nexa
End If
If BSD <= TED And BED >= TSD Then
BSCol = GetTimelineDateColumnNo(sRng, BSD, cpt.TimelineStart): BECol = GetTimelineDateColumnNo(sRng, BED, cpt.TimelineEnd)
Call DrawShapeForRowD(cRow, BSCol, BECol, "B")
End If
End If
End If
End If
If typeofbar = actBars Or typeofbar = allBars Then
If st.ShowActualBar Then
If IsDate(vArrAllData(cRow, cpg.ASD)) And IsDate(vArrAllData(cRow, cpg.AED)) Then
ASD = Int(CDbl(CDate(vArrAllData(cRow, cpg.ASD)))): AED = Int(CDbl(CDate(vArrAllData(cRow, cpg.AED))))
Select Case cView
Case Is = "W"
ASD = GetFirstDateOfWeek(ASD): AED = GetFirstDateOfWeek(AED)
Case Is = "M"
ASD = DateSerial(Year(ASD), Month(ASD), 1): AED = DateSerial(Year(AED), Month(AED), 1)
Case Is = "Q"
ASD = GetFirstDateOfQuarter(ASD): AED = GetFirstDateOfQuarter(AED)
Case Is = "HY"
BSD = ASD = GetFirstDateOfHalfYearly(ASD): AED = GetFirstDateOfHalfYearly(AED)
Case Is = "Y"
ASD = GetFirstDateOfYear(ASD): AED = GetFirstDateOfYear(AED)
End Select
If AED < TSD Or ASD > TED Then
If bDeleteAllAndDrawGB = False Then
On Error Resume Next
ActiveSheet.Shapes("S_TA_" & Cells(cRow, cpg.TID)).Delete
On Error GoTo 0
End If
GoTo nexa
End If
If ASD <= TED And AED >= TSD Then
ASCol = GetTimelineDateColumnNo(sRng, ASD, cpt.TimelineStart): AECol = GetTimelineDateColumnNo(sRng, AED, cpt.TimelineEnd)
Call DrawShapeForRowD(cRow, ASCol, AECol, "A")
End If
End If
End If
End If
nexa:
If bNoBar Then bNoBar = False: Call setColors(cRow, strGEtype, "E")
Next cRow
final:
tlog "DrawGanttBars: " & typeofbar & "-" & familytype
ReDim vArrAllData(1, 1)
If typeofbar = estBars Or typeofbar = allBars Then If st.ShowDepLines Then Call DrawDependencyLines
If bDeleteAllAndDrawGB Or cView = "HH" Then Call DrawTGBAll
Call OrderShapes: Call AddLineForToday: ActiveWindow.Zoom = lzoomLevel: bDeleteAllAndDrawGB = False: Call UpdateShapeColumns: Call HideNonWorkingColumns
End Sub

Sub DrawShapeForRowD(ByVal cRow As Long, ByVal sCol As Long, ByVal eCol As Long, sBarType As String)
Dim ts As Shape, s As Shape: Dim strShapeInfo As String, sConstructed As String, strGEtype As String: Dim c As Range
Dim bParentTask As Boolean, bShapeFound As Boolean, bNewShape As Boolean, bOverdue As Boolean, bIsTask As Boolean, bPCompleted As Boolean
Dim pleft As Double, pLeftGap As Double, pWidthGap As Double, pwidth As Double, pTop As Double, pHeight As Double
Dim savedShapeL As Double, savedShapeT As Double, savedShapeR As Double, savedShapeB As Double, roundedL As Double, roundedT As Double, roundedR As Double, roundedB As Double:
bShapeFound = False: bNewShape = True:
If IsOverdue(cRow) And bShowOverdueBars Then bOverdue = True Else bOverdue = False
bIsTask = IsTask(cRow): strGEtype = getGEtype(cRow): Set c = Cells(cRow, sCol):
If bIsTask Then
sConstructed = "S_T" & sBarType & "_" & vArrAllData(cRow, cpg.TID): pwidth = Range(c, Cells(cRow, eCol)).Width
Else
sConstructed = "S_M" & sBarType & "_" & vArrAllData(cRow, cpg.TID): pwidth = c.Width
End If
On Error Resume Next
Set ts = ActiveSheet.Shapes(sConstructed)
On Error GoTo 0
If ts Is Nothing Then bShapeFound = False Else bShapeFound = True
If bIsTask Then
If st.ShowDepLines Then
If vArrAllData(cRow, cpg.Dependency) = "" And vArrAllData(cRow, cpg.Dependents) = "" Then
pLeftGap = 1: pwidth = pwidth - (pLeftGap * 2): GoTo gapwidth
End If
If vArrAllData(cRow, cpg.Dependency) <> "" And vArrAllData(cRow, cpg.Dependents) <> "" Then
pLeftGap = 5: pWidthGap = 5:pwidth = pwidth - (pWidthGap * 2): GoTo gapwidth
End If
If vArrAllData(cRow, cpg.Dependency) <> "" And vArrAllData(cRow, cpg.Dependents) = "" Then
If InStr(1, vArrAllData(cRow, cpg.Dependency), "FS") > 0 Or InStr(1, vArrAllData(cRow, cpg.Dependency), "SS") > 0 Then
pLeftGap = 5: pWidthGap = 5: pwidth = pwidth - (pWidthGap + 1):
GoTo gapwidth
End If
If InStr(1, vArrAllData(cRow, cpg.Dependency), "FF") > 0 Or InStr(1, vArrAllData(cRow, cpg.Dependency), "SF") > 0 Then
pLeftGap = 1: pWidthGap = 5: pwidth = pwidth - (pWidthGap + 1):
GoTo gapwidth
End If
End If
If vArrAllData(cRow, cpg.Dependents) <> "" And vArrAllData(cRow, cpg.Dependency) = "" Then
pLeftGap = 5:pwidth = pwidth - (pLeftGap - 1): pLeftGap = 1:GoTo gapwidth
End If
Else
pLeftGap = 1: pwidth = pwidth - (pLeftGap * 2): GoTo gapwidth
End If
End If
gapwidth:
If bIsTask Then
pleft = c.Left + pLeftGap: pHeight = c.Height * BHeightRatio: pTop = c.Top + pHeight: pHeight = c.Height - (pHeight * 2): bParentTask = IsParentTask(cRow)
If sBarType = "B" Then
pHeight = pHeight * lheightFactorOthers
ElseIf sBarType = "A" Then
pTop = pTop + (pHeight * (1 - lheightFactorOthers)): pHeight = pHeight * lheightFactorOthers
End If
Else
pTop = c.Top + (c.Height * MHeightRatio): pHeight = c.Height - (c.Height * MHeightRatio * 2): pleft = c.Left + ((c.Width - pHeight) / 2):
End If
If bShapeFound Then
If sBarType = "E" Then strShapeInfo = CStr(vArrAllData(cRow, cpg.ShapeInfoE))
If sBarType = "B" Then strShapeInfo = CStr(vArrAllData(cRow, cpg.ShapeInfoB))
If sBarType = "A" Then strShapeInfo = CStr(vArrAllData(cRow, cpg.ShapeInfoA))
If strShapeInfo <> "" Then ' compares shapeinfo with new coordinates
savedShapeL = getSavedShapeL(strShapeInfo): savedShapeT = getSavedShapeT(strShapeInfo): savedShapeR = getSavedShapeR(strShapeInfo): savedShapeB = getSavedShapeB(strShapeInfo)
End If
If bIsTask Then 'new coordinates
roundedL = WorksheetFunction.RoundUp(pleft, 1): roundedT = WorksheetFunction.RoundUp(pTop, 1): roundedR = WorksheetFunction.RoundUp(pleft + pwidth, 1): roundedB = WorksheetFunction.RoundUp(pTop + pHeight, 1)
Else
roundedL = WorksheetFunction.RoundUp(pleft, 1): roundedT = WorksheetFunction.RoundUp(pTop, 1): roundedR = savedShapeR: roundedB = savedShapeB ' not important
End If
If sBarType = "P" Then
If sCol = ts.TopLeftCell.column And Int(pwidth) = Int(ts.Width) Then Set s = ts: bNewShape = False Else ts.Delete: bNewShape = True
Else
If roundedL = savedShapeL And roundedT = savedShapeT And roundedR = savedShapeR And roundedB = savedShapeB Then Set s = ts: bNewShape = False Else ts.Delete: bNewShape = True
End If
Else
bNewShape = True
End If
If bNewShape Then
If bIsTask Then
If bParentTask Then
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeLeftRightArrow, Left:=pleft, Top:=pTop, Width:=pwidth, Height:=pHeight)
Else
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=pleft, Top:=pTop, Width:=pwidth, Height:=pHeight)
End If
s.Name = "S_T" & sBarType & "_" & vArrAllData(cRow, cpg.TID)
Else
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeFlowchartDecision, Left:=pleft, Top:=pTop, Width:=pHeight, Height:=pHeight)
s.Name = "S_M" & sBarType & "_" & vArrAllData(cRow, cpg.TID):
End If
s.Line.visible = msoFalse:
If st.CurrentView = "D" Then s.OnAction = "ShapeClicked"
If bDeleteAllAndDrawGB = False And st.ShowTextinBars And sBarType = "E" Then Call DrawTGBAll(s)
End If
Call setColors(cRow, strGEtype, sBarType, s)
Last:
If bDeleteAllAndDrawGB = False And st.ShowTextinBars And sBarType = "E" Then Call DrawTGBAll(s)'limitation - very slow
Set s = Nothing: pleft = 0: pHeight = 0: pTop = 0: pHeight = 0: savedShapeL = 0: savedShapeT = 0: savedShapeR = 0: savedShapeB = 0: roundedL = 0: roundedT = 0: roundedR = 0: roundedB = 0
End Sub
Sub ShapeClicked()
DA:
If st.CurrentView <> "D" Then Cells(rowsix, cpg.LC) = "": MsgBox msg(69): Exit Sub
Dim s As Shape: Dim TID As Long, tidRow As Long: Dim rngSName As Range: Dim sType As String, strName As String:
Dim bIsTask As Boolean: Set rngSName = Cells(rowsix, cpg.LC)
ActiveSheet.Shapes(Application.Caller).Select: rngSName.value = Application.Caller
Set s = ActiveSheet.Shapes(rngSName.value): sType = getShapeType(s): TID = getTIDfromShape(s): tidRow = getTIDRow(TID)
If IsParentTask(tidRow) Then MsgBox msg(1): GoTo Last
If getGETypefromShape(s) = "T" Then bIsTask = True: strName = "S_TE_" Else bIsTask = False: strName = "S_ME_"
If sType = "TP" Or sType = "MP" Or sType = "TGB" Then
If sType = "TGB" Then ActiveSheet.Shapes("S_TGB_" & TID).Delete
If sType = "TP" Then ActiveSheet.Shapes("S_TP_" & TID).Delete
If shapeExists(strName & TID) Then Set s = ActiveSheet.Shapes(strName & TID) Else MsgBox msg(69)
End If
s.Select: sType = getShapeType(s): rngSName.value = s.Name:

If sType = "TE" Or sType = "TB" Or sType = "TA" Or sType = "ME" Or sType = "MB" Or sType = "MA" Then
Call EA: Call CheckForMove(s)
End If
Last:
Cells(tidRow, cpg.ED).Select: EA
End Sub

Sub CheckForMove(s As Shape) ' uses clicked pos and moved pos
Call DA
tlog "checkmove"
Dim rngSName As Range: Set rngSName = Cells(rowsix, cpg.LC): rngSName.Font.Color = st.cPRC
Dim r As Object
If rngSName.value = Empty Then Exit Sub
Dim bMoved As Boolean, bStartDateChanged As Boolean, bEndDateChanged As Boolean: Dim strShape As String, sType As String
Dim leftpos As Double, toppos As Double, rightpos As Double, ShapeInfoL As Double, ShapeInfoT As Double, ShapeInfoR As Double, ShapeInfoD As Double
Dim newscol As Long, tidr As Long: Dim curDown As Double, orgStartHrs As Double, orgEndHrs As Double
Dim origLeft As Double, origTop As Double, origRight As Double, origDown As Double, curleft As Double, curTop As Double, curRight As Double
Dim dSCol As Long, dECol As Long: Dim sTopLeftColDate As Date, sBotRigColDate As Date, startDate As Date, endDate As Date, dTSD As Date, dTED As Date
Dim vstr1 As Variant: Dim curWs As Worksheet: Set curWs = ActiveSheet
Set s = ActiveSheet.Shapes(rngSName.value): s.Select: tidr = getTIDRowfromShape(s): sType = getShapeType(s)
origLeft = WorksheetFunction.RoundUp(s.Left, 1): origTop = WorksheetFunction.RoundUp(s.Top, 1):
origRight = WorksheetFunction.RoundUp((s.Left + s.Width), 1): origDown = WorksheetFunction.RoundUp((s.Top + s.Height), 1)
If sType = "TE" Or sType = "ME" Then vstr1 = Split(Cells(tidr, cpg.ShapeInfoE), DepSeperator): dSCol = cpg.ESD: dECol = cpg.EED
If sType = "TB" Or sType = "MB" Then vstr1 = Split(Cells(tidr, cpg.ShapeInfoB), DepSeperator): dSCol = cpg.BSD: dECol = cpg.BED
If sType = "TA" Or sType = "MA" Then vstr1 = Split(Cells(tidr, cpg.ShapeInfoA), DepSeperator): dSCol = cpg.ASD: dECol = cpg.AED
ShapeInfoL = vstr1(2): ShapeInfoT = vstr1(3): ShapeInfoR = vstr1(4): ShapeInfoD = vstr1(5)
dTSD = CDate(Cells(rowsix, cpt.TimelineStart)): dTED = CDate(Cells(rowsix, cpt.TimelineEnd))
Call EA
For CountNum = 1 To 50000
DoEvents
If ThisWorkbook.Name <> ActiveWorkbook.Name Then GoTo endcheck
If curWs.Name <> ActiveSheet.Name Then GoTo endcheck
If rngSName.value = Empty Then GoTo endcheck
Set r = Selection
If TypeName(r) = "Range" Then GoTo endcheck
curleft = WorksheetFunction.RoundUp(s.Left, 1): curTop = WorksheetFunction.RoundUp(s.Top, 1):
curRight = WorksheetFunction.RoundUp((s.Left + s.Width), 1): curDown = WorksheetFunction.RoundUp((s.Top + s.Height), 1)
If curleft <> origLeft And curRight = origRight Then ' only l pos changed
If s.TopLeftCell.column < cpt.TimelineStart Then
startDate = dTSD
Else
startDate = CDate(Cells(rowsix, s.TopLeftCell.column))
End If
bMoved = True: bStartDateChanged = True
End If
If curleft = origLeft And curRight <> origRight Then ' only r pos changed
If getGETypefromShape(s) = "T" Then
If s.BottomRightCell.column > cpt.TimelineEnd Then
endDate = dTED: bEndDateChanged = True
Else
endDate = CDate(Cells(rowsix, s.BottomRightCell.column)): bEndDateChanged = True
End If
ElseIf getGETypefromShape(s) = "M" Then
If s.BottomRightCell.column > cpt.TimelineEnd Then
startDate = CDate(Cells(rowsix, s.TopLeftCell.column)): bStartDateChanged = True
End If
End If
bMoved = True:
End If
If curleft <> origLeft And curRight <> origRight Then ' both position changed
If getGETypefromShape(s) = "T" Then
If s.TopLeftCell.column < cpt.TimelineStart Then
startDate = dTSD:
ElseIf s.TopLeftCell.column > cpt.TimelineEnd Then
startDate = dTED
Else
startDate = CDate(Cells(rowsix, s.TopLeftCell.column))
End If
bMoved = True: bStartDateChanged = True: ' bEndDateChanged = True
ElseIf getGETypefromShape(s) = "M" Then
If s.TopLeftCell.column < cpt.TimelineStart Then
startDate = dTSD:
ElseIf s.TopLeftCell.column > cpt.TimelineEnd Then
startDate = dTED
Else
startDate = CDate(Cells(rowsix, s.TopLeftCell.column))
End If
bMoved = True: bStartDateChanged = True:
End If
End If
If curTop <> origTop Or curDown <> origDown Then
bMoved = True
End If
If bMoved Then rngSName.ClearContents:
Next CountNum
endcheck:
rngSName.ClearContents
Call EA
If bMoved Then
Call DA: Call UpdateShapeColumns: Call EA
If bMoved And st.CurrentView <> "D" Then MsgBox msg(69): GoTo Last
If bStartDateChanged Then
If st.HGC Then
orgStartHrs = sArr.ResourceP(0, 10): Cells(tidr, dSCol) = startDate + orgStartHrs
Else
Cells(tidr, dSCol) = startDate:
End If
End If
If bEndDateChanged Then
If st.HGC Then
orgEndHrs = sArr.ResourceP(0, 11): Cells(tidr, dECol) = endDate + orgEndHrs
Else
Cells(tidr, dECol) = endDate
End If
End If
If startDate = False And endDate = False Then GoTo Last
End If
Exit Sub
Last:
Call DA: Call DrawGanttBars(tidr): Call EA
tlog "checkmove"
End Sub

Sub DrawGanttBarsHours(typeofbar As String, resname As String, rownum As Long, taskstartdate As Date, taskEndDate As Date)
Dim bParentTask As Boolean: Dim drawrange As Range: Dim sAs Shape
Dim dTSD As Double, TimelineED As Double, TimelineSD As Double, dTED As Double, noofdays As Double, pwidth As Double
Dim findStartDate As Variant, findEndDate As Variant, findStartTime As Variant, findEndTime As Variant: Dim newtaskstartdate As Date, taskpercentmarkdate As Date, neweed As Date
Dim percED As Long, i As Long, cRow As Long, x As Long, y As Long, drawfromcol As Long, drawtillcol As Long, searchfromcol As Long, searchtillcol As Long, minutesstarthour As Long, minutesendhour As Long, wkday As Long
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer:
cRow = rownum: dTSD = taskstartdate: dTED = taskEndDate: TimelineSD = Int(CDbl(Cells(rowsix, cpt.TimelineStart).value)): TimelineED = Int(CDbl(Cells(rowsix, cpt.TimelineEnd).value))
If IsParentTask(cRow) Then bParentTask = True
Call getResValue(resname, sArr.ResourceP): resStartHour = Format(sArr.ResourceP(resvalue, 10), "h"): resEndHour = Format(sArr.ResourceP(resvalue, 11), "h"):
If resEndHour = 0 Then resEndHour = 24
daystartHour = Format(dTSD, "h"): dayendHour = Format(dTED, "h"): minutesstarthour = Minute(dTSD): minutesendhour = Minute(dTED)
If minutesstarthour >= 30 Then dTSD = Int(CDbl(dTSD)) + TimeSerial(Hour(dTSD), 0, 0) Else dTSD = Int(CDbl(dTSD)) + TimeSerial(Hour(dTSD), 0, 0)
If minutesendhour >= 30 Then dTED = Int(CDbl(dTED)) + TimeSerial(Hour(dTED), 60, 0) Else dTED = Int(CDbl(dTED)) + TimeSerial(Hour(dTED), 0, 0)
findStartDate = Application.Match(dTSD, Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)), 0)
findEndDate = Application.Match(dTED, Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)), 0)
If IsError(findStartDate) And IsError(findEndDate) Then
If dTSD < TimelineSD And dTED > TimelineED Then searchfromcol = cpt.TimelineStart: searchtillcol = cpt.TimelineEnd: GoTo startdraw Else Exit Sub
End If
If IsError(findStartDate) = True And IsError(findEndDate) = False Then searchfromcol = cpt.TimelineStart: searchtillcol = cpt.TimelineStart + findEndDate - 2: GoTo startdraw
If IsError(findStartDate) = False And IsError(findEndDate) = True Then searchfromcol = cpt.TimelineStart + findStartDate - 1: searchtillcol = cpt.TimelineEnd: GoTo startdraw
If IsError(findStartDate) = False And IsError(findEndDate) = False Then searchfromcol = cpt.TimelineStart + findStartDate - 1: searchtillcol = cpt.TimelineStart + findEndDate - 2: GoTo startdraw
startdraw:'Draw base shape
drawfromcol = searchfromcol: drawtillcol = searchtillcol:
If drawtillcol = drawfromcol - 1 Then drawtillcol = drawfromcol
Set drawrange = Range(Cells(cRow, drawfromcol), Cells(cRow, drawtillcol))
If vArrAllData(cRow, 1) = "M" Then
Call DrawShapeForRowH(drawrange, msoShapeFlowchartDecision, cRow, typeofbar, "S_ME_" & vArrAllData(cRow, cpg.TID)): GoTo drawPerc
Else
If bParentTask Then
Call DrawShapeForRowH(drawrange, msoShapeLeftRightArrow, cRow, typeofbar & "basepar", "S_TE_" & vArrAllData(cRow, cpg.TID))
If vArrAllData(cRow, cpg.PercentageCompleted) <> 0 And st.ShowPercBar Then Call DrawShapeForRowH(drawrange, msoShapeLeftRightArrow, cRow, typeofbar & "PP", "S_TP_" & vArrAllData(cRow, cpg.TID))
Exit Sub
Else
Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar & "base", "S_TH_" & vArrAllData(cRow, cpg.TID))
End If
End If
If typeofbar = "E" And vArrAllData(cRow, 1) = "T" Then 'DRAW BREAKS
noofdays = RoundUp(Int(CDbl(taskEndDate)) - Int(CDbl(taskstartdate))) + 1
If noofdays <= 1 Then
If daystartHour > resStartHour Then
findStartTime = Application.Match(daystartHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol)), 0): findEndTime = Application.Match(dayendHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol + 2)), 0)
If IsError(findEndTime) Then drawtillcol = cpt.TimelineEnd Else drawfromcol = searchfromcol + findStartTime - 1: drawtillcol = searchfromcol + findEndTime - 2
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol)), 0): findEndTime = Application.Match(dayendHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol + 2)), 0)
If IsError(findEndTime) Then drawtillcol = cpt.TimelineEnd Else drawfromcol = searchfromcol + findStartTime - 1: drawtillcol = searchfromcol + findEndTime - 2
End If
Set drawrange = Range(Cells(cRow, drawfromcol), Cells(cRow, drawtillcol)):
Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar, "S_TE_" & vArrAllData(cRow, cpg.TID))
GoTo ContinueNext
End If
x = cpt.TimelineStart: y = x + 24: newtaskstartdate = taskstartdate
For i = 1 To noofdays
If x > cpt.TimelineEnd Then Exit For
If TimelineSD > Int(CDbl(newtaskstartdate)) Then
If Int(CDbl(newtaskstartdate)) <= TimelineSD Then newtaskstartdate = newtaskstartdate + 1
GoTo ContinueNext
End If
wkday = Weekday(Cells(rowsix, x), 2)
If Cells(rowsix, x).value < Int(CDbl(dTSD)) Then x = x + 24: y = y + 24: i = i - 1: GoTo ContinueNext
If IsDateAHoliday(resvalue, Cells(rowsix, x)) = True Or IsDateAWorkday(resvalue, wkday) = False Then x = x + 24: y = y + 24: GoTo ContinueNext
If noofdays > 1 Then
If i = 1 And daystartHour >= resEndHour Then x = x + 24: y = y + 24: GoTo ContinueNext
If i = 1 And daystartHour > resStartHour Then
findStartTime = Application.Match(daystartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
If resEndHour = 24 Then findEndTime = 25 Else findEndTime = Application.Match(resEndHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
If resEndHour = 24 Then findEndTime = 25 Else findEndTime = Application.Match(resEndHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
End If
If i = noofdays Then
If Cells(rowsix, x).value >= dTED Then
GoTo ContinueNext
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
findEndTime = Application.Match(dayendHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
End If
End If
If IsError(findStartTime) Or IsError(findEndTime) Then
drawfromcol = cpt.TimelineStart + ((i - 1) * 24) + findStartTime - 1: drawtillcol = cpt.TimelineEnd
Else
drawfromcol = x + findStartTime - 1: drawtillcol = x + findEndTime - 2
If findEndTime <= findStartTime Then GoTo ContinueNext
End If
Set drawrange = Range(Cells(cRow, drawfromcol), Cells(cRow, drawtillcol)):
Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar, "S_TE_" & vArrAllData(cRow, cpg.TID) & "_" & "i" & i)
Else
End If
x = x + 24: y = y + 24
ContinueNext:
Next
End If
drawPerc:
i = 0
If vArrAllData(cRow, cpg.PercentageCompleted) <> 0 And st.ShowPercBar Then
If vArrAllData(cRow, cpg.GEtype) = "M" And vArrAllData(cRow, cpg.PercentageCompleted) <> 1 Then Exit Sub
If noofdays > 1 Then
percED = vArrAllData(cRow, cpg.ED) * vArrAllData(cRow, cpg.PercentageCompleted)
If percED <= 1 Then percED = 1
neweed = CalEEDHrs(Cells(cRow, cpg.Resource), taskstartdate, percED): Call DrawGanttBarsHoursPerc("E", resname, cRow, taskstartdate, neweed)
Else
If vArrAllData(cRow, cpg.GEtype) = "T" Then Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar & "PP", "S_TP_" & vArrAllData(cRow, cpg.TID)) Else Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar & "PP", "S_MP_" & vArrAllData(cRow, cpg.TID))
End If
End If
End Sub

Sub DrawShapeForRowH(drawrange As Range, shapeType As MsoAutoShapeType, cRow As Long, typeofbar As String, shapeName As String)
Dim leftfrom As Double, topfrom As Double, widshp As Double, hgtshp As Double, leftoffset As Double, widoffset As Double, cellwidth As Double:
Dim s As Shape: Dim bOverdue As Boolean: Dim strGEtype As String
strGEtype = getGEtype(cRow)
If IsOverdue(cRow) And bShowOverdueBars Then bOverdue = True Else bOverdue = False
With drawrange
leftfrom = .Left:topfrom = .Top + 3:widshp = .Width:hgtshp = .Height
End With
hgtshp = hgtshp * 0.6
If hgtshp <= 0 Then hgtshp = 0
If st.ShowDepLines And Cells(cRow, 1) = "T" = True Then
cellwidth = Columns(cpt.TimelineStart).Width: leftoffset = cellwidth * 0.1: widoffset = leftoffset * 3: leftfrom = leftfrom + leftoffset: widshp = widshp - widoffset
End If
If Cells(cRow, 1) = "T" Then
If typeofbar = "EPP" Then
If Cells(cRow, cpg.PercentageCompleted) > 0 And Cells(cRow, cpg.PercentageCompleted) <= 1 Then widshp = widshp * Cells(cRow, cpg.PercentageCompleted)
End If
End If
If Cells(cRow, 1) = "T" Then Set s = ActiveSheet.Shapes.AddShape(shapeType, leftfrom + leftoffset, topfrom, widshp, hgtshp) Else Set s = ActiveSheet.Shapes.AddShape(msoShapeFlowchartDecision, leftfrom + 3, topfrom + 1, widshp - 5, hgtshp)
With s
.Name = shapeName:.Line.visible = msoFalse
End With
Call setColors(cRow, strGEtype, typeofbar, s)
End Sub

Sub colorShape(s As Shape, lngColor As Long, cRow As Long)
With s
With .Fill
.Solid: .ForeColor.RGB = lngColor
If bFreeVersion And cRow >= 20 Then .visible = msoFalse Else .visible = msoTrue
End With
End With
End Sub

Sub OrderShapes()
Dim ShAs Shape
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 4) = "S_TP" Then Sh.ZOrder msoBringToFront
Next Sh
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 4) = "S_TB" Then Sh.ZOrder msoBringToFront
Next Sh
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 4) = "S_TA" Then Sh.ZOrder msoBringToFront
Next Sh
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 5) = "S_TGB" Then Sh.ZOrder msoBringToFront
Next Sh
End Sub
Sub DrawGanttBarsHoursPerc(typeofbar As String, resname As String, rownum As Long, taskstartdate As Date, percEndDate As Date)
Dim drawrange As Range: Dim sAs Shape
Dim dTSD As Double, TimelineED As Double, TimelineSD As Double, dTED As Double, noofdays As Double
Dim findStartDate As Variant, findEndDate As Variant, findStartTime As Variant, findEndTime As Variant: Dim newtaskstartdate As Date
Dim i As Long, cRow As Long, x As Long, y As Long, drawfromcol As Long, drawtillcol As Long, searchfromcol As Long, searchtillcol As Long, minutesstarthour As Long, minutesendhour As Long, wkday As Long
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer: Dim drawpercentbar As Boolean
cRow = rownum: dTSD = taskstartdate: dTED = percEndDate: TimelineSD = Int(CDbl(Cells(rowsix, cpt.TimelineStart).value)): TimelineED = Int(CDbl(Cells(rowsix, cpt.TimelineEnd).value))
Call getResValue(resname, sArr.ResourceP): resStartHour = Format(sArr.ResourceP(resvalue, 10), "h"): resEndHour = Format(sArr.ResourceP(resvalue, 11), "h")
If resEndHour = 0 Then resEndHour = 24
daystartHour = Format(dTSD, "h"): dayendHour = Format(dTED, "h"): minutesstarthour = Minute(dTSD): minutesendhour = Minute(dTED)
If minutesstarthour >= 30 Then dTSD = Int(CDbl(dTSD)) + TimeSerial(Hour(dTSD), 0, 0) Else dTSD = Int(CDbl(dTSD)) + TimeSerial(Hour(dTSD), 0, 0)
If minutesendhour >= 30 Then dTED = Int(CDbl(dTED)) + TimeSerial(Hour(dTED), 60, 0) Else dTED = Int(CDbl(dTED)) + TimeSerial(Hour(dTED), 0, 0)
findStartDate = Application.Match(dTSD, Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)), 0): findEndDate = Application.Match(dTED, Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)), 0)
If IsError(findStartDate) And IsError(findEndDate) Then
If dTSD < TimelineSD And dTED > TimelineED Then
searchfromcol = cpt.TimelineStart: searchtillcol = cpt.TimelineEnd
If searchtillcol < searchfromcol Then searchtillcol = searchfromcol
GoTo startdraw
Else
Exit Sub
End If
End If
If IsError(findStartDate) = True And IsError(findEndDate) = False Then
searchfromcol = cpt.TimelineStart: searchtillcol = cpt.TimelineStart + findEndDate - 2:
If searchtillcol < searchfromcol Then searchtillcol = searchfromcol
GoTo startdraw
End If
If IsError(findStartDate) = False And IsError(findEndDate) = True Then
searchfromcol = cpt.TimelineStart + findStartDate - 1: searchtillcol = cpt.TimelineEnd
If searchtillcol < searchfromcol Then searchtillcol = searchfromcol
GoTo startdraw
End If
If IsError(findStartDate) = False And IsError(findEndDate) = False Then
searchfromcol = cpt.TimelineStart + findStartDate - 1: searchtillcol = cpt.TimelineStart + findEndDate - 1
If searchtillcol < searchfromcol Then searchtillcol = searchfromcol
GoTo startdraw
End If
startdraw:
If typeofbar = "E" And Cells(cRow, 1) = "T" Then
noofdays = RoundUp(Int(CDbl(percEndDate)) - Int(CDbl(taskstartdate))) + 1
If noofdays <= 1 Then
If daystartHour > resStartHour Then
findStartTime = Application.Match(daystartHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol)), 0)
findEndTime = Application.Match(dayendHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol + 2)), 0)
If IsError(findEndTime) Then drawtillcol = cpt.TimelineEnd Else drawfromcol = searchfromcol + findStartTime - 1: drawtillcol = searchfromcol + findEndTime - 2
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol)), 0)
findEndTime = Application.Match(dayendHour, Range(Cells(roweight, searchfromcol), Cells(roweight, searchtillcol + 2)), 0)
If IsError(findEndTime) Then drawtillcol = cpt.TimelineEnd Else drawfromcol = searchfromcol + findStartTime - 1: drawtillcol = searchfromcol + findEndTime - 2'2
End If
If drawtillcol < drawfromcol Then drawtillcol = drawtillcol + 1
Set drawrange = Range(Cells(cRow, drawfromcol), Cells(cRow, drawtillcol)):
Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar & "P", "S_TP_" & Cells(cRow, cpg.TID))
GoTo ContinueNext
End If
x = cpt.TimelineStart: y = x + 24:newtaskstartdate = taskstartdate
For i = 1 To noofdays
If x > cpt.TimelineEnd Then Exit For
If TimelineSD > Int(CDbl(newtaskstartdate)) Then
If Int(CDbl(newtaskstartdate)) <= TimelineSD Then newtaskstartdate = newtaskstartdate + 1
GoTo ContinueNext
End If
wkday = Weekday(Cells(rowsix, x), 2)
If Cells(rowsix, x).value < Int(CDbl(dTSD)) Then
x = x + 24: y = y + 24: i = i - 1: GoTo ContinueNext
End If
If IsDateAHoliday(resvalue, Cells(rowsix, x)) = True Or IsDateAWorkday(resvalue, wkday) = False Then
x = x + 24: y = y + 24: GoTo ContinueNext
End If
If noofdays > 1 Then
If i = 1 And daystartHour >= resEndHour Then x = x + 24: y = y + 24: GoTo ContinueNext
If i = 1 And daystartHour > resStartHour Then
findStartTime = Application.Match(daystartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
If resEndHour = 24 Then findEndTime = 25 Else findEndTime = Application.Match(resEndHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
If resEndHour = 24 Then findEndTime = 25 Else findEndTime = Application.Match(resEndHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
End If
If i = noofdays Then
If Cells(rowsix, x).value >= dTED Then
GoTo ContinueNext
Else
findStartTime = Application.Match(resStartHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
findEndTime = Application.Match(dayendHour, Range(Cells(roweight, x), Cells(roweight, y)), 0)
End If
End If
If IsError(findStartTime) Or IsError(findEndTime) Then drawfromcol = cpt.TimelineStart + ((i - 1) * 24) + findStartTime - 1: drawtillcol = cpt.TimelineEnd Else drawfromcol = x + findStartTime - 1: drawtillcol = x + findEndTime - 2
Set drawrange = Range(Cells(cRow, drawfromcol), Cells(cRow, drawtillcol)):
Call DrawShapeForRowH(drawrange, msoShapeRectangle, cRow, typeofbar & "P", "S_TP_" & Cells(cRow, cpg.TID) & "_" & "i" & i)
Else
End If
x = x + 24: y = y + 24
ContinueNext:
Next
i = 0
End If
End Sub

Sub AddLineForToday()
On Error Resume Next
ActiveSheet.Shapes("ST_Today_LineLeft").Delete:
On Error GoTo 0
On Error Resume Next
ActiveSheet.Shapes("ST_Today_LineRight").Delete:
On Error GoTo 0
If st.ShowTodayLines = False Then Exit Sub
Dim ShAs Shape: Dim r As Range, f As Range, tRng As Range:Dim pleftAs Double, pwidth As Double, pTop As Double, pHeight As Double
Dim lSOffset As Long, lEoffset As Long, lrow As Long, cPos As Long: Dim tDate As Date
Set tRng = Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd))
If st.CurrentView = "HH" Then
tDate = DateAdd("h", Hour(Now), Date)
ElseIf st.CurrentView = "D" Then
tDate = Date
ElseIf st.CurrentView = "W" Then
tDate = GetFirstDateOfWeek(Date)
ElseIf st.CurrentView = "M" Then
tDate = CDate(DateSerial(Year(Date), Month(Date), 1))
ElseIf st.CurrentView = "Q" Then
tDate = CDate(GetFirstDateOfQuarter(Date))
ElseIf st.CurrentView = "HY" Then
tDate = CDate(GetFirstDateOfHalfYearly(Date))
ElseIf st.CurrentView = "Y" Then
tDate = CDate(GetFirstDateOfYear(Date))
End If
Call LoadCalDateValuesArray: cPos = GetTimelineDateColumnNo(tRng, tDate, 0)
If cPos = 0 Then GoTo Last
lrow = GetLastRow
If lrow = rownine Then Exit Sub
Set r = Cells(rowseven, cPos)
If cPos = cpt.TimelineStart Then lSOffset = 0.6
If cPos = cpt.TimelineEnd Then lEoffset = 0.6
Set Sh = ActiveSheet.Shapes.AddShape(Type:=msoConnectorStraight, Left:=r.Left + lSOffset, Top:=r.Offset(1, 0).Top, Width:=0.1, Height:=Cells(lrow + 1, cPos).Top - r.Offset(1, 0).Top)
Sh.Name = "ST_Today_LineLeft"
Sh.ZOrder msoBringToFront
With Sh.Line
.visible = msoTrue:.ForeColor.RGB = st.cTLC:.Weight = 0.1
End With
Set Sh = ActiveSheet.Shapes.AddShape(Type:=msoConnectorStraight, Left:=r.Left + r.Width - lEoffset - 1, Top:=r.Offset(1, 0).Top, Width:=0.1, Height:=Cells(lrow + 1, cPos).Top - r.Offset(1, 0).Top)
Sh.Name = "ST_Today_LineRight"
Sh.ZOrder msoBringToFront
With Sh.Line
.visible = msoTrue:.ForeColor.RGB = st.cTLC:.Weight = 0.1
End With
Last:
Set Sh = Nothing:Set r = Nothing
On Error GoTo 0
End Sub

Sub DrawDependencyLines(Optional t As Boolean)
If ActiveSheet.AutoFilterMode Or st.ShowCompleted = 0 Or st.ShowInProgress = 0 Or st.ShowPlanned = 0 Or st.ShowDepLines = False Then Call DeleteShape("S_De", 4): Exit Sub
If bDeleteAllAndDrawGB = False Then Call DeleteShape("S_De", 4)
Dim cRow As Long, dIDFrom As Long, dIDTo As Long, lrow As Long, dcount As Long, i As Long: Dim vStr As Variant, dStr As Variant
Dim dShFrom As Shape, dShTo As Shape, dSh As Shape: Dim dTypeAs String
lrow = GetLastRow: dcount = WorksheetFunction.CountA(Range(Cells(rownine, cpg.Dependency), Cells(lrow, cpg.Dependency)))
If dcount < 2 Then Exit Sub
cRow = firsttaskrow: dcount = dcount - 1:
ReDim vArrAllData(1 To lrow, 1 To cpg.LC):vArrAllData = Range(Cells(1, 1), Cells(lrow, cpg.LC)).FormulaLocal
Do Until dcount = 0
If vArrAllData(cRow, cpg.Dependency) <> vbNullString Then
dcount = dcount - 1:vStr = Split(vArrAllData(cRow, cpg.Dependency), DepSeperator)
For i = 0 To UBound(vStr) - 1
dStr = Split(vStr(i), "_"): dType = dStr(1): dIDFrom = dStr(0): dIDTo = vArrAllData(cRow, cpg.TID)
On Error Resume Next
Set dShFrom = Nothing: Set dShTo = Nothing
If st.CurrentView = "HH" Then
Set dShFrom = ActiveSheet.Shapes("S_TE_" & dIDFrom)
If dShFrom Is Nothing Then Set dShFrom = ActiveSheet.Shapes("S_TH_" & dIDFrom)
If dShFrom Is Nothing Then Set dShFrom = ActiveSheet.Shapes("S_ME_" & dIDFrom)
Set dShTo = ActiveSheet.Shapes("S_TE_" & dIDTo)
If dShTo Is Nothing Then Set dShTo = ActiveSheet.Shapes("S_TH_" & dIDTo)
If dShTo Is Nothing Then Set dShTo = ActiveSheet.Shapes("S_ME_" & dIDTo)
Else
Set dShFrom = ActiveSheet.Shapes("S_TE_" & dIDFrom)
If dShFrom Is Nothing Then Set dShFrom = ActiveSheet.Shapes("S_ME_" & dIDFrom)
Set dShTo = ActiveSheet.Shapes("S_TE_" & dIDTo)
If dShTo Is Nothing Then Set dShTo = ActiveSheet.Shapes("S_ME_" & dIDTo)
End If
On Error GoTo 0
If dShFrom Is Nothing Or dShTo Is Nothing Then GoTo nextrow
Set dSh = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1, 1, 1, 1)
With dSh
.Name = "S_De_" & dType & "_" & dIDFrom & "_" & dIDTo: .ShapeStyle = msoLineStylePreset1: .Line.ForeColor.RGB = st.cDLC: .ZOrder msoSendToBack
End With
With dSh.Line
.BeginArrowheadLength = msoArrowheadLengthMedium: .BeginArrowheadWidth = msoArrowheadWidthMedium: .BeginArrowheadStyle = msoArrowheadNone
.EndArrowheadLength = msoArrowheadShort: .EndArrowheadWidth = msoArrowheadNarrow: .EndArrowheadStyle = msoArrowheadTriangle
End With
With dSh.ConnectorFormat
If IsParentTask(getTIDRow(dIDFrom)) Then
Select Case dType
Case Is = "FS"
.BeginConnect dShFrom, 8: .EndConnect dShTo, 2
Case Is = "SS"
.BeginConnect dShFrom, 4: .EndConnect dShTo, 2
Case Is = "SF"
.BeginConnect dShFrom, 4: .EndConnect dShTo, 4
Case Is = "FF"
.BeginConnect dShFrom, 8: .EndConnect dShTo, 4
End Select
Else
Select Case dType
Case Is = "FS"
.BeginConnect dShFrom, 4: .EndConnect dShTo, 2
Case Is = "SS"
.BeginConnect dShFrom, 2: .EndConnect dShTo, 2
Case Is = "SF"
.BeginConnect dShFrom, 2: .EndConnect dShTo, 4
Case Is = "FF"
.BeginConnect dShFrom, 4: .EndConnect dShTo, 4
End Select
End If
End With
nextrow:
Next
End If
cRow = cRow + 1
Loop
ReDim vArrAllData(1, 1)
End Sub
Sub FormatTaskIcons(Optional cRowOnly As Long, Optional field As String, Optional familytype As String) 'called from settings only
Exit Sub
'Dim startRow As Long, endRow As Long, crow As Long, lrow As Long: Dim pervalue As Double: Dim csp As String, cs As String:
'Dim rngStatus As Range
'If field = "" Then field = allFields
'If familytype = "" Then familytype = allRows
'tlog "FormatTasks: " & field & familytype
'lrow = GetLastRow: startRow = getStartRow(cRowOnly, familytype): endRow = getEndRow(cRowOnly, familytype)
'ReDim vArrAllData(1 To endRow, 1 To cpg.LC): vArrAllData = Range(Cells(1, 1), Cells(endRow, cpg.LC)).value:
'For crow = startRow To endRow
Call setColors(crow, CStr(vArrAllData(crow, 1)), "E")
'Next
End Sub
Sub setColors(cRow As Long, strTaskType As String, sBarType As String, Optional s As Shape)
Dim bPCompleted As Boolean: Dim bcrowCSE As Boolean, bcrowCSP As Boolean, bcrowCSB As Boolean, bcrowCSA As Boolean, bOverdue As Boolean
Dim customEColor As Long, customPColor As Long, customBColor As Long, customAColor As Long, percValue As Double: Dim rngStatus As Range
If sBarType = "EPP" Or sBarType = "EP" Then sBarType = "P"
percValue = Cells(cRow, cpg.PercentageCompleted)
If bShowPBar Then
If sBarType = "E" Or sBarType = "P" Then If percValue = 1 Then bPCompleted = True Else bPCompleted = False
End If
If sBarType = "E" Then
If bPCompleted Then
If Filled(Cells(cRow, cpg.TPColor)) Then bcrowCSP = True: customPColor = Cells(cRow, cpg.TPColor).Interior.Color
Else
If Filled(Cells(cRow, cpg.TColor)) Then bcrowCSE = True: customEColor = Cells(cRow, cpg.TColor).Interior.Color
End If
End If
If sBarType = "P" Then If Filled(Cells(cRow, cpg.TPColor)) Then bcrowCSP = True: customPColor = Cells(cRow, cpg.TPColor).Interior.Color
If sBarType = "B" Then If Filled(Cells(cRow, cpg.BLColor)) Then bcrowCSB = True: customBColor = Cells(cRow, cpg.BLColor).Interior.Color
If sBarType = "A" Then If Filled(Cells(cRow, cpg.ACColor)) Then bcrowCSA = True: customAColor = Cells(cRow, cpg.ACColor).Interior.Color

If bShowOverdueBars = True And IsOverdue(cRow) Then bOverdue = True

If sBarType = "Ebase" Then sColor = st.cEBase: GoTo recolorFree
If sBarType = "E" Then
If bOverdue Then sColor = st.cOBC: GoTo Colored
If bShowPBar And bPCompleted Then
If bcrowCSP Then sColor = customPColor Else sColor = st.cPBC
Else
If strTaskType = "T" Then
If bcrowCSE Then sColor = customEColor Else sColor = st.cEBC
Else
If bcrowCSE Then sColor = customEColor Else sColor = st.cEMC
End If
End If
GoTo recolorFree
End If
If sBarType = "P" And bShowPBar Then
If bcrowCSP Then sColor = customPColor Else sColor = st.cPBC: GoTo recolorFree
End If
If sBarType = "B" Then
If bcrowCSB Then sColor = customBColor Else sColor = st.cBBC: GoTo recolorFree
End If
If sBarType = "A" Then
If bcrowCSA Then sColor = customAColor Else sColor = st.cABC: GoTo recolorFree
End If
recolorFree:
If bFreeVersion And cRow > 14 Then
If sBarType = "E" Then
If bShowPBar And bPCompleted Then sColor = rgbGray: GoTo Colored Else sColor = rgbLightGray: GoTo Colored
End If
If sBarType = "P" Then
If bShowPBar Then sColor = rgbGray: GoTo Colored
End If
End If
Colored:
If s Is Nothing Then
Cells(cRow, cpg.TaskIcon).Font.Color = sColor
Else
Call colorShape(s, sColor, cRow)
If sBarType = "E" Then Cells(cRow, cpg.TaskIcon).Font.Color = sColor
End If
Last:
bcrowCSE = False: bcrowCSP = False: bcrowCSB = False: bcrowCSA = False
End Sub

Sub ChangeShapeColor(shapeType As String)
tlog "ChangeShapeColor"
Dim Sh As Shape: Dim lrow As Long, cRow As Long: Dim bOverdue As Boolean, bGanttBar As Boolean: Dim sBarType As String, strGEtype As String
lrow = GetLastRow: bShowPBar = st.ShowPercBar
If st.ShowOverdueBar Then bShowOverdueBars = True Else bShowOverdueBars = False
ReDim vArrAllData(1 To lrow, 1 To cpg.LC): vArrAllData = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 3) = "S_M" Or Left(Sh.Name, 3) = "S_T" Then bGanttBar = True Else bGanttBar = False
If bGanttBar Then
cRow = getTIDRowfromShape(Sh): strGEtype = getGEtype(cRow)
If Left(Sh.Name, 5) <> "S_TGB" Then sBarType = "TGB" ' check this code
If Left(Sh.Name, 4) = "S_TE" Or Left(Sh.Name, 4) = "S_ME" Then sBarType = "E"
If Left(Sh.Name, 4) = "S_TP" Or Left(Sh.Name, 4) = "S_MP" Then sBarType = "P"
If Left(Sh.Name, 4) = "S_TB" Or Left(Sh.Name, 4) = "S_MB" Then sBarType = "B"
If Left(Sh.Name, 4) = "S_TA" Or Left(Sh.Name, 4) = "S_MA" Then sBarType = "A"
If Left(Sh.Name, 4) = "S_TH" Then sBarType = "Ebase"
End If
If shapeType = "Ebase" Or shapeType = "all" Then
If Left(Sh.Name, 4) = "S_TH" Then Call setColors(cRow, strGEtype, sBarType, Sh)
End If
If shapeType = "M" Or shapeType = "all" Then
If Left(Sh.Name, 4) = "S_ME" Or Left(Sh.Name, 4) = "S_MP" Then Call setColors(cRow, strGEtype, sBarType, Sh)
End If
If shapeType = "TE" Or shapeType = "all" Then
If Left(Sh.Name, 4) = "S_TE" Or Left(Sh.Name, 4) = "S_TP" Then Call setColors(cRow, strGEtype, sBarType, Sh)
End If
If shapeType = "TB" Or shapeType = "all" Then
If Left(Sh.Name, 4) = "S_TB" Or Left(Sh.Name, 4) = "S_MB" Then Call setColors(cRow, strGEtype, sBarType, Sh)
End If
If shapeType = "TA" Or shapeType = "all" Then
If Left(Sh.Name, 4) = "S_TA" Or Left(Sh.Name, 4) = "S_MA" Then Call setColors(cRow, strGEtype, sBarType, Sh)
End If

If shapeType = "TLC" Or shapeType = "all" Then If Left(Sh.Name, 6) = "ST_Tod" Then Sh.Line.ForeColor.RGB = st.cTLC: GoTo nextshape
If shapeType = "TGB" Or shapeType = "all" Then If Left(Sh.Name, 5) = "S_TGB" Then Sh.TextFrame.Characters.Font.Color = st.cTGB: GoTo nextshape
If shapeType = "DLC" Or shapeType = "all" Then If Left(Sh.Name, 4) = "S_De" Then Sh.Line.ForeColor.RGB = st.cDLC: GoTo nextshape
If shapeType = "Mo" Or shapeType = "all" Then
If Left(Sh.Name, 5) = "ST_Mo" Then Sh.TextFrame.Characters.Font.Color = st.cPRC: Sh.Line.ForeColor.RGB = st.cPRC: GoTo nextshape
If Left(Sh.Name, 10) = "SG_Project" Then Sh.TextFrame.Characters.Font.Color = st.cPRC: Sh.Line.ForeColor.RGB = st.cPRC: GoTo nextshape
End If
nextshape:
bOverdue = False
Next Sh
If shapeType = "Mo" Or shapeType = "all" Then ActiveSheet.Cells(rowsix, cpg.LC).Font.Color = st.cPRC
Last:
ReDim vArrAllData(1, 1)
tlog "ChangeShapeColor"
End Sub

Sub DeleteShape(shName As String, leng As Long, Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
Dim ShAs Shape
For Each Sh In ws.Shapes
If Left(Sh.Name, leng) = shName Then Sh.Delete
Next Sh
End Sub
Sub DeleteAllTimelineShapes()
Dim ShAs Shape
For Each Sh In ActiveSheet.Shapes
If Left(Sh.Name, 2) = "S_" Or Left(Sh.Name, 2) = "ST" Then Sh.Delete
Next Sh
End Sub
Sub DeleteAllGanttShapes()
tlog "DeleteAllGanttShapes"
Dim Sh As Shape
ActiveSheet.Rectangles.Delete
For Each Sh In ActiveSheet.Shapes
If Sh.AutoShapeType = 37 Or Sh.AutoShapeType = -2 Or Sh.AutoShapeType = 63 Then Sh.Delete
Next Sh
tlog "DeleteAllGanttShapes"
End Sub
Sub LoadCalDateValuesArray()
ReDim vArrDatesValues(1 To 1, 1 To cpt.TimelineEnd)
vArrDatesValues = Range(Cells(rowsix, 1), Cells(rowsix, cpt.TimelineEnd)).value
End Sub
Function GetTimelineDateColumnNo(sRng As Range, sdate As Date, colNoIfNotFound As Long) As Long ' Dont move this to functions
Dim i As Integer:GetTimelineDateColumnNo = colNoIfNotFound
For i = cpt.TimelineStart To cpt.TimelineEnd
If sdate = vArrDatesValues(1, i) Then GetTimelineDateColumnNo = i: Exit For
Next
End Function

Sub UpdateShapeColumns()
If st.CurrentView = "H" Then Exit Sub ' works for Daily view only
If Cells(firsttaskrow, cpg.Task) = sAddTaskPlaceHolder Then Exit Sub
tlog "UpdateShapeColumns"
Dim arrAllV(): Dim arrShapeInfoE(): Dim arrShapeInfoB(): Dim arrShapeInfoA(): Dim lrow As Long, i As Long, cRow As Long: Dim s As Shape
lrow = GetLastRow: ReDim arrAllV(1 To lrow, 1 To cpg.LC): arrAllV = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
ReDim arrShapeInfoE(firsttaskrow To lrow, 1 To 1): ReDim arrShapeInfoB(firsttaskrow To lrow, 1 To 1): ReDim arrShapeInfoA(firsttaskrow To lrow, 1 To 1)
Dim sLeftc As Double, sTopc As Double, sRightc As Double, sDownc As Double: Dim sType As String, strShapeInfo As String: Dim Sh As Shape:
For Each s In ActiveSheet.Shapes
If Left(s.Name, 5) = "S_TGB" Then GoTo nexs
If Left(s.Name, 3) = "S_T" Or Left(s.Name, 3) = "S_M" Then
sType = Left(s.Name, 4):
sLeftc = WorksheetFunction.RoundUp(s.Left, 1): sTopc = WorksheetFunction.RoundUp(s.Top, 1)
sRightc = WorksheetFunction.RoundUp((s.Left + s.Width), 1): sDownc = WorksheetFunction.RoundUp((s.Top + s.Height), 1)
cRow = getTIDRowfromShape(s)
strShapeInfo = sType & DepSeperator & getTIDfromShape(s) & DepSeperator & sLeftc & DepSeperator & sTopc & DepSeperator & sRightc & DepSeperator & sDownc
Select Case sType
Case Is = "S_TE"
arrShapeInfoE(cRow, 1) = strShapeInfo
Case Is = "S_TB"
arrShapeInfoB(cRow, 1) = strShapeInfo
Case Is = "S_TA"
arrShapeInfoA(cRow, 1) = strShapeInfo
Case Is = "S_ME"
arrShapeInfoE(cRow, 1) = strShapeInfo
Case Is = "S_MB"
arrShapeInfoB(cRow, 1) = strShapeInfo
Case Is = "S_MA"
arrShapeInfoA(cRow, 1) = strShapeInfo
End Select
End If
nexs:
Next s
If ActiveSheet.AutoFilterMode Then
Call ReApplyAutoFilter(arrShapeInfoE, lrow, "ShapeInfoEst"):
Call ReApplyAutoFilter(arrShapeInfoB, lrow, "ShapeInfoBas")
Call ReApplyAutoFilter(arrShapeInfoA, lrow, "ShapeInfoAct")
Else
Call ArrayToRange(arrShapeInfoE, lrow, "ShapeInfoEst")
Call ArrayToRange(arrShapeInfoB, lrow, "ShapeInfoBas")
Call ArrayToRange(arrShapeInfoA, lrow, "ShapeInfoAct")
End If
tlog "UpdateShapeColumns"
End Sub
Option Explicit
Option Private Module
Private bCreateTimeline As Boolean
Private vArrAllData()

Sub CreateTimeline()
Call ReadSettings: Set gs = setGSws
Dim cols As Long, hourlycol As Long, noofdays As Long: Dim sy As Integer, ey As Integer: Dim cd, NewTSD As Date, NewTED As Date, newTSDw As Date, newTEDw As Date
Dim newTSDm As Date, newTEDm As Date, newTSDq As Date, newTEDq As Date, newTSDhy As Date, newTEDhy As Date, newTSDy As Date, newTEDy As Date
NewTSD = st.TSD: NewTED = st.TED: newTSDw = GetFirstDateOfWeek(st.TSD): newTEDw = GetFirstDateOfWeek(st.TED) + 6:
newTSDm = DateSerial(Year(NewTSD), Month(NewTSD), 1): newTEDm = DateSerial(Year(NewTED), Month(NewTED) + 1, 0): newTSDq = GetFirstDateOfQuarter(NewTSD):
newTEDq = GetFirstDateOfQuarter(NewTED) + 90: newTSDhy = GetFirstDateOfHalfYearly(NewTSD): newTEDhy = GetFirstDateOfHalfYearly(NewTED) + 180
newTSDy = GetFirstDateOfYear(NewTSD): newTEDy = GetFirstDateOfYear(NewTED):
cols = (st.TED - st.TSD + 1) * 24:
If cols <= 1 Then cols = 1 'h
If st.HideNonWorkingHours = False Then If cols > maxHCol Then cols = maxHCol
gs.Cells(rowtwo, cps.HCOL).value = cols

cols = NewTED - NewTSD: 'D
If cols <= 1 Then cols = 1
If cols > maxDCol Then cols = maxDCol
gs.Cells(rowtwo, cps.dCol).value = cols

cols = Application.WorksheetFunction.Round((newTEDw - newTSDw) / 7, 0) - 1: 'W
If cols <= 1 Then cols = 1
gs.Cells(rowtwo, cps.WCOL).value = cols

cols = Application.WorksheetFunction.Round(((newTEDm - newTSDm) / 30), 0) - 1: 'M
If cols <= 1 Then cols = 1
gs.Cells(rowtwo, cps.MCOL).value = cols

cols = (((newTEDq - newTSDq) / 30) / 3) - 1: 'Q
If cols <= 1 Then cols = 1
gs.Cells(rowtwo, cps.QCOL).value = cols

sy = Year(gs.Cells(rowtwo, cps.TSD).value): ey = Year(gs.Cells(rowtwo, cps.TED).value) 'HY
cols = (((newTEDhy - newTSDhy) / 30) / 6) - 1:
If cols <= 1 Then cols = 1
gs.Cells(rowtwo, cps.HYCOL).value = cols
cols = (ey - sy):
If cols <= 1 Then cols = 1 'Y
gs.Cells(rowtwo, cps.YCOL).value = cols
Call ReadSettings
Select Case st.CurrentView
Case "HH"
Call BuildView("HH")
Case "D"
Call BuildView("D")
Case "W"
Call BuildView("W")
Case "M"
Call BuildView("M")
Case "Q"
Call BuildView("Q")
Case "HY"
Call BuildView("HY")
Case "Y"
Call BuildView("Y")
End Select
End Sub

Sub BuildView(View As String)
Dim tendcol As Long: Set gs = setGSws
If FreeVersion Then If View = "Q" Or View = "HY" Or View = "Y" Then ShowLimitation: Exit Sub
If st.ShowRefreshTimeline = False Then Call TimelineMsg: Exit Sub
Call ClearOldTimeline
If View = "HH" Then
gs.Cells(rowtwo, cps.CurrentView) = "HH": tendcol = st.HCOL + cpt.TimelineStart - 1
ElseIf View = "D" Then
gs.Cells(rowtwo, cps.CurrentView) = "D": tendcol = st.dCol + cpt.TimelineStart:
ElseIf View = "W" Then
gs.Cells(rowtwo, cps.CurrentView) = "W": tendcol = st.WCOL + cpt.TimelineStart:
ElseIf View = "M" Then
gs.Cells(rowtwo, cps.CurrentView) = "M":tendcol = st.MCOL + cpt.TimelineStart:
ElseIf View = "Q" Then
gs.Cells(rowtwo, cps.CurrentView) = "Q":tendcol = st.QCOL + cpt.TimelineStart:
ElseIf View = "HY" Then
gs.Cells(rowtwo, cps.CurrentView) = "HY":tendcol = st.HYCOL + cpt.TimelineStart:
ElseIf View = "Y" Then
gs.Cells(rowtwo, cps.CurrentView) = "Y":tendcol = st.YCOL + cpt.TimelineStart:
End If
Call ReadSettings: Cells(1, cpt.TimelineEnd).value = vbNullString: Cells(1, cpt.LLC).value = vbNullString:
If Cells(1, tendcol).value = "TimelineStart" Then
Cells(1, tendcol + 1).value = "TimelineEnd": Cells(1, tendcol + 2).value = "LLC"
Else
Cells(1, tendcol).value = "TimelineEnd": Cells(1, tendcol + 1).value = "LLC"
End If
Call CalcColPosTimeline: Columns(cpt.LLC).ColumnWidth = 2
If View = "HH" Then
Call BuildHourlyView
ElseIf View = "D" Then
Call BuildDailyView
ElseIf View = "W" Then
Call BuildWeeklyView
ElseIf View = "M" Then
Call BuildMonthlyView
ElseIf View = "Q" Then
Call BuildQuarterlyView
ElseIf View = "HY" Then
Call BuildHalfYearlyView
ElseIf View = "Y" Then
Call BuildYearlyView
End If
Call DrawTNavIcons: Call DrawFilterIcon: Call DrawFreezeIcon
Call ColorTimeline: Call HiliteHolidays: Call HiliteWorkOffDays: Call HiliteHNWDperrow:
Call DelnDrawAllGanttBars: Call FormatTimeline
End Sub

Sub BuildHourlyView()
Dim icol As Long, sPos As Long, wkdaynum As Long: Dim i As Double:
Dim cd As Date, sd As Date, td As Date, startDate As Date, timestart As Date, thour As Date: Dim dName As String, timedisplay As String, resname As String
Dim resStartHour As Integer, daystartHour As Integer, resEndHour As Integer, dayendHour As Integer:
Call RememberResArrays: ResArraysReady = True
Call getResValue(resname, sArr.ResourceP): resStartHour = Format(sArr.ResourceP(resvalue, 10), "h"): resEndHour = Format(sArr.ResourceP(resvalue, 11), "h")
cd = st.TSD: startDate = Int(CDbl(cd)): Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).NumberFormat = "@"
i = 0
For icol = cpt.TimelineStart To cpt.TimelineEnd
Cells(rowsix, icol) = startDate + i: i = i + (1 / 24)
Next
timestart = TimeValue("00:00:00"): thour = TimeValue("01:00:00")
i = 0
For icol = cpt.TimelineStart To cpt.TimelineEnd
Cells(roweight, icol) = i: i = i + 1:
If i = 24 Then i = 0
timedisplay = Format(timestart, "hh AM/PM"):Cells(rownine, icol) = Mid(timedisplay, 1, 4):timestart = timestart + thour
Next
For icol = cpt.TimelineStart To cpt.TimelineEnd
If Cells(roweight, icol) = resStartHour Then Cells(rowseven, icol) = "'" & Format(Cells(rowsix, icol), "ddd - dd-mmm-yy")
Next
Range(Cells(roweight, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.HWID
End Sub
Sub BuildDailyView()
Set gs = setGSws
Dim icol As Long, noofcolordays As Long: Dim sAs Shape
Dim cd As Date, sd As Date: Dim dName As String: Dim leftfromAs Double, widshp As Double, topfrom As Double, hgtshp As Double
cd = st.TSD: sd = cd:
For icol = cpt.TimelineStart To cpt.TimelineEnd
Cells(rowsix, icol) = sd + (icol - cpt.TimelineStart)
Cells(rownine, icol) = Left(Format(sd + (icol - cpt.TimelineStart), "ddd"), 1)
Cells(roweight, icol) = Day(sd + (icol - cpt.TimelineStart))
If Application.WorksheetFunction.Weekday(Cells(rowsix, icol), 2) = gs.Cells(rowtwo, cps.WeekStartDay) Then
Cells(rowseven, icol) = "W" & WeekNumVBA(sd + icol - cpt.TimelineStart, gs.Cells(rowtwo, cps.WeekNumType))
End If
Next
For icol = cpt.TimelineStart To cpt.TimelineEnd
noofcolordays = NoOfDaysInMonth(Cells(rowsix, icol)) - Cells(roweight, icol)
If Month(Cells(rowsix, icol)) = Month(Cells(rowsix, cpt.TimelineEnd)) And Year(Cells(rowsix, icol)) = Year(Cells(rowsix, cpt.TimelineEnd)) Then
noofcolordays = CDate(Cells(rowsix, cpt.TimelineEnd)) - CDate(Cells(rowsix, icol))
End If
With Range(Cells(rowsix, icol), Cells(rowsix, icol + noofcolordays))
leftfrom = .Left:topfrom = .Top + 4:widshp = .Width - 0.5:hgtshp = .Height - 4
End With
If widshp <= 0 Then Exit Sub
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRound2SameRectangle, Left:=leftfrom, Top:=topfrom, Width:=widshp, Height:=hgtshp)
With s
.Name = "ST_Mo_" & Cells(rowsix, icol): .Line.visible = msoTrue: .Line.ForeColor.RGB = st.cPRC
With .TextFrame.Characters
If widshp < 300 Then .Text = Format(Cells(rowsix, icol), "mmm-yy") Else .Text = Format(Cells(rowsix, icol), "mmmm-yyyy")
.Font.Color = st.cPRC: .Font.size = 14
End With
.TextFrame.VerticalAlignment = xlVAlignCenter:.TextFrame.HorizontalAlignment = xlHAlignLeft
End With
With s.Fill
.Solid: .ForeColor.RGB = vbWhite
End With
icol = icol + noofcolordays
Next
Range(Cells(roweight, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.DWID
End Sub

Sub BuildWeeklyView()
Dim icol As Long: Dim cd As Date, sd As Date, td As Date
Set gs = setGSws
cd = st.TSD: icol = cpt.TimelineStart
If cd = 0 Then cd = Date
sd = (cd - (Weekday(cd, vbMonday) - 1)) + (7 * (1 - 1))
Do Until icol = cpt.TimelineEnd + 1
td = sd + ((icol - cpt.TimelineStart) * 7)
Cells(rowsix, icol) = td
Cells(roweight, icol) = WeekNumVBA(td, gs.Cells(rowtwo, cps.WeekNumType))
If icol = cpt.TimelineStart Then
Cells(rowseven, icol) = " " & Format(td, "MMM-YY")
ElseIf Month(Cells(rowsix, icol)) <> Month(Cells(rowsix, icol - 1)) Then
Cells(rowseven, icol) = " " & Format(td, "MMM-YY")
End If
icol = icol + 1
Loop
Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.WWID
End Sub
Sub BuildMonthlyView()
Dim icol As Long: Dim cd As Date, sd As Date, td As Date
sd = DateSerial(Year(st.TSD), Month(st.TSD), 1): icol = cpt.TimelineStart: cd = st.TSD
If cd = 0 Then cd = Date
sd = DateAdd("m", 1 - 1, cd)
Do Until icol = cpt.TimelineEnd + 1
td = DateAdd("m", icol - cpt.TimelineStart, sd)
Cells(rowsix, icol) = DateSerial(Year(td), Month(td), 1)
If icol = cpt.TimelineStart Then
Cells(rowseven, icol) = "'" & Year(td)
ElseIf Year(Cells(rowsix, icol)) <> Year(Cells(rowsix, icol - 1)) Then
Cells(rowseven, icol) = "'" & Year(td)
End If
Cells(roweight, icol) = Format(td, "MMM"):Cells(rownine, icol) = Month(td):Cells(rownine, icol).NumberFormat = "General"
icol = icol + 1
Loop
Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.MWID
End Sub
Sub BuildQuarterlyView()
Dim icol As Long: Dim cd As Date, sd As Date, td As Date
cd = st.TSD: icol = cpt.TimelineStart:
If cd = 0 Then cd = Date
sd = GetFirstDateOfQuarter(cd): sd = DateAdd("q", (1 - 1), sd)
icol = cpt.TimelineStart
Do Until icol = cpt.TimelineEnd + 1
td = DateAdd("m", (icol - cpt.TimelineStart) * 3, sd)
Cells(rowsix, icol) = td
Cells(roweight, icol) = "Q" & (Month(td) + 2) \ 3
If icol = cpt.TimelineStart Then
Cells(rowseven, icol) = "'" & Year(td)
ElseIf Year(td) <> Year(Cells(rowsix, icol - 1)) Then
Cells(rowseven, icol) = "'" & Year(td)
End If
icol = icol + 1
Loop
Cells(rowsix, cpt.TimelineEnd + 1) = DateAdd("m", 3, sd)
Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.QWID
End Sub
Sub BuildHalfYearlyView()
Dim icol As Long, tcol As Long: Dim cd As Date, sd As Date, td As Date
cd = st.TSD: icol = cpt.TimelineStart: tcol = cpt.TimelineStart
If cd = 0 Then cd = Date
sd = GetFirstDateOfHalfYearly(cd): sd = DateAdd("Q", (1 - 1) * 2, sd)
icol = cpt.TimelineStart
Do Until icol = cpt.TimelineEnd + 1
td = DateAdd("m", (icol - cpt.TimelineStart) * 6, sd)
Cells(rowsix, icol) = td
Cells(roweight, icol) = "H" & (CBool(Month(td) > 6) * -1 + 1)
If icol = cpt.TimelineStart Then
Cells(rowseven, icol) = "'" & Year(td)
ElseIf Year(td) <> Year(Cells(rowsix, icol - 1)) Then
Cells(rowseven, icol) = "'" & Year(td)
End If
icol = icol + 1
Loop
Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.HYWID
End Sub
Sub BuildYearlyView()
Dim icol As Long, lrow As Long: Dim cd As Date, sd As Date, td As Date
lrow = GetLastRow
cd = st.TSD
If cd = 0 Then cd = Date
sd = GetFirstDateOfYear(cd)
sd = DateAdd("YYYY", (1 - 1), sd)
icol = cpt.TimelineStart
Do Until icol = cpt.TimelineEnd + 1
td = DateAdd("YYYY", (icol - cpt.TimelineStart), sd)
Cells(rowsix, icol) = td
Cells(roweight, icol) = "'" & Year(td)
icol = icol + 1
Loop
Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).ColumnWidth = st.YWID
End Sub

Sub FormatTimeline()
tlog "FormatTimeline"
With Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd))
.ClearFormats:.Interior.Color = st.cPRC:.Font.Color = st.cPRC
End With
With Range(Cells(rowseven, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd))
.Font.Bold = False
.Font.Color = vbBlack:.Font.size = 8:.Font.Name = "Arial":.VerticalAlignment = xlCenter:.HorizontalAlignment = xlCenter
End With
If st.CurrentView = "HH" Then
With Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd))
.Font.size = 8:.Font.Name = "Calibri":.Interior.Color = st.cCR3C
End With
End If
Range(Cells(rowseven, cpt.TimelineStart), Cells(rowseven, cpt.TimelineEnd)).HorizontalAlignment = xlLeft
'Range(Cells(rowsix, cpt.TimelineEnd + 1), Cells(1, 1000)).EntireColumn.clear
'Range(Cells(rowsix, cpt.TimelineEnd + 1), Cells(GetLastTaskRowNo, cpt.TimelineEnd + 1000)).Delete

#If Mac Then
#Else
Range(Cells(1, cpt.LLC + 1), Cells(1, Columns.Count)).EntireColumn.ColumnWidth = 0
#End If
Call DrawTimelineBorders: Call DrawTasksBorders
tlog "FormatTimeline"
End Sub

Sub ColorTimeline()
tlog "ColorTimeline"
If st.ShowRefreshTimeline = False Then Exit Sub
Dim icol As Long: Dim bOn As Boolean: icol = cpt.TimelineStart:
Range(Cells(rowsix, cpg.SS), Cells(rowsix, cpt.LLC)).Interior.Color = st.cPRC:
Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.LLC)).Font.Color = st.cPRC
Range(Cells(roweight, cpt.TimelineStart), Cells(roweight, cpt.TimelineEnd)).Interior.Color = st.cCR2C
Range(Cells(rownine, cpg.SS), Cells(rownine, cpg.LC)).Interior.Color = st.cHBC: Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).Interior.Color = st.cCR3C
If st.CurrentView = "HH" Then
With Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd))
.NumberFormat = "@"'.Font.size = 7
End With
bOn = False
For icol = cpt.TimelineStart To cpt.TimelineEnd
If Cells(roweight, icol) = 0 Then
With Range(Cells(rowseven, icol), Cells(rowseven, icol + 23))
If bOn Then .Interior.Color = st.cCR12C Else .Interior.Color = st.cCR1C
bOn = Not (bOn)
End With
icol = icol + 23
End If
Next
GoTo Last:
End If
Dim x As Long, y As Long:bOn = False: y = 1
For icol = cpt.TimelineStart To cpt.TimelineEnd
If Cells(rowseven, icol) <> vbNullString Then
For x = icol + 1 To cpt.TimelineEnd
If Cells(rowseven, x) = vbNullString Then y = y + 1 Else Exit For
Next x
Else
For x = icol + 1 To cpt.TimelineEnd
If Cells(rowseven, x) = vbNullString Then y = y + 1 Else Exit For
Next x
End If
With Range(Cells(rowseven, icol), Cells(rowseven, icol + y - 1))
.HorizontalAlignment = xlLeft
If bOn Then .Interior.Color = st.cCR12C Else .Interior.Color = st.cCR1C
bOn = Not (bOn)
End With
icol = icol + y - 1: y = 1
Next
Last:
tlog "ColorTimeline"
End Sub

Sub ClearOldTimeline()
tlog "ClearOldTimeline"
Call DeleteShape("ST_Mo", 5)
Range(Cells(rowsix, cpt.TimelineStart), Cells(GetLastRow + 2, cpt.TimelineEnd + 10)).Clear
Range(Cells(1, cpt.TimelineStart), Cells(1, Columns.Count)).EntireColumn.ColumnWidth = 10
'Range(Cells(rowone, cpt.TimelineStart), Cells(GetLastRow + 2, cpt.TimelineEnd + 10)).clear
'Cells(rowone, cpg.LC + 1) = "TimelineStart": Cells(rowone, cpg.LC + 3) = "TimelineEnd": Cells(rowone, cpg.LC + 4) = "LLC": Call CalcColPosGCT
'Range(Cells(1, cpt.TimelineStart), Cells(1, Columns.Count)).EntireColumn.ColumnWidth = 10
tlog "ClearOldTimeline"
End Sub

Sub ClearBorders(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
Dim lrow As Long, colGEtype As Long, colTE As Long: colGEtype = getColGEType(ws): colTE = getColTE(ws): lrow = GetLastRow(ws)
ws.Range(ws.Cells(rowsix, colGEtype), ws.Cells(GetLastRow + 100, colTE + 10)).Borders.LineStyle = xlNone
End Sub


Sub DrawTasksBorders(Optional cRow As Long, Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
Dim c As Range: Dim lrow As Long, colSS As Long, colLC As Long, colTE As Long
lrow = GetLastRow(ws) + 2: colSS = getColSS(ws): colLC = getColLC(ws): colTE = getColTE(ws)
If cRow = 0 Then
If st.ShowRefreshTimeline = False Then
Set c = ws.Range(ws.Cells(roweight, colSS), ws.Cells(lrow - 1, colLC))
Else
Set c = ws.Range(ws.Cells(roweight, colSS), ws.Cells(lrow - 1, colTE))
End If
Else
If st.ShowRefreshTimeline = False Then
Set c = ws.Range(ws.Cells(cRow - 1, colSS), ws.Cells(cRow + 1, colLC))
Else
Set c = ws.Range(ws.Cells(cRow - 1, colSS), ws.Cells(cRow + 1, colTE))
End If
End If
With c.Borders(xlInsideHorizontal)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = st.cTBC
End With
If st.VerticalBorders Then
With ws.Range(ws.Cells(roweight, colSS), ws.Cells(lrow - 1, colLC))
With .Borders(xlInsideVertical)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = st.cTBC
End With
End With
ws.Range(ws.Cells(roweight, colSS), ws.Cells(rownine, colLC)).Borders(xlInsideVertical).LineStyle = xlNone
ws.Range(ws.Cells(lrow - 1, colSS), ws.Cells(lrow - 1, colLC)).Borders(xlInsideVertical).LineStyle = xlNone
End If
End Sub
Sub DrawTimelineBorders(Optional cRow As Long, Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
If st.ShowRefreshTimeline = False Or st.ShowTimelineGrid = False Then Exit Sub
Dim c As Range: Dim lrow As Long, colTS As Long, colTE As Long
lrow = GetLastRow(ws) + 2: colTS = getColTS(ws): colTE = getColTE(ws)
If cRow = 0 Then
Set c = ws.Range(ws.Cells(roweight, colTS - 1), ws.Cells(lrow - 2, colTE + 1))
Else
Set c = ws.Range(ws.Cells(cRow, colTS - 1), ws.Cells(cRow, colTE + 1))
End If
With c.Borders(xlInsideVertical)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = st.cGBC
End With
End Sub

Sub HiliteHolidays()
Call FormatHolidays("Hilite")
End Sub
Sub HideHolidays()
Call FormatHolidays("Hide")
End Sub
Sub FormatHolidays(Optional doo As String)
If st.ShowRefreshTimeline = False Then Exit Sub
If st.CurrentView <> "D" And st.CurrentView <> "HH" Then Exit Sub
'If st.HiliteHolidays = False And st.HideHolidays = False Then Exit Sub
Dim cCol As Long, maxnoofholidays As Long, i As Long: Dim c As Range: Dim bFound As Boolean
Call RememberResArrays: ResArraysReady = True: maxnoofholidays = UBound(newHolidaysArray(), 2)
For cCol = cpt.TimelineStart To cpt.TimelineEnd
For i = 1 To maxnoofholidays
If newHolidaysArray(0, i) = CDate(Cells(rowsix, cCol)) Then bFound = True: Exit For
Next
If st.CurrentView = "HH" Then Set c = Range(Cells(rowseven, cCol), Cells(rowseven, cCol + 23)): cCol = cCol + 23 Else Set c = Cells(roweight, cCol)
If bFound = True And doo = "Hilite" Then If st.HiliteHolidays Then c.Interior.Color = st.cHC: bFound = False
If bFound = True And doo = "Hide" Then If st.HideHolidays Then c.Columns.EntireColumn.ColumnWidth = 0: bFound = False
Next cCol
ResArraysReady = False: 'Exit Sub
End Sub
Sub HiliteWorkOffDays()
Call FormatWorkOffDays("Hilite")
End Sub
Sub HideWorkOffDays()
Call FormatWorkOffDays("Hide")
End Sub
Sub FormatWorkOffDays(Optional doo As String)
If st.ShowRefreshTimeline = False Then Exit Sub
If st.CurrentView <> "D" And st.CurrentView <> "HH" Then Exit Sub
'If st.HiliteWorkOffDays = False Or st.HideWorkOffDays Then Exit Sub:
Dim cCol, wkdaynum As Long: Dim c As Range: Dim bColorDayName As Boolean
Call RememberResArrays: ResArraysReady = True:
For cCol = cpt.TimelineStart To cpt.TimelineEnd
wkdaynum = Weekday(Cells(rowsix, cCol).value, vbMonday)
If wkdaynum = 7 Then
If newWorkdaysArray(0, 7) = 0 Then bColorDayName = True
ElseIf wkdaynum = 1 Then
If newWorkdaysArray(0, 1) = 0 Then bColorDayName = True
ElseIf wkdaynum = 2 Then
If newWorkdaysArray(0, 2) = 0 Then bColorDayName = True
ElseIf wkdaynum = 3 Then
If newWorkdaysArray(0, 3) = 0 Then bColorDayName = True
ElseIf wkdaynum = 4 Then
If newWorkdaysArray(0, 4) = 0 Then bColorDayName = True
ElseIf wkdaynum = 5 Then
If newWorkdaysArray(0, 5) = 0 Then bColorDayName = True
ElseIf wkdaynum = 6 Then
If newWorkdaysArray(0, 6) = 0 Then bColorDayName = True
End If
If bColorDayName Then
If st.CurrentView = "HH" Then
Set c = Range(Cells(rowseven, cCol), Cells(rowseven, cCol + 23)): cCol = cCol + 23
Else
Set c = Cells(rownine, cCol)
End If
bColorDayName = False:
If st.HiliteWorkOffDays And doo = "Hilite" Then c.Interior.Color = st.cWC
If st.HideWorkOffDays And doo = "Hide" Then c.Columns.EntireColumn.ColumnWidth = 0
End If
Next cCol
ResArraysReady = False: 'Exit Sub
'last:
'Range(Cells(rownine, cpt.TimelineStart), Cells(rownine, cpt.TimelineEnd)).Interior.Color = st.cCR3C
End Sub
Sub HideNonWorkingColumns()
Call HideHolidays: Call HideWorkOffDays: Call HideNonWorkingHours
End Sub

Sub HideNonWorkingHours()
If st.ShowRefreshTimeline = False Then Exit Sub
If st.HideNonWorkingHours = False Or st.CurrentView <> "HH" Then Exit Sub
Dim resname As String: Dim resStartHour As Integer, resEndHour As Integer: Dim cCol As Long
Call RememberResArrays: ResArraysReady = True
If st.CurrentView = "HH" Then
Call getResValue(resname, sArr.ResourceP): resStartHour = Format(sArr.ResourceP(resvalue, 10), "h"): resEndHour = Format(sArr.ResourceP(resvalue, 11), "h"):
If resEndHour = 0 Then resEndHour = 24
For cCol = cpt.TimelineStart To cpt.TimelineEnd
If Cells(roweight, cCol) < resStartHour Or Cells(roweight, cCol) >= resEndHour Then Columns(cCol).EntireColumn.ColumnWidth = 0
Next cCol
End If
ResArraysReady = False
End Sub

Sub HiliteHNWDperrow(Optional cRowOnly As Long, Optional familytype As String)
tlog "HiliteHNWDperrow"
If st.ShowRefreshTimeline = False Then Exit Sub
Dim cRow As Long, wkdaynum As Long, lrow As Long, cCol As Long, n As Long, maxnoofholidays As Long, i As Long, StartRow As Long, EndRow As Long, findcol As Long
Dim c As Range: Dim bFound As Boolean, bColorDayName As Boolean: Dim resname As String: Dim rngTimeline As Range
If st.CurrentView <> "D" And st.CurrentView <> "HH" Then Exit Sub
lrow = GetLastRow:
If st.HiliteHolidaysPR = False And st.HiliteWorkOffDaysPR = False Then Range(Cells(firsttaskrow, cpt.TimelineStart), Cells(lrow, cpt.TimelineEnd)).Interior.Color = xlNone: GoTo Last
ReDim vArrAllData(1 To lrow, 1 To cpt.TimelineEnd): vArrAllData = Range(Cells(1, 1), Cells(lrow, cpt.TimelineEnd)).value
newHolidaysArray = sArr.HolidaysP: newResourcesArray = sArr.ResourceP: newWorkdaysArray = sArr.WorkdaysP
maxnoofholidays = UBound(newHolidaysArray(), 2)
If familytype = "" Then familytype = allRows
Set rngTimeline = Range(Cells(rowsix, cpt.TimelineStart), Cells(rowsix, cpt.TimelineEnd)): Call LoadCalDateValuesArray
Range(Cells(firsttaskrow, cpt.TimelineStart), Cells(lrow, cpt.TimelineEnd)).Interior.Color = xlNone
StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
For cRow = StartRow To EndRow
resname = LCase(vArrAllData(cRow, cpg.Resource))
If resname = vbNullString Or resname = "organization" Or InStr(1, resname, ",") > 0 Then GoTo nexcrow
If IsParentTask(cRow) Then GoTo nexcrow
Call getResValue(resname, newResourcesArray)
If st.HiliteHolidaysPR = False Then GoTo startweekendoff
If UBound(newHolidaysArray(), 2) = 0 Then GoTo startweekendoff
If newHolidaysArray(resvalue, 1) = vbNullString Then GoTo startweekendoff
For i = 1 To maxnoofholidays
If newHolidaysArray(resvalue, i) = vbNullString Then GoTo startweekendoff
If newHolidaysArray(resvalue, i) < CDate(vArrAllData(rowsix, cpt.TimelineStart)) Or newHolidaysArray(resvalue, i) > CDate(vArrAllData(rowsix, cpt.TimelineEnd)) Then GoTo checknexthol
findcol = GetTimelineDateColumnNo(rngTimeline, CDate(newHolidaysArray(resvalue, i)), 0)
If st.CurrentView = "D" Then Set c = Cells(cRow, findcol) Else Set c = Range(Cells(cRow, findcol), Cells(cRow, findcol + 23))
c.Interior.Color = st.cHCPR: findcol = 0
GoTo checknexthol
checknexthol:
Next i
startweekendoff:
If st.HiliteWorkOffDaysPR = False Then GoTo nexcrow
For cCol = cpt.TimelineStart To cpt.TimelineEnd
wkdaynum = Weekday(vArrAllData(rowsix, cCol), vbMonday)
If wkdaynum = 7 Then
If newWorkdaysArray(resvalue, 7) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 1 Then
If newWorkdaysArray(resvalue, 1) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 2 Then
If newWorkdaysArray(resvalue, 2) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 3 Then
If newWorkdaysArray(resvalue, 3) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 4 Then
If newWorkdaysArray(resvalue, 4) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 5 Then
If newWorkdaysArray(resvalue, 5) = 0 Then bColorDayName = True: GoTo colthisday
ElseIf wkdaynum = 6 Then
If newWorkdaysArray(resvalue, 6) = 0 Then bColorDayName = True: GoTo colthisday
End If
colthisday:
Set c = Cells(cRow, cCol)
If bColorDayName Then
If st.CurrentView = "D" Then Set c = Cells(cRow, cCol) Else Set c = Range(Cells(cRow, cCol), Cells(cRow, cCol + 23)): cCol = cCol + 23
c.Interior.Color = st.cWCPR: bColorDayName = False
End If
Next cCol
nexcrow:
Next cRow
Last:
ReDim vArrAllData(1, 1)
tlog "HiliteHNWDperrow"
End Sub

Sub ScrollTimelineBackS()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Call DA: Set gs = setGSws
Dim rTSD As Range, rTED As Range: Set rTSD = gs.Cells(rowtwo, cps.TSD): Set rTED = gs.Cells(rowtwo, cps.TED)
Dim dTSD As Date, dTED As Date, newDate As Date, dRTSD As Date: dTSD = rTSD.value: dTED = rTED.value: Dim wkday As Long: Dim resname As String
Dim cView As String: Set gs = setGSws: cView = st.CurrentView: dRTSD = CDate(Int(Cells(rowsix, cpt.TimelineStart))): newDate = dRTSD - 1
If cView = "HH" Then
If st.HideHolidays Then
dTSD = GetPreWDAfHol(Org, newDate, 1)
Else
dTSD = newDate
End If
If st.HideWorkOffDays Then
dTSD = GetPreWDAfWO(Org, dTSD, 1)
Else
dTSD = dTSD
End If
ElseIf cView = "D" Then
dTSD = GetFirstDateOfWeek(dTSD - 7): dTSD = GetFirstDateOfWeek(dTSD)
ElseIf cView = "W" Then
dTSD = GetFirstDateOfWeek(dTSD - 7): dTSD = GetFirstDateOfWeek(dTSD)
ElseIf cView = "M" Then
dTSD = GetFirstDateOfMonth(dTSD) - 1: dTSD = GetFirstDateOfMonth(dTSD)
ElseIf cView = "Q" Then
dTSD = GetFirstDateOfQuarter(dTSD) - 1: dTSD = GetFirstDateOfQuarter(dTSD)
ElseIf cView = "HY" Then
dTSD = GetFirstDateOfHalfYearly(dTSD) - 1: dTSD = GetFirstDateOfHalfYearly(dTSD)
ElseIf cView = "Y" Then
dTSD = GetFirstDateOfYear(dTSD) - 1: dTSD = GetFirstDateOfYear(dTSD)
End If
rTSD.value = dTSD: rTED.value = dTED:
Call CreateTimeline: Call EA
End Sub

Sub ScrollTimelineFrontS()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Call DA: Set gs = setGSws
Dim rTSD As Range, rTED As Range: Set rTSD = gs.Cells(rowtwo, cps.TSD): Set rTED = gs.Cells(rowtwo, cps.TED)
Dim dTSD As Date, dTED As Date, newDate As Date, dRTSD As Date: dTSD = rTSD.value: dTED = rTED.value: Dim wkday As Long: Dim resname As String
Dim cView As String: Set gs = setGSws: cView = st.CurrentView: dRTSD = CDate(Int(Cells(rowsix, cpt.TimelineStart))): newDate = dRTSD + 1
If cView = "HH" Then
If st.HideHolidays Then
dTSD = GetNxtWDAfHol(Org, newDate, 1)
Else
dTSD = newDate
End If
If st.HideWorkOffDays Then
dTSD = GetNxtWDAfWO(Org, dTSD, 1)
Else
dTSD = dTSD
End If
If dTED > dTSD + 1 Then 'min 2
Else
dTED = dTSD + 1:
If st.HideHolidays Then
dTED = GetNxtWDAfHol(Org, dTED, 1)
Else
dTED = dTED
End If
If st.HideWorkOffDays Then
dTED = GetNxtWDAfWO(Org, dTED, 1)
Else
dTED = dTED
End If
End If
ElseIf cView = "D" Then
dTSD = GetLastDateOfWeek(dTSD) + 1
If GetFirstDateOfWeek(dTED) <= GetFirstDateOfWeek(dTSD) Then
dTED = GetLastDateOfWeek(dTSD) + 1: dTED = GetLastDateOfWeek(dTED)
End If
ElseIf cView = "W" Then
dTSD = GetLastDateOfWeek(dTSD) + 1
If GetFirstDateOfWeek(dTED) <= GetFirstDateOfWeek(dTSD) Then
dTED = GetLastDateOfWeek(dTSD) + 1: dTED = GetLastDateOfWeek(dTED)
End If
ElseIf cView = "M" Then
dTSD = GetLastDateOfMonth(dTSD) + 1
If GetFirstDateOfMonth(dTED) <= GetFirstDateOfMonth(dTSD) Then
dTED = GetLastDateOfMonth(dTSD) + 1: dTED = GetLastDateOfMonth(dTED)
End If
ElseIf cView = "Q" Then
dTSD = GetLastDateOfQuarter(dTSD) + 1
If GetFirstDateOfQuarter(dTED) <= GetFirstDateOfQuarter(dTSD) Then
dTED = GetLastDateOfQuarter(dTSD) + 1: dTED = GetLastDateOfQuarter(dTED)
End If
ElseIf cView = "HY" Then
dTSD = GetLastDateOfHalfYearly(dTSD) + 1
If GetFirstDateOfHalfYearly(dTED) <= GetFirstDateOfHalfYearly(dTSD) Then
dTED = GetLastDateOfHalfYearly(dTSD) + 1: dTED = GetLastDateOfHalfYearly(dTED)
End If
ElseIf cView = "Y" Then
dTSD = GetLastDateOfYear(dTSD) + 1
If GetFirstDateOfYear(dTED) <= GetFirstDateOfYear(dTSD) Then
dTED = GetLastDateOfYear(dTSD) + 1: dTED = GetLastDateOfYear(dTED)
End If
End If
rTSD.value = dTSD: rTED.value = dTED:
Call CreateTimeline: Call EA
End Sub


Sub ScrollTimelineBackE()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Call DA: Set gs = setGSws
Dim rTSD As Range, rTED As Range: Set rTSD = gs.Cells(rowtwo, cps.TSD): Set rTED = gs.Cells(rowtwo, cps.TED)
Dim dTSD As Date, dTED As Date, newDate As Date, dRTED As Date: dTSD = rTSD.value: dTED = rTED.value: Dim wkday As Long: Dim resname As String
Dim cView As String: Set gs = setGSws: cView = st.CurrentView: dRTED = CDate(Int(Cells(rowsix, cpt.TimelineEnd))): newDate = dRTED - 1
If cView = "HH" Then
If st.HideHolidays Then
dTED = GetPreWDAfHol(Org, newDate, 1)
Else
dTED = newDate
End If
If st.HideWorkOffDays Then
dTED = GetPreWDAfWO(Org, dTED, 1)
Else
dTED = dTED
End If
If dTED > dTSD + 1 Then 'min 2
Else
dTSD = dTED - 1:
If st.HideHolidays Then
dTSD = GetPreWDAfHol(Org, dTSD, 1)
Else
dTSD = dTSD
End If
If st.HideWorkOffDays Then
dTSD = GetPreWDAfWO(Org, dTSD, 1)
Else
dTSD = dTSD
End If
End If
ElseIf cView = "D" Then
dTED = GetFirstDateOfWeek(dTED) - 1
If dTSD >= dTED - 7 Then
dTSD = GetFirstDateOfWeek(dTED - 7): dTSD = GetFirstDateOfWeek(dTSD):
End If
ElseIf cView = "W" Then
dTED = GetFirstDateOfWeek(dTED) - 1
If dTSD >= dTED - 7 Then
dTSD = GetFirstDateOfWeek(dTED - 7): dTSD = GetFirstDateOfWeek(dTSD):
End If
ElseIf cView = "M" Then
dTED = GetFirstDateOfMonth(dTED) - 1
If GetFirstDateOfMonth(dTSD) >= GetFirstDateOfMonth(dTED) Then
dTSD = GetFirstDateOfMonth(dTED) - 1: dTSD = GetFirstDateOfMonth(dTSD)
End If
ElseIf cView = "Q" Then
dTED = GetFirstDateOfQuarter(dTED) - 1:
If GetFirstDateOfQuarter(dTSD) >= GetFirstDateOfQuarter(dTED) Then
dTSD = GetFirstDateOfQuarter(dTED) - 1: dTSD = GetFirstDateOfQuarter(dTSD)
End If
ElseIf cView = "HY" Then
dTED = GetFirstDateOfHalfYearly(dTED) - 1:
If GetFirstDateOfHalfYearly(dTSD) >= GetFirstDateOfHalfYearly(dTED) Then
dTSD = GetFirstDateOfHalfYearly(dTED) - 1: dTSD = GetFirstDateOfHalfYearly(dTSD)
End If
ElseIf cView = "Y" Then
dTED = GetFirstDateOfYear(dTED) - 1
If GetFirstDateOfYear(dTSD) >= GetFirstDateOfYear(dTED) Then
dTSD = GetFirstDateOfYear(dTED) - 1: dTSD = GetFirstDateOfYear(dTSD)
End If
End If
rTSD.value = dTSD: rTED.value = dTED:
Call CreateTimeline: Call EA
End Sub

Sub ScrollTimelineFrontE()
tlog "ScrollTimelineFrontE"
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Call DA: Set gs = setGSws
Dim rTSD As Range, rTED As Range: Set rTSD = gs.Cells(rowtwo, cps.TSD): Set rTED = gs.Cells(rowtwo, cps.TED)
Dim dTSD As Date, dTED As Date, newDate As Date, dRTED As Date: dTSD = rTSD.value: dTED = rTED.value: Dim wkday As Long: Dim resname As String
Dim cView As String: Set gs = setGSws: cView = st.CurrentView: dRTED = CDate(Int(Cells(rowsix, cpt.TimelineEnd))): newDate = dRTED + 1
If cView = "HH" Then
If st.HideHolidays Then
dTED = GetNxtWDAfHol(Org, newDate, 1)
Else
dTED = newDate
End If
If st.HideWorkOffDays Then
dTED = GetNxtWDAfWO(Org, dTED, 1)
Else
dTED = dTED
End If
ElseIf cView = "D" Then
dTED = GetLastDateOfWeek(dTED) + 1: dTED = GetLastDateOfWeek(dTED)
ElseIf cView = "W" Then
dTED = GetLastDateOfWeek(dTED) + 1: dTED = GetLastDateOfWeek(dTED)
ElseIf cView = "M" Then
dTED = GetLastDateOfMonth(dTED) + 1: dTED = GetLastDateOfMonth(dTED)
ElseIf cView = "Q" Then
dTED = GetLastDateOfQuarter(dTED) + 1: dTED = GetLastDateOfQuarter(dTED)
ElseIf cView = "HY" Then
dTED = GetLastDateOfHalfYearly(dTED) + 1: dTED = GetLastDateOfHalfYearly(dTED)
ElseIf cView = "Y" Then
dTED = GetLastDateOfYear(dTED) + 1: dTED = GetLastDateOfYear(dTED)
End If
rTSD.value = dTSD: rTED.value = dTED:
Call CreateTimeline:
tlog "ScrollTimelineFrontE"
Call EA

End Sub

Sub sts()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Set gs = setGSws
Dim noofdays As Long: noofdays = st.TED - st.TSD
If st.CurrentView = "HH" Then
gs.Cells(rowtwo, cps.TSD).value = Int(WorksheetFunction.Min(Columns(cpg.ESD)))
Else
gs.Cells(rowtwo, cps.TSD).value = Int(GetFirstDateOfWeek(WorksheetFunction.Min(Columns(cpg.ESD)))) 'GetWeekStartDate
End If
gs.Cells(rowtwo, cps.TED).value = gs.Cells(rowtwo, cps.TSD).value + noofdays: Call CreateTimeline
End Sub
Sub stt()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Set gs = setGSws
Dim noofdays As Long: noofdays = st.TED - st.TSD
If st.CurrentView = "HH" Then
gs.Cells(rowtwo, cps.TSD).value = Date
Else
gs.Cells(rowtwo, cps.TSD).value = GetFirstDateOfWeek(Date) ' 'GetWeekStartDate
End If
gs.Cells(rowtwo, cps.TED).value = gs.Cells(rowtwo, cps.TSD).value + noofdays: Call CreateTimeline
End Sub
Sub ste()
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
Set gs = setGSws
Dim noofdays As Long: noofdays = st.TED - st.TSD
If gs.Cells(rowtwo, cps.CurrentView) = "HH" Then
gs.Cells(rowtwo, cps.TSD).value = WorksheetFunction.Max(Columns(cpg.EED))
Else
gs.Cells(rowtwo, cps.TSD).value = GetFirstDateOfWeek(WorksheetFunction.Max(Columns(cpg.EED))) 'GetWeekStartDate
End If
gs.Cells(rowtwo, cps.TED).value = gs.Cells(rowtwo, cps.TSD).value + noofdays: Call CreateTimeline
End Sub

Sub turnOffTimeline()
If Not GanttChart Then Exit Sub
Call DA
Set gs = setGSws
gs.Cells(rowtwo, cps.TimelineVisible) = False: gs.Cells(rowtwo, cps.RefreshTimeline) = False: Call ReadSettings
Call DeleteAllTimelineShapes
Range(Cells(rowsix, cpt.TimelineStart), Cells(GetLastRow + 2, cpt.TimelineEnd + 1)).Clear
Range(Cells(rowsix, cpt.TimelineStart), Cells(GetLastRow + 2 + 5, cpt.TimelineEnd + 1)).Borders.LineStyle = xlNone
Call EA
End Sub
Sub turnOnTimeline()
Set gs = setGSws
gs.Cells(rowtwo, cps.TimelineVisible) = True: gs.Cells(rowtwo, cps.RefreshTimeline) = True: Call ReadSettings
If Not GanttChart Or st.ShowTimeline = False Then Exit Sub
Call DA:
Call DeleteShape("ST_Timeline_Notice", 18)
Call CreateTimeline: Call colorAllPrioritySS: Call EA
End Sub
Sub dontRefreshTimeline()
Call DA:
Dim leftfrom As Double, topfrom As Double, widshp As Double, hgtshp As Double: Dim s As Shape
Set gs = setGSws
gs.Cells(rowtwo, cps.RefreshTimeline) = False: Call ReadSettings
With Range(Cells(firsttaskrow, cpt.TimelineStart), Cells(firsttaskrow, cpt.TimelineEnd))
leftfrom = .Left:topfrom = .Top:widshp = 300:hgtshp = .Height * 3
End With
topfrom = topfrom + 4:widshp = 300:hgtshp = 100
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRound2SameRectangle, Left:=leftfrom, Top:=topfrom, Width:=widshp, Height:=hgtshp)
With s
.Name = "ST_Timeline_Notice": .Line.visible = msoFalse: .TextFrame.Characters.Text = "Timeline Refresh - Paused"
With .TextFrame.Characters.Font
.Color = vbWhite: .size = 14
End With
.TextFrame.VerticalAlignment = xlVAlignCenter:.TextFrame.HorizontalAlignment = xlHAlignLeft
With s.Fill
.Solid: .ForeColor.RGB = vbRed
End With
End With
Call EA
End Sub

Sub HiliteRow(rownum As Long)
Dim rowrange As Range
If st.ShowRefreshTimeline = False Then
Set rowrange = Range(Cells(rownum, cpg.SS), Cells(rownum, cpg.LC))
Else
Set rowrange = Range(Cells(rownum, cpg.SS), Cells(rownum, cpt.TimelineEnd))
End If
With rowrange
With .Borders(xlEdgeTop)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = vbRed
End With
With .Borders(xlEdgeBottom)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = vbRed
End With
End With
End Sub

Sub UnHiliteRow(Optional cRow As Long)
Dim c As Range
If st.ShowRefreshTimeline = False Then
Set c = Range(Cells(cRow - 1, cpg.SS), Cells(cRow + 1, cpg.LC))
Else
If Cells(cRow, cpg.Task) = sAddTaskPlaceHolder Then
Set c = Range(Cells(cRow, cpg.SS), Cells(cRow, cpt.TimelineEnd))
Else
Set c = Range(Cells(cRow - 1, cpg.SS), Cells(cRow + 1, cpt.TimelineEnd))
End If
End If

If Cells(cRow, cpg.Task) = sAddTaskPlaceHolder Then
With c
With .Borders(xlEdgeTop)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = st.cTBC
End With
With .Borders(xlEdgeBottom)
.LineStyle = xlNone
End With
End With
Else
With c
With .Borders(xlInsideHorizontal)
.LineStyle = xlContinuous: .Weight = xlThin: .Color = st.cTBC
End With
End With
End If
End Sub

Sub TimelineMsg()
MsgBox msg(9)
End Sub

Sub AdjustSizeForMac(farmname As Object)
Dim ControlOnForm As Object
Const SizeCoefForMac = 1.2
With farmname
.Width = .Width * SizeCoefForMac
.Height = .Height * SizeCoefForMac
For Each ControlOnForm In .Controls
With ControlOnForm
.Top = .Top * SizeCoefForMac
.Left = .Left * SizeCoefForMac
.Width = .Width * SizeCoefForMac
.Height = .Height * SizeCoefForMac
On Error Resume Next
.Font.size = .Font.size * SizeCoefForMac
On Error GoTo 0
End With
Next
End With
End Sub

Sub ShowHideTasks()
Call DA:
Dim lrow As Long, cRow As Long, i As Long: Dim b As Boolean, bShowParent As Boolean: Dim percValue As Double
Call ReadSettings: lrow = GetLastRow: ReDim vArrAllData(1 To lrow, 1 To cpg.LC): vArrAllData = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
ActiveSheet.AutoFilterMode = False: bShowParent = False

For cRow = firsttaskrow To lrow
If vArrAllData(cRow, cpg.GEtype) = vbNullString Then GoTo nextrow
percValue = vArrAllData(cRow, cpg.PercentageCompleted)
If IsParentTask(cRow) Then
For i = GetFirstRowOfFamily(cRow) + 1 To GetLastRowOfFamily(cRow)
If st.ShowCompleted Then If percValue = 1 Then bShowParent = True
If st.ShowInProgress Then If percValue > 0 And percValue < 1 Then bShowParent = True
If st.ShowPlanned Then If percValue = 0 Then bShowParent = True
Next i
If bShowParent = False Then Call hiderow(cRow, True) Else Call hiderow(cRow, False)
GoTo nextrow
Else
If percValue = 1 And st.ShowCompleted = False Then
Call hiderow(cRow, True):GoTo nextrow
ElseIf percValue > 0 And percValue < 1 And st.ShowInProgress = False Then
Call hiderow(cRow, True):GoTo nextrow
ElseIf percValue = 0 And st.ShowPlanned = False Then
Call hiderow(cRow, True):GoTo nextrow
Else
Call hiderow(cRow, False):
End If
End If
nextrow:
bShowParent = False
Next cRow
If st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then
lrow = WorksheetFunction.CountA(Range("A" & Cells.Rows.Count & ":A" & Cells.Rows.Count)) + firsttaskrow
If Cells(lrow, cpg.Task) = sAddTaskPlaceHolder Then Cells(lrow, cpg.Task) = vbNullString
Call DeleteShape("S_De", 4)
Else
lrow = GetLastRow + 1
If Cells(lrow, cpg.GEtype) = vbNullString Then
AddNewTaskPlaceholder
End If
End If
Call DrawAllGanttBars: Call colorAllPrioritySS
Call EA
End Sub

Sub hiderow(cRow As Long, b As Boolean)
If b Then Rows(cRow).Hidden = True Else Rows(cRow).Hidden = False
End Sub

Sub HighlightParentTasks()
Dim cRow As Long, lrow As Long, StartRow As Long, EndRow As Long, i As Long: lrow = GetLastRow: i = 1
ReDim vArrAllData(1 To lrow, 1 To cpg.LC): vArrAllData = Range(Cells(1, 1), Cells(lrow, cpg.LC)).value
ReDim tilarr(1 To lrow)
For cRow = firsttaskrow To lrow
If cRow = lrow Then Exit For
If vArrAllData(cRow, cpg.TIL) < vArrAllData(cRow + 1, cpg.TIL) Then tilarr(i) = cRow: i = i + 1
Next cRow
Range(Cells(firsttaskrow, cpg.WBS), Cells(lrow, cpg.TColor - 1)).Font.Bold = False
For cRow = LBound(tilarr) To UBound(tilarr)
If tilarr(cRow) = "" Then GoTo Last
Range(Cells(tilarr(cRow), cpg.WBS), Cells(tilarr(cRow), cpg.TColor - 1)).Font.Bold = True
Cells(tilarr(cRow), cpg.Done).Font.Bold = False
Next cRow
Last:
Range(Cells(firsttaskrow, cpg.Priority), Cells(lrow, cpg.Priority)).Font.Bold = True
End Sub

Sub colorAllPrioritySS()
Call FormatTasks(, allFields, allRows)
End Sub
Sub colorAllPriority()
Call FormatTasks(, fPriority, allRows)
End Sub
Sub colorAllSS()
Call FormatTasks(, fPerc, allRows)
End Sub
Sub FormatTasks(Optional cRowOnly As Long, Optional field As String, Optional familytype As String)
If checkSheetError Then Exit Sub
Dim StartRow As Long, EndRow As Long, cRow As Long, lrow As Long: Dim pervalue As Double: Dim rngStatus As Range: Dim vArrAllD()
If field = "" Then field = allFields
If familytype = "" Then familytype = allRows
tlog "FormatTasks: " & field & familytype
lrow = GetLastRow: StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
ReDim vArrAllD(1 To EndRow, 1 To cpg.LC): vArrAllD = Range(Cells(1, 1), Cells(EndRow, cpg.LC)).value
For cRow = StartRow To EndRow
If field = fPriority Or field = allFields Then
With Cells(cRow, cpg.Priority)
If LCase(vArrAllD(cRow, cpg.Priority)) = "high" Then .Interior.Color = st.cTPCH
If LCase(vArrAllD(cRow, cpg.Priority)) = "normal" Then .Interior.Color = st.cTPCN
If LCase(vArrAllD(cRow, cpg.Priority)) = "low" Then .Interior.Color = st.cTPCL
End With
End If
If field = "cb" Then
With Cells(cRow, cpg.Done)
.HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
End With
End If
If field = fPerc Or field = allFields Then
pervalue = vArrAllD(cRow, cpg.PercentageCompleted): Set rngStatus = Cells(cRow, cpg.Status)
If pervalue = 1 Then
Cells(cRow, cpg.SS).Interior.Color = st.cTSCC
If containsFormula(rngStatus) = False Then rngStatus = "Completed"
Cells(cRow, cpg.Done) = 100
GoTo nexa
End If
If pervalue > 0 And vArrAllD(cRow, cpg.PercentageCompleted) < 1 Then
Cells(cRow, cpg.SS).Interior.Color = st.cTSCI
If containsFormula(rngStatus) = False Then rngStatus = "In Progress"
Cells(cRow, cpg.Done) = 0
GoTo nexa
End If
If pervalue = 0 Then
Cells(cRow, cpg.SS).Interior.Color = st.cTSCP
If containsFormula(rngStatus) = False Then rngStatus = "Planned"
Cells(cRow, cpg.Done) = 0
GoTo nexa
End If
End If
nexa:
Next
tlog "FormatTasks: " & field & familytype
Call DrawPercDataBar: Call DrawCompCheck
End Sub


Sub FormatImportedProject(Optional cRowOnly As Long, Optional field As String, Optional familytype As String)
Dim StartRow As Long, EndRow As Long, cRow As Long, lrow As Long: Dim pervalue As Double: Dim csp As String, cs As String: Dim vArrAllD()
If field = "" Then field = allFields
If familytype = "" Then familytype = allRows
lrow = GetLastRow: StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)
ReDim vArrAllD(1 To EndRow, 1 To cpg.LC): vArrAllD = Range(Cells(1, 1), Cells(EndRow, cpg.LC)).value
For cRow = StartRow To EndRow
Rows(cRow).Font.Color = rgbBlack
With Range(Cells(cRow, cpg.SS), Cells(cRow, cpg.LC))
If Cells(cRow + 1, cpg.TIL) <= Cells(cRow, cpg.TIL) Then .Font.Bold = False Else .Font.Bold = True
End With
With Cells(cRow, cpg.TaskIcon)
.value = "u":
If Cells(cRow, cpg.GEtype) = "T" Then .Font.Name = "Wingdings 3" Else .Font.Name = "Wingdings": .Font.size = 11
End With
With Cells(cRow, cpg.Priority)
.Font.Bold = True:.Font.size = 10:.Font.Color = vbWhite
End With
With Cells(cRow, cpg.Status)
.HorizontalAlignment = xlCenter:.Font.size = 8
End With
With Cells(cRow, cpg.Done)
.value = 0:
End With
Next
With Range(Cells(firsttaskrow, cpg.GEtype), Cells(lrow, cpg.LC))
.RowHeight = taskRowHeight
End With
End Sub

Sub DrawPercDataBar()
Call ReadSettings
Dim bDraw As Boolean: Dim lrow As Long, totalPerc As Double: Dim arrPer As Variant: lrow = GetLastRow:
ReDim arrPer(1 To lrow)
arrPer = Range(Cells(firsttaskrow, cpg.PercentageCompleted), Cells(lrow, cpg.PercentageCompleted))
With Application.WorksheetFunction
totalPerc = .Sum(.index(arrPer, 0, 1))
End With
If st.ShowPercDataBar = False Or totalPerc = 0 Then
Range(Cells(firsttaskrow, cpg.PercentageCompleted), Cells(lrow, cpg.PercentageCompleted)).FormatConditions.Delete
Exit Sub
End If
With Range(Cells(firsttaskrow, cpg.PercentageCompleted), Cells(lrow, cpg.PercentageCompleted))
.FormatConditions.Delete: .FormatConditions.AddDatabar: .FormatConditions(.FormatConditions.Count).ShowValue = True
.FormatConditions(.FormatConditions.Count).SetFirstPriority
If Not 2007 Then ' NEW for 2007
With .FormatConditions(1)
.MinPoint.Modify newtype:=xlConditionValueAutomaticMin: .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
End With
End If
With .FormatConditions(1).barColor
.Color = st.cPDC
End With
If Not Is2007 Then .FormatConditions(1).BarFillType = xlDataBarFillGradient: .FormatConditions(1).Direction = xlContext: .FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
End With
End Sub

Sub DrawCompCheck()
Dim lrow As Long: lrow = GetLastRow:
Range(Cells(firsttaskrow, cpg.Done), Cells(lrow, cpg.Done)).FormatConditions.Delete
With Range(Cells(firsttaskrow, cpg.Done), Cells(lrow, cpg.Done))
.FormatConditions.AddIconSetCondition
.FormatConditions(.FormatConditions.Count).SetFirstPriority
With .FormatConditions(1)
.ReverseOrder = False
.ShowIconOnly = True
.IconSet = ActiveWorkbook.IconSets(xl3Symbols2)
End With
With .FormatConditions(1).IconCriteria(2)
.Type = xlConditionValueNumber
.value = 0
.Operator = 7
If Not Is2007 Then .Icon = xlIconNoCellIcon
End With
With .FormatConditions(1).IconCriteria(3)
.Type = xlConditionValueNumber
.value = 99
.Operator = 7
End With
End With
End Sub
Option Explicit

Option Explicit
Option Explicit
Option Explicit
Option Private Module
Private arrAllData()
Private arrAllDataSorted()

Sub ImportLic(wbname As String)
Dim ImportWB As Workbook
Dim ImportGC As Worksheet
Set ImportWB = Workbooks(wbname): Set ImportGC = ImportWB.Worksheets("GanttSettingsTemplate")
Call uhs
ImportWB.Worksheets("GanttSettingsTemplate").Range("LA1:LZ2").Copy
ThisWorkbook.Worksheets("GanttSettingsTemplate").Range("LA1").PasteSpecial Paste:=xlPasteValues
Call hs: Call CalcColPosGST: ThisWorkbook.Save: Call RefreshRibbon
MsgBox "License Copied"
End Sub

Sub ImportGC(wbname As String, wsname As String)
Call DA
Application.ScreenUpdating = True
frmStatus.show
frmStatus.lblStatusMsg.Caption = "Importing Project - Please wait for a few seconds..."
DoEvents
Application.ScreenUpdating = False
Dim colSS As Long, colLC As Long, colGRT As Long, colWBS As Long, NTR As Long, newPID As Long, OrigGClrow As Long, i As Long, j As Long
Dim colEnableCosts As Long, colBasBudget As Long, colEstBudget As Long, colTSD As Long, colTED As Long, colOrigs As Long
Dim colODBar As Long, colPercBar As Long, colBasBar As Long, colActBar As Long, colDepCon As Long, colEnableTGB As Long
Dim colSelTheme As Long, colHCOL As Long, colDateFormat As Long, colCurrency As Long, colPercentageEntryMode As Long
Dim colPercentageCalculationType As Long
Dim OrigWB As Workbook, newWB As Workbook: Dim GSTrange As Range, rngOrigColors As Range, rngGSColors As Range
Dim startdatehour As Date, dtTSD As Date, dtTED As Date
Dim ws As Worksheet, OrigGC As Worksheet, OrigRS As Worksheet, OrigGS As Worksheet, newGC As Worksheet, newGS As Worksheet, newRS As Worksheet
Dim orgStartHrs As Double, NewProjBasBudget As Double, NewProjEstBudget As Double
Dim NewWorksheetName As String, NewProjectName As String: Dim NewProjectType As String: Dim NewProjectLead As String
Dim bEnableCosts As Boolean, mp As Boolean, wksfound As Boolean, GSMissing As Boolean, RSMissing As Boolean, valDone As Boolean
Dim bShowODBar As Boolean, bShowPercBar As Boolean, bShowBasBar As Boolean, bShowActBar As Boolean, bShowDepCon As Boolean, bShowTGB As Boolean
Set OrigWB = Workbooks(wbname): Set newWB = ThisWorkbook: Set OrigGC = OrigWB.Worksheets(wsname):
OrigWB.Activate: Call LockWB(False): OrigGC.Activate:
colSS = Application.WorksheetFunction.Match("SS", OrigGC.Range("1:1"), 0): colWBS = Application.WorksheetFunction.Match("WBS", OrigGC.Range("1:1"), 0)
colLC = Application.WorksheetFunction.Match("LC", OrigGC.Range("1:1"), 0): OrigGClrow = GetLastRow(OrigGC): NTR = 9
If InStr(1, OrigGC.Cells(rowtwo, colSS), DepSeperator, vbTextCompare) > 0 Then
If CheckSheet(getGSname(OrigGC)) = False Then GSMissing = True Else Set OrigGS = OrigWB.Worksheets(getGSname(OrigGC))
If CheckSheet(getRSname(OrigGC)) = False Then RSMissing = True Else Set OrigRS = OrigWB.Worksheets(getRSname(OrigGC))
Else
If CheckSheet("GS" & OrigGC.Cells(rowtwo, colSS).value) = False Then GSMissing = True Else Set OrigGS = OrigWB.Worksheets("GS" & ActiveSheet.Cells(rowtwo, colSS).value)
If GSMissing = False Then
Set GSTrange = OrigGS.Range("1:1"): colGRT = Application.WorksheetFunction.Match("GRT", GSTrange.value, 0)
If CheckSheet(OrigWB.Worksheets(OrigGS.Name).Cells(rowtwo, colGRT).value) = False Then
RSMissing = True
Else
Set OrigRS = OrigWB.Worksheets(OrigWB.Worksheets(OrigGS.Name).Cells(rowtwo, colGRT).value)
End If
End If
End If
If GSMissing Then MsgBox "Error - Imported Gantt Chart Settings Sheet does not exist."
If RSMissing Then MsgBox "Error - Imported Gantt Chart Resourse Sheet does not exist."
NewWorksheetName = OrigGC.Name: NewProjectName = OrigGC.Cells(6, colWBS)
If GSMissing Then
NewProjectType = InputBox("Setings sheet missing - Type the word Days if the original was created in days or else Hours")
Else
If OrigGS.Cells(2, 1) = "ds082" Then NewProjectType = "Days" Else NewProjectType = "Hours"
Set GSTrange = OrigGS.Range("1:1")
colGRT = Application.WorksheetFunction.Match("GRT", GSTrange.value, 0)
colEnableCosts = Application.WorksheetFunction.Match("EC", GSTrange, 0): bEnableCosts = OrigGS.Cells(2, colEnableCosts)
colBasBudget = Application.WorksheetFunction.Match("BaselineBudget", GSTrange.value, 0): NewProjBasBudget = CDbl(OrigGS.Cells(2, colBasBudget))
colEstBudget = Application.WorksheetFunction.Match("EstimatedBudget", GSTrange.value, 0): NewProjEstBudget = CDbl(OrigGS.Cells(2, colEstBudget))
colTSD = Application.WorksheetFunction.Match("TSD", GSTrange.value, 0): dtTSD = CDate(OrigGS.Cells(2, colTSD))
colTED = Application.WorksheetFunction.Match("TED", GSTrange.value, 0): dtTED = CDate(OrigGS.Cells(2, colTED))
colODBar = Application.WorksheetFunction.Match("ShowOverdueBar", GSTrange.value, 0): bShowODBar = OrigGS.Cells(2, colODBar)
colPercBar = Application.WorksheetFunction.Match("ShowPercBar", GSTrange.value, 0): bShowPercBar = OrigGS.Cells(2, colPercBar)
colBasBar = Application.WorksheetFunction.Match("ShowBaselineBar", GSTrange.value, 0): bShowBasBar = OrigGS.Cells(2, colBasBar)
colActBar = Application.WorksheetFunction.Match("ShowActualBar", GSTrange.value, 0): bShowActBar = OrigGS.Cells(2, colActBar)
colDepCon = Application.WorksheetFunction.Match("ShowDependencyConnector", GSTrange.value, 0): bShowDepCon = OrigGS.Cells(2, colDepCon)
colEnableTGB = Application.WorksheetFunction.Match("EnableBarText", GSTrange.value, 0): bShowTGB = OrigGS.Cells(2, colEnableTGB)
colSelTheme = Application.WorksheetFunction.Match("SelectedTheme", GSTrange.value, 0)
colHCOL = Application.WorksheetFunction.Match("HCOL", GSTrange.value, 0)
colDateFormat = Application.WorksheetFunction.Match("DateFormat", GSTrange.value, 0)
colCurrency = Application.WorksheetFunction.Match("CurrencySymbol", GSTrange.value, 0)
colPercentageEntryMode = Application.WorksheetFunction.Match("PercentageEntryMode", GSTrange.value, 0)
colPercentageCalculationType = Application.WorksheetFunction.Match("PercentageCalculationType", GSTrange.value, 0)
End If
NewProjectLead = Trim(Replace(OrigGC.Cells(rowseven, colWBS).value, "Project Lead:", vbNullString))
ThisWorkbook.Activate: Call ProjectCountPlusOne: newPID = getProjectCounter
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "GS" & newPID:
Set newGS = ThisWorkbook.Worksheets("GS" & newPID)
GST.UsedRange.Copy (newGS.Cells(1, 1))
ActiveSheet.Range("LL1", "LZ2").Clear: Call hideWorksheet(ActiveSheet): Call hideWorksheet(GST)
If LCase(NewProjectType) = "days" Then Call createGCsheet("Days") Else Call createGCsheet("Hours")
If Len(Trim(NewWorksheetName) & " P-" & newPID) > 30 Then
ActiveSheet.Name = Left(Trim(NewWorksheetName), 10) & " P-" & newPID
Else
ActiveSheet.Name = Trim(NewWorksheetName) & " P-" & newPID
End If
Set newGC = ActiveSheet
If RSMissing Then ' do the RS
Call createRSsheet: ActiveSheet.Name = "RS" & newPID: Call hideWorksheet(ActiveSheet)
Else
OrigWB.Activate: OrigRS.visible = xlSheetVisible: newWB.Activate
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "RS" & newPID:
Set newRS = newWB.Sheets("RS" & newPID)
OrigRS.UsedRange.Copy (newRS.Cells(1, 1))
Call hideWorksheet(ActiveSheet): OrigWB.Activate: Call hideWorksheet(OrigRS)
End If
OrigWB.Activate: Call LockWB(True)
ThisWorkbook.Activate: newGC.Activate: Call CalcColPosGCT: Call CalcColPosTimeline
Dim cnt As Integer, indx As Integer
Dim colOrdr()
colOrdr = OrigGC.Range(OrigGC.Cells(1, 1), OrigGC.Cells(1, colLC - 1)).value: cnt = 1
For indx = LBound(colOrdr, 2) To UBound(colOrdr, 2)
If colOrdr(1, indx) = "tType" Then colOrdr(1, indx) = "GEtype"
If colOrdr(1, indx) = "tID" Then colOrdr(1, indx) = "TID"
If colOrdr(1, indx) = "tDependency" Then colOrdr(1, indx) = "Dependency"
If colOrdr(1, indx) = "tDependents" Then colOrdr(1, indx) = "Dependents"
If colOrdr(1, indx) = "tStartConstrain" Then colOrdr(1, indx) = "StartConstrain"
If colOrdr(1, indx) = "tEndConstrain" Then colOrdr(1, indx) = "EndConstrain"
If colOrdr(1, indx) = "TaskType" Then colOrdr(1, indx) = "TaskIcon"
If colOrdr(1, indx) = "TaskLead" Then colOrdr(1, indx) = "Resource"
If colOrdr(1, indx) = "TaskLeadCost" Then colOrdr(1, indx) = "ResourceCost"
If colOrdr(1, indx) = "Completed" Then colOrdr(1, indx) = "Done"
colOrigs = Application.WorksheetFunction.Match(colOrdr(1, indx), ActiveSheet.Range("1:1"), 0)
If colOrigs > 0 Then
If colOrigs <> cnt Then
Columns(colOrigs).EntireColumn.Cut: Columns(cnt).Insert Shift:=xlToRight: Application.CutCopyMode = False
End If
cnt = cnt + 1
End If
Next indx
Call CalcColPosGCT: Call CalcColPosTimeline
Call setGSRSname("GS" & newPID, "RS" & newPID, ActiveSheet, newPID): Set gs = setGSws: Call CalcColPosGST
gs.Cells(rowtwo, cps.GRT) = "RS" & newPID: sArr.LoadAllArrays
If LCase(NewProjectType) = "days" Then
gs.Cells(rowtwo, cps.gcType).value = "ds082"
gs.Cells(rowtwo, cps.CurrentView) = "D": Cells(firsttaskrow, cpg.ESD).value = Date
Cells(firsttaskrow, cpg.ED).value = 10: Cells(firsttaskrow, cpg.EED).value = CDate(GetEndDateFromWorkDays("", Date, 10))
Else
gs.Cells(rowtwo, cps.gcType).value = "s0n84"
gs.Cells(rowtwo, cps.CurrentView) = "HH"
orgStartHrs = sArr.ResourceP(0, 10): startdatehour = Date + orgStartHrs
Cells(firsttaskrow, cpg.ESD) = startdatehour
Cells(firsttaskrow, cpg.ED).value = 6: Cells(firsttaskrow, cpg.EED).value = CalEEDHrs("organization", Cells(firsttaskrow, cpg.ESD), Cells(firsttaskrow, cpg.ED))
End If
gs.Cells(rowtwo, cps.TSD).value = dtTSD: gs.Cells(rowtwo, cps.TED).value = dtTED
gs.Cells(rowtwo, cps.ShowOverdueBar).value = bShowODBar: gs.Cells(rowtwo, cps.ShowPercBar).value = bShowPercBar
gs.Cells(rowtwo, cps.ShowBaselineBar).value = bShowBasBar: gs.Cells(rowtwo, cps.ShowActualBar).value = bShowActBar
gs.Cells(rowtwo, cps.ShowDependencyConnector).value = bShowDepCon: gs.Cells(rowtwo, cps.BarTextEnable).value = bShowTGB
gs.Cells(rowtwo, cps.SelectedTheme).value = OrigGS.Cells(rowtwo, colSelTheme).value
gs.Cells(rowtwo, cps.DateFormat).value = OrigGS.Cells(rowtwo, colDateFormat).value
gs.Cells(rowtwo, cps.PercentageEntryMode).value = OrigGS.Cells(rowtwo, colPercentageEntryMode).value
gs.Cells(rowtwo, cps.PercentageCalculationType).value = OrigGS.Cells(rowtwo, colPercentageCalculationType).value
gs.Cells(rowtwo, cps.CurrencySymbol).value = OrigGS.Cells(rowtwo, colCurrency).value
Set rngOrigColors = OrigGS.Range(OrigGS.Cells(rowsix, colSelTheme + 1), OrigGS.Cells(rowseven, colHCOL - 1))
Set rngGSColors = gs.Range(gs.Cells(rowsix, cps.SelectedTheme + 1), gs.Cells(rowseven, cps.HCOL - 1))
rngOrigColors.Copy 'rngGSColors.Parent.Activate
rngGSColors.PasteSpecial xlPasteFormats
Application.CutCopyMode = False
Cells(rowsix, cpg.WBS).value = Trim(NewProjectName): Cells(rowseven, cpg.WBS).value = "Project Lead: " & NewProjectLead
If bEnableCosts Then
gs.Cells(rowtwo, cps.EnableCostsModule).value = True
If IsNumeric(NewProjBasBudget) Then gs.Cells(rowtwo, cps.BaselineBudget).value = CDbl(NewProjBasBudget)
If IsNumeric(NewProjEstBudget) Then gs.Cells(rowtwo, cps.EstimatedBudget).value = CDbl(NewProjEstBudget)
Call ReCalculateBudgetLineCosts
Else
gs.Cells(rowtwo, cps.EnableCostsModule).value = False:Cells(roweight, cpg.WBS).value = vbNullString
End If
Call ReadSettings: valDone = False
selectvalformula:
If valDone = False Then
arrAllData = OrigGC.Range(OrigGC.Cells(1, 1), OrigGC.Cells(OrigGClrow, colLC)).value
Else
arrAllData = OrigGC.Range(OrigGC.Cells(1, 1), OrigGC.Cells(OrigGClrow, colLC)).Formula
End If
For j = 1 To colLC
If arrAllData(1, j) = newGC.Cells(1, j) Then
newGC.Cells(rownine, j) = arrAllData(rownine, j)
End If
Next j
ReDim arrAllDataSorted(1 To OrigGClrow, cpg.GEtype To cpg.LC)
For i = firsttaskrow To UBound(arrAllData())
For j = 1 To colLC
If arrAllData(1, j) = newGC.Cells(1, cpg.GEtype) Or arrAllData(1, j) = "tType" Then
arrAllDataSorted(i - NTR, cpg.GEtype) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.TID) Or arrAllData(1, j) = "tID" Then
arrAllDataSorted(i - NTR, cpg.TID) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.Dependency) Or arrAllData(1, j) = "tDependency" Then
arrAllDataSorted(i - NTR, cpg.Dependency) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.Dependents) Or arrAllData(1, j) = "tDependents" Then
arrAllDataSorted(i - NTR, cpg.Dependents) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.StartConstrain) Or arrAllData(1, j) = "tStartConstrain" Then
arrAllDataSorted(i - NTR, cpg.StartConstrain) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.EndConstrain) Or arrAllData(1, j) = "tEndConstrain" Then
arrAllDataSorted(i - NTR, cpg.EndConstrain) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.TIL) Or arrAllData(1, j) = "til" Then arrAllDataSorted(i - NTR, cpg.TIL) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.SS) Then arrAllDataSorted(i - NTR, cpg.SS) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.TaskIcon) Or arrAllData(1, j) = "GEType" Then
arrAllDataSorted(i - NTR, cpg.TaskIcon) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.WBS) Then arrAllDataSorted(i - NTR, cpg.WBS) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Task) Then arrAllDataSorted(i - NTR, cpg.Task) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Priority) Then arrAllDataSorted(i - NTR, cpg.Priority) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Status) Then arrAllDataSorted(i - NTR, cpg.Status) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Resource) Or arrAllData(1, j) = "TaskLead" Then
arrAllDataSorted(i - NTR, cpg.Resource) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.ResourceCost) Or arrAllData(1, j) = "TaskLeadCost" Then
arrAllDataSorted(i - NTR, cpg.ResourceCost) = arrAllData(i, j)
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.BSD) Then arrAllDataSorted(i - NTR, cpg.BSD) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.BED) Then arrAllDataSorted(i - NTR, cpg.BED) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.BD) Then arrAllDataSorted(i - NTR, cpg.BD) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.ESD) Then arrAllDataSorted(i - NTR, cpg.ESD) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.EED) Then arrAllDataSorted(i - NTR, cpg.EED) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.ED) Then arrAllDataSorted(i - NTR, cpg.ED) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Work) Then arrAllDataSorted(i - NTR, cpg.Work) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Done) Then arrAllDataSorted(i - NTR, cpg.Done) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.PercentageCompleted) Then arrAllDataSorted(i - NTR, cpg.PercentageCompleted) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.ASD) Then arrAllDataSorted(i - NTR, cpg.ASD) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.AED) Then arrAllDataSorted(i - NTR, cpg.AED) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.AD) Then arrAllDataSorted(i - NTR, cpg.AD) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.BCS) Then arrAllDataSorted(i - NTR, cpg.BCS) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.ECS) Then arrAllDataSorted(i - NTR, cpg.ECS) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.ACS) Then arrAllDataSorted(i - NTR, cpg.ACS) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Notes) Then arrAllDataSorted(i - NTR, cpg.Notes) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.TColor) Then
If OrigGC.Cells(i, j).Interior.ColorIndex > 0 Then newGC.Cells(i, cpg.TColor).Interior.Color = OrigGC.Cells(i, j).Interior.Color
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.TPColor) Then
If OrigGC.Cells(i, j).Interior.ColorIndex > 0 Then newGC.Cells(i, cpg.TPColor).Interior.Color = OrigGC.Cells(i, j).Interior.Color
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.BLColor) Then
If OrigGC.Cells(i, j).Interior.ColorIndex > 0 Then newGC.Cells(i, cpg.BLColor).Interior.Color = OrigGC.Cells(i, j).Interior.Color
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.ACColor) Then
If OrigGC.Cells(i, j).Interior.ColorIndex > 0 Then newGC.Cells(i, cpg.ACColor).Interior.Color = OrigGC.Cells(i, j).Interior.Color
End If
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom1) Then arrAllDataSorted(i - NTR, cpg.Custom1) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom2) Then arrAllDataSorted(i - NTR, cpg.Custom2) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom3) Then arrAllDataSorted(i - NTR, cpg.Custom3) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom4) Then arrAllDataSorted(i - NTR, cpg.Custom4) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom5) Then arrAllDataSorted(i - NTR, cpg.Custom5) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom6) Then arrAllDataSorted(i - NTR, cpg.Custom6) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom7) Then arrAllDataSorted(i - NTR, cpg.Custom7) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom8) Then arrAllDataSorted(i - NTR, cpg.Custom8) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom9) Then arrAllDataSorted(i - NTR, cpg.Custom9) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom10) Then arrAllDataSorted(i - NTR, cpg.Custom10) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom11) Then arrAllDataSorted(i - NTR, cpg.Custom11) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom12) Then arrAllDataSorted(i - NTR, cpg.Custom12) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom13) Then arrAllDataSorted(i - NTR, cpg.Custom13) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom14) Then arrAllDataSorted(i - NTR, cpg.Custom14) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom15) Then arrAllDataSorted(i - NTR, cpg.Custom15) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom16) Then arrAllDataSorted(i - NTR, cpg.Custom16) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom17) Then arrAllDataSorted(i - NTR, cpg.Custom17) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom18) Then arrAllDataSorted(i - NTR, cpg.Custom18) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom19) Then arrAllDataSorted(i - NTR, cpg.Custom19) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.Custom20) Then arrAllDataSorted(i - NTR, cpg.Custom20) = arrAllData(i, j)
If arrAllData(1, j) = newGC.Cells(1, cpg.LC) Then arrAllDataSorted(i - NTR, cpg.LC) = arrAllData(i, j)
Next j
Next i
newGC.Range(newGC.Cells(firsttaskrow, cpg.GEtype), newGC.Cells(OrigGClrow, cpg.LC)) = arrAllDataSorted
For j = 1 To colLC - 1
If OrigGC.Columns(j).Hidden = True Then newGC.Columns(j).Hidden = True Else newGC.Columns(j).Hidden = False
If OrigGC.Columns(j).ColumnWidth <> newGC.Columns(j).ColumnWidth Then newGC.Columns(j).ColumnWidth = OrigGC.Columns(j).ColumnWidth
Next j
If valDone = False Then valDone = True: GoTo selectvalformula
Call DrawProjectName: Call IndentTaskfromTIL: Call WBSNumbering: Call AddNewTaskPlaceholder: Call FormatAllDates: Call FormatAllCosts
Call FormatImportedProject(, allFields, allRows): Call mRefreshGC: Call DA 'Call sts
MsgBox "Import Complete for Project: " & Cells(6, cpg.WBS) & ". All Systems Go. Good Luck!"
Last:
Call RefreshRibbon: Call EA
End Sub

Sub exportXLSX(exType As String)
Call DA
Dim GCWB As Workbook, newWB As Workbook, newwbname As Workbook: Dim gc As Worksheet, rs As Worksheet, gs As Worksheet, newGCsheet As Worksheet, ws As Worksheet, gsDel As Worksheet, rsDel As Worksheet
Dim thiswbname As String, thisGCName As String, newfullpath As String, newfullfolder As String, File_Name As String, foldername As String: Dim answer As Integer
Set GCWB = ThisWorkbook: Set gc = ActiveSheet: thiswbname = ThisWorkbook.Name: thisGCName = ActiveSheet.Name
Set gs = setGSws: Set rs = setRSws
File_Name = "Export - " & thiswbname & " " & Format(Now, "dd-mmm-yy-hhmmss")
#If Mac Then
foldername = GetFolderMac()
#Else
foldername = GetFolder()
#End If
If foldername = "" Then MsgBox "No folder was selected.": Call EA: Exit Sub
File_Name = foldername & "\" & File_Name: Workbooks.Add.SaveAs Filename:=File_Name & ".xlsx": Set newWB = ActiveWorkbook: GCWB.Activate: Call uhs
gc.Activate: gc.Range("A1").Select
If exType = "single" Then
ActiveSheet.Copy After:=newWB.Sheets(newWB.Sheets.Count): Set newGCsheet = ActiveSheet
GCWB.Activate
gs.Copy After:=newWB.Worksheets(thisGCName): ActiveSheet.Range(Cells(rowone, cps.tLicenseVal), Cells(rowtwo, cps.tFirstSavedDate)).Clear: ActiveSheet.visible = xlSheetVeryHidden
GCWB.Activate
rs.Copy After:=newWB.Worksheets(thisGCName): ActiveSheet.visible = xlSheetVeryHidden
newGCsheet.Activate: newfullpath = ActiveWorkbook.FullName: newfullfolder = ActiveWorkbook.Path
With ActiveSheet.Range(Cells(firsttaskrow, cpg.WBS), Cells(GetLastRow + 2, cpg.WBS))
.NumberFormat = "General":.value = .value:.HorizontalAlignment = xlLeft
End With
ActiveSheet.Cells(GetLastRow + 1, cpg.Task).value = "Exported from Gantt Excel on " & Format(Now, "dd-mmm-yy hh:mm:ss")
End If
If exType = "all" Then
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then
ws.Activate: Call CalcColPosGCT: Call CalcColPosTimeline
Set gsDel = setGSws: Set rsDel = setRSws
rsDel.Activate: rsDel.Copy After:=newWB.Sheets(newWB.Sheets.Count): ActiveSheet.visible = xlSheetVeryHidden
gsDel.Activate: gsDel.Copy After:=newWB.Sheets(newWB.Sheets.Count): ActiveSheet.visible = xlSheetVeryHidden
ws.Activate: ws.Copy After:=newWB.Sheets(newWB.Sheets.Count)
newWB.Activate: newfullpath = ActiveWorkbook.FullName: newfullfolder = ActiveWorkbook.Path
With ActiveSheet.Range(Cells(firsttaskrow, cpg.WBS), Cells(GetLastRow + 2, cpg.WBS))
.NumberFormat = "General": .value = .value: .HorizontalAlignment = xlLeft
End With
ActiveSheet.Cells(GetLastRow + 1, cpg.Task).value = "Exported from Gantt Excel on " & Format(Now, "dd-mmm-yy hh:mm:ss")
End If
GCWB.Activate
Next ws
End If
newWB.Save: newWB.Close: GCWB.Activate:
gc.Activate: gc.Cells(firsttaskrow, cpg.Task).Select
Call hs
answer = MsgBox("Backup Complete - The file has been saved to " & newfullpath & vbNewLine & "Do you want to open the file?", vbYesNo + vbQuestion, "Export Complete")
If answer = vbYes Then
#If Mac Then
Call Workbooks.Open(File_Name & ".xlsx")
#Else
Shell "explorer.exe " & newfullpath, vbNormalFocus
#End If
Else
End If
Call EA
End Sub

Sub exportPDF()
Dim wsa As Worksheet
Dim wbA As Workbook
Dim strTime As String
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile As Variant
On Error GoTo errHandler
Set wbA = ActiveWorkbook
Set wsa = ActiveSheet 'ActiveSheet.Cells(rowsix, cpg.ESD).Value = "Exported from Gantt Excel on " & Format(Now, "dd-mmm-yy hh:mm:ss")
strTime = Format(Now(), "yyyymmdd\_hhmmss")
strPath = wbA.Path 'get active workbook folder, if saved
If strPath = "" Then
strPath = Application.DefaultFilePath
End If
strPath = strPath & "\"
'replace spaces and periods in sheet name
strName = Replace(wsa.Name, " ", "")
strName = Replace(strName, ".", "_")
'create default name for savng file
strFile = strName & "_" & strTime & ".pdf"
strPathFile = strPath & strFile
Dim foldername As String
#If Mac Then
foldername = GetFolderMac()
myFile = foldername & strFile
#Else
myFile = Application.GetSaveAsFilename _
(InitialFileName:=strPathFile, _
FileFilter:="PDF Files (*.pdf), *.pdf", _
Title:="Select Folder and FileName to save")
#End If
If myFile <> "False" Then 'export to PDF if a folder was selected
wsa.ExportAsFixedFormat _
Type:=xlTypePDF, _
Filename:=myFile, _
Quality:=xlQualityStandard, _
IncludeDocProperties:=True, _
IgnorePrintAreas:=False, _
OpenAfterPublish:=False

#If Mac Then
Call Workbooks.Open(vbCrLf & myFile)
#Else
Dim answer As Integer
answer = MsgBox("PDF file has been created" & vbNewLine & "Do you want to open the file?", vbYesNo + vbQuestion, "Export Complete")
If answer = vbYes Then
Shell "explorer.exe " & vbCrLf & myFile, vbNormalFocus
Else
End If
#End If
Else
End If
exitHandler:
ActiveSheet.Cells(rowsix, cpg.ESD).value = ""
Exit Sub
errHandler:
MsgBox "Could not create PDF file"
Resume exitHandler
ActiveSheet.Cells(rowsix, cpg.ESD).value = ""
End Sub
Option Explicit
Public bReDrawDependencies As Boolean
Sub MakeTaskParent(Optional t As Boolean)
If ActiveSheet.AutoFilterMode Or st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then
MsgBox "You cannot set this as a parent task when the filter mode is enabled", vbInformation, "Information": Exit Sub
End If
Dim rSel As Range, crng As Range: Dim ir As Long, selRowsCount As Long: Dim bIndentedTasks As Boolean, bMsg As Boolean: Dim iLvl As Integer, tiLvl As Integer, answer As Integer
Set rSel = Selection: selRowsCount = rSel.Rows.Count
For ir = selRowsCount To 1 Step -1
Set crng = rSel.Rows(ir): iLvl = Cells(crng.Row, cpg.Task).IndentLevel: tiLvl = Cells(crng.Row + 1, cpg.Task).IndentLevel
If Cells(crng.Row, cpg.GEtype) = "M" And (iLvl - 1 < tiLvl) Then
bMsg = True
Else
If Cells(crng.Row, cpg.Task).IndentLevel >= 1 And Cells(crng.Row, cpg.GEtype) <> vbNullString Then
If Cells(crng.Row, cpg.Dependents) <> "" Or Cells(crng.Row, cpg.Dependency) <> "" Then
answer = MsgBox(msg(53), vbYesNo + vbQuestion, "Confirm")
If answer = vbNo Then Exit Sub
End If
Cells(crng.Row, cpg.TIL) = Cells(crng.Row, cpg.Task).IndentLevel - 1
Cells(crng.Row, cpg.Task).IndentLevel = Cells(crng.Row, cpg.TIL): bIndentedTasks = True
If Cells(crng.Row + 1, cpg.TIL) > Cells(crng.Row, cpg.TIL) Then
If Cells(crng.Row, cpg.Dependents) <> "" Then
Call DeleteDependencies(crng.Row)
Cells(crng.Row, cpg.Dependents) = vbNullString: Cells(crng.Row, cpg.WBSSuccessors) = vbNullString
End If
If Cells(crng.Row, cpg.Dependency) <> "" Then
Call DeleteDependents(crng.Row)
Cells(crng.Row, cpg.Dependency) = vbNullString: Cells(crng.Row, cpg.WBSPredecessors) = vbNullString
End If
End If
End If
End If
nexir:
Next ir
Last:
Call WBSNumbering: Call HighlightParentTasks: Call PopParentTasks(, allFields): Call DelnDrawAllGanttBars
Set rSel = Nothing: Set crng = Nothing
End Sub
Sub MakeTaskChild(Optional t As Boolean)
If ActiveSheet.AutoFilterMode Or st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then
MsgBox "You cannot set this as a child task when the filter mode is enabled", vbInformation, "Information": Exit Sub
End If
Dim rSel As Range, crng As Range, tRng As Range: Dim ir As Long, selRowsCount As Long:Dim bIndentedTasks As Boolean, bMsg As Boolean:
Dim iLvl As Integer, tiLvl As Integer, answer As Integer
Set rSel = Selection: selRowsCount = rSel.Rows.Count
For ir = 1 To selRowsCount
Set crng = rSel.Rows(ir)
If crng.Row = firsttaskrow Then GoTo nextir
Set tRng = crng.Offset(-1, 0): iLvl = Cells(crng.Row, cpg.Task).IndentLevel: tiLvl = Cells(tRng.Row, cpg.Task).IndentLevel
If Cells(tRng.Row, cpg.GEtype) = "M" And (iLvl + 1 > tiLvl) Then MsgBox msg(58): Exit Sub
If tiLvl >= 0 And (iLvl - tiLvl) <> 1 And Cells(crng.Row, cpg.GEtype) <> vbNullString Then
If Cells(crng.Row, cpg.Dependents) <> "" Or Cells(crng.Row, cpg.Dependency) <> "" Then
answer = MsgBox(msg(53), vbYesNo + vbQuestion, "Confirm")
If answer = vbNo Then Exit Sub
End If
Cells(crng.Row, cpg.TIL) = iLvl + 1: Cells(crng.Row, cpg.Task).IndentLevel = Cells(crng.Row, cpg.TIL): bIndentedTasks = True
If Cells(crng.Row - 1, cpg.TIL) < Cells(crng.Row, cpg.TIL) Then
If Cells(crng.Row - 1, cpg.Dependents) <> "" Then
Call DeleteDependencies(crng.Row - 1)
Cells(crng.Row - 1, cpg.Dependents) = vbNullString
Cells(crng.Row - 1, cpg.WBSSuccessors) = vbNullString
End If
If Cells(crng.Row - 1, cpg.Dependency) <> "" Then
Call DeleteDependents(crng.Row - 1)
Cells(crng.Row - 1, cpg.Dependency) = vbNullString
Cells(crng.Row - 1, cpg.WBSPredecessors) = vbNullString
End If
End If
End If
nextir:
Next ir
Last:
If bIndentedTasks Then'if current is child task and the previous is a parent of this
If Cells(rSel.Row, cpg.Task).IndentLevel > Cells(rSel.Row - 1, cpg.Task).IndentLevel Then
'If The parent task above is being driven then remove it
If Cells(rSel.Row - 1, cpg.Dependency) <> vbNullString And rSel.Row - 1 <> rownine And Cells(rSel.Row, cpg.Dependency) = vbNullString Then
RemoveDependenciesOnTaskChild rSel.Row - 1, True
End If
End If
Call WBSNumbering: Call HighlightParentTasks:
Call ReCalculateDates(, estDates, allRows) ' me thinks not required
Call PopParentTasks(, allFields): Call DelnDrawAllGanttBars
Set rSel = Nothing: Set crng = Nothing
End Sub

Sub WBSNumbering(Optional t As Boolean)
tlog "WBSNumbering"
If st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then Exit Sub
On Error Resume Next
Dim bLocked As Boolean: Dim cIndent As Long, i As Long, j As Long, r As Long, arrWBS() As Long, lrow As Long:Dim WBS As String
r = firsttaskrow:i = 0: lrow = GetLastRow
ReDim arrWBS(0 To 0) As Long
Columns(cpg.WBS).NumberFormat = "@"
Do While Cells(r, cpg.GEtype) <> vbNullString
If Cells(r, cpg.Task) <> vbNullString Then
cIndent = Cells(r, cpg.TIL): Cells(r, cpg.Task).IndentLevel = cIndent'cIndent = Cells(r, cpg.Task).IndentLevel
If cIndent = 0 Then
i = i + 1:WBS = CStr(i):ReDim arrWBS(0 To 0)
Else
ReDim Preserve arrWBS(0 To cIndent) As Long:cIndent = cIndent - 1
If arrWBS(cIndent) <> 0 Then arrWBS(cIndent) = arrWBS(cIndent) + 1 Else arrWBS(cIndent) = 1
If arrWBS(cIndent + 1) <> 0 Then
For j = cIndent + 1 To UBound(arrWBS)
arrWBS(j) = 0
Next j
End If
WBS = CStr(i)
For j = 0 To cIndent
WBS = WBS & "." & CStr(arrWBS(j))
Next j
End If
Cells(r, cpg.WBS).value = WBS
If InStr(1, WBS, ".0") > 0 Then
Cells(r, cpg.TIL) = Cells(r, cpg.Task).IndentLevel - 1: Cells(r, cpg.Task).IndentLevel = Cells(r, cpg.TIL)
Cells(r, cpg.WBS).value = Replace(WBS, ".0", vbNullString)
End If
Cells(r, cpg.WBS).Errors(xlNumberAsText).Ignore = True
End If
nextrow:
r = r + 1
Loop
On Error GoTo 0
Range(Cells(rownine, cpg.WBS), Cells(lrow, cpg.WBS)).Columns.AutoFit
GenerateOutLineGroups
Application.ErrorCheckingOptions.NumberAsText = False
tlog "WBSNumbering"
End Sub
Sub MoveTaskUp(Optional t As Boolean)
Call DA: Call MoveUp: Call EA
End Sub

Sub MoveUp(Optional t As Boolean)
If ActiveSheet.AutoFilterMode Or st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then
MsgBox "You cannot move tasks when the filter mode is enabled", vbInformation, "Information": Exit Sub
End If
If Selection.Cells.Rows.Count > 1 Then MsgBox "You can move only one task or milestone at a time.", vbInformation, "Information":Exit Sub
If Cells(Selection.Row, cpg.GEtype) = vbNullString Then Exit Sub
If Selection.Row <= firsttaskrow Then Exit Sub
If Cells(Selection.Row, cpg.Task).IndentLevel = 1 And Cells(Selection.Row - 1, cpg.Task).IndentLevel = 0 Then MsgBox "Child tasks can be moved only within the parent task.": Exit Sub
If Cells(Selection.Row, cpg.Task).IndentLevel > Cells(Selection.Row - 1, cpg.Task).IndentLevel Then MsgBox "Child tasks can be moved only within the parent task.": Exit Sub
If IsDataCollapsed = True Then Exit Sub
Dim r As Range: Dim cLevel As Long, tlrow As Long, iRow As Long
Set r = Selection:cLevel = Cells(r.Row, cpg.Task).IndentLevel
Find if selection has sub tasks
If cLevel < Cells(r.Row + 1, cpg.Task).IndentLevel Then
Has SubTasks, Find last row where this parent task ends
tlrow = r.Row + 1
Do Until Cells(tlrow, cpg.GEtype) = vbNullString
If Cells(tlrow, cpg.Task).IndentLevel <= cLevel Then
Exit Do
End If
tlrow = tlrow + 1
Loop
tlrow = tlrow - 1
Else
tlrow = r.Row
End If
Find the row where the selected task has to be cut and pasted
iRow = r.Row - 1
Do Until Cells(iRow, cpg.Task).IndentLevel <= cLevel Or Cells(iRow, cpg.Task).IndentLevel = 0
iRow = iRow - 1
Loop
Dim sArr() As Double, fArr() As Double, p As Long, k As Long
ReDim sArr(1 To tlrow - r.Row + 1)
ReDim fArr(1 To iRow + tlrow - r.Row + 1)
k = 1
For p = r.Row To tlrow
sArr(k) = Rows(p).RowHeight:k = k + 1
Next p
Rows(r.Row & ":" & tlrow).Cut
k = 1
For p = iRow To (iRow + tlrow - r.Row)
fArr(k) = Rows(p).RowHeight:k = k + 1
Next p
Rows(iRow).Insert Shift:=xlDown

Last:
Application.CutCopyMode = False
Call WBSNumbering
LastEnd:
Call DrawAllGanttBars: Call colorAllPrioritySS: r.Select: Set r = Nothing
End Sub
Sub MoveTaskDown(Optional t As Boolean)
Call DA: Call MoveDown:Call EA
End Sub

Sub MoveDown(Optional t As Boolean)
Set gs = setGSws
If ActiveSheet.AutoFilterMode Or _
(gs.Cells(rowtwo, cps.ShowCompleted) = 0 Or gs.Cells(rowtwo, cps.ShowInProgress) = 0 Or gs.Cells(rowtwo, cps.ShowPlanned) = 0) Then
MsgBox "You cannot move tasks when the filter mode is enabled", vbInformation, "Information"
Exit Sub
End If
If Selection.Cells.Rows.Count > 1 Then MsgBox "You can move only one task at a time.", vbInformation, "Information":Exit Sub
If Cells(Selection.Row + 1, cpg.Task) = sAddTaskPlaceHolder Then Exit Sub
If Cells(Selection.Row, cpg.GEtype) = vbNullString Then Exit Sub
If Selection.Row <= rownine Then Exit Sub
If IsDataCollapsed = True Then Exit Sub
Dim r As Range: Dim cLevel As Long, tlrow As Long, iRow As Long
Set r = Selection: cLevel = Cells(r.Row, cpg.Task).IndentLevel
Find if selection has sub tasks
If cLevel < Cells(r.Row + 1, cpg.Task).IndentLevel Then
Has SubTasks, Find last row where this parent task ends
tlrow = r.Row + 1
Do Until Cells(tlrow, cpg.GEtype) = vbNullString
If Cells(tlrow, cpg.Task).IndentLevel <= cLevel Then Exit Do
tlrow = tlrow + 1
Loop
tlrow = tlrow - 1
Else
tlrow = r.Row
End If
Section is selected to be moved down and this is the last section , so no movement needed
If Cells(tlrow + 1, cpg.GEtype) = vbNullString And Cells(r.Row, cpg.GEtype) <> vbNullString Then Exit Sub
If Cells(tlrow + 1, cpg.Task).IndentLevel <> cLevel Then MsgBox "Child tasks can be moved only within the parent task": Exit Sub
tlRow+1 is the position where we may have to move the above task, But as tlrow+1 may have sub tasks, we need to find the last subtask row of tlrow+1
iRow = tlrow + 1
If Cells(iRow, cpg.Task).IndentLevel < Cells(iRow + 1, cpg.Task).IndentLevel Then
has subtasks'Loop this last subtask
iRow = iRow + 1
Do Until Cells(iRow, cpg.GEtype) = vbNullString
If Cells(iRow, cpg.Task).IndentLevel <= Cells(tlrow + 1, cpg.Task).IndentLevel Then Exit Do
iRow = iRow + 1
Loop
Else
iRow = tlrow + 2
End If
Rows(r.Row & ":" & tlrow).Cut
Rows(iRow).Insert Shift:=xlDown
Call PopParentTasks(, allFields)
Last:
Application.CutCopyMode = False
Call WBSNumbering
LastEnd:
Call DrawAllGanttBars: Call colorAllPrioritySS: r.Select: Set r = Nothing
End Sub

Sub DeleteTasks(Optional t As Boolean)
Dim x As Long, selFirstRow As Long, selLastRow As Long, selTIL As Long, vArr() As Long, AffectedTasksCount As Long, cActualAffectedTasks As Long
Dim bvArrHasValues As Boolean
Dim arrAffectedTIDS()
Dim arrAffectedTrows()
Dim vStr, i As Long, j As Long, StartRow As Long, EndRow As Long, cRow As Long, cLevel As Long, tidr As Long
EndRow = GetLastRow: AffectedTasksCount = 0: cActualAffectedTasks = 0:
If Cells(Selection.Row, cpg.GEtype) = vbNullString Or Selection.Row < firsttaskrow Or Selection.Row > EndRow Then MsgBox msg(50): GoTo Last
If ActiveSheet.AutoFilterMode Or st.ShowCompleted = False Or st.ShowInProgress = False Or st.ShowPlanned = False Then
MsgBox "You cannot delete a task when Filter mode is on", vbInformation, "Information": GoTo Last
End If
If MsgBox(msg(52), vbQuestion + vbYesNo, "Delete Confirmation") <> vbYes Then GoTo Last
If Selection.Areas.Count = 1 Then ' multirow delete
If Selection.Rows.Count > 1 Then
selFirstRow = Selection.Rows(1).Row: selLastRow = Selection.Rows.Count + selFirstRow - 1
If selLastRow < firsttaskrow Or selLastRow > EndRow Then MsgBox msg(50): GoTo Last
selTIL = Cells(selFirstRow, cpg.TIL)
For x = selFirstRow To selLastRow
If selTIL <> Cells(x, cpg.TIL) Then GoTo oops
Next x
If WorksheetFunction.CountA(Range(Cells(selFirstRow, cpg.Dependency), Cells(selLastRow, cpg.Dependents))) = 0 Then
Selection.EntireRow.Delete: Call WBSNumbering:
If selTIL > 0 Then
Call PopParentTasks(, allFields): Call HighlightParentTasks: Call DelnDrawAllGanttBars:
Call colorAllPrioritySS: GoTo Last
Else
Call ReCalculateBudgetLineCosts: GoTo Last
End If
Else
MsgBox msg(51): GoTo Last
End If
End If
End If
oops:
If Selection.Areas.Count > 1 Then MsgBox msg(50): GoTo Last
If Selection.Rows.Count > 1 Then MsgBox msg(50): GoTo Last
If Cells(Selection.Row, cpg.TIL) = 0 And Cells(Selection.Row + 1, cpg.TIL) = 0 Then ' for normal tasks without dep
If Cells(Selection.Row, cpg.Dependency) = "" And Cells(Selection.Row, cpg.Dependents) = "" Then
Selection.EntireRow.Delete:: Call WBSNumbering: Call ReCalculateBudgetLineCosts: GoTo Last
End If
End If
If Cells(Selection.Row + 1, cpg.TIL) = Cells(Selection.Row, cpg.TIL) Then ' for child tasks with same indent wo dep
If Cells(Selection.Row, cpg.Dependency) = "" And Cells(Selection.Row, cpg.Dependents) = "" Then
Selection.EntireRow.Delete: Call WBSNumbering: Call PopParentTasks(, allFields): Call FormatTasks(, fStatus, allRows)
Call DelnDrawAllGanttBars: Call colorAllPrioritySS: Call ReCalculateBudgetLineCosts: GoTo Last
End If
End If
StartRow = Selection.Row: cLevel = Cells(StartRow, cpg.Task).IndentLevel
ReDim ArrInd(1 To EndRow)
ReDim arrAffectedTIDS(1 To EndRow)
i = StartRow + 1
Do Until i > EndRow Or Cells(i, cpg.Task).IndentLevel <= cLevel
i = i + 1
Loop
EndRow = i - 1
ReDim vArr(1 To 1)
For cRow = StartRow To EndRow
If Cells(cRow, cpg.Dependents) <> vbNullString Then
bvArrHasValues = True: vStr = Split(Cells(cRow, cpg.Dependents), DepSeperator): 'j = UBound(vArr) 'ReDim vArr(1 To UBound(vStr) + j)
For i = LBound(vStr) To UBound(vStr) - 1 'get affected task ids aka driven tasks
AffectedTasksCount = AffectedTasksCount + 1: ReDim Preserve arrAffectedTIDS(1 To AffectedTasksCount)
arrAffectedTIDS(AffectedTasksCount) = CLng(vStr(i)):
Next i
Call DeleteDependencies(cRow) 'Delete all sucessors
End If
If Cells(cRow, cpg.Dependency) <> vbNullString Then Call DeleteDependents(cRow) ' delete all drivers
Next cRow
Rows(StartRow & ":" & EndRow).EntireRow.Delete: Call WBSNumbering: Call CalcDepFormulas
If bvArrHasValues Then
arrAffectedTIDS = RemoveDuplicatesFromArray(arrAffectedTIDS)
ReDim arrAffectedTrows(LBound(arrAffectedTIDS) To UBound(arrAffectedTIDS))
For i = LBound(arrAffectedTIDS) To UBound(arrAffectedTIDS) 'get affected task rows into array
If arrAffectedTIDS(i) <> "" And getRowNoForTaskID(CLng(arrAffectedTIDS(i))) > 0 Then
cActualAffectedTasks = cActualAffectedTasks + 1:
arrAffectedTrows(i) = getRowNoForTaskID(CLng(arrAffectedTIDS(i)))
End If
Next i
If cActualAffectedTasks > 0 Then 'recalcDep for affected task rows
For i = LBound(arrAffectedTrows) To UBound(arrAffectedTrows)
If arrAffectedTrows(i) = "" Then GoTo checkedAffectedTasks
Call ReCalcDepFormulas(CLng(arrAffectedTrows(i)), True)
Next
End If
End If
checkedAffectedTasks:
Call ClearDepFormulas: Call PopParentTasks(, allFields): Call DelnDrawAllGanttBars: Call colorAllPrioritySS: Call HighlightParentTasks
Last:
End Sub

Function getRowNoForTaskID(TID As Long) As Long
Dim rngFound As Range
Set rngFound = Range(Cells(rowone, cpg.TID), Cells(10000, cpg.TID)).Find(TID, , xlFormulas, xlWhole)
If Not rngFound Is Nothing Then ' some error handler
getRowNoForTaskID = rngFound.Row
Else
getRowNoForTaskID = 0
End If
End Function

Function RemoveDuplicatesFromArray(sourceArray As Variant)
Dim duplicateFound As Boolean
Dim arrayIndex As Integer, i As Integer, j As Integer
Dim deduplicatedArray() As Variant
arrayIndex = -1: deduplicatedArray = Array(1)
For i = LBound(sourceArray) To UBound(sourceArray)
duplicateFound = False
For j = LBound(deduplicatedArray) To UBound(deduplicatedArray)
If sourceArray(i) = deduplicatedArray(j) Then
duplicateFound = True
Exit For
End If
Next j
If duplicateFound = False Then
arrayIndex = arrayIndex + 1
ReDim Preserve deduplicatedArray(arrayIndex)
deduplicatedArray(arrayIndex) = sourceArray(i)
End If
Next i
RemoveDuplicatesFromArray = deduplicatedArray
End Function

Sub DeleteDependents(cRow As Long)
Dim dIDs As String, vStr As Variant
Dim tidr As Long
dIDs = Cells(cRow, cpg.Dependency)
If dIDs <> vbNullString Then
Dim i As Long, j As Long
Dim tRng As Range
Dim dStr As Variant, Fstr As String
 Set tRng = Range(Cells(firsttaskrow, cpg.TID), Cells(Cells.Rows.Count, cpg.TID))
vStr = Split(dIDs, DepSeperator)
For i = LBound(vStr) To UBound(vStr) - 1
ltemp = tRng.Find(Left(vStr(i), InStr(1, vStr(i), "_") - 1), , xlFormulas, xlWhole).Row
tidr = getTIDRow(Left(vStr(i), InStr(1, vStr(i), "_") - 1))
dStr = Split(Cells(tidr, cpg.Dependents), DepSeperator)
dStr = Split(Cells(tRng.Find(Left(vStr(i), InStr(1, vStr(i), "_") - 1), , xlFormulas, xlWhole).Row, cpg.Dependents), DepSeperator)
For j = 0 To UBound(dStr) - 1
If CLng(dStr(j)) <> CLng(Cells(cRow, cpg.TID)) Then
Fstr = Fstr & dStr(j) & DepSeperator
End If
Next j
tidr = getTIDRow(Left(vStr(i), InStr(1, vStr(i), "_") - 1))
Cells(tidr, cpg.Dependents) = Fstr
Cells(tRng.Find(Left(vStr(i), InStr(1, vStr(i), "_") - 1), , xlFormulas, xlWhole).Row, cpg.Dependents) = Fstr
Fstr = vbNullString
Next i
End If
End Sub
Sub DeleteDependencies(cRow As Long)
Dim dIDs As String, newdepstring As String, Fstr As String
Dim foundrow As Long, tidr As Long, i As Long, noOfDepInRow As Long, x As Long, j As Long
Dim dStr As Variant, vStr As Variant, vstrNew As Variant, splitdep As Variant 'Dim tRng As Range
dIDs = Cells(cRow, cpg.Dependents)
If dIDs <> vbNullString Then
Set tRng = Range(Cells(firsttaskrow, cpg.TID), Cells(Cells.Rows.Count, cpg.TID))
vStr = Split(dIDs, DepSeperator)
For i = LBound(vStr) To UBound(vStr) - 1
tidr = getTIDRow(CLng(vStr(i))): dStr = Split(Cells(tidr, cpg.Dependency), DepSeperator)
dStr = Split(Cells(tRng.Find(vStr(i), , xlFormulas, xlWhole).Row, cpg.Dependency), DepSeperator)
For j = 0 To UBound(dStr) - 1
If CLng(Left(dStr(j), InStr(1, dStr(j), "_") - 1)) <> CLng(Cells(cRow, cpg.TID)) Then
Fstr = Fstr & dStr(j) & DepSeperator
End If
Next j
foundrow = getTIDRow(CLng(vStr(i)))
foundrow = tRng.Find(vStr(i), , xlFormulas, xlWhole).Row
Cells(foundrow, cpg.Dependency) = Fstr
Last:
Fstr = vbNullString
Next i
End If
End Sub
Sub DuplicateTasks(Optional t As String)
tlog "DuplicateTasks"
Set gs = setGSws
If ActiveSheet.AutoFilterMode Or _
(gs.Cells(rowtwo, cps.ShowCompleted) = 0 Or gs.Cells(rowtwo, cps.ShowInProgress) = 0 Or gs.Cells(rowtwo, cps.ShowPlanned) = 0) Then
MsgBox "You cannot duplicate tasks when the filter mode is enabled", vbInformation, "Information"
Exit Sub
End If
If IsDataCollapsed = True Or Cells(Selection.Row, cpg.GEtype) = vbNullString Or Selection.Row <= rownine Then Exit Sub
If Selection.Cells.Rows.Count > 1 Then MsgBox "You can duplicate only one task at a time.", vbInformation, "Information":Exit Sub
If FreeVersion Then If Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A")) - 3 >= cFreeVersionTasksCount Then sTempStr1 = msg(80) & msg(82): frmBuyPro.show: Exit Sub
Dim r As Range: Dim cLevel As Long, tlrow As Long, iRow As Long, dRow As Long, k As Long
Set r = Selection:cLevel = Cells(r.Row, cpg.Task).IndentLevel
Find if selection has sub tasks
If cLevel < Cells(r.Row + 1, cpg.Task).IndentLevel Then
Has SubTasks, Find last row where this parent task ends
tlrow = r.Row + 1
Do Until Cells(tlrow, cpg.GEtype) = "S" Or Cells(tlrow, cpg.GEtype) = vbNullString
If Cells(tlrow, cpg.Task).IndentLevel <= cLevel Then Exit Do
tlrow = tlrow + 1
Loop
tlrow = tlrow - 1
Else
tlrow = r.Row
End If
Rows(r.Row & ":" & tlrow).Copy:Rows(tlrow + 1).Insert Shift:=xlDown
k = 0
For dRow = tlrow + 1 To tlrow + 1 + (tlrow - r.Row)
Cells(dRow, cpg.TID) = GetNextIDNumber:
Cells(dRow, cpg.Dependency) = vbNullString: Cells(dRow, cpg.Dependents) = vbNullString: Cells(dRow, cpg.StartConstrain) = vbNullString: Cells(dRow, cpg.EndConstrain) = vbNullString
Rows(dRow).RowHeight = Rows(r.Offset(k, 0).Row).RowHeight:k = k + 1
Call DrawTasksBorders(dRow)
Next
Call DeleteExtrasRowsInFree:
Call PopParentTasks(, allFields)
Last:
Application.CutCopyMode = False
Call WBSNumbering: Call DelnDrawAllGanttBars: Call colorAllPrioritySS
tlog "DuplicateTasks"
End Sub
Sub ClearOutLineGroups(Optional t As Boolean)
Call ReadSettings
Dim bLock As Boolean
ActiveSheet.Cells.EntireRow.ClearOutline
Range(Cells(1, cpt.TimelineEnd + 1), Cells(Cells.Rows.Count, Cells.Columns.Count)).EntireColumn.ColumnWidth = 0
call DrawAllGanttBars
End Sub
Sub GenerateOutLineGroups(Optional t As Boolean)
Call ReadSettings
If st.ShowGrouping = False Then Exit Sub
Dim gOutlineRow(1 To 8) As Long, i As Long, j As Long, cLevel As Long, lrow As Long
With ActiveSheet
.Cells.EntireRow.ClearOutline
lrow = GetLastRow:i = rownine + 2
Do Until i > lrow
cLevel = Cells(i, cpg.Task).IndentLevel + 1
If cLevel < 9 Then Rows(i).OutlineLevel = cLevel
nextrow:
i = i + 1
Loop
.Outline.SummaryRow = xlAbove
End With
Range(Cells(1, cpt.TimelineEnd + 1), Cells(Cells.Rows.Count, Cells.Columns.Count)).EntireColumn.ColumnWidth = 0
Last:
call DrawAllGanttBars
End Sub

Function CountOfLevel(WBS As String) As Integer
CountOfLevel = Len(WBS) - Len(Replace(WBS, ".", vbNullString)) + 1
End Function
Sub ExpandAlLGroups(Optional t As Boolean)
If FreeVersion Then Call ShowLimitation:Exit Sub
ActiveSheet.Outline.ShowLevels RowLevels:=8
Call ClearFilters: Call DrawAllGanttBars
End Sub
Sub CollapseAllGroups(Optional t As Boolean)
If FreeVersion Then Call ShowLimitation:Exit Sub
ActiveSheet.Outline.ShowLevels RowLevels:=1
Call DrawAllGanttBars: Call DeleteShape("S_De", 4)
End Sub
Option Explicit

Sub AddToCellMenu(Optional t As String)
Set gs = setGSws
If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
#If Mac Then
Exit Sub
#End If
On Error Resume Next
Call DeleteRightClickGanttMenu 'Delete the controls first to avoid duplicates
If gs.Cells(rowtwo, cps.ShowCompleted) = 0 Or gs.Cells(rowtwo, cps.ShowInProgress) = 0 Or gs.Cells(rowtwo, cps.ShowPlanned) = 0 Then Exit Sub
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Set ContextMenu = Application.CommandBars("Cell")
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=1)
.OnAction = "mAddTaskAtSelection"
.Caption = "Add Task at Selection"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=2)
.OnAction = "mAddTaskBelowSelection"
.Caption = "Add Task below Selection"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=3)
.OnAction = "mTriggerAddMilestone"
.Caption = "Add Milestone"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=4)
.OnAction = "mEditTask"
.Caption = "Edit Task / Milestone"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=5)
.OnAction = "mDeleteTask"
.Caption = "Delete Task/ Milestone"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=6)
.OnAction = "mDuplicateTask"
.Caption = "Duplicate Task / Milestone"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=7)
.OnAction = "mMakeParent"
.Caption = "Make Task Parent"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=8)
.OnAction = "mMakeChild"
.Caption = "Make Task Child"
End With
With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=9)
.OnAction = "mRefreshGC"
.Caption = "Refresh Gantt Chart"
End With
ContextMenu.Controls(10).BeginGroup = True
On Error GoTo 0
End Sub
Sub mTriggerAddMilestone(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
bAddMilestone = True: Call AddNewTask("AtSelection")
EA
End If
End Sub

Sub DeleteRightClickGanttMenu(Optional t As String)
#If Mac Then
Exit Sub
#End If
Application.CommandBars("Cell").reset
End Sub
Sub mMakeParent(Optional t As String)
DA
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then Call MakeTaskParent
EA
End Sub
Sub mMakeChild(Optional t As String)
DA
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then Call MakeTaskChild
EA
End Sub

Sub mRefreshGC(Optional t As String)
tlog "RefreshGC"
If GanttChart Then
Call DA
If checkSheetError Then GoTo Last
Application.ScreenUpdating = True
frmStatus.show
frmStatus.lblStatusMsg.Caption = "Refreshing Gantt Chart"
DoEvents
Application.ScreenUpdating = False
Call CalcColPosGCT: Call CalcColPosTimeline: Call CalcColPosGST: Call ReadSettings: Call sArr.LoadAllArrays
Application.ScreenUpdating = True
frmStatus.lblStatusMsg.Caption = "Updating Resources"
DoEvents
Application.ScreenUpdating = False
Call RemoveInvalidResourcesNamesFromTasks
Application.ScreenUpdating = True
frmStatus.lblStatusMsg.Caption = "Recalculating Dates"
DoEvents
Application.ScreenUpdating = False
Call ReCalculateDates(, allFields, allRows): Call CalcAllResCost: Call PopParentTasks(, allFields)
Application.ScreenUpdating = True
frmStatus.lblStatusMsg.Caption = "Redrawing Timeline"
DoEvents
Application.ScreenUpdating = False
Call CreateTimeline: Call colorAllPrioritySS
Application.ScreenUpdating = True
frmStatus.lblPleaseWait.Caption = "Auto closing status popup"
frmStatus.lblStatusMsg.ForeColor = vbGreen
Unload frmStatus
Call DeleteShape("S_RefreshButton", 15)
End If
Last: Call EA
tlog "RefreshGC"
End Sub
Sub mAddTaskAtSelection(Optional t As String)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
Call AddNewTask("AtSelection")
EA
End If
End Sub
Sub mAddTaskBelowSelection(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
Call AddNewTask("BelowSelection")
EA
End If
End Sub
Sub mEditTask(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
Call EditExistingTask
EA
End If
End Sub
Sub mDeleteTask(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
Call DA: Call DeleteTasks: Call EA
End If
End Sub
Sub mDuplicateTask(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
Call DA: Call DuplicateTasks:Call EA
End If
End Sub
Function IsMac2011(Optional t As Boolean) As Boolean
#If Mac Then
#If MAC_OFFICE_VERSION >= 15 Then
#Else
If val(Application.Version) < 15 Then
IsMac2011 = True
End If
#End If
#End If
End Function
Function IsLanguageSettingEnglish(Optional t As Boolean) As Boolean
If Application.localizedlanguage = 1033 Then
IsLanguageSettingEnglish = True
End If
End Function

Sub SetupRightClickGanttMenu(Optional b As Boolean)
If GanttChart Then AddToCellMenu Else DeleteRightClickGanttMenu
End Sub
Option Explicit
Option Private Module

Function msg(msgNo As Long) As String
Select Case msgNo
Case Is = 1
msg = "This is a Parent task. Parent task dates are autocalculated based on child tasks."
Case Is = 2
msg = "This is a Parent Task. Percent Complete is autocalculated based on child tasks."
Case Is = 3
msg = "This is a Parent Task (summary task). It is recommended not to assign resources to summary tasks. While you can assign a resource, this resource will not have an influence over the task dates and resource cost."
Case Is = 4
msg = "This is a Parent Task. Parent duration is autocalculated based on child tasks. You can change the duration calculation in the Settings form."
Case Is = 5
msg = "This is a Parent Task. Parent Tasks autocalculate costs based on child tasks. You can turn this off in the Settings window."
Case Is = 6
msg = "This is a Parent Task. Parent Tasks autocalculate resource costs based on child tasks. You can turn this off in the Settings window."
Case Is = 7
msg = "WBS is autocalculated. Please do not edit this cell."
Case Is = 8
msg = "Resource cost is automatically calculated based on resource cost * task duration. You can turn this off in the Settings window."
Case Is = 9
msg = "Please enable Show Timeline and Refresh Timeline in the Gantt Excel Menu."
Case Is = 10
msg = "Percentage Complete calculation is set to Automatic. You can turn this off in the Settings window."
Case Is = 11
msg = "While you can assign more than one resource to a task, this task will use the Organization holidays and workdays for calculating the Estimated End Date."
Case Is = 12
msg = "Task Status is automatically set based on Task Percentage Completed."
Case Is = 13
msg = "Milestone is a zero duration task. The duration will reset to zero."
Case Is = 14
msg = "Please enter numbers only."
Case Is = 15
msg = "Please type a task name. If you want to DELETE the task please use the Delete Task button."
Case Is = 16
msg = "Please type a task name."
Case Is = 17
msg = "You cannot add a new task when the data is filtered."
Case Is = 18
msg = "The resource name(s) entered is not in the resources list." & vbNewLine & "Please click on the Resources button in the Gantt menu to add resources."
Case Is = 19
msg = "Task priority can be set as High, Normal or Low."
Case Is = 20
msg = "Formulas may affect calculations. We will try to retain the formula however please note that there is a possibility that the automation may delete it."
Case Is = 21
msg = "Formulas are not allowed in this particular column."
Case Is = 22
msg = "End Date is before the Start Date. Please check dates."
Case Is = 23
msg = "Please enter a valid Start Date."
Case Is = 24
msg = "Please enter a valid End Date."
Case Is = 25
msg = "Please enter the Baseline Start Date first."
Case Is = 26
msg = "Please enter the Actual Start Date first."
Case Is = 27
msg = "Please check the duration."
Case Is = 28
msg = "The Start date of this task is dependent on other tasks and hence can't be changed. Please edit the dependency/ lag/ lead for this task."
Case Is = 29
msg = "The Finish date of this task is dependent on other tasks and hence can't be changed. Please edit the dependency/ lag/ lead for this task."
Case Is = 30
msg = "Sorry this task icon cell is not editable."
Case Is = 31
msg = " worksheet is missing or has been deleted. Please send this file to support@ganttexcel.com"
Case Is = 32
msg = "This is the Daily Planner Version. If you want to schedule task duration in HOURS, please purchase the Hourly Planner Version. "
Case Is = 33
msg = "This is the Hourly Planner Version. If you want to schedule task duration in DAYS, please purchase the Daily Planner Version. "
Case Is = 34 'reuse
msg = "Gantt Resource Template worksheet has been deleted. Please send this file to support@ganttexcel.com"
Case Is = 35
msg = "Excel 2011 on Mac does not support Gantt Excel - Please upgrade to Excel 2016 and above"
Case Is = 36
msg = "Please do not insert/delete cells, rows and columns. You can insert tasks by clicking the Add task button. You can insert columns in the Settings window."
Case Is = 37
msg = "Please do not edit these cells. Editing these cells will cause automation errors."
Case Is = 38
msg = "Please do not edit the timeline section manually. Any edits will be deleted when the timeline is redrawn."
Case Is = 39
msg = "We will make an attempt to save your file now. Please send it to support@ganttexcel.com"
Case Is = 40
msg = "File Saved: Please send it to support@ganttexcel.com"
Case Is = 41
msg = "This is a milestone and the estimated end date is set automatically. Please set the Estimated Start Date"
Case Is = 42
msg = "This is a milestone and the baseline end date is set automatically. Please set the Baseline Start Date"
Case Is = 43
msg = "This is a milestone and the actual end date is set automatically. Please set the Actual Start Date"
Case Is = 44
msg = "Please note that this may affect performance if you have more than 1000 tasks."
Case Is = 45
msg = "This is a Parent task. Baseline Parent task dates are autocalculated based on child tasks. You can turn this off in the Settings window."
Case Is = 46
msg = "This is a Parent task. Baseline duration is autocalculated based on child tasks. You can turn this off in the Settings window."
Case Is = 47
msg = "This is a Parent task. Actual Parent task dates are autocalculated based on child tasks. You can turn this off in the Settings window."
Case Is = 48
msg = "This is a Parent task. Actual duration is autocalculated based on child tasks. You can turn this off in the Settings window."
Case Is = 49
msg = "This column is left intentionally blank for aesthetic reasons. Please insert a column in the Settings window to add other data."
Case Is = 50
msg = "Please select a single task or milestone to delete it."
Case Is = 51
msg = "One or more selected tasks have a task dependency. Please select a single task or milestone to delete it."
Case Is = 52
msg = "Are you sure you want to permanently delete the selected task and all its child tasks if any?"
Case Is = 53
msg = "This action will remove task dependencies if any from this task. Continue?"
Case Is = 54
msg = "The above parent task will lose task dependencies if any. Continue?"
Case Is = 55
msg = "Do you want to convert this task to a milestone?"
Case Is = 56
msg = "Important: Please note that if the date calculation is lesser than the start year or if it is greater than the end year then you may see errors. Please change this only if it is necessary."
Case Is = 57
msg = "The Start Date and Finish Date of this task is dependent on other tasks and hence duration can't be changed. Please edit the dependency/ lag/ lead for this task."
Case Is = 58
msg = "Milestones cannot have child tasks under them."
Case Is = 59
msg = "Sorry the Task Done column cell is not editable. Please double click the cell to mark the task complete or to remove the checkmark."
Case Is = 60
msg = "Please upgrade to the Pro version."
Case Is = 61
msg = "Please note that filtering is an experimental feature. Some buttons in the Gantt menu will be temporarily disabled when filters are turned on. Kindly Turn Off filters once you have reviewed data. You can also click the Export to XLSX button and then add filters you like in the XLSX file." & vbNewLine & "On Mac OS please make a selection of the tasks info with headers before you add filters." & vbNewLine & "Please DO NOT SORT THE DATA."
Case Is = 62
msg = "This is a Parent task. Work is autocalculated based on child tasks."
Case Is = 63
msg = "This column is a read-only column and the dependency data is auto populated."
Case Is = 64
msg = "The headers for columns cannot be set as blank. Please type the header name."
Case Is = 65
msg = " called from non gantt chart worksheet"
Case Is = 66
msg = " Please send this file to support@ganttexcel.com"
Case Is = 67
msg = "Some cells contain errors due to incorrect user formulas. "
Case Is = 68
msg = "Percentage Bar cannot be moved"
Case Is = 69
msg = "Moving Timeline bars in the Daily view is a new experimental feature. We will add this feature to other views in the future."
Case Is = 70
msg = "License code validated. "
Case Is = 71
msg = "Thank you for downloading this free version. "
Case Is = 72
msg = "Thank you for upgrading to "
Case Is = 73
msg = "You already have the Pro Version."
Case Is = 74
msg = "Click on the Add Gantt Chart button from the Gantt Menu to add a new Gantt Chart"
Case Is = 75
msg = "Contact support@ganttexcel.com for help."
Case Is = 76
msg = "Please recheck formula. User entered error in cell "
Case Is = 77
msg = "This column is a read-only column and the shapeinfo is auto populated."
Case Is = 78 ' REUSE
msg = "This is an experimental feature designed to work only in daily view."
Case Is = 79
msg = "Your Excel version - 2007 does not support Dashboard as the required pivots and slicers do not work on Excel 2007"
Case Is = 80
msg = "This free version is limited to 15 tasks. "
Case Is = 81
msg = "This feature is not available in the Free version. "
Case Is = 82
msg = "Please purchase the Pro version to create Unlimited Gantt Charts with Unlimited tasks."
Case Is = 83
msg = "This version does not work on MAC OS. Please upgrade to Pro MAC for cross compatibility between Windows and MAC OS."
Case Is = 84
msg = "You cannot add/edit resources when the data is filtered."
Case Is = 85
msg = "This workbook has been locked to protect the digital signature. Nothing will be affected by the discarded signature. Everything will still work exactly as programmed. The digital signature has served its purpose — to give you confidence that the code was not tampered with in between the time that Gantt Excel signed it and when it was delivered to you. " & vbNewLine & "Do you want to proceed unlocking the workbook? "
Case Is = 86
msg = "We hope you have backed up your file. "
Case Is = 87
msg = "This free version is limited to One Gantt Chart only. "
Case Is = 88
msg = "Please note that this action will invalidate the digital signature. Nothing will be affected by the discarded signature. Everything will still work exactly as programmed. The digital signature has served its purpose — to give you confidence that the code was not tampered with in between the time that Gantt Excel signed it and when it was delivered to you."
Case Is = 89
msg = "The Dashboard Data worksheet and the Pivot Sheet are now unhidden"
Case Is = 90
msg = "The Dashboard Data worksheet and the Pivot Sheet are now hidden"
Case Is = 91
msg = "This field contains a user entered formula in its respective cell."
Case Is = 92
msg = "This worksheet has user entered formulas that contain one or more errors. Please recheck the formulas or delete them."
Case Is = 93
msg = "Update Failed. Please recheck the formulas or delete them to continue."
End Select
End Function

Sub myMsgBox(cRow As Long, cCol As Long, msgNo As Long)
Dim leftfrom As Double, topfrom As Double, widshp As Double, hgtshp As Double: Dim s As Shape
Call DeleteShape("S_W", 3)
With Range(Cells(cRow, cCol), Cells(cRow, cCol + 1))
leftfrom = .Left:topfrom = .Top:widshp = .Width:hgtshp = .Height
End With
leftfrom = leftfrom - 40: topfrom = topfrom - 70: widshp = 300: hgtshp = 60
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangularCallout, Left:=leftfrom, Top:=topfrom, Width:=widshp, Height:=hgtshp)
With s
.Name = "S_W"
.ShapeStyle = msoShapeStylePreset5'.Shadow.Type = msoShadow21
.ZOrder msoBringToFront
.TextFrame.Characters.Text = msg(msgNo)
.TextFrame.Characters.Font.Color = vbBlack
End With
With s.Line
.visible = msoTrue
.Weight = 2
End With
With s.Fill
.visible = msoTrue'.ForeColor.Brightness = 0.8
.Solid
End With
End Sub



Option Explicit
Option Explicit
Option Private Module

Sub RememberResArrays()
newHolidaysArray = sArr.HolidaysP: newResourcesArray = sArr.ResourceP: newWorkdaysArray = sArr.WorkdaysP
End Sub

Sub RemoveInvalidResourcesNamesFromTasks(Optional bPopulateResourceCost As Boolean)
Call RememberResArrays: ResArraysReady = True
Dim cRow As Long, lrow As Long, rcount As Long: Dim sReso As String, sresources(), vR, i As Integer, j As Integer
Dim bFound As Boolean: Dim sNewResource As String
lrow = GetLastRow
rcount = UBound(newResourcesArray())
For cRow = firsttaskrow To lrow
sReso = Cells(cRow, cpg.Resource).value
If sReso = vbNullString Then
Cells(cRow, cpg.ResourceCost) = vbNullString
Else
vR = Split(sReso & sResourceSeperator, sResourceSeperator)
sNewResource = vbNullString
For i = 0 To UBound(vR) - 1
bFound = False
For j = 0 To rcount
If LCase(newResourcesArray(j, 0)) = LCase(vR(i)) Then
sNewResource = sNewResource & vR(i) & sResourceSeperator
bFound = True
End If
Next
Next i
If sNewResource = vbNullString Then
Cells(cRow, cpg.Resource) = sNewResource
Else
Cells(cRow, cpg.Resource) = Left(sNewResource, Len(sNewResource) - 2)
End If
End If
Next cRow
ResArraysReady = False
End Sub

Sub CalcAllResCost(Optional cRowOnly As Long, Optional familytype As String)
Call RememberResArrays: ResArraysReady = True
If st.CalResCosts = False Then GoTo Last
If familytype = "" Then familytype = allRows
Dim cRow As Long, StartRow As Long, EndRow As Long: Dim sresources As String: Dim resnewcost As Double, lDur As Double: Dim vR, i As Integer
StartRow = getStartRow(cRowOnly, familytype): EndRow = getEndRow(cRowOnly, familytype)

For cRow = StartRow To EndRow
If IsParentTask(cRow) Then GoTo nextrow
sresources = Cells(cRow, cpg.Resource)
If sresources = vbNullString Or sresources = "Organization" Then sresources = "Organization"
lDur = Cells(cRow, cpg.ED)
vR = Split(sresources & sResourceSeperator, sResourceSeperator)
For i = 0 To UBound(vR) - 1
Call getResValue(vR(i), newResourcesArray)
resnewcost = resnewcost + newResourcesArray(resvalue, 1)
Next
If resnewcost <= 0 Then Cells(cRow, cpg.ResourceCost) = "": GoTo nextrow
Cells(cRow, cpg.ResourceCost) = resnewcost * lDur
resnewcost = 0
nextrow:
Next cRow
Last:
ResArraysReady = False
End Sub
#If Mac = False Or MAC_OFFICE_VERSION >= 15 Then
Public gobjRibbon As IRibbonUI
Public bSecTask As Boolean
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
Set gobjRibbon = ribbon
End Sub

Public Sub OnActionButton(control As IRibbonControl)
Select Case control.ID
Case Is = "btnAddGanttChart"
TriggerAddNewSheet
Case Is = "btnQuarterlyView"
DA
Call BuildView("Q")
EA
Case Is = "btnHalfYearlyView"
DA
Call BuildView("HY")
EA
Case Is = "btnYearlyView"
DA
Call BuildView("Y")
EA
End Select
End Sub

Sub OnActionCheckBox(control As IRibbonControl, pressed As Boolean)
If ActiveSheet.Cells(1, 1) <> "GEType" Then Exit Sub
DA
Call ReadSettings
Set gs = setGSws
Select Case control.ID
Case "chkShowCompleted"
If st.ShowCompleted = False Then
gs.Cells(rowtwo, cps.ShowCompleted) = 1
Else
If st.ShowPlanned = False And st.ShowInProgress = False Then
MsgBox "Atleast one option has to be checked", vbInformation, "Information"
gs.Cells(rowtwo, cps.ShowCompleted) = 1
GoTo Last
Else
gs.Cells(rowtwo, cps.ShowCompleted) = 0
End If
End If
Call ShowHideTasks
Case "chkShowInProgress"
If st.ShowInProgress = False Then
gs.Cells(rowtwo, cps.ShowInProgress) = 1
Else
If st.ShowCompleted = False And st.ShowPlanned = False Then
MsgBox "Atleast one option has to be checked", vbInformation, "Information"
gs.Cells(rowtwo, cps.ShowInProgress) = 1
GoTo Last
Else
gs.Cells(rowtwo, cps.ShowInProgress) = 0
End If
End If
Call ShowHideTasks
Case "chkShowPlanned"
If st.ShowPlanned = False Then
gs.Cells(rowtwo, cps.ShowPlanned) = 1
Else
If st.ShowCompleted = False And st.ShowInProgress = False Then
MsgBox "Atleast one option has to be checked", vbInformation, "Information"
gs.Cells(rowtwo, cps.ShowPlanned) = 1
GoTo Last
Else
gs.Cells(rowtwo, cps.ShowPlanned) = 0
End If
End If
Call ShowHideTasks
Case "ShowTimeline"
If st.ShowTimeline = False Then Call turnOnTimeline Else Call turnOffTimeline
Case "RefreshTimeline"
If st.RefreshTimeline = False Then Call turnOnTimeline Else Call dontRefreshTimeline
Case "chkEnableGrouping"
If FreeVersion Then
ShowLimitation
GoTo Last
End If
If st.ShowGrouping = False Then
gs.Cells(rowtwo, cps.ShowGrouping) = 1: Call DA: Call GenerateOutLineGroups: Call DrawAllGanttBars: Call EA
Else
gs.Cells(rowtwo, cps.ShowGrouping) = 0: Call DA: Call ClearOutLineGroups: Call DrawAllGanttBars: Call EA
End If
End Select
Last:
ReadSettings
RefreshRibbon
AddToCellMenu
End Sub


Sub GetPressedCheckBox(control As IRibbonControl, ByRef bReturn)
Dim bSC As Boolean, bSI As Boolean, bSP As Boolean, bSG As Boolean, bST As Boolean, bRT As Boolean
If ActiveSheet.Cells(1, 1) <> "GEType" Then Exit Sub
Call ReadSettings
If st.ShowCompleted = True Then bSC = True Else bSC = False
If st.ShowInProgress = True Then bSI = True Else bSI = False
If st.ShowPlanned = True Then bSP = True Else bSP = False
If st.ShowTimeline = True Then bST = True Else bST = False
If st.RefreshTimeline = True Then bRT = True Else bRT = False
If st.ShowGrouping = True Then bSG = True Else bSG = False

Select Case control.ID
Case "chkShowCompleted"
If bSC Then bReturn = True Else bReturn = False
Exit Sub
Case "chkShowInProgress"
If bSI Then bReturn = True Else bReturn = False
Exit Sub
Case "chkShowPlanned"
If bSP Then bReturn = True Else bReturn = False
Exit Sub
Case "ShowTimeline"
If bST Then bReturn = True Else bReturn = False
Exit Sub
Case "RefreshTimeline"
If bRT Then bReturn = True Else bReturn = False
Exit Sub
Case "chkEnableGrouping"
If bSG Then bReturn = True Else bReturn = False
Exit Sub
End Select
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
Dim bAutoFilterMode As Boolean
If control.ID = "btnAddGanttChart" Then enabled = True
If GST.Cells(rowtwo, cps.tLicenseVal) = 0 And (control.ID = "btnAddGanttChart" Or control.ID = "grpGanttCharts") Then enabled = False
If control.ID = "btnUpgrade" Then
pstrLicType = GetLicType
If pstrLicType = pstrDP Or pstrLicType = pstrDPM Or pstrLicType = pstrHP Or pstrLicType = pstrHPM Or pstrLicType = pstrHD Or pstrLicType = pstrHDM Then enabled = False Else enabled = True
End If
If ActiveSheet.Cells(1, 1) <> "GEType" Then Exit Sub
bAutoFilterMode = ActiveSheet.AutoFilterMode
If control.ID = "grpAboutGanttExcel" Then
enabled = True
ElseIf control.ID = "grpGanttCharts" Then
If GST.Cells(rowtwo, cps.tLicenseVal) Then enabled = True Else enabled = False
Else
If Cells(rowone, cpg.WBS) = cHeaderName Then
enabled = True
If control.ID = "btnAddSection" Or control.ID = "btnAddTaskAtSelection" Or control.ID = "btnAddTaskBelowSelection" Or _
control.ID = "btnDuplicate" Or control.ID = "btnDelete" Or control.ID = "btnMoveUp" Or control.ID = "btnMoveDown" Or _
control.ID = "btnSetRowHeight" Or control.ID = "btnDashboard" Or control.ID = "btnParent" Or control.ID = "btnChild" Or control.ID = "btnSetMilestone" Then

If st.ShowPlanned = False Or st.ShowInProgress = False Or st.ShowCompleted = False Or bAutoFilterMode = True Then
enabled = False
End If
End If
Select Case control.ID
Case Is = "chkShowCompleted"
If st.ShowGrouping = True Or bAutoFilterMode Then enabled = False Else enabled = True
Case Is = "chkShowPlanned"
If st.ShowGrouping = True Or bAutoFilterMode Then enabled = False Else enabled = True
Case Is = "chkShowInProgress"
If st.ShowGrouping = True Or bAutoFilterMode Then enabled = False Else enabled = True
Case Is = "btnCollapseAllGroups"
If st.ShowPlanned = False Or st.ShowInProgress = False Or st.ShowCompleted = False Or bAutoFilterMode Then
enabled = False
Else
If st.ShowGrouping = True Or bAutoFilterMode Then enabled = True Else enabled = False
End If
Case Is = "btnExpandAllGroups"
If st.ShowPlanned = False Or st.ShowInProgress = False Or st.ShowCompleted = False Or bAutoFilterMode Then
enabled = False
Else
If st.ShowGrouping = True Or bAutoFilterMode Then enabled = True Else enabled = False
End If
Case Is = "chkEnableGrouping"
If st.ShowPlanned = False Or st.ShowInProgress = False Or st.ShowCompleted = False Or bAutoFilterMode Then enabled = False Else enabled = True
End Select
Else
enabled = False
End If
End If
End Sub
Public Sub HideFilterGroups(control As IRibbonControl, ByRef visible)
If GanttChart Then
If control.ID = "grpFilter" Then visible = True
Else
If control.ID = "GroupSortFilter" Then visible = True
End If
End Sub
Public Sub HideProofingGroups(control As IRibbonControl, ByRef visible)
If GanttChart Then
If control.ID = "grpSpellCheck" Then visible = True
Else
If control.ID = "GroupProofing" Then visible = True
End If
End Sub
Public Sub GetVisible(control As IRibbonControl, ByRef visible)
If control.ID = "btnAbout" Then visible = True
If control.ID = "btnUpgrade" Or control.ID = "Upgrade" Then
pstrLicType = GetLicType
If pstrLicType = pstrDP Or pstrLicType = pstrDPM Or pstrLicType = pstrHP Or pstrLicType = pstrHPM Or pstrLicType = pstrHD Or pstrLicType = pstrHDM Then visible = False Else visible = True
ElseIf GetLicType = vbNullString Then
visible = False
Else
visible = True
End If
End Sub
Public Sub GetVisiblev(control As IRibbonControl, ByRef visible)
If ActiveSheet.Cells(1, 1) <> "GEType" Then Exit Sub
If control.ID = "btnHourly" Then
If ActiveSheet.Cells(1, 1) <> "GEType" Then visible = False: Exit Sub Else visible = True
If st.HGC Then visible = True Else visible = False
End If
End Sub

Sub getlabelF(control As IRibbonControl, ByRef label)
If control.ID = "btnUpgrade" Then
label = getlabelTrigger(control.ID)
End If
End Sub
Sub btnEditGanttChart(control As IRibbonControl)
If Not GanttChart Then MsgBox "Select a Gantt Worksheet": Exit Sub
LoadNewGanttFormOnDblClick
End Sub
Sub btnDashboardRib(control As IRibbonControl)
Call CheckDupSettingsSheets: Call CreateDashboard(getPID)
End Sub
Sub AddTaskAtSelection(control As IRibbonControl)
mAddTaskAtSelection
End Sub
Sub AddTaskBelowSelection(control As IRibbonControl)
mAddTaskBelowSelection
End Sub
Sub AddMilestone(control As IRibbonControl)
mTriggerAddMilestone
End Sub
Sub EditTask(control As IRibbonControl)
mEditTask
End Sub
Sub DuplicateTaskRib(control As IRibbonControl)
mDuplicateTask
End Sub
Sub DeleteTask(control As IRibbonControl)
mDeleteTask
End Sub
Sub tIndentLeft(control As IRibbonControl)
mMakeParent
End Sub
Sub tIndentRight(control As IRibbonControl)
mMakeChild
End Sub
Sub MoveUpAboveRib(control As IRibbonControl)
MoveTaskUp
End Sub
Sub MoveDownBelowRib(control As IRibbonControl)
MoveTaskDown
End Sub
Sub HourlyViewRib(control As IRibbonControl)
DA
Call BuildView("HH")
EA
End Sub
Sub DailyViewRib(control As IRibbonControl)
DA
Call BuildView("D")
EA
End Sub
Sub SetupTimelineRib(control As IRibbonControl)
If st.ShowRefreshTimeline = False Then TimelineMsg: Exit Sub
DA
frmTimeline.show
EA
End Sub
Sub WeeklyViewRib(control As IRibbonControl)
DA
Call BuildView("W")
EA
End Sub
Sub MonthlyViewRib(control As IRibbonControl)
DA
Call BuildView("M")
EA
End Sub
Sub GotoStartRib(control As IRibbonControl)
DA
sts
EA
End Sub
Sub GotoTodayRib(control As IRibbonControl)
DA
stt
EA
End Sub
Sub GotoEndRib(control As IRibbonControl)
DA
ste
EA
End Sub
Sub SetRowHeightRib(control As IRibbonControl)
SetRowHeight
End Sub
Sub OutlineShowDetailRib(control As IRibbonControl)
Call DA: ExpandAlLGroups: Call EA
End Sub
Sub OutlineHideDetailRib(control As IRibbonControl)
Call DA: CollapseAllGroups: Call EA
End Sub
Sub SetColumnWidthRib(control As IRibbonControl)
SetColumnWidth
End Sub
Sub ExportToPDF(control As IRibbonControl)
Call exportPDF
End Sub
Sub ExportToXLSX(control As IRibbonControl)
Call DA: Call exportXLSX("single"): Call EA
End Sub
Sub AboutPopUpRib(control As IRibbonControl)
AboutPopUpTrigger
End Sub

Sub ShowHolidayList(control As IRibbonControl)
ShowHolidayListTrigger
End Sub
Sub SpellChecking(control As IRibbonControl)
DoSpellCheck
End Sub

Sub AddFilterToTasks(control As IRibbonControl)
Call AddFilterToTasksTrigger
End Sub

Sub ClearFilterToTasks(control As IRibbonControl)
ClearFilters
End Sub

Sub ShowResources(control As IRibbonControl)
ShowResourcesTrigger
End Sub

Sub UpgradeLicense(control As IRibbonControl)
AddNewLicense
End Sub
Sub ShowSettings(control As IRibbonControl)
ShowSettingsTrigger
End Sub

Sub OpenContactUs(control As IRibbonControl)
OpenHyperlink contactURL
End Sub

#End If

Sub RefreshRibbon(Optional t As Boolean)
#If Mac = False Or MAC_OFFICE_VERSION >= 15 Then
If gobjRibbon Is Nothing Then
MsgBox "The file ribbon has been reset. Please re-open the file again", vbInformation, "Information"
Else
gobjRibbon.Invalidate
End If
If t = 0 Then If GanttChart Then AddToCellMenu Else DeleteRightClickGanttMenu
#Else
AddMenusForMac
#End If
End Sub

Sub ShowSettingsTrigger(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
Call CheckDupSettingsSheets
frmSettings.show
EA
End If
End Sub
Sub ShowResourcesTrigger(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
DA
Call CheckDupSettingsSheets
frmResources.show
EA
End If
End Sub
Sub ShowHolidayListTrigger(Optional t As Boolean)
If GST.Cells(rowtwo, cps.tLicenseVal) And Cells(rowone, cpg.WBS) = cHeaderName Then
If TimelineVisible = True Then
DA
frmTimeline.show
EA
Else
End If
End If
End Sub

Function getlabelTrigger(sControl) As String
If sControl = "btnUpgrade" Then
If FreeVersion Then
getlabelTrigger = "Upgrade"
Else
getlabelTrigger = "Activate License"
End If
End If
End Function
Sub AboutPopUpTrigger(Optional t As String)
frmAbout.show
End Sub

Option Explicit
Option Private Module
Sub DrawEnableMacrosButton(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
If Not GanttChart(ws) Then Exit Sub
Dim leftfrom As Double, topfrom As Double: Dim s As Shape: Dim colTask As Long: colTask = getColTask(ws)
Call DeleteShape("SG_EnableMacrosButton", 21, ws)
With ws.Range(ws.Cells(firsttaskrow, colTask), ws.Cells(firsttaskrow, colTask))
leftfrom = .Left:topfrom = .Top
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=leftfrom, Top:=topfrom + 4, Width:=600, Height:=50)
With s
.Name = "SG_EnableMacrosButton":.Line.visible = msoFalse:
.TextFrame.Characters.Text = "Automation Off - Enable Macros to Continue. Click Here for Help."
With .TextFrame.Characters.Font
.Color = vbWhite:.size = 14
End With
.TextFrame.VerticalAlignment = xlVAlignCenter: .TextFrame.HorizontalAlignment = xlHAlignLeft
With s.Fill
.Solid:.ForeColor.RGB = vbRed
End With
End With
ws.Hyperlinks.Add Anchor:=s, Address:="https://www.ganttexcel.com/blocked-macros-enable-macros-in-microsoft-excel-microsoft-office/", ScreenTip:="Enable Macros"
End Sub

Function getColGEType(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColGEType = Application.WorksheetFunction.Match("GEType", r.value, 0)
End Function
Function getColSS(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColSS = Application.WorksheetFunction.Match("SS", r.value, 0)
End Function
Function getColTask(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColTask = Application.WorksheetFunction.Match("Task", r.value, 0)
End Function
Function getColLC(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColLC = Application.WorksheetFunction.Match("LC", r.value, 0)
End Function
Function getColTS(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColTS = Application.WorksheetFunction.Match("TimelineStart", r.value, 0)
End Function
Function getColTE(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColTE = Application.WorksheetFunction.Match("TimelineEnd", r.value, 0)
End Function
Function getColLLC(ws As Worksheet) As Long
Dim r As Range: Set r = ws.Range("1:1")
getColLLC = Application.WorksheetFunction.Match("LLC", r.value, 0)
End Function

Sub DrawFilterIcon(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
If Not GanttChart(ws) Then Exit Sub
Dim leftfrom As Double, topfrom As Double: Dim s As Shape
Call DeleteShape("SG_Filter", 9, ws)
With ws.Cells(rownine, getColTask(ws))
leftfrom = .Left + .Width - 20: topfrom = .Top + 5
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=leftfrom, Top:=topfrom, Width:=8, Height:=6)
With s
.Rotation = 180:.Name = "SG_Filter":.Line.visible = msoFalse:.OnAction = "AddFilterToTasksTrigger"
End With
With s.Fill
.Solid:.ForeColor.RGB = rgbGray
End With
End Sub

Sub DrawTNavIcons(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
If Not GanttChart(ws) Then Exit Sub
Dim leftfrom As Double, topfrom As Double: Dim s As Shape: Dim colLC As Long, colLLC As Long
Call DeleteShape("ST_TSBackIcon", 13, ws): Call DeleteShape("ST_TEBackIcon", 13, ws):
Call DeleteShape("ST_TSFrontIcon", 14, ws): Call DeleteShape("ST_TEFrontIcon", 14, ws)
colLC = getColLC(ws): colLLC = getColLLC(ws)
With ws.Cells(rowsix, colLC)
leftfrom = .Left + 2: topfrom = .Top + 6
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=leftfrom, Top:=topfrom + 2, Width:=8, Height:=8)
With s
.Rotation = -90:.Name = "ST_TSBackIcon":.Line.visible = msoFalse:.OnAction = "ScrollTimelineBackS"
End With
With s.Fill
.Solid:.ForeColor.RGB = vbWhite
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=leftfrom, Top:=topfrom + 12, Width:=8, Height:=8)
With s
.Rotation = 90:.Name = "ST_TSFrontIcon":.Line.visible = msoFalse:.OnAction = "ScrollTimelineFrontS"
End With
With s.Fill
.Solid:.ForeColor.RGB = vbWhite
End With
With ws.Cells(rowsix, colLLC)
leftfrom = .Left + 2: topfrom = .Top + 6
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=leftfrom, Top:=topfrom + 2, Width:=8, Height:=8)
With s
.Rotation = -90:.Name = "ST_TEBackIcon":.Line.visible = msoFalse:.OnAction = "ScrollTimelineBackE"
End With
With s.Fill
.Solid:.ForeColor.RGB = vbWhite
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeIsoscelesTriangle, Left:=leftfrom, Top:=topfrom + 12, Width:=8, Height:=8)
With s
.Rotation = 90:.Name = "ST_TEFrontIcon":.Line.visible = msoFalse:.OnAction = "ScrollTimelineFrontE"
End With
With s.Fill
.Solid:.ForeColor.RGB = vbWhite
End With
End Sub

Sub DrawProjectName()
Dim leftfromAs Double, widshp As Double, topfrom As Double, hgtshp As Double: Dim s As Shape
Call DeleteShape("SG_Project", 10, ActiveSheet)
With Range(Cells(rowsix, cpg.WBS), Cells(rowsix, cpg.LC - 1))
leftfrom = .Left - 6:topfrom = .Top:widshp = .Width:hgtshp = .Height
End With
topfrom = topfrom + 6:widshp = widshp - 0.5:hgtshp = hgtshp - 5
Set s = ActiveSheet.Shapes.AddShape(Type:=msoShapeRound2SameRectangle, Left:=leftfrom, Top:=topfrom, Width:=400, Height:=hgtshp)
With s
.Name = "SG_Project": .OnAction = "LoadNewGanttFormOnDblClick"
.Line.visible = msoFalse: .TextFrame.Characters.Text = Cells(rowsix, cpg.WBS)
With .TextFrame.Characters.Font
.Color = st.cPRC: .size = 18
End With
.TextFrame.VerticalAlignment = xlVAlignCenter:.TextFrame.HorizontalAlignment = xlHAlignLeft
End With
With s.Fill
.Solid: .ForeColor.RGB = vbWhite
End With
s.Select
With Selection.ShapeRange.ThreeD
.SetPresetCamera (msoCameraOrthographicFront):.LightAngle = 10:.PresetLighting = msoLightRigBrightRoom
.BevelTopType = msoBevelAngle:.BevelTopInset = 3:.BevelTopDepth = 4.5:.BevelBottomType = msoBevelNone
End With
If Not Is2007 Then
With Selection.ShapeRange.Shadow
.Type = msoShadow25:.visible = msoTrue:.Style = msoShadowStyleOuterShadow:.Blur = 4.55
.OffsetX = 1.5647190602:.OffsetY = 2.0764523261:.RotateWithShape = msoTrue:.ForeColor.RGB = RGB(0, 0, 0)
.Transparency = 0.6999999881:.size = 100
End With
End If
Cells(firsttaskrow, cpg.Task).Select
End Sub

Sub DrawFreezeIcon(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
If Not GanttChart(ws) Then Exit Sub
Dim leftfrom As Double, topfrom As Double: Dim s As Shape: Dim colLC As Long
Call DeleteShape("ST_Freeze", 9, ws): colLC = getColLC(ws)
With ws.Cells(rownine, colLC)
leftfrom = .Left + 1: topfrom = .Top + 1
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeCorner, Left:=leftfrom, Top:=topfrom, Width:=10, Height:=10)
With s
.Name = "ST_Freeze":.Line.visible = msoFalse:.OnAction = "ToggleFreeze": .Rotation = 90:
End With
With s.Fill
.Solid:.ForeColor.RGB = rgbLightGray
End With
End Sub

Sub DrawRefreshButton(Optional ws As Worksheet)
If ws Is Nothing Then Set ws = ActiveSheet
Dim s As Shape: Dim leftfrom As Double, topfrom As Double: Dim colTS As Long: colTS = getColTS(ws)
If Not GanttChart(ws) Then Exit Sub
If GST.Cells(2, cps.SavedDate) = Date Then Exit Sub

Call DeleteShape("S_RefreshButton", 15, ws)
With ws.Cells(firsttaskrow, colTS)
leftfrom = .Left:topfrom = .Top + 4
End With
Set s = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=leftfrom, Top:=topfrom, Width:=150, Height:=30)
With s
.Name = "S_RefreshButton": .Line.visible = msoFalse: .TextFrame.Characters.Text = "Refresh Project": .OnAction = "mRefreshGC"
With .TextFrame.Characters.Font
.Color = vbWhite: .size = 14
End With
.TextFrame.VerticalAlignment = xlVAlignCenter:.TextFrame.HorizontalAlignment = xlHAlignCenter
With s.Fill
.Solid: .ForeColor.RGB = vbRed
End With
End With
End Sub
Sheet3
Option Explicit
Option Private Module
Private arrAllData()
Private arrAllDataSorted()

Sub ActivateGanttChart()
If GanttChart Then
Set gs = setGSws: Set rs = setRSws
Call CalcColPosGCT: Call CalcColPosTimeline: Call CalcColPosGST: Call ReadSettings: sArr.LoadAllArrays
End If
End Sub

Sub PrepAllGanttCharts(Optional bActivate As Boolean)
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
If GanttChart(ws) Then
ws.Activate: Call CalcColPosGCT: Call CalcColPosTimeline:
If bActivate Then
Call DeleteShape("SG_EnableMacrosButton", 21, ws):'Call DrawTasksBorders(, ws): Call DrawTimelineBorders(, ws)
Call DrawFilterIcon(ws): Call DrawTNavIcons(ws):
Call DrawFreezeIcon(ws): Call DrawRefreshButton(ws):
Call Freeze(ws, False): Call Freeze(ws, True)
Else
Call DeleteShape("SG_Filter", 9, ws): Call DeleteShape("ST_Freeze", 9, ws): 'Call ClearBorders(ws)
Call DeleteShape("ST_TSBackIcon", 13, ws): Call DeleteShape("ST_TEBackIcon", 13, ws): Call DeleteShape("ST_TSFrontIcon", 14, ws): Call DeleteShape("ST_TEFrontIcon", 14, ws)
Call DrawEnableMacrosButton(ws): Call Freeze(ws, False)
End If
End If
Next ws
End Sub

Sub createGCsheet(bPlanner As String)
Worksheets.Add: Call SetupGCColumns("GC"): Call SetupGCRowColWidth(bPlanner): Call HideColumns(True)
Call SetupGCFirstRow: Call SetupGCAlignmentnFormat:
Call DrawFilterIcon(ActiveSheet)
ActiveWindow.DisplayGridlines = False: ActiveWindow.DisplayHeadings = False
End Sub

Sub SetupGCColumns(wstype As String)
Dim i As Long, x As Long, cCol As Long: cCol = 1
If wstype = "GDD" Then x = 2
For i = 1 To UBound(GCcolumns())
If wstype = "GDD" Then
If GCcolumns(i) = "TColor" Or GCcolumns(i) = "TPColor" Or GCcolumns(i) = "BLColor" Or GCcolumns(i) = "ACColor" Or GCcolumns(i) = "ShapeInfoE" Or GCcolumns(i) = "ShapeInfoB" Or GCcolumns(i) = "ShapeInfoA" Then GoTo nexi
End If
Cells(rowone, cCol + x) = GCcolumns(i)
Cells(rowtwo, cCol + x) = GCcolumns(i)
Cells(rownine, cCol + x) = GCcolumnsEngName(i)
If Left(GCcolumns(i), 6) <> "Custom" Then Cells(rowtwo, cCol + x) = "System"
cCol = cCol + 1
nexi:
Next i
If wstype = "GC" Then
Cells(1, UBound(GCcolumns()) + 1) = "TimelineStart": Cells(1, UBound(GCcolumns()) + 6) = "TimelineEnd": Cells(1, UBound(GCcolumns()) + 7) = "LLC"
End If
If wstype = "GDD" Then
Cells(1, 1) = "ProjectID": Cells(1, 2) = "dType": Range("2:10").Clear
End If
If wstype = "GC" Then Call CalcColPosGCT: Call CalcColPosTimeline
If wstype = "GDD" Then Call CalcColPosGDD
End Sub

Sub SetupGCFirstRow()
Cells(firsttaskrow, cpg.GEtype) = "T": Cells(firsttaskrow, cpg.TID) = 1: Cells(firsttaskrow, cpg.WBS) = 1
Cells(firsttaskrow, cpg.Task) = "Type here or double click to edit in form": Cells(firsttaskrow, cpg.PercentageCompleted) = "0%"
End Sub

Sub SetupGCRowColWidth(bPlanner As String)
Rows(rowsix).RowHeight = 30: Rows(rownine).RowHeight = 24: Rows(firsttaskrow).RowHeight = taskRowHeight
Columns(cpg.Priority).ColumnWidth = 10: Columns(cpg.Status).ColumnWidth = 10:
Range(Cells(1, cpg.Resource), Cells(1, cpg.LC)).ColumnWidth = 15
Columns(cpg.SS).ColumnWidth = 2: Columns(cpg.TaskIcon).ColumnWidth = 2:
Columns(cpg.WBS).ColumnWidth = 6: Columns(cpg.Task).ColumnWidth = 45
Columns(cpg.Done).ColumnWidth = 5: Columns(cpg.LC).ColumnWidth = 2
Columns(cpg.PercentageCompleted).ColumnWidth = 10: Columns(cpg.ED).ColumnWidth = 10:
If bPlanner = "Hours" Then Columns(cpg.ESD).ColumnWidth = 21: Columns(cpg.EED).ColumnWidth = 21
End Sub

Sub SetupGCAlignmentnFormat()
With Range(Columns(cpg.GEtype), Columns(cpg.LC))
.HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Font.Color = rgbBlack: .Font.size = 10
End With
Columns(cpg.WBS).HorizontalAlignment = xlLeft
Columns(cpg.Task).HorizontalAlignment = xlLeft
Columns(cpg.Resource).HorizontalAlignment = xlLeft
Columns(cpg.BSD).HorizontalAlignment = xlRight: Columns(cpg.BED).HorizontalAlignment = xlRight: Columns(cpg.BD).HorizontalAlignment = xlRight
Columns(cpg.ESD).HorizontalAlignment = xlRight: Columns(cpg.EED).HorizontalAlignment = xlRight: Columns(cpg.ED).HorizontalAlignment = xlRight
Columns(cpg.ASD).HorizontalAlignment = xlRight: Columns(cpg.AED).HorizontalAlignment = xlRight: Columns(cpg.AD).HorizontalAlignment = xlRight
Columns(cpg.BCS).HorizontalAlignment = xlRight: Columns(cpg.ECS).HorizontalAlignment = xlRight: Columns(cpg.ACS).HorizontalAlignment = xlRight
Columns(cpg.Work).HorizontalAlignment = xlRight
Columns(cpg.WBSPredecessors).HorizontalAlignment = xlLeft: Columns(cpg.WBSSuccessors).HorizontalAlignment = xlLeft
Columns(cpg.ResourceCost).HorizontalAlignment = xlRight
With Columns(cpg.Priority)
.Font.size = 9: .Font.Bold = True:
End With
With Range(Cells(rowsix, cpg.WBS), Cells(rowsix, cpg.LC - 1))
.Font.size = 12: .Font.Bold = True: .Font.Color = rgbWhite: .HorizontalAlignment = xlLeft
End With

With Range(Cells(rowseven, cpg.WBS), Cells(rowseven, cpg.LC - 1))
.Font.size = 10: .Font.Color = rgbGray: .HorizontalAlignment = xlLeft
End With
With Range(Cells(roweight, cpg.WBS), Cells(roweight, cpg.LC - 1))
.Font.size = 10: .Font.Color = rgbGray: .HorizontalAlignment = xlLeft
End With
With Range(Cells(rownine, cpg.GEtype), Cells(rownine, cpg.LC))
.Font.size = 10: .HorizontalAlignment = xlCenter: .Font.Bold = False
End With
With Columns(cpg.PercentageCompleted)
.NumberFormat = "0%"
End With
With Columns(cpg.Notes)
.NumberFormat = "General"
End With
With Columns(cpg.ShapeInfoE)
.WrapText = True
End With
With Columns(cpg.ShapeInfoB)
.WrapText = True
End With
With Columns(cpg.ShapeInfoA)
.WrapText = True
End With

End Sub

Sub HideColumns(Optional show As Boolean)
ActiveSheet.Range(Cells(1, cpg.GEtype), Cells(1, cpg.TIL)).EntireColumn.Hidden = show
ActiveSheet.Range(Cells(1, cpg.GEtype), Cells(5, cpg.LC)).EntireRow.Hidden = show
Columns(cpg.Status).EntireColumn.Hidden = show: Columns(cpg.ResourceCost).EntireColumn.Hidden = show: Columns(cpg.Work).EntireColumn.Hidden = show
Columns(cpg.BSD).EntireColumn.Hidden = show: Columns(cpg.BED).EntireColumn.Hidden = show: Columns(cpg.BD).EntireColumn.Hidden = show
Columns(cpg.ASD).EntireColumn.Hidden = show: Columns(cpg.AED).EntireColumn.Hidden = show: Columns(cpg.AD).EntireColumn.Hidden = show
Columns(cpg.BCS).EntireColumn.Hidden = show: Columns(cpg.ACS).EntireColumn.Hidden = show: Columns(cpg.ECS).EntireColumn.Hidden = show
Columns(cpg.Notes).EntireColumn.Hidden = show: Columns(cpg.TColor).EntireColumn.Hidden = show: Columns(cpg.TPColor).EntireColumn.Hidden = show
Columns(cpg.BLColor).EntireColumn.Hidden = show: Columns(cpg.ACColor).EntireColumn.Hidden = show
Columns(cpg.WBSPredecessors).EntireColumn.Hidden = show: Columns(cpg.WBSSuccessors).EntireColumn.Hidden = show:
Columns(cpg.ShapeInfoE).EntireColumn.Hidden = show: Columns(cpg.ShapeInfoB).EntireColumn.Hidden = show: Columns(cpg.ShapeInfoA).EntireColumn.Hidden = show:
Range(Cells(1, cpg.Custom1), Cells(1, cpg.Custom20)).Columns.EntireColumn.Hidden = show
End Sub

Sub uhs()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
ws.visible = xlSheetVisible
Next ws
End Sub
Sub hs()

Dim ws As Worksheet
Call hideWorksheet(GST): Call hideWorksheet(GDT): Call hideWorksheet(GDD): Call hideWorksheet(PVS)
For Each ws In ThisWorkbook.Sheets
If Left(ws.Name, 2) = "GS" And ws.Cells(1, 1) = "GCTYPE" Then Call hideWorksheet(ws)
If Left(ws.Name, 2) = "RS" And ws.Cells(1, 1) = "Resource" Then Call hideWorksheet(ws)
Next ws
End Sub

Sub hideWorksheet(ws As Worksheet)
#If Mac Then
ws.visible = xlHidden
#Else
ws.visible = xlVeryHidden
#End If
End Sub

Sub urc()
If GanttChart = False Then Exit Sub
Range("G:G").ColumnWidth = 3: Range("B:F").ColumnWidth = 8: Range("1:5").RowHeight = taskRowHeight: Range("A:A").ColumnWidth = 6:
ActiveWindow.DisplayHeadings = True
End Sub
Sub hrc()
If GanttChart = False Then Exit Sub
Range("A:G").ColumnWidth = 0: Range("1:5").RowHeight = 0
Cells(firsttaskrow, cpg.Task).Select
ActiveWindow.DisplayHeadings = False
End Sub
Sub ugsrs()
Dim gcws As Worksheet, gsws As Worksheet, rsws As Worksheet
Call DA: Set gcws = ActiveSheet:
Set gsws = setGSws: Set rsws = setRSws
rsws.visible = True: gsws.visible = True: Call EA:gcws.Activate
End Sub
Sub hgsrs()
Dim gcws As Worksheet, gsws As Worksheet, rsws As Worksheet
Call DA: Set gcws = ActiveSheet:
Set gsws = setGSws: Set rsws = setRSws
rsws.visible = xlSheetVeryHidden: gsws.visible = xlSheetVeryHidden: Call EA:gcws.Activate
End Sub
Sub SetRowHeight(Optional t As Boolean)
sRowOrWidth = "R"
frmRowColumnHeightWidth.show
End Sub
Sub SetColumnWidth(Optional t As Boolean)
If Selection.Columns.Count > 1 Then
MsgBox "Select only one column to set the column width", vbInformation, "Information"
Exit Sub
End If
If Selection.column = 1 Then Exit Sub
If Selection.column >= cpt.TimelineStart Then
MsgBox "Column width for the timeline cannot be changed", vbInformation, "Information"
Exit Sub
End If
sRowOrWidth = "C"
frmRowColumnHeightWidth.show
End Sub

Sub EnableShortCutKeys(Optional t As Boolean)
If IsAnyOpenWorkbookProtected Then Exit Sub
On Error Resume Next
Application.OnKey "%{RIGHT}", "mMakeChild"
Application.OnKey "%{LEFT}", "mMakeParent"
Application.OnKey "%{UP}", "MoveTaskUp"
Application.OnKey "%{DOWN}", "MoveTaskDown"
Application.OnKey "%{RETURN}", "mEditTask"
Application.OnKey "{F7}", "DoSpellCheck"
Application.OnKey "+^L", "AddFilterToTasksTrigger"
On Error GoTo 0
End Sub
Sub DisableShortCutKeys(Optional t As Boolean)
On Error Resume Next
Application.OnKey "%{RIGHT}"
Application.OnKey "%{LEFT}"
Application.OnKey "%{UP}"
Application.OnKey "%{DOWN}"
Application.OnKey "%{RETURN}"
Application.OnKey "{F7}"
Application.OnKey "+^L"
On Error GoTo 0
End Sub
Sub EnableCtrlDCtrlRKeys(Optional t As Boolean)
If IsAnyOpenWorkbookProtected Then Exit Sub
On Error Resume Next
Application.OnKey "^d"
Application.OnKey "^D"
Application.OnKey "^r"
Application.OnKey "^R"
On Error GoTo 0
End Sub
Sub DisableCtrlDCtrlRKeys(Optional t As Boolean)
If IsAnyOpenWorkbookProtected Then Exit Sub
On Error Resume Next
Application.OnKey "^d", "DisableKeysDummy"
Application.OnKey "^D", "DisableKeysDummy"
Application.OnKey "^r", "DisableKeysDummy"
Application.OnKey "^R", "DisableKeysDummy"
On Error GoTo 0
End Sub

Sub TriggerWorkbookClose(Optional t As Boolean)
Set cps = Nothing
frmSaveClose.show
End Sub

Sub makeFree()
Call DA: GST.Cells(rowtwo, cps.tLiType) = pstrFree: Call EA: Call RefreshRibbon
End Sub

Sub daea()
Debug.Print "Application.ScreenUpdating: " & Application.ScreenUpdating
Debug.Print "Application.EnableEvents: " & Application.EnableEvents
Debug.Print "Application.DisplayAlerts: " & Application.DisplayAlerts
Debug.Print "Application.Calculation: " & Application.Calculation
End Sub

Sub chkThisWB()
If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate: currentSheet.Activate
End Sub

Sub PopTIL()
Dim cRow As Long:cRow = 10
Do While Cells(cRow, cpg.GEtype) <> vbNullString
Cells(cRow, cpg.TIL) = Cells(cRow, cpg.Task).IndentLevel:cRow = cRow + 1
Loop
End Sub

Sub IndentTaskfromTIL()
Dim cRow As Long:cRow = 10
Do While Cells(cRow, cpg.GEtype) <> vbNullString
Cells(cRow, cpg.Task).IndentLevel = Cells(cRow, cpg.TIL):cRow = cRow + 1
Loop
End Sub

Sub createRSsheet()
ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "newgc"
Cells(1, 1) = "Resource"
Cells(1, 2) = "Cost"
Cells(1, 3) = "Use Org Holidays"
Cells(1, 4) = "Department"
Cells(1, 11) = "Workhours Start": Cells(1, 12) = "Workhours End"
Cells(1, 13) = "WorkdaySun": Cells(1, 14) = "WorkdayMon": Cells(1, 15) = "WorkdayTue": Cells(1, 16) = "WorkdayWed": Cells(1, 17) = "WorkdayThu": Cells(1, 18) = "WorkdayFri": Cells(1, 19) = "WorkdaySat"
Cells(1, 20) = "Holidays"
Cells(2, 1) = "Organization"
Cells(2, 2) = ""
Cells(2, 3) = "FALSE"
Cells(2, 5) = "DO NOT MODIFY THIS SHEET OR INSERT COLUMNS": Cells(2, 5).Font.Color = rgbRed
Cells(2, 11) = "09:00 AM": Cells(2, 12) = "06:00 PM"
Cells(2, 13) = 0: Cells(2, 14) = 1: Cells(2, 15) = 1: Cells(2, 16) = 1: Cells(2, 17) = 1: Cells(2, 18) = 1: Cells(2, 19) = 0
End Sub

Sub ToggleFreeze()
With Application.ActiveWindow
If .FreezePanes = True Then .FreezePanes = False: Exit Sub
If .FreezePanes = False Then Cells(firsttaskrow, cpg.LC).Activate: .FreezePanes = True: Exit Sub
End With
End Sub

Sub Freeze(Optional ws As Worksheet, Optional bfreeze As Boolean)
If ws Is Nothing Then Set ws = ActiveSheet
ws.Activate
With Application.ActiveWindow
If bfreeze Then
Call CalcColPosGCT: ws.Cells(firsttaskrow, cpg.LC).Activate: .FreezePanes = bfreeze
Else
.FreezePanes = bfreeze
End If
End With
ws.Cells(firsttaskrow, cpg.Task).Select
End Sub

Sub MagicCommands(comm As String)
Select Case comm
Case Is = "rc"
If Columns("A").Hidden = True Then Call DA: Call urc: Call EA Else Call DA: Call hrc: Call EA
Case Is = "gsrs"
Set gs = setGSws:
If gs.visible = True Then Call hgsrs Else Call ugsrs
Case Is = "check"
Call DA: Call check: Call EA
Case Is = "da"
Call DA
Case Is = "debug"
If tlogg = True Then Call doff Else Call don
Case Is = "uhs"
Call uhs
Case Is = "hs"
Call hs
Case Is = "af"
bAllowFormulas = True
Case Is = "resetws"
Call resetws
Case Is = "resethelp"
Call resetHelp
Case Is = "resetall"
Call resetall
End Select
Last:
Cells(rownine, cpg.TaskIcon).Select
End Sub

Sub resetws()
Call reset("ws")
End Sub

Sub resetall()
Call reset("all")
End Sub

Sub reset(rtype As String)
Dim ws As Worksheet: Dim bHelpSheetExists As Boolean
Call DA
If rtype = "lic" Or rtype = "all" Then Call RemoveLicense
Call DA ' needed
If rtype = "ws" Or rtype = "all" Then
bHelpSheetExists = False
For Each ws In ThisWorkbook.Sheets
If ws.Name = "Help" Then bHelpSheetExists = True: GoTo helpsheetchecked
Next ws
helpsheetchecked:
If bHelpSheetExists Then Worksheets("Help").Activate Else Worksheets.Add
For Each ws In ThisWorkbook.Sheets
If ws.Name = GDT.Name Or ws.Name = GST.Name Or ws.Name = GDD.Name Or ws.Name = PVS.Name Or LCase(ws.Name) = "help" Then GoTo nexa
ws.visible = xlSheetVisible: ws.Delete 'del all other ws
nexa:
Next
GST.Cells(rowtwo, cps.SSN) = 0: Call hs
End If
If rtype = "help" Or rtype = "all" Then If bHelpSheetExists Then Call resetHelp
Call EA: Call RefreshRibbon
End Sub
Sub LockWB(lok As Boolean)
If lok Then
If IsWorkbookProtected = False Then ActiveWorkbook.Protect wbPass
Else
If IsWorkbookProtected = True Then ActiveWorkbook.Unprotect wbPass
End If
End Sub

Function IsWorkbookProtected() As Boolean
With ActiveWorkbook
If .ProtectWindows Or .ProtectStructure Then IsWorkbookProtected = True Else IsWorkbookProtected = False
End With
End Function

Sub ImportGanttChart()
bImportGC = True: frmImportGC.show: bImportGC = False
End Sub

Sub resetHelp()
Dim murl As String, durl As String, qurl As String
Worksheets("Help").Activate:
murl = siteURL & docURL & macroURL & strSource & "EnableMacros" & strMedium
durl = siteURL & docURL & strSource & "Doc" & strMedium
qurl = siteURL & howtoURL & strSource & "QSG" & strMedium
Worksheets("Help").Range("E11").Hyperlinks.Delete
Worksheets("Help").Hyperlinks.Add Anchor:=Range("E11"), Address:=murl, ScreenTip:="Enable Macros", TextToDisplay:="Click here for help to Enable Macros"
Worksheets("Help").Range("E22").Hyperlinks.Delete
Worksheets("Help").Hyperlinks.Add Anchor:=Range("E22"), Address:=qurl, ScreenTip:="Quick Start Guide", TextToDisplay:="Quick Start Guide"
Worksheets("Help").Range("E23").Hyperlinks.Delete
Worksheets("Help").Hyperlinks.Add Anchor:=Range("E23"), Address:=durl, ScreenTip:="Click Here For Documentation", TextToDisplay:="CLICK HERE FOR DOCUMENTATION"
Worksheets("Help").Range("D5").Select
End Sub
Option Explicit

Private Sub Workbook_Activate()
If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
#If Mac Then
#Else
Needed when you come from other workbook to gantt workbook
Call SetupRightClickGanttMenu: Call DisableCtrlDCtrlRKeys
#End If
End Sub

Private Sub Workbook_Deactivate()
#If Mac Then
#Else
Call DeleteRightClickGanttMenu: Call DisableShortCutKeys: Call EnableCtrlDCtrlRKeys'Needed when you you go to other workbook from gantt workbook
#End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call EnableCtrlDCtrlRKeys: Call DisableShortCutKeys
#If Mac Then
Call DA
Dim curWs As Worksheet: Set curWs = ActiveSheet: Call PrepAllGanttCharts(False): curWs.Activate:
Call EA
#Else
If bClosing Then
Else
Cancel = True
Call TriggerWorkbookClose
Exit Sub
End If
ThisWorkbook.Saved = True
bClosing = False
#End If
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Application.EnableEvents = False:GST.Cells(2, cps.SavedDate) = Date
With GST.Cells(rowtwo, cps.tFirstSavedDate)
If .value = vbNullString Then If FreeVersion Then .value = Date
End With
Application.EnableEvents = True
End Sub

Private Sub Workbook_Open()
On Error Resume Next
If Left$(DPB.Caption, 4) <> "DPB=" Then
MsgBox "This file has experienced a problem and must be closed!" & vbCrLf & "Please contact support with error code: 789", vbCritical
Application.DisplayAlerts = False: ThisWorkbook.Close False: Application.DisplayAlerts = True: Exit Sub
End If
On Error GoTo 0
If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
Call checkReqWorksheets
#If Mac Then
If Not FreeVersion And GST.Cells(2, cps.tLiType) <> pstrDPM And GST.Cells(2, cps.tLiType) <> pstrHPM And GST.Cells(2, cps.tLiType) <> pstrHDM And GST.Cells(2, cps.tLiType) <> vbNullString Then
MsgBox msg(83), vbInformation, "Gantt Excel"
Application.DisplayAlerts = False: ThisWorkbook.Close: Application.DisplayAlerts = True: Exit Sub
End If
If IsMac2011 Then MsgBox msg(35): Application.DisplayAlerts = False: ThisWorkbook.Close: Application.DisplayAlerts = True: Exit Sub
#Else
Call SetupRightClickGanttMenu: Call EnableShortCutKeys: Call DisableCtrlDCtrlRKeys
#End If
sTempStr = "OnStartUp"
If GST.Cells(rowtwo, cps.tliky) = vbNullString Or GST.Cells(rowtwo, cps.tliky) = "-" Then
GoTo Last:
Else
If GetLicType <> pstrFree Then frmAbout.show
If FreeVersion Then
Call PrepAllGanttCharts(True)
With GST.Cells(rowtwo, cps.tFirstSavedDate)
If .value = vbNullString Then frmWelcome.show Else OpenURLOnStartup
End With
End If
End If
Dim curWs As Worksheet: Set curWs = ActiveSheet
Call DA
Call uhs: Call hs ' Needs Research why
Call PrepAllGanttCharts(True)
sTempStr = vbNullString: curWs.Activate: Call ActivateGanttChart:
Last:
Call EA
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
Call ActivateGanttChart
Call RefreshRibbon
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
If Not GanttChart Then Exit Sub
bHideTime = False: dbcCol = Target.column: dbcRow = Target.Row
Dim lrow As Long: lrow = GetLastRow
If dbcRow = lrow + 1 And dbcCol = cpg.Task Then GoTo taskplaceholderok
If dbcRow < firsttaskrow Or dbcRow > lrow Then Cancel = True:Exit Sub
If dbcCol >= cpt.TimelineStart And dbcCol <= cpt.TimelineEnd Then
If st.HGC Then
Dim orgStartHrs As Double:
If st.CurrentView = "HH" Then
orgStartHrs = sArr.ResourceP(0, 10): Cells(dbcRow, cpg.ESD) = CDate(Cells(rowsix, dbcCol))
Else
orgStartHrs = sArr.ResourceP(0, 10): Cells(dbcRow, cpg.ESD) = CDate(Cells(rowsix, dbcCol)) + orgStartHrs
End If
Else
Cells(dbcRow, cpg.ESD) = CDate(Cells(rowsix, dbcCol))
End If
Cells(dbcRow, cpg.ED).Select: Exit Sub
End If
taskplaceholderok:
Cancel = True ' important
If parCheck(dbcRow, dbcCol, "dc") = True Then Call DA: Cells(dbcRow, cpg.Task).Select: GoTo Last
If dbcCol = cpg.Resource Then Call DA: ResSelectorDB = True: Call RememberResArrays: ResArraysReady = True: frmResourceSelector.show: Exit Sub
If dbcCol = cpg.ResourceCost And st.CalResCosts Then Call DA: Call myMsgBox(dbcRow, dbcCol, 8):GoTo Last
If dbcCol = cpg.PercentageCompleted And st.PercAuto Then Call DA: Call myMsgBox(dbcRow, dbcCol, 10): GoTo Last
If dbcCol = cpg.Priority Then Call DA: SelTaskRow = dbcRow:Call frmPriority.show: GoTo Last
If dbcCol = cpg.Done Then Call DA: SelTaskRow = dbcRow:Call CompletedAction: GoTo Last

If dbcCol = cpg.ESD Or dbcCol = cpg.BSD Or dbcCol = cpg.ASD Or dbcCol = cpg.EED Or dbcCol = cpg.BED Or dbcCol = cpg.AED And Cells(dbcRow, 1) <> vbNullString Then
Call DA: SelTaskRow = dbcRow: Call RememberResArrays: ResArraysReady = True:
Select Case dbcCol
Case Is = cpg.ESD
bESD = True: StartDateCheck = True
Case Is = cpg.EED
bEED = True: Cells(dbcRow, cpg.ED).Select
Case Is = cpg.BSD
bBSD = True: StartDateCheck = True
Case Is = cpg.ASD
bASD = True: StartDateCheck = True
Case Is = cpg.BED
bBED = True: Cells(dbcRow, cpg.BD).Select
If Cells(dbcRow, cpg.BSD) = "" Then Call myMsgBox(dbcRow, dbcCol, 25): Call EA: Exit Sub
Case Is = cpg.AED
bAED = True: Cells(dbcRow, cpg.AD).Select
If Cells(dbcRow, cpg.ASD) = "" Then Call myMsgBox(dbcRow, dbcCol, 26): Call EA: Exit Sub
End Select
frmDateSelector.show: Call EA
Else
Call DA: Call LoadFormOnDblClick: GoTo Last
End If
Last:
Call EA
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
Dim cRow As Long, cCol As Long, lrow As Long: lrow = GetLastRow
If Sh.Cells(1, 1) <> "GEType" Then Exit Sub
On Error Resume Next
ActiveSheet.Shapes("S_W").Delete
On Error GoTo 0
If Target.Row = rownine And Target.column = cpg.SS Then frmMagic.show: Exit Sub
If ((Target.Address = Target.EntireRow.Address Or Target.Address = Target.EntireColumn.Address)) Then MsgBox msg(36): Exit Sub 'nopasstech
If Target.column < cpg.SS Or Target.Row < 3 Then MsgBox msg(37): Cells(firsttaskrow, cpg.Task).Select: Exit Sub
If Target.column > cpt.TimelineEnd Then Exit Sub
If Target.Row = lrow + 1 And Target.column = cpg.Task Then GoTo approved
If Target.Row > lrow Then
If Target.Row = lrow + 1 Then Rows(Target.Row).Borders(xlEdgeBottom).LineStyle = xlNone: Exit Sub
If Target.Row > lrow + 1 Then Rows(Target.Row).Borders(xlEdgeTop).LineStyle = xlNone: Rows(Target.Row).Borders(xlEdgeBottom).LineStyle = xlNone: Exit Sub
End If

approved:
If Selection.CountLarge > 1 Then Exit Sub
'If Not Intersect(Target, Range("J6:L8")) Is Nothing Then Call LoadNewGanttFormOnDblClick: Call EA
If Target.Row >= rowsix And Target.Row <= roweight Then
If Target.column >= cpg.WBS And Target.column <= cpg.Task + 2 Then Call LoadNewGanttFormOnDblClick: Call EA
End If
If Intersect(Target, Range(Cells(firsttaskrow, cpg.SS), Cells(GetLastRow + 1, cpt.TimelineEnd))) Is Nothing Then Exit Sub
'check if mac and win are same here
If Target.column >= cpg.SS And Target.column <= cpt.TimelineEnd Then
If Application.CutCopyMode = False Then
If SelectedRow <> 0 And SelectedRow <> Target.Row Then Call DA: Call UnHiliteRow(SelectedRow): Call EA
SelectedRow = Target.Row: Call DA: Call HiliteRow(Target.Row):Call EA
End If
End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
If DashboardSheet(Sh) Then If Target.Address = "$B$2" Then Call DA: Call RefreshDashboard: Call EA:Exit Sub
If Sh.Cells(1, 1) <> "GEType" Then Exit Sub
'nopasstech
If ((Target.Address = Target.EntireRow.Address Or Target.Address = Target.EntireColumn.Address)) Then
With Application
.EnableEvents = False:.Undo:MsgBox msg(36):.EnableEvents = True
End With
Exit Sub
End If
If IsLicValid(0, 1) = False Then Exit Sub
If Target.column > cpg.LC Then MsgBox msg(38)
If Target.Row < rownine Then Exit Sub
Call TriggerCellValueChanged(Target)
End Sub
Option Explicit
Option Private Module

Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#End If

Public Function TimeToMillisecond() As String
Dim tSystem As SYSTEMTIME:Dim sRet
On Error Resume Next
GetSystemTime tSystem
sRet = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & _
"." & Format(tSystem.wMilliseconds, "000")
TimeToMillisecond = sRet
End Function

Public Function tlog(s As String)
If tlogg = False Then Exit Function
Dim cws As Worksheet: Set cws = ActiveSheet
If s = "UnHiliteRow" Then Exit Function

Dim wb As Workbook: Dim i As Long: Dim foundDebugFile As Boolean: Dim milli As String
For Each wb In Application.Workbooks
If wb.Windows(1).visible Then
If wb.Name = ThisWorkbook.Name Then GoTo nextwb
If wb.Name = "debug.xlsm" Then foundDebugFile = True
End If
nextwb:
Next

milli = TimeToMillisecond
If foundDebugFile = False Then Exit Function
Dim dwb As Workbook: Set dwb = Workbooks("debug.xlsm"): Dim wss As Worksheet: Set wss = dwb.Worksheets("TimeLog")
Dim noofitems As Long:noofitems = Application.WorksheetFunction.CountA(wss.Range("A:A"))
wss.Cells(noofitems + 1, 1) = s:
With wss.Cells(noofitems + 1, 2)
.value = milli
.NumberFormat = "hh:mm:ss.000"
End With
ThisWorkbook.Activate: cws.Activate

End Function

Public Function dlog(s As String)
If dlogg = False Then Exit Function
Dim cws As Worksheet: Set cws = ActiveSheet
If s = "UnHiliteRow" Then Exit Function

Dim wb As Workbook: Dim i As Long: Dim foundDebugFile As Boolean: Dim milli As String:
For Each wb In Application.Workbooks
If wb.Windows(1).visible Then
If wb.Name = ThisWorkbook.Name Then GoTo nextwb
If wb.Name = "debug.xlsm" Then foundDebugFile = True
End If
nextwb:
Next
If foundDebugFile = False Then Exit Function
Dim dwb As Workbook: Set dwb = Workbooks("debug.xlsm"): Dim ws As Worksheet: Set ws = Workbooks("debug.xlsm").Worksheets("Analysis")
Dim noofitems As Long:noofitems = Application.WorksheetFunction.CountA(ws.Range("A:A"))
ws.Cells(noofitems + 1, 1) = s:
With ws.Cells(noofitems + 1, 1)
.Font.Color = rgbDeepPink
End With
ThisWorkbook.Activate: cws.Activate
End Function

Sub don()
tlogg = True: dlogg = True
End Sub
Sub doff()
tlogg = False: dlogg = False
End Sub
