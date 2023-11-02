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
