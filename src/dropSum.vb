
Sub screen_cap()
'https://forums.autodesk.com/t5/robot-structural-analysis-forum/api-screen-capture-of-results-and-save-in-jpeg/td-p/8309084
Set RobApp = New RobotApplication

Dim mavueRobot As IRobotView3
Set mavueRobot = RobApp.Project.ViewMngr.GetView(1)
Dim ScPar As RobotViewScreenCaptureParams
Set ScPar = RobApp.CmpntFactory.Create(I_CT_VIEW_SCREEN_CAPTURE_PARAMS)

ScPar.Name = "capture"
ScPar.UpdateType = I_SCUT_CURRENT_VIEW
ScPar.Resolution = I_VSCR_4096
ScPar.UpdateType = I_SCUT_COPY_TO_CLIPBOARD
mavueRobot.MakeScreenCapture ScPar

ActiveSheet.Range("A1").Select
ActiveSheet.Paste

End Sub

Sub clr_range()

    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Value = ""
    Range("A6").Select
    
End Sub

Sub papaya_extract()

    Dim rob As IRobotApplication
    Dim str As IRobotStructure
   
    Dim membSel As IRobotSelection
    Dim caseSel As IRobotSelection
    Dim ulsCases As IRobotSelection
    
    Dim origin As Range
    Dim dataRange As Range
    Set origin = Range("A6")
    
    Dim membNo As Long
    Dim caseNo As Long
    Dim totalCount As Long
    Dim forces() As Double
     
    Set rob = New RobotApplication
    Set str = rob.Project.Structure
    Set membSel = str.Selections.Get(I_OT_BAR)
    
    If membSel.count < 1 Then
        MsgBox "Kulang ka ba sa puti?", vbCritical, "No Bar Element selected"
        Exit Sub
    End If
    
    ' Manual selection of load cases:
    Set caseSel = str.Selections.Get(I_OT_CASE)
    
    'matic selection based on this model ONLY
    Set ulsCases = str.Selections.Create(I_OT_CASE)
    ulsCases.AddText ("1001to1149")
    
    Dim Res As IRobotResultQueryReturnType
    Dim RobResQueryParams As RobotResultQueryParams
    Dim RobResRowSet As New RobotResultRowSet

    Set RobResQueryParams = rob.CmpntFactory.Create(I_CT_RESULT_QUERY_PARAMS)
    RobResQueryParams.Selection.Set I_OT_BAR, membSel
    RobResQueryParams.Selection.Set I_OT_CASE, ulsCases
    RobResQueryParams.SetParam I_RPT_BAR_ELEMENT_DIV_COUNT, 11

    RobResQueryParams.ResultIds.SetSize (5)
    RobResQueryParams.ResultIds.Set 1, I_EVT_FORCE_BAR_FX
    RobResQueryParams.ResultIds.Set 2, I_EVT_FORCE_BAR_FY
    RobResQueryParams.ResultIds.Set 3, I_EVT_FORCE_BAR_FZ
    'RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MX
    RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MY
    RobResQueryParams.ResultIds.Set 5, I_EVT_FORCE_BAR_MZ
    
    Dim resultsCount As Long
    Dim count As Long
    Dim reset As Boolean
    
    Do
        Res = str.Results.Query(RobResQueryParams, RobResRowSet)
        Dim ok As Boolean
        ok = RobResRowSet.MoveFirst()
        resultsCount = 0
    While ok
        resultsCount = resultsCount + 1
        ok = RobResRowSet.MoveNext()
    Wend
    
    Loop While Res = I_RQRT_MORE_AVAILABLE
    
    reset = RobResRowSet.MoveFirst()
    
    If reset = True Then
        ReDim forces(resultsCount, 0 To 6)
        For count = 0 To resultsCount
            If reset = False Then
                Exit For
            End If
            forces(count, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_BAR_ELEMENT)
            forces(count, 1) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(1)) / 1000
            forces(count, 2) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(2)) / 1000
            forces(count, 3) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(3)) / 1000
            forces(count, 4) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(4)) / 1000
            forces(count, 5) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(5)) / 1000
            forces(count, 6) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            'forces(count, 7) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE_CMPNT)
            reset = RobResRowSet.MoveNext()
            
        Next count
    End If
    
    Set dataRange = Range(origin, Cells(origin.Row + resultsCount - 1, 7))
    dataRange.Value = forces
     
    Set RobResQueryParams = Nothing
    Set ulsCases = Nothing
    Set caseSel = Nothing
    Set membSel = Nothing
    Set str = Nothing
    Set rob = Nothing
    Set dataRange = Nothing
    Set origin = Nothing
 
End Sub


Function verify_count(ByVal count As Long)
    
    Dim msg As String
    
    If count = Range("J17").Value Then
        msg = "Data range is verified. Ok."
    Else
        msg = "Error. Please Check Data Range. IGOT!!!"
        MsgBox msg
    End If
    
End Function


Sub clear_output()

    Range("J6: N15").Value = ""

End Sub

Sub Summary()
    
    Dim dForcesArr(9, 4) As Double
    Dim pVal As Object
    Dim rawData() As Variant

    Dim WF As Object
    Set WF = WorksheetFunction
    
    Dim dataOrigin, dataRange, outputRange As Range
    Dim pCol, v2Col, v3Col, m2Col, m3Col As Range
    Dim i, j, k, rowCount As Long
    
    Set dataOrigin = Cells(6, 1)
    If IsEmpty(dataOrigin) Then
        MsgBox "Please provide results first", vbCritical, "Bungks!"
        Exit Sub
    End If
    
    rowCount = Range(dataOrigin, dataOrigin.End(xlDown)).count
    verify_count (rowCount)
    
    Set dataRange = Range(Range("A6"), Cells(6 + rowCount - 1, 10))
    Set outputRange = Range("j6: n15")
    
    ReDim rawData(rowCount, dataRange.Columns.count)
    rawData = dataRange.Value
    
    dForcesArr(0, 0) = WF.min(dataRange.Columns(2))
    dForcesArr(1, 0) = WF.max(dataRange.Columns(2))
    dForcesArr(2, 1) = WF.min(dataRange.Columns(3))
    dForcesArr(3, 1) = WF.max(dataRange.Columns(3))
    dForcesArr(4, 2) = WF.min(dataRange.Columns(4))
    dForcesArr(5, 2) = WF.max(dataRange.Columns(4))
    dForcesArr(6, 3) = WF.min(dataRange.Columns(5))
    dForcesArr(7, 3) = WF.max(dataRange.Columns(5))
    dForcesArr(8, 4) = WF.min(dataRange.Columns(6))
    dForcesArr(9, 4) = WF.max(dataRange.Columns(6))
    
    For i = 0 To rowCount - 1
        For j = 0 To 1
            'pMin and pMax
            If dForcesArr(j, 0) = dataRange(i, 2) Then
                dForcesArr(j, j + 1) = dataRange(i, 3 + j)
                dForcesArr(j, j + 3) = dataRange(i, 5 + j)
            End If
            'v2Min and v2Max
            If dForcesArr(j + 2, 1) = dataRange(i, 3) Then
                dForcesArr(j + 2, 0) = dataRange(i, 2)
                dForcesArr(j + 2, 2) = dataRange(i, 4)
                dForcesArr(j + 2, 3) = dataRange(i, 5)
                dForcesArr(j + 2, 4) = dataRange(i, 6)
            End If
            'v3Min and v3Max
            If dForcesArr(j + 4, 2) = dataRange(i, 4) Then
                dForcesArr(j + 4, 0) = dataRange(i, 2)
                dForcesArr(j + 4, 1) = dataRange(i, 3)
                dForcesArr(j + 4, 3) = dataRange(i, 5)
                dForcesArr(j + 4, 4) = dataRange(i, 6)
            End If
            'm2Min and m2Max
            If dForcesArr(j + 6, 3) = dataRange(i, 5) Then
                dForcesArr(j + 6, 0) = dataRange(i, 2)
                dForcesArr(j + 6, 1) = dataRange(i, 3)
                dForcesArr(j + 6, 2) = dataRange(i, 4)
                dForcesArr(j + 6, 4) = dataRange(i, 6)
            End If
            'm3Min and m3Max
            If dForcesArr(j + 8, 4) = dataRange(i, 6) Then
                dForcesArr(j + 8, 0) = dataRange(i, 2)
                dForcesArr(j + 8, 1) = dataRange(i, 3)
                dForcesArr(j + 8, 2) = dataRange(i, 4)
                dForcesArr(j + 8, 3) = dataRange(i, 5)
            End If
        Next j
    Next i
    
    txt = WF.max(dataRange.Columns(5))
    
    outputRange = dForcesArr

    Set WF = Nothing
End Sub


