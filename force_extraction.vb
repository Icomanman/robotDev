
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
    ulsCases.AddText ("1001")
    
    Dim Res As IRobotResultQueryReturnType
    Dim RobResQueryParams As RobotResultQueryParams
    Dim RobResRowSet As New RobotResultRowSet

    Set RobResQueryParams = rob.CmpntFactory.Create(I_CT_RESULT_QUERY_PARAMS)
    RobResQueryParams.Selection.Set I_OT_BAR, membSel
    RobResQueryParams.Selection.Set I_OT_CASE, ulsCases
    RobResQueryParams.SetParam I_RPT_BAR_ELEMENT_DIV_COUNT, 11

    RobResQueryParams.ResultIds.SetSize (6)

    
    RobResQueryParams.ResultIds.Set 1, I_EVT_FORCE_BAR_FX
    RobResQueryParams.ResultIds.Set 2, I_EVT_FORCE_BAR_FY
    RobResQueryParams.ResultIds.Set 3, I_EVT_FORCE_BAR_FZ
    'RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MX
    RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MY
    RobResQueryParams.ResultIds.Set 5, I_EVT_FORCE_BAR_MZ
    
    Dim resultsCount As Integer
    Dim count As Integer
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
        ReDim forces(resultsCount, 0 To 5)
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
            reset = RobResRowSet.MoveNext()
            
        Next count
    End If
    
    Set dataRange = Range(origin, Cells(origin.Row + resultsCount - 1, 6))
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


Private Sub CommandButton2_Click()

    Dim RobApp As New RobotOM.RobotApplication

    Dim Res As IRobotResultQueryReturnType
    Dim RobResQueryParams As RobotResultQueryParams
    Dim RobResRowSet As New RobotResultRowSet

    Dim SelBar As RobotSelection
    Dim SelCas As RobotSelection

    Set SelBar = RobApp.Project.Structure.Selections.Create(I_OT_BAR)
    Set SelCas = RobApp.Project.Structure.Selections.Create(I_OT_CASE)
    SelBar.AddText ("1")
    SelCas.AddText ("1")

    Set RobResQueryParams = RobApp.CmpntFactory.Create(I_CT_RESULT_QUERY_PARAMS)
    RobResQueryParams.Selection.Set I_OT_BAR, SelBar
    RobResQueryParams.Selection.Set I_OT_CASE, SelCas
    RobResQueryParams.SetParam I_RPT_BAR_ELEMENT_DIV_COUNT, 11

    RobResQueryParams.ResultIds.SetSize (6)

    RobResQueryParams.ResultIds.Set 1, I_EVT_FORCE_BAR_FX
    RobResQueryParams.ResultIds.Set 2, I_EVT_FORCE_BAR_FY
    RobResQueryParams.ResultIds.Set 3, I_EVT_FORCE_BAR_FZ
    RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MX
    RobResQueryParams.ResultIds.Set 5, I_EVT_FORCE_BAR_MY
    RobResQueryParams.ResultIds.Set 6, I_EVT_FORCE_BAR_MZ

    Dim v As Double

    Dim max As Double
    Dim min As Double
    max = 0
    min = 0

    Do
        Res = RobApp.Project.Structure.Results.Query(RobResQueryParams, RobResRowSet)
        Dim ok As Boolean
        ok = RobResRowSet.MoveFirst()

    While ok
        v = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(5))
        If (min > v) Then min = v
        If (max < v) Then max = v
        ok = RobResRowSet.MoveNext()
    Wend
    
    Loop While Res = I_RQRT_MORE_AVAILABLE

    MsgBox ("min = " & min & " max = " & max)
End Sub
