Attribute VB_Name = "results"
Option Explicit
Option Base 1

Sub pullResults()
    'ByRef rob As IRobotApplication, ByRef str As IRobotStructure
    MsgBox "The API will extract forces from " & " members. This might take some time", vbOKOnly, "Babala!"

    Dim Res As IRobotResultQueryReturnType
    Dim RobResQueryParams As RobotResultQueryParams
    Dim RobResRowSet As New RobotResultRowSet
    
    Dim membSel As IRobotSelection
    Dim ulsCases As IRobotSelection
    Dim storeySel As IRobotSelection
    Dim storey As IRobotStorey
    Dim selBarNames As String
    
    Set storey = utils.get_storey("2F")
    Set storeySel = utils.get_by_storey(storey)
    selBarNames = storeySel.ToText
    Debug.Print "Selected Bars: " & selBarNames
    Set membSel = utils.get_bars(selBarNames)
    Set ulsCases = utils.get_cases("1101to1149")
    
    Set RobResQueryParams = rob.CmpntFactory.Create(I_CT_RESULT_QUERY_PARAMS)
    RobResQueryParams.Selection.Set I_OT_BAR, membSel
    RobResQueryParams.Selection.Set I_OT_CASE, ulsCases
    RobResQueryParams.SetParam I_RPT_BAR_ELEMENT_DIV_COUNT, 11

    RobResQueryParams.ResultIds.SetSize (3) ' ommitted some parts that are not needed at this point; see numbering below
    'RobResQueryParams.ResultIds.Set 1, I_EVT_FORCE_BAR_FX
    'RobResQueryParams.ResultIds.Set 2, I_EVT_FORCE_BAR_FY
    RobResQueryParams.ResultIds.Set 1, I_EVT_FORCE_BAR_FZ
    'RobResQueryParams.ResultIds.Set 4, I_EVT_FORCE_BAR_MX
    RobResQueryParams.ResultIds.Set 2, I_EVT_FORCE_BAR_MY
    RobResQueryParams.ResultIds.Set 3, I_EVT_FORCE_BAR_MZ
    
    Dim resultsCount As Long
    Dim count As Long
    Dim reset As Boolean
    
    Do
        Res = struc.results.Query(RobResQueryParams, RobResRowSet)
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
        ReDim forces(resultsCount, 0 To 2)
        For count = 1 To resultsCount
            If reset = False Then
                Exit For
            End If
            'forces(count, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_BAR)
            'forces(count, 1) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(1)) / 1000
            'forces(count, 2) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(2)) / 1000
            'forces(count, 3) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(3)) / 1000
            'forces(count, 4) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(4)) / 1000
            'forces(count, 5) = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(5)) / 1000
            'forces(count, 6) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            'forces(count, 7) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE_CMPNT)
            reset = RobResRowSet.MoveNext()
        Next count
    End If
    Debug.Print "Total Extracted Results: " & resultsCount - 1

    release RobResQueryParams
    release struc
    release rob
End Sub


Sub get_beam_forces(ByVal selDesc As String, Optional ByVal selectedCases As String = "1")
    
    Dim ulsCases As IRobotSelection
    
    Dim membNo As Long
    Dim caseNo As Long
    Dim totalCount As Long
    Dim forces() As Double
    
    Set membSel = str.Selections.Create(I_OT_BAR)
    membSel.AddText (selDesc)
    
    totalCount = membSel.count
    
    'matic selection based on this model ONLY
    Set ulsCases = str.Selections.Create(I_OT_CASE)
    ulsCases.AddText (selectedCases)
    
    'Set ulsCases = Nothing
    'Set membSel = Nothing

End Sub

