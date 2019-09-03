

Sub clr_range()

    Range("A6:G6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Value = ""
    Range("A6").Select
    
End Sub

Sub papaya_extract()

    Dim rob As IRobotApplication
    Dim str As IRobotStructure
   
    Dim panSel As IRobotSelection
    Dim caseSel As IRobotSelection
    Dim ulsCases As IRobotSelection
    
    Dim origin As Range
    Dim dataRange As Range
    
    Dim panelNo As Long
    Dim caseNo As Long
    Dim totalCount As Long
    Dim forces() As Double
    Dim bool As Long
    
    Set rob = New RobotApplication

    If Not rob.Visible Then
        MsgBox "No instance of Robot was found.", vbCritical, "Robot API"
        Debug.Print "Robot Flag: " & rob.Visible
        Debug.Print "No instance of Robot was found. Exiting..."
        Exit Sub
    Else
        If Not I_RST_AVAILABLE = 1 Then
            MsgBox "Results are not available. Please run you calculations.", vbCritical, "Bobot"
            Exit Sub
        End If
    End If
    Set str = rob.Project.Structure
    Set panSel = str.Selections.Get(I_OT_PANEL)
    Debug.Print "Selected Panels: " & panSel.count
    If panSel.count < 1 Then
        MsgBox "Kulang ka ba sa puti?", vbCritical, "No Panel Element selected"
        Exit Sub
    End If
    
    ' Manual selection of load cases:
    'Set caseSel = str.Selections.Get(I_OT_CASE)
    
    'matic selection based on this model ONLY
    Set ulsCases = str.Selections.Create(I_OT_CASE)
    ulsCases.AddText ("1001to1149")
    'ulsCases.AddText ("11")
    
    Dim Res As IRobotResultQueryReturnType
    Dim RobResQueryParams As RobotResultQueryParams
    Dim RobResRowSet As RobotResultRowSet
    Dim vecDir(2) As Double
    vecDir(2) = 1

    Set RobResQueryParams = rob.CmpntFactory.Create(I_CT_RESULT_QUERY_PARAMS)
    RobResQueryParams.Selection.Set I_OT_PANEL, panSel
    RobResQueryParams.Selection.Set I_OT_CASE, ulsCases

    Debug.Print "Selected Panel ID(s): " & panSel.ToText
    Debug.Print "Forces will be based on global vector: " _
    & vecDir(0) & ", " & vecDir(1) & ", " & vecDir(2)
    RobResQueryParams.SetParam I_RPT_MAX_BUFFER_SIZE, 2000000
    RobResQueryParams.SetParam I_RPT_SMOOTHING, I_FRS_SMOOTHING_WITHIN_A_PANEL
    RobResQueryParams.SetParam I_RPT_LAYER, I_FLT_ABSOLUTE_MAXIMUM
    RobResQueryParams.SetParam I_RPT_DIR_X_DEFTYPE, I_OLXDDT_CARTESIAN
    RobResQueryParams.SetParam I_RPT_DIR_X, vecDir

    RobResQueryParams.ResultIds.SetSize (4)
    RobResQueryParams.ResultIds.Set 1, I_FRT_DETAILED_NXX '492
    RobResQueryParams.ResultIds.Set 2, I_FRT_DETAILED_NYY '493
    RobResQueryParams.ResultIds.Set 3, I_FRT_DETAILED_NXY '494
    RobResQueryParams.ResultIds.Set 4, I_FRT_DETAILED_MXX '501
    
    Dim NXX As Double
    Dim NYY As Double
    Dim NXY As Double
    Dim MXX As Double
    
    Dim xF(7, 5) As Double
    Dim i As Integer
    Dim j As Integer

    ' Initialise values to 0:
    For i = 0 To 7
        For j = 1 To 4
            xF(i, j) = 0.01
        Next j
    Next i

    str.Results.Any.SetDirX I_OLXDDT_CARTESIAN, 1, 0, 0

    Do
        Set RobResRowSet = New RobotResultRowSet
        Res = str.Results.Query(RobResQueryParams, RobResRowSet)
        starter = RobResRowSet.MoveFirst()
        If starter = False Then
            Debug.Print "Early exit...results row set failed"
            MsgBox "Early exit...results row set failed", vbCritical, "Bobot"
            Exit Sub
        End If
        
        While starter
            NXX = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(1)) / 1000
            NYY = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(2)) / 1000
            NXY = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(3)) / 1000
            MXX = RobResRowSet.CurrentRow.GetValue(RobResRowSet.ResultIds.Get(4)) / 1000

        ' Set proper values:
            ' NXX Min
            If xF(0, 1) > NXX Then
                xF(0, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(0, 1) = NXX
                xF(0, 2) = NYY
                xF(0, 3) = NXY
                xF(0, 4) = MXX
                xF(0, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' NXX Max
            If xF(1, 1) < NXX Then
                xF(1, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(1, 1) = NXX
                xF(1, 2) = NYY
                xF(1, 3) = NXY
                xF(1, 4) = MXX
                xF(1, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' NYY Min
            If xF(2, 2) > NYY Then
                xF(2, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(2, 1) = NXX
                xF(2, 2) = NYY
                xF(2, 3) = NXY
                xF(2, 4) = MXX
                xF(2, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' NYY Max
            If xF(3, 2) < NYY Then
                xF(3, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(3, 1) = NXX
                xF(3, 2) = NYY
                xF(3, 3) = NXY
                xF(3, 4) = MXX
                xF(3, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' NXY Min
            If xF(4, 3) > NXY Then
                xF(4, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(4, 1) = NXX
                xF(4, 2) = NYY
                xF(4, 3) = NXY
                xF(4, 4) = MXX
                xF(4, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' NXY Max
            If xF(5, 3) < NXY Then
                xF(5, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(5, 1) = NXX
                xF(5, 2) = NYY
                xF(5, 3) = NXY
                xF(5, 4) = MXX
                xF(5, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' MXX Min
            If xF(6, 4) > MXX Then
                xF(6, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(6, 1) = NXX
                xF(6, 2) = NYY
                xF(6, 3) = NXY
                xF(6, 4) = MXX
                xF(6, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
            ' MXX Max
            If xF(7, 4) < MXX Then
                xF(7, 0) = RobResRowSet.CurrentRow.GetParam(I_RPT_PANEL)
                xF(7, 1) = NXX
                xF(7, 2) = NYY
                xF(7, 3) = NXY
                xF(7, 4) = MXX
                xF(7, 5) = RobResRowSet.CurrentRow.GetParam(I_RPT_LOAD_CASE)
            End If
        starter = RobResRowSet.MoveNext()
        Wend

    Loop While Res = I_RQRT_MORE_AVAILABLE
    
    Set origin = Range("I6")
    Set dataRange = Range(origin, Range("N13"))
    dataRange.Value = xF
 
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

    Range("I6: N13").Value = ""

End Sub


Sub papaya_holder()
    
    MsgBox "Currently not working; function shifted to 'Extract Summary'", vbInformation, "Papaya!"

End Sub
