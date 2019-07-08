
' Note: Object types TOO LONG; it can be shortened (aliased) but Autodesk has reasons they put them that way

Sub setBeamSec(ByRef shape As IRobotBarSectionData, ByRef widthType As IRobotBarSectionConcreteDataValue, ByRef depthType As IRobotBarSectionConcreteDataValue)
    ' Setting Shape Type to Beam
    shape.shapeType() = I_BSST_CONCR_BEAM_RECT 'Code: -3
    widthType = I_BSCDV_BEAM_B 'Code: 1
    depthType = I_BSCDV_BEAM_H 'Code: 0
End Sub

Sub setColSec(ByRef shapeType As IRobotBarSectionData, ByRef widthType As IRobotBarSectionConcreteDataValue, ByRef depthType As IRobotBarSectionConcreteDataValue)
    ' Setting Shape Type to Column
    shapeType.shapeType() = I_BSST_CONCR_COL_R 'Code: -108
    widthType = I_BSCDV_COL_B 'Code: 1
    depthType = I_BSCDV_COL_H 'Code: 0
End Sub

Sub extractData(ByRef data As range) 'ByRef range As range,
    On Error Resume Next
    Set data = Application.InputBox("Please select the data range.", "Data Range", "B1", Type:=8)
    If data Is Nothing Then
        MsgBox "No input was selected!"
        Exit Sub
    End If
End Sub

Sub addSections()

    ' Validation before proceeding to Robot
    Dim dbRange As range
    Call extractData(dbRange)
    
    If dbRange Is Nothing Then
        Exit Sub
    Else
        Dim rob As IRobotApplication
        Dim botLabServ As IRobotLabelServer
        Dim barSec As IRobotLabel
        Dim barDat As IRobotBarSectionData
        Dim concThk As IRobotBarSectionConcreteData
    
        Dim widthType As IRobotBarSectionConcreteDataValue
        Dim depthType As IRobotBarSectionConcreteDataValue
    
        'Spreadsheet Data
        Dim dataArr() As Variant
        Dim iColCount As Integer
        Dim iRowCount As Integer
        Dim i As Integer
        Dim secName As String
        iColCount = dbRange.Columns.Count
        iRowCount = dbRange.Rows.Count
    
        ReDim dataArr(iRowCount, iColCount)
        dataArr = dbRange.Value
        
        Set rob = New RobotApplication
        Set botLabServ = rob.Project.Structure.Labels
        
        For i = 1 To iRowCount
            secName = dataArr(i, 1)
            Set barSec = botLabServ.Create(I_LT_BAR_SECTION, secName)
            Set barDat = barSec.data
          
            'Setting Material Type (has to be available within the Robot Database)
            barDat.MaterialName = "FC27"
            
            'Check if the section is a Beam or a Girder; defaults to Column section
            If InStr(1, secName, "B", vbBinaryCompare) <> 0 Then
                Call setBeamSec(barDat, widthType, depthType)
                Set concThk = barDat.Concrete
                concThk.SetReduction True, 0.35, 0.35, 1 'checked/unchecked, Ix, Iy, Iz
            ElseIf InStr(1, secName, "G", vbBinaryCompare) <> 0 Then
                Call setBeamSec(barDat, widthType, depthType)
                Set concThk = barDat.Concrete
                concThk.SetReduction True, 0.35, 0.35, 1 'checked/unchecked, Ix, Iy, Iz
            Else
                Call setColSec(barDat, widthType, depthType)
                Set concThk = barDat.Concrete
                concThk.SetReduction True, 0.7, 0.7, 1 'checked/unchecked, Ix, Iy, Iz
            End If
            Debug.Print "Stop"
            If CStr(dataArr(i, 5)) = "Rectangular" Then 'Section Shape
                concThk.SetValue depthType, dataArr(i, 7)
                concThk.SetValue widthType, dataArr(i, 9)
            Else
                'Assuming all circular sections are columns
                barDat.shapeType() = I_BSST_CONCR_COL_C
                ' Set reductions again; shape type is overidden
                ' and it seems all other values are reset also
                concThk.SetReduction True, 0.7, 0.7, 1
                concThk.SetValue I_BSCDV_COL_DE, dataArr(i, 7)
            End If
            
            barDat.CalcNonstdGeometry
            botLabServ.Store barSec
            
            Set concThk = Nothing
            Set barDat = Nothing
            Set barSec = Nothing
        Next i
        
        Set botLabServ = Nothing
        Set rob = Nothing
    
    End If
    
End Sub

Sub addThickness()

    ' Validation before proceeding to Robot
    Dim dbRange As range
    Call extractData(dbRange)
    
    If dbRange Is Nothing Then
        Exit Sub
    Else
        Dim rob As IRobotApplication
        Dim botLabServ As IRobotLabelServer
        Dim panelThk As IRobotLabel
        Dim panelDat As IRobotThicknessData
        Dim panelThkData As IRobotThicknessHomoData
    
        'Spreadsheet Data
        Dim dataArr() As Variant
        Dim iColCount As Integer
        Dim iRowCount As Integer
        Dim i As Integer
        Dim thkName As String
        iColCount = dbRange.Columns.Count
        iRowCount = dbRange.Rows.Count
    
        ReDim dataArr(iRowCount, iColCount)
        dataArr = dbRange.Value
        
        Set rob = New RobotApplication
        Set botLabServ = rob.Project.Structure.Labels
        
        For i = 1 To iRowCount
            thkName = CStr(dataArr(i, 1)) + CStr(dataArr(i, 2)) + CStr(dataArr(i, 3))
            Set panelThk = botLabServ.Create(I_LT_PANEL_THICKNESS, thkName)
            Set panelDat = panelThk.data
          
            'Setting Material Type (has to be available within the Robot Database)
            panelDat.MaterialName = "FC27"
            panelDat.ThicknessType = I_TT_HOMOGENEOUS
            Set panelThkData = panelDat.data
            
            'Setting thickness
            panelThkData.ThickConst = CDbl(dataArr(i, 4))
            
            'Setting reductions
            If InStr(1, dataArr(i, 1), "PT", vbBinaryCompare) <> 0 Then
                panelThkData.SetReduction False, 1#  'checked/unchecked, overall stiffness reduction
            ElseIf InStr(1, dataArr(i, 1), "W", vbBinaryCompare) <> 0 Then
                panelThkData.SetReduction True, 0.7 'checked/unchecked, overall stiffness reduction
            Else
                panelThkData.SetReduction True, 0.25 'checked/unchecked, overall stiffness reduction
            End If
            botLabServ.Store panelThk
            
            Set panelThkDat = Nothing
            Set panelDat = Nothing
            Set panelThk = Nothing
        Next i
        
        Set botLabServ = Nothing
        Set rob = Nothing
    
    End If
    
End Sub

Sub mano()

    Dim db As range
    Call extractData(db)
    
    Dim lim As Integer
    Dim i As Integer
    lim = db.Rows.Count
    Dim arr() As Variant
    Dim data As String
    
    ReDim arr(lim)
    arr = db.Value
    
    For i = 1 To lim
        
        data = CStr(arr(i, 1)) + CStr(arr(i, 3)) + CStr(arr(i, 5))
        MsgBox data
    
    Next i

End Sub
