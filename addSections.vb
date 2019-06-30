
' Note: Object types TOO LONG; it can be shortened (aliased) but Autodesk has reasons they put them that way

Sub setBeamSec(ByRef shapeType As IRobotBarSectionData, ByRef widthType As IRobotBarSectionConcreteDataValue, ByRef depthType As IRobotBarSectionConcreteDataValue)
    ' Setting Shape Type to Beam
    shapeType.shapeType() = I_BSST_CONCR_BEAM_RECT 'Code: -3
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
        Dim robLab As IRobotLabelServer
        Dim barSec As IRobotLabel
        Dim barDat As IRobotBarSectionData
        Dim concSec As IRobotBarSectionConcreteData
    
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
        Set robLab = rob.Project.Structure.Labels
        
        For i = 1 To iRowCount
            Set barSec = robLab.Create(I_LT_BAR_SECTION, secName)
            Set barDat = barSec.data
          
            
            'Setting Material Type (has to be available within the Robot Database)
            barDat.MaterialName = "FC27"
            secName = Cells(i, 2)
            
            'Check if the section is a Beam or a Girder; defaults to Column section
            If InStr(1, secName, "B", vbBinaryCompare) <> 0 Then
                Call setBeamSec(barDat, widthType, depthType)
                Set concSec = barDat.Concrete
                concSec.SetReduction True, 0.35, 0.35, 1 'checked/unchecked, Ix, Iy, Iz
            ElseIf InStr(1, secName, "G", vbBinaryCompare) <> 0 Then
                Call setBeamSec(barDat, widthType, depthType)
                Set concSec = barDat.Concrete
                concSec.SetReduction True, 0.35, 0.35, 1 'checked/unchecked, Ix, Iy, Iz
            Else
                Call setColSec(barDat, widthType, depthType)
                Set concSec = barDat.Concrete
                concSec.SetReduction True, 0.7, 0.7, 1 'checked/unchecked, Ix, Iy, Iz
            End If
            
            If Cells(i, 6) = "Rectangular" Then 'Section Shape
                concSec.SetValue widthType, Cells(i, 10)
                concSec.SetValue depthType, Cells(i, 8)
            Else
                'Assuming all circular sections are columns
                barDat.shapeType() = I_BSST_CONCR_COL_C
                concSec.SetValue I_BSCDV_COL_DE, Cells(i, 8)
            End If
            
            barDat.CalcNonstdGeometry
            robLab.Store barSec
            
            Set concSec = Nothing
            Set barDat = Nothing
            Set barSec = Nothing
        Next i
        
        Set robLab = Nothing
        Set rob = Nothing
    
    End If
    
End Sub
