Attribute VB_Name = "geometry"
Option Explicit
Option Base 1

Public Sub write_txt(ByVal contents As Variant, Optional ByVal fileName As String = "dat.csv")
    'Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\System32\scrrun.dll"

    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Dim filePath As String
    filePath = Application.ActiveWorkbook.Path
    
    
    Set oFile = fso.CreateTextFile(filePath & "\" & fileName, True)
    
    Dim rows As Long
    Dim cols As Integer
    
    For rows = LBound(contents) To UBound(contents)
        For cols = 1 To 10
            oFile.Write Chr(34) & contents(rows, cols) & Chr(34) & "," '& vbTab
            'oFile.Write contents(rows, cols) & "," '& vbTab
        Next cols
        oFile.Write Chr(34) & contents(rows, 11) & Chr(34) & vbCrLf
        'oFile.Write contents(rows, 11) & vbCrLf

    Next rows
            
    oFile.Close
    MsgBox fileName, vbOKOnly
    Set fso = Nothing
    Set oFile = Nothing
    
End Sub


Sub get_section_data(ByVal str As IRobotStructure, ByVal selectedBar As RobotBar, ByRef barDims() As Double)
    
    Dim barName As RobotLabel
    Dim barDat As IRobotBarSectionData
    Dim barWidth As Double
    Dim barDepth As Double
    
    Set barName = str.Labels.Get(I_LT_BAR_SECTION, selectedBar.GetLabelName(I_LT_BAR_SECTION))
    Set barDat = barName.data
    barDepth = barDat.GetValue(I_BSDV_D)
    barWidth = barDat.GetValue(I_BSDV_BF)
    
    ReDim barDims(2)
    barDims(1) = barWidth
    barDims(2) = barDepth
    
End Sub

Function vectorise(ByRef twoDimArr() As Double, ByVal col As Integer) As Double()
    Dim vectorCont() As Double
    Dim count As Long
    Dim i As Long
    count = UBound(twoDimArr, 1)
    ReDim vectorCont(count)
    
    For i = 1 To count
        vectorCont(i) = twoDimArr(i, col)
    Next i
    
    ReDim vectorise(count)
    
    vectorise = vectorCont
    
End Function

Sub extract_min_max(ByRef ext_val As Double, ByRef data() As Double, Optional ByVal criterion As String = "max")
    If IsArray(data) Then
    
        If criterion = "min" Then
            ext_val = data.Min()
        Else
            ext_val = data.Max()
        End If
    
    Else
    
        MsgBox "Data has to be an array of doubles."
        Exit Sub
    
    End If

End Sub

Function get_beam_data() As Variant()

    Dim storeyName As String
    storeyName = "2F"
    
    robInit
    
    'Dim modelStories() As String
    'ReDim modelStories(str.storeys.Count)
    'modelStories = get_storeys(str, str.storeys.Count)
    
    Dim dataRange As Range
    
    Dim barSelGroup As IRobotSelection ' per story
    Dim barData As Variant
    Dim barSelCount As Long
    Dim nodeId(2) As Long
    Dim barId As Long
    
    Dim nodeServ As IRobotNodeServer
    Dim barServ As IRobotBarServer
    
    Dim selBar As IRobotBar
    Dim startJt As IRobotNode
    Dim endJt As IRobotNode
    
    Dim i As Integer
    Dim j As Integer
 
    Set barSelGroup = str.Selections.CreateByStorey(I_OT_BAR, storeyName)
    barSelCount = barSelGroup.count
                
    Set nodeServ = str.Nodes
    Set barServ = str.Bars
        
End Sub
