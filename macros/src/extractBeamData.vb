' Created for Autodesk Robot Structural Analysis Professional 2017
' Reference: robotom.tlb (under Robot SDK)

Option Explicit
Option Base 1

Dim rob As RobotApplication
Dim str As IRobotStructure

Sub robInit()
    
    Set rob = New RobotApplication
    Set str = rob.Project.Structure
    
End Sub

Sub robSever()

    Set str = Nothing
    Set rob = Nothing

End Sub

Function get_storeys(ByVal str As IRobotStructure, ByVal noOfSty As Long) As String()
    Dim i As Long
    Dim storeys() As String
    ReDim storeys(noOfSty)
    For i = 1 To noOfSty
        If str.storeys.Get(i).Name = "" Then
            Exit For
        End If
        storeys(i) = str.storeys.Get(i).Name
    Next i
    
    get_storeys = storeys
End Function

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
 
    Set barSelGroup = str.Selections.CreateByStorey(I_OT_BAR, storeyName)
    barSelCount = barSelGroup.Count
                
    Set nodeServ = str.Nodes
    Set barServ = str.Bars
        
    ReDim barData(barSelCount, 11)
    For i = 1 To barSelCount
        barId = barSelGroup.Get(i)
        Set selBar = barServ.Get(barId)
        Set startJt = nodeServ.Get(selBar.startNode)
        Set endJt = nodeServ.Get(selBar.endNode)
        barData(i, 1) = barId
        barData(i, 2) = selBar.startNode
        barData(i, 3) = startJt.X
        barData(i, 4) = startJt.Y
        barData(i, 5) = startJt.Z
        barData(i, 6) = selBar.endNode
        barData(i, 7) = endJt.X
        barData(i, 8) = endJt.Y
        barData(i, 9) = endJt.Z
        barData(i, 10) = selBar.Length
        barData(i, 11) = storeyName
        Set selBar = Nothing
        Set startJt = Nothing
        Set endJt = Nothing
    Next i
    
    MsgBox "Bars selected: " & barSelCount
    Sheet3.Activate
    
    Set dataRange = Range(Cells(2, 1), Cells(1 + barSelCount, 11))
    
    dataRange.Value = barData
    get_beam_data = barData

    Set dataRange = Nothing
    Set barServ = Nothing
    Set nodeServ = Nothing
    Set selBar = Nothing
    robSever
    
End Function

Function vecDir(ByVal data As Variant) As Double

    Dim startCoords(3) As Double
    Dim endCoords(3) As Double
    

End Function

Sub main()

    vecDir (get_beam_data)
    
End Sub
