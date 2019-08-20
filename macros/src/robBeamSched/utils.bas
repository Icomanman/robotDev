Attribute VB_Name = "utils"
Option Explicit

Public rob As IRobotApplication
Public struc As IRobotStructure

Public Function robInit() As IRobotApplication
    Set rob = New RobotApplication
    Set robInit = rob
End Function

Public Function strInit() As IRobotStructure
    Set struc = rob.Project.Structure
    Set strInit = struc
End Function

Public Sub robSever()
    Set str = Nothing
    Set rob = Nothing
End Sub

Public Sub release(ByRef obj As Variant)
    Set obj = Nothing
End Sub

Public Function get_bar_id(ByVal selectedBar As IRobotSelection, ByVal pos As Long) As Long
    get_bar_id = selectedBar.Get(pos) ' pos = index within the selection
End Function

Public Function get_bars(ByVal barNames As String) As IRobotSelection
    Dim membSel As IRobotSelection
    Set membSel = struc.Selections.Create(I_OT_BAR)
    membSel.AddText (barNames)
    Set get_bars = membSel
    Debug.Print membSel.count & " members are selected."
End Function

Public Function get_by_storey(ByVal storey As IRobotStorey, Optional ByVal elem As Long = 1) As IRobotSelection
    '0: I_OT_NODE
    '1: I_OT_BAR
    '2: I_OT_CASE
    '...etc (see reference)
    Dim membSel As IRobotSelection
    Dim storeyName As String
    storeyName = storey.Name
    Set membSel = struc.Selections.CreateByStorey(elem, storeyName)
    Set get_by_storey = membSel
    Debug.Print "Selections created for storey: " & storeyName
End Function

Public Function get_storey(ByVal storeyName As String) As IRobotStorey
    Dim storey As IRobotStorey
    Dim storeyId As Long
    storeyId = struc.storeys.Find(storeyName)
    
    If storeyId < 1 Then
        MsgBox "Requested Storey not found in the model."
        Exit Function
    End If
    Set storey = struc.storeys.Get(storeyId)
    Set get_storey = storey
    Debug.Print "Storey Extracted: " & storeyName
End Function

Public Function get_cases(ByVal caseNames As String) As IRobotSelection
    Dim selCases As IRobotSelection
    Set selCases = struc.Selections.Create(I_OT_CASE)
    selCases.AddText (caseNames)
    
    Set get_cases = selCases
    Debug.Print "Selected Load Cases: " & caseNames
End Function
