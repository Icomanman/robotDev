Option Explicit

Sub AddReference()
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = ActiveWorkbook.VBProject

    '~~> Check if "Microsoft VBScript Regular Expressions 5.5" is already added
    For Each chkRef In vbProj.References
        If chkRef.Name = "VBScript_RegExp_55" Then
            BoolExists = True
            GoTo CleanUp
        End If
    Next

    vbProj.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"

CleanUp:
    If BoolExists = True Then
        MsgBox "Reference already exists"
    Else
        MsgBox "Reference Added Successfully"
    End If

    Set vbProj = Nothing
    Set VBAEditor = Nothing
End Sub