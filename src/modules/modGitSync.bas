Attribute VB_Name = "modGitSync"
Option Explicit

Public Sub ExportVBA()

    Dim comp As Object
    Dim basePath As String
    
    basePath = ThisWorkbook.path & "\..\src\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        Select Case comp.Type
        
            Case 1 ' Standard Module
                comp.Export basePath & "modules\" & comp.Name & ".bas"
                
            Case 2 ' Class Module
                comp.Export basePath & "classes\" & comp.Name & ".cls"
                
            Case 3 ' UserForm
                comp.Export basePath & "forms\" & comp.Name & ".frm"
                
        End Select
        
    Next comp
    
    MsgBox "VBA exportado com sucesso!", vbInformation

End Sub


Public Sub ImportVBA()

    Dim path As String
    Dim file As String
    
    path = ThisWorkbook.path & "\..\src\modules\"
    file = Dir(path & "*.bas")

    Do While file <> ""
        ThisWorkbook.VBProject.VBComponents.Import path & file
        file = Dir
    Loop

    path = ThisWorkbook.path & "\..\src\classes\"
    file = Dir(path & "*.cls")

    Do While file <> ""
        ThisWorkbook.VBProject.VBComponents.Import path & file
        file = Dir
    Loop

    path = ThisWorkbook.path & "\..\src\forms\"
    file = Dir(path & "*.frm")

    Do While file <> ""
        ThisWorkbook.VBProject.VBComponents.Import path & file
        file = Dir
    Loop

    MsgBox "VBA importado com sucesso!", vbInformation

End Sub
