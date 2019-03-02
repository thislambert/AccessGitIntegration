VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GIT_Rep_Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : GIT_Rep_Tools
' Author    : Lambert
' Date      : 2/27/2019
' Purpose   : A class to handle some GIT repository functionality in Access.
'           : all forlders paths are assumed to be the current path wherre the .git folder resides
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const NOT_A_MODULE = -1
Private GIT_OBJECTS_TABLE   As String
Private Mod_ThisProject As VBProject

Public Function InitializeProperties(ObjectTable As String) As Boolean

' first check to see if there is a .git folder in the application folder
    Dim sFile As Variant
    Dim fso As Object
    Dim folder As Object
    Dim IsRepository As Boolean
    ' Using vbscript to look for the hidden .git folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsRepository = fso.folderexists(CurrentProject.Path & "\.git")
    If IsRepository Then
        GIT_OBJECTS_TABLE = ObjectTable
        ' Here we want to get the first project in the VBE, which is
        ' the one which will be importing/exporting files for the Git rpository it sits in.
        Set Mod_ThisProject = Application.VBE.VBProjects(1)
    Else
        Debug.Print "The current file " & CurrentProject.Name & " is not in a GIT repository."
    End If
    InitializeProperties = IsRepository

End Function


Public Sub exportModuleOrClass(sName As String)
'TODO
    Dim component As VBComponent
    Set component = Mod_ThisProject.VBComponents(sName)
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        If Not ModuleSaveState(sName) Then
            DoCmd.Save acMacro, sName
        End If
        component.Export CurrentProject.Path & "\" & component.Name & ToFileExtension(component.Type)
        GitSaveModuleState sName
    End If
End Sub


Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
        ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
        ToFileExtension = vbNullString
    End Select

End Function
Function ExportAllModulesAndClasses(Optional bAskToSave As Boolean = True) As Boolean
' this routine will export classes and vba modules
    Dim component As VBComponent
    Dim n As Long
    
    If bAskToSave Then
        'First check if there are any code modules that have not been saved
        Dim c As New Collection
        Dim msg As String
        For Each component In Mod_ThisProject.VBComponents
            If Not ModuleSaveState(component.Name) Then
                c.Add component.Name
                msg = msg & IIf(msg & "" > "", vbCrLf, vbNullString) & component.Name
            End If
        Next
        msg = IIf(c.Count > 1, "There are ", "There is ") & c.Count & " unsaved modules/classes:" & vb2CrLf & msg
        msg = msg & vb2CrLf & "Do you want to save them now?"

        If MsgBox(msg, vbYesNo Or vbQuestion) = vbNo Then
            ExportAllModulesAndClasses = False
            Exit Function
        Else
            For n = 1 To c.Count
                DoCmd.Save acModule, c(n)
            Next n
        End If
    End If
    For Each component In Mod_ThisProject.VBComponents
        exportModuleOrClass component.Name
    Next
    ExportAllModulesAndClasses = True
    Debug.Print "done"
End Function

Public Sub ExportAllObjects()
    If ExportAllModulesAndClasses() Then
'    ExportAllForms
'    ExportAllReports
'    exportAllQueries
    End If
End Sub

'----------------------------------
'---------------------------------------------------------------------------------------
' Procedure : GitSaveAllModuleStates
' Author    : Lambert
' Date      : 2/25/2019
' Purpose   : Scan all the modules in the current database and record their
            ' last modified data and saved status
'---------------------------------------------------------------------------------------
Sub GitSaveAllModuleStates()
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim ModuleName As String
'    Dim dLastModified As Date
'    Dim IsSaved As Boolean
'    Dim i As Long
'
'    On Error GoTo GitSaveAllModuleStates_Error
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
'    If Not IsDestroyed(rs) Then
'        Set Mod_ThisProject = Application.VBE.VBProjects(GetFileBaseName(CurrentProject.name))
'        For i = 0 To CodeDb.Containers("Modules").Documents.Count - 1
'            ModuleName = CodeDb.Containers("Modules").Documents(i).name
'            dLastModified = ModuleModified(ModuleName)
'            IsSaved = ModuleSaveState(ModuleName)
'            With rs
'                .FindFirst "ObjectName=" & quote(ModuleName) & " AND ObjectType=" & acModule
'                If .NoMatch Then
'                    .AddNew
'                Else
'                    .Edit
'                End If
'                !ObjectName = ModuleName
'                !LastUpdated = dLastModified
'                !ObjectType = acModule
'                !ModuleType = ModuleType(ModuleName)
'                !SavedState = IsSaved
'                '!LastGitExport = Whatever
'                .Update
'            End With
'
'            Debug.Print ModuleName & " - Last Updated: " & dLastModified & " Is saved: " & IsSaved
'            '
'        Next i
'
'    Else
'        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
'    End If
'
'GitSaveAllModuleStates_Exit:
'    If Not IsDestroyed(rs) Then
'        rs.Close
'        Set rs = Nothing
'        Set db = Nothing
'    End If
'    On Error GoTo 0
'    Exit Sub
'
'GitSaveAllModuleStates_Error:
'    Dim err_other_info As Variant
'    Select Case Err.Number
'    Case 0    ' No Error
'        DoEvents
'    Case Else
'        #If Not DEBUGGING Then
'            logError Err.Number, Err.Description, "GitSaveAllModuleStates", "Module Git_tools_mod", Erl, err_other_info
'            Resume GitSaveAllModuleStates_Exit
'        #Else
'            ' Next 3 lines only for debugging
'            MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "GitSaveAllModuleStates"
'            Stop
'            Resume
'        #End If
'    End Select
End Sub



Private Sub GitSaveModuleState(ByVal ModuleName As String)
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dLastModified As Date
    Dim IsSaved As Boolean
    Dim i As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
    If Not IsDestroyed(rs) Then
        Set Mod_ThisProject = Application.VBE.VBProjects(1)
        ''''For i = 0 To CodeDb.Containers("Modules").Documents.Count - 1
            'ModuleName = CodeDb.Containers("Modules").Documents(i).name
            dLastModified = ModuleModified(ModuleName)
            IsSaved = ModuleSaveState(ModuleName)
            With rs
                .FindFirst "ObjectName=" & quote(ModuleName) & " AND ObjectType=" & acModule
                If .NoMatch Then
                    .AddNew
                Else
                    .Edit
                End If
                !ObjectName = ModuleName
                !LastUpdated = dLastModified
                !ObjectType = acModule
                !ModuleType = ModuleType(ModuleName)
                !SavedState = IsSaved
                !LastGitExport = Now()
                .Update
            End With

            Debug.Print ModuleName & " - Last Updated: " & dLastModified & " Is saved: " & IsSaved
            '
        ''''Next i
    
    Else
        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
    End If
End Sub
'---------------------------------------------------------------------------------------
' Function  : ModuleType
' Author    : Lambert
' Date      : 2/27/2019
' Purpose   : determins what type of module we are looking at.
' Returns   : acClassModule or acStandardModule for a Module
'           : Else returns
'---------------------------------------------------------------------------------------
'
Private Function ModuleType(ByVal sName As String) As Long
    Dim component As VBComponent
    Dim theProject As VBProject
    Set theProject = Application.VBE.VBProjects(1)
    Set component = theProject.VBComponents(sName)
    Select Case component.Type
    Case vbext_ct_ClassModule
        ModuleType = acClassModule
    Case vbext_ct_StdModule
        ModuleType = acStandardModule
    Case Else
        ModuleType = NOT_A_MODULE
    End Select
    Set component = Nothing
    Set theProject = Nothing
End Function

Function ModuleSaveState(ModuleName)
    Dim component As VBComponent
    If IsDestroyed(Mod_ThisProject) Then
        Set Mod_ThisProject = Application.VBE.VBProjects(1)
    End If
    ModuleSaveState = Mod_ThisProject.VBComponents(ModuleName).Saved
End Function

Function ModuleModified(sModuleName As String) As Date
Dim n As Long
    ModuleModified = CodeDb.Containers("Modules").Documents(sModuleName).LastUpdated
    Exit Function
    Debug.Print CodeDb.Containers("Modules").Documents(sModuleName).Properties.Count
    For n = 0 To CodeDb.Containers("Modules").Documents(sModuleName).Properties.Count - 1
        Debug.Print CodeDb.Containers("Modules").Documents(sModuleName).Properties(n).Name, CodeDb.Containers("Modules").Documents(sModuleName).Properties(n)
    Next n
End Function
