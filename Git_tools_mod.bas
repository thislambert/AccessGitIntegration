Attribute VB_Name = "Git_tools_mod"
'Option Compare Database
'Option Explicit
Const ZERO = 0
'
'
'Sub Export_All()
'Dim oG As GIT_Rep_Tools
'    Set oG = Create_GIT_Rep_Tools()
'    If Not IsDestroyed(oG) Then
'    'oG.gitSaveFormState ("fORM1")
'
'    End If
'    Set oG = Nothing
'End Sub


''
''E:\Users\Lambert\Documents\@Programming\Databases\Libs\Lib Additions Dev\TableClasses_dev.accdb
''TEMP
'Private Const GIT_OBJECTS_TABLE = "ff"
'
'Private Const NOT_A_MODULE = -1
'
'Private Mod_ThisProject As VBProject
'
'Sub GitExportAll()
'' first check to see if there is a .git folder in the application folder
'    Dim sFile As Variant
'    Dim fso As Object
'    Dim folder As Object
'    Dim expObj As cExportObjects
'    ' Using vbscript to look for the hidden .git folder
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If fso.folderexists(CurrentProject.Path & "\.git") Then
'        Set expObj = Create_cExportObjects(CurrentProject.Path)
'        expObj.ExportAllObjects
'        Set expObj = Nothing
'    Else
'        Debug.Print "The current file " & CurrentProject.name & " is not in a GIT repository."
'    End If
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : GitSaveAllModuleStates
'' Author    : Lambert
'' Date      : 2/25/2019
'' Purpose   : Scan all the modules in the current database and record their
'            ' last modified data and saved status
''---------------------------------------------------------------------------------------
'Sub GitSaveAllModuleStates()
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
'End Sub
'
'Sub GitSaveModuleState(ByVal ModuleName As String)
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim dLastModified As Date
'    Dim IsSaved As Boolean
'    Dim i As Long
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
'    If Not IsDestroyed(rs) Then
'        Set Mod_ThisProject = Application.VBE.VBProjects(GetFileBaseName(CurrentProject.name))
'        ''''For i = 0 To CodeDb.Containers("Modules").Documents.Count - 1
'            'ModuleName = CodeDb.Containers("Modules").Documents(i).name
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
'                !LastGitExport = Now()
'                .Update
'            End With
'
'            Debug.Print ModuleName & " - Last Updated: " & dLastModified & " Is saved: " & IsSaved
'            '
'        ''''Next i
'
'    Else
'        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
'    End If
'End Sub
'
'Function ModuleModified(sModuleName As String) As Date
'Dim n As Long
'    ModuleModified = CodeDb.Containers("Modules").Documents(sModuleName).LastUpdated
'    Exit Function
'    Debug.Print CodeDb.Containers("Modules").Documents(sModuleName).Properties.Count
'    For n = 0 To CodeDb.Containers("Modules").Documents(sModuleName).Properties.Count - 1
'        Debug.Print CodeDb.Containers("Modules").Documents(sModuleName).Properties(n).name, CodeDb.Containers("Modules").Documents(sModuleName).Properties(n)
'    Next n
'End Function
'
'Function ModuleSaveState(ModuleName)
'    Dim component As VBComponent
'    If IsDestroyed(Mod_ThisProject) Then
'        Set Mod_ThisProject = Application.VBE.VBProjects(GetFileBaseName(CurrentProject.name))
'    End If
'    ModuleSaveState = Mod_ThisProject.VBComponents(ModuleName).Saved
'End Function
'
''---------------------------------------------------------------------------------------
'' Function  : ModuleType
'' Author    : Lambert
'' Date      : 2/27/2019
'' Purpose   : determins what type of module we are looking at.
'' Returns   : acClassModule or acStandardModule for a Module
''           : Else returns
''---------------------------------------------------------------------------------------
''
'Function ModuleType(ByVal sName As String) As Long
'    Dim component As VBComponent
'    Dim theProject As VBProject
'    Set theProject = Application.VBE.VBProjects(GetFileBaseName(CurrentDb.name))
'    Set component = theProject.VBComponents(sName)
'    Select Case component.Type
'    Case vbext_ct_ClassModule
'        ModuleType = acClassModule
'    Case vbext_ct_StdModule
'        ModuleType = acStandardModule
'    Case Else
'        ModuleType = NOT_A_MODULE
'    End Select
'    Set component = Nothing
'    Set theProject = Nothing
'End Function
'
