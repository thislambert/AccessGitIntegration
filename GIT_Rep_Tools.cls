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
Private GitFolder As String

Public Function InitializeProperties(ObjectTable As String) As Boolean
' first check to see if there is a .git folder in the application folder
    Dim sFile As Variant
    Dim fso As Object
    Dim folder As Object
    Dim IsRepository As Boolean
    ' Using vbscript to look for the hidden .git folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    GitFolder = CurrentProject.Path & "\"
    IsRepository = fso.folderexists(GitFolder & ".git")
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

Public Function ExportAllModulesAndClasses() As Boolean    'Optional bAskToSave As Boolean = True) As Boolean
' this routine will export classes and vba modules
    Dim component As VBComponent
    Dim N As Long
    Dim Saved As Boolean
    Dim mType As Long

    'First check if there are any code modules that have not been saved
    Dim c As New Collection
    Dim msg As String

    For Each component In Mod_ThisProject.VBComponents
        mType = moduleType(component.Name)
        If mType = acStandardModule Or mType = acClassModule Then
            Saved = component.Saved
            If Saved = False Then
                c.Add component.Name
                msg = msg & IIf(msg & "" > "", vbCrLf, vbNullString) & component.Name
            End If
        End If
    Next
    msg = IIf(c.Count > 1, "There are ", "There is ") & c.Count & " unsaved modules/classes:" & vb2CrLf & msg
    msg = msg & vb2CrLf & "Do you want to save them now?"
    If c.Count > 0 Then
        If MsgBox(msg, vbYesNo Or vbQuestion) = vbNo Then
            FlashMessage "Export cancelled"
            ExportAllModulesAndClasses = False
            Exit Function
        End If
    End If
    For N = 1 To c.Count
        Debug.Print c(N)
        DoCmd.Save acModule, c(N)
    Next N

    For Each component In Mod_ThisProject.VBComponents
        mType = moduleType(component.Name)
        If mType = acStandardModule Or mType = acClassModule Then
            ExportModuleOrClass component.Name
        End If
    Next

    ExportAllModulesAndClasses = True
    Debug.Print "done"
End Function

Public Sub ExportAllObjects()
    If ExportAllModulesAndClasses() Then
    ExportAllForms
'    ExportAllReports
'    exportAllQueries
    End If
End Sub

Public Sub ExportAForm(ByVal formName As String, Optional PromptToSave As Boolean = False)

    Select Case SysCmd(acSysCmdGetObjectState, acForm, formName)
    Case Is < 2
        DoEvents    ' the form is either closed or open but unchanged
    Case 3    ' the form has been modified
        If PromptToSave Then
            If MsgBox("Form " & formName & " has been modified. Save the design before exporting?", vbYesNo Or vbQuestion) = vbYes Then
                DoCmd.Save acForm, formName
            End If
        Else
            DoCmd.Save acForm, formName
        End If
    Case 5
        ' a new form, don't do anything
        Exit Sub
    End Select
    ' now export the design of the form object
    Application.SaveAsText acForm, formName, GitFolder & formName & ".frm"
    gitSaveFormState formName
End Sub

Public Sub ExportAllForms()
    Dim N As Long
    Dim db As DAO.Database
    Set db = CurrentDb

    For N = 0 To db.Containers!Forms.Documents.Count - 1
        ExportAForm db.Containers!Forms.Documents(N).Name
    Next N
    Set db = Nothing
End Sub

Public Sub ExportModuleOrClass(sName As String)
    Dim component As VBComponent
    Set component = Mod_ThisProject.VBComponents(sName)
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        Dim Saved As Boolean
         Saved = moduleSaveState(sName)
        If Not Saved Then
        
            ' code is modified. Save it now.
            DoCmd.Save acMacro, sName
        End If
        component.Export GitFolder & component.Name & toFileExtension(component.Type)
        gitSaveModuleState sName
    End If
End Sub

' These gitSavexxxState routines log data about the objects which
' will be used in a future addition to automate the selection of
' the objects that are exported to text file. This will resudce the time
' taken to output a project to the repository, ready to be commited.

Private Sub gitSaveFormState(ByVal formName As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dLastModified As Date
    Dim IsSaved As Boolean
    Dim i As Long
    Dim f As Form

Set db = CurrentDb
    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
    If Not IsDestroyed(rs) Then
        Set Mod_ThisProject = Application.VBE.VBProjects(1)
            dLastModified = formLastUpdated(formName)
            IsSaved = True
            With rs
                .FindFirst "ObjectName=" & quote(formName) & " AND ObjectType=" & acForm
                If .NoMatch Then
                    .AddNew
                Else
                    .Edit
                End If
                'TODO only update fields that change
                !ObjectName = formName
                !LastUpdated = dLastModified
                !ObjectType = acForm
                !moduleType = acClassModule
                !SavedState = IsSaved
                'TODO, this should only be called in the export routine ExportAForm
                !LastGitExport = Now()
                .Update
            End With

            Debug.Print formName & " - Last Updated: " & dLastModified & " Is saved: " & IsSaved
    Else
        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
    End If
End Sub

Private Sub gitSaveModuleState(ByVal ModuleName As String)
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dLastModified As Date
    Dim IsSaved As Boolean
    Dim i As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
    If Not IsDestroyed(rs) Then
        Set Mod_ThisProject = Application.VBE.VBProjects(1)
            dLastModified = moduleLastUpdated(ModuleName)
            IsSaved = moduleSaveState(ModuleName)
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
                !moduleType = moduleType(ModuleName)
                !SavedState = IsSaved
                !LastGitExport = Now()
                .Update
            End With

            Debug.Print ModuleName & " - Last Updated: " & dLastModified & " Is saved: " & IsSaved
    Else
        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
    End If
End Sub

Private Sub gitSaveMacroState(ByVal MacroName As String)
End Sub

Private Sub gitSaveReportState(ByVal ReportName As String)
End Sub

Private Sub gitSaveQueryState(ByVal QueryName As String)
End Sub

Private Sub gitSaveTableState(ByVal TableName As String)
End Sub

Private Function moduleLastUpdated(sModuleName As String) As Date
    moduleLastUpdated = CodeDb.Containers("Modules").Documents(sModuleName).LastUpdated
End Function

Private Function formLastUpdated(ByVal formName As String) As Date
    Dim N As Long
    Dim db As DAO.Database
    Set db = CurrentDb

    For N = 0 To db.Containers!Forms.Documents.Count - 1
        If formName = db.Containers!Forms.Documents(N).Name Then
            formLastUpdated db.Containers!Forms.Documents(N).LastUpdated
            Exit Function
        End If
    Next N
    Set db = Nothing
End Function

Private Function moduleSaveState(ModuleName As String) As Boolean
    Dim IsSaved As Boolean
    If Mod_ThisProject.VBComponents(ModuleName).Saved = False Then
        moduleSaveState = False
    Else
        moduleSaveState = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : ModuleType
' Author    : Lambert
' Date      : 2/27/2019
' Purpose   : determins what type of module we are looking at.
' Returns   : acClassModule or acStandardModule for a Module
'           : Else returns
'---------------------------------------------------------------------------------------
'
Private Function moduleType(ByVal sName As String) As Long
    Dim component As VBComponent
    Dim theProject As VBProject
    Set theProject = Application.VBE.VBProjects(1)
    Set component = theProject.VBComponents(sName)
    Select Case component.Type
    Case vbext_ct_ClassModule
        moduleType = acClassModule
    Case vbext_ct_StdModule
        moduleType = acStandardModule
    Case Else
        moduleType = NOT_A_MODULE
    End Select
    Set component = Nothing
    Set theProject = Nothing
End Function

Private Function toObjType(ObjName As String) As AcObjectType
    If Left(ObjName, 5) = "Form_" Then
        If isDbObject(acForm, Mid(ObjName, 6)) Then
            toObjType = acForm
        End If
    ElseIf isDbObject(acModule, ObjName) Then
        toObjType = acModule
    ElseIf isDbObject(acQuery, ObjName) Then
        toObjType = acQuery
    ElseIf isDbObject(acReport, ObjName) Then
        toObjType = acReport
    ElseIf isDbObject(acTable, ObjName) Then
        toObjType = acTable
    Else
        toObjType = acDefault
    End If
End Function


Private Function toFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
        toFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        toFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        toFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
        toFileExtension = vbNullString
    End Select

End Function

