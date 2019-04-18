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
'           : all folders paths are assumed to be the current path where the .git folder resides
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const NOT_A_MODULE = -1
Private Const NULL_DATE = #1/1/1700#

Private GIT_OBJECTS_TABLE   As String
Private theProject As VBProject
Private theGitFolder As String

'---------------------------------------------------------------------------------------
' Procedure : InitializeProperties
' Author    : Lambert
' Date      : 4/16/2019
' Purpose   : Creates the GIT object that will managed the database objects in the
            ' development project.
' Parameters: ObjectTable - the name of the table that will be used to log information
            ' about database objects
'---------------------------------------------------------------------------------------
'
Public Function InitializeProperties(ObjectTable As String) As Boolean
' first check to see if the application folder is set up as
' a GIT repository

    Dim sFile As Variant
    Dim fso As Object
    Dim folder As Object
    Dim IsRepository As Boolean
    ' Using vbscript to look for the hidden .git folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    theGitFolder = CurrentProject.Path & "\"
    IsRepository = fso.folderexists(theGitFolder & ".git")
    If IsRepository Then
        GIT_OBJECTS_TABLE = ObjectTable
        ' Here we want to get the first project in the VBE, which is
        ' the one which will be importing/exporting files for the Git rpository it sits in.
        Set theProject = Application.VBE.VBProjects(1)
    Else
        Debug.Print "The current file " & CurrentProject.Name & " is not in a GIT repository."
    End If
    InitializeProperties = IsRepository

End Function

Public Sub ExportAllObjects()
    ExportForms
    ExportlModulesAndClasses
End Sub

Public Sub ExportForm(ByVal formName As String, Optional PromptToSave As Boolean = False)

    Select Case SysCmd(acSysCmdGetObjectState, acForm, formName)
    Case Is < 2
        DoEvents    ' the form is either closed or open but unchanged
    Case 3    ' the form has been modified
        If PromptToSave Then
            If MsgBox("Form " & formName & " has been modified. Save the design before exporting?", vbYesNo Or vbQuestion) = vbYes Then
                DoCmd.Save acForm, formName
            End If
        Else ' default is to save without asking
            DoCmd.Save acForm, formName
        End If
    Case 5
        ' a new form, don't do anything
        mbox "'" & formName & "'" & vb2CrLf & "This form is new. Please save it manually first.", vbOKOnly
        Exit Sub
    End Select
    ' now export the design of the form object
    Application.SaveAsText acForm, formName, theGitFolder & formName & ".frm"
    gitSaveObjectState acForm, formName
End Sub

Public Sub ExportForms()
    Dim N As Long
    Dim db As DAO.Database
    Set db = CurrentDb
    ' TODO check each form's SavedState to see if we need to export it at all
    For N = 0 To db.Containers!Forms.Documents.Count - 1
        ExportForm db.Containers!Forms.Documents(N).Name
    Next N
    Set db = Nothing
End Sub

Public Function ExportlModulesAndClasses() As Boolean    'Optional bAskToSave As Boolean = True) As Boolean
' this routine will export classes and vba modules
    Dim component As VBComponent
    Dim N As Long
    Dim Saved As Boolean
    Dim mType As Long

    Dim c As New Collection
    Dim msg As String

    'First check if there are any code modules that have not been saved
    For Each component In theProject.VBComponents
        mType = moduleType(component.Name)
        If mType = acStandardModule Or mType = acClassModule Then
            Saved = component.Saved
            If Saved = False Then
                c.Add component.Name
                msg = msg & IIf(msg & "" > "", vbCrLf, vbNullString) & component.Name
            End If
        End If
    Next
    If c.Count > 0 Then
        msg = IIf(c.Count > 1, "There are ", "There is ") & c.Count & " unsaved modules/classes:" & vb2CrLf & msg _
              & vb2CrLf & "Do you want to save them now?"
        If MsgBox(msg, vbYesNo Or vbQuestion) = vbNo Then
            FlashMessage "Export cancelled"
            ExportlModulesAndClasses = False
            Exit Function
        End If
        ' save the modules
        For N = 1 To c.Count
            Debug.Print c(N)
            DoCmd.Save acModule, c(N)
        Next N
    End If
    ' Now we can export
    For Each component In theProject.VBComponents
        mType = moduleType(component.Name)
        If mType = acStandardModule Or mType = acClassModule Then
            ExportModuleOrClass component.Name
        End If
    Next
    ExportlModulesAndClasses = True
    Debug.Print "done"
End Function

Public Sub ExportModuleOrClass(sName As String)
'TODO check the savedstate to see if we need to export at all
    Dim component As VBComponent
    Set component = theProject.VBComponents(sName)
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        Dim Saved As Boolean
        Saved = isModuleSaved(sName)
        If Not Saved Then
            ' code is modified. Save it now.
            DoCmd.Save acMacro, sName
        End If
        ' TODO test this date export check
        If dObjModified(acModule, sName) > gitExportDate(acModule, sName) Then
            component.Export theGitFolder & component.Name & toFileExtension(component.Type)
        End If
        gitSaveObjectState acModule, sName
    End If
End Sub

' The getSaveObjectState routine logs data about the objects which
' will be used in a future edition to automate the selection of
' the objects that are exported to text files. This will reduce the time
' taken to output a project to the repository, ready to be commited.

'---------------------------------------------------------------------------------------
' Procedure : getSaveObjectState
' Author    : Lambert
' Date      : 4/18/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub gitSaveObjectState(ByVal oType As AcObjectType, ByVal oName As String)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dLastModified As Date
    Dim IsSaved As Boolean
    Dim i As Long

    On Error GoTo getSaveObjectState_Error
    dLastModified = dObjModified(oType, oName)
    Set db = CurrentDb
    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
    If Not IsDestroyed(rs) Then
        Set theProject = Application.VBE.VBProjects(1)

        With rs
            .FindFirst "ObjectName=" & quote(oName) & " AND ObjectType=" & oType
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            !ObjectName = oName
            !LastUpdated = dLastModified
            !ObjectType = oType
            Select Case oType
            Case acForm
                !moduleType = acClassModule
            Case acModule
                !moduleType = moduleType(oName)
                IsSaved = isModuleSaved(oName)
            Case Else
                !moduleType = Null
            End Select
            !SavedState = IsSaved
            'TODO, this should only be called in the export routine ExportForm
            !LastGitExport = Now()
            .Update
        End With
        rs.Close
        Set rs = Nothing
    Else
        Debug.Print "Failed to open the table " & GIT_OBJECTS_TABLE
    End If
getSaveObjectState_Exit:
    If Not IsDestroyed(rs) Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
    Exit Sub

getSaveObjectState_Error:
    Dim err_other_info As Variant
    Select Case Err.Number
    Case 0    ' No Error
        Resume getSaveObjectState_Exit    ' We should never get here
    Case Else
        #If Not DEBUGGING Then
            logError Err.Number, Err.Description, "getSaveObjectState", "Class Module GIT_Rep_Tools", Erl, err_other_info
            Resume getSaveObjectState_Exit
        #Else
            ' Next 3 lines only for debugging
            MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "getSaveObjectState"
            Stop
            Resume
        #End If
    End Select
End Sub


Private Function isModuleSaved(ModuleName As String) As Boolean
    Dim IsSaved As Boolean
    If theProject.VBComponents(ModuleName).Saved = False Then
        isModuleSaved = False
    Else
        isModuleSaved = True
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

Private Function gitExportDate(nType As AcObjectType, sName As String) As Date
Dim db As DAO.Database
Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(GIT_OBJECTS_TABLE, dbOpenDynaset)
    With rs
        .FindFirst "ObjectType=" & nType & " and ObjectName='" & sName & "'"
        If .NoMatch Then
            gitExportDate = NULL_DATE
        Else
            gitExportDate = !LastGitExport
        End If
        .Close
    End With
    Set rs = Nothing
    Set db = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : dObjModified
' Author    : Lambert
' Date      : 4/16/2019
' Purpose   : Returns the last modified date of the specified object, or the date 1/1/1700 on error
' Paramaters: nType - what kind of object, Form, Table etc.
'           : sName - its name
'---------------------------------------------------------------------------------------
'
Private Function dObjModified(nType As AcObjectType, sName As String) As Date
    Dim dbs As Object
    Dim obj As AccessObject
    On Error GoTo dObjModified_Error
    Set dbs = CurrentProject
    Select Case nType
        '# FORM
    Case acForm
        dObjModified = dbs.AllForms(sName).DateModified
        '# MODULE
    Case acModule
        For Each obj In dbs.AllModules
            If obj.Name = sName Then
                dObjModified = obj.DateModified
            End If
        Next obj
        '# QUERY
    Case acQuery
        dObjModified = CurrentDb.QueryDefs(sName).LastUpdated
        '# REPORT
    Case acReport
        dObjModified = dbs.AllReports(sName).DateModified
        '# MACOR
    Case acMacro
        dObjModified = dbs.AllMacros(sName).DateModified
        '# TABLE
    Case acTable
        dObjModified = CurrentDb.TableDefs(sName).LastUpdated
    Case Else
        DoEvents
    End Select
dObjModified_Exit:
    Set dbs = Nothing
    Set obj = Nothing
    On Error GoTo 0
    Exit Function

dObjModified_Error:
    Dim msg As String
    Dim t As String
    Dim err_other_info As Variant
    Select Case Err.Number
    Case 0    ' No Error
        Resume dObjModified_Exit    ' We should never get here
    Case 2467    ' The expression you entered refers to an object that is closed or doesn't exist.
        msg = "No such | named '" & sName & "'"
        Select Case nType
        Case acForm
            t = "form"
            '# MODULE
        Case acModule
            t = "module"
            '# QUERY
        Case acQuery
            t = "query"
            '# REPORT
        Case acReport
            t = "report"
            '# MACOR
        Case acMacro
            t = "macro"
            '# TABLE
        Case acTable
            t = "table"
        Case Else
        End Select
        msg = Replace(msg, "|", t)
        mbox msg, vbOKOnly
        dObjModified = NULL_DATE
        Resume dObjModified_Exit

    Case Else
        #If Not DEBUGGING Then
            logError Err.Number, Err.Description, "dObjModified", "Module Git_tools_test_mod", Erl, err_other_info
            dObjModified = NULL_DATE
            Resume dObjModified_Exit
        #Else
            ' Next 3 lines only for debugging
            MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "dObjModified"
            Stop
            Resume
        #End If
    End Select
End Function

