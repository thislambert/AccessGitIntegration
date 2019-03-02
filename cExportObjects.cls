VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExportObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private theDb As DAO.Database
Private theProject As VBProject
Private theCodeProject As CodeProject
Private theDest As String

Public Sub InitializeProperties(destpath As String, Optional strSourceDb As String = vbNullString)
    theDest = QualifyPath(destpath)
    If strSourceDb = vbNullString Then
        Set theDb = CurrentDb
    Else
        Set theDb = OpenDatabase(strSourceDb)
    End If
    Set theProject = Application.VBE.VBProjects(1) '(GetFileBaseName(theDb.name))
End Sub


Public Sub exportAForm(sName As String)
    Application.SaveAsText acForm, sName, theDest & sName & ".frm"
    Debug.Print sName
End Sub

Public Sub ExportAllForms()
    Dim obj As Document
    
    For Each obj In theDb.Containers!Forms.Documents
        If Left$(obj.Name, 1) <> "~" Then
            exportAForm obj.Name
        End If
    Next obj
End Sub

Public Sub exportAQuery(sName As String)
    Dim qd As DAO.QueryDef
    For Each qd In theDb.QueryDefs
        If Left$(qd.Name, 1) <> "~" Then
            Application.SaveAsText acQuery, qd.Name, theDest & qd.Name & ".qry"
        End If
    Next qd
End Sub

Public Sub exportAllQueries()
    Dim qd As DAO.QueryDef
    For Each qd In theDb.QueryDefs
        exportAQuery qd.Name
    Next qd
End Sub

Public Sub exportAReport(sName As String)
    Application.SaveAsText acReport, sName, theDest & sName & ".frm"
End Sub

Private Sub ExportAllReports()
    Dim obj As Document
    For Each obj In theDb.Containers!Reports.Documents
        If Left$(obj.Name, 1) <> "~" Then
            exportAReport obj.Name
        End If
    Next obj
End Sub

'Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
'    Select Case vbeComponentType
'    Case vbext_ComponentType.vbext_ct_ClassModule
'        ToFileExtension = ".cls"
'    Case vbext_ComponentType.vbext_ct_StdModule
'        ToFileExtension = ".bas"
'    Case vbext_ComponentType.vbext_ct_MSForm
'        ToFileExtension = ".frm"
'    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
'    Case vbext_ComponentType.vbext_ct_Document
'    Case Else
'        ToFileExtension = vbNullString
'    End Select
'
'End Function

Private Sub Class_Terminate()
    theDb.Close
    Set theDb = Nothing
    Set theProject = Nothing
End Sub
