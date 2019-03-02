Attribute VB_Name = "Git_Tools_Class_factory"
'---------------------------------------------------------------------------------------
' Module    : Class_factory
' Author    : Lambert
' Date      : 1/29/2019

' Purpose   : VBA Classes do not have a Constructor method as such. This module
' addresses this limit by defining functions which create class instances
' and then call a method whose purpose is to initialize the class properties
' These functions take as many arguments, of whatever type are desired, and they
' and then assign them to the instance's various properties, constructor style

'           In this documentation the class name MYCLASS will be used.

'           The function template is...

'Public Function Create_MYCLASS(paramter1, parameter2...) As MYCLASS
'    Dim MYCLASS_var As MYCLASS
'    Dim ErrorStr as string
'    Set MYCLASS_var = New MYCLASS 'create the object
'    With MYCLASS_var
'        If Not .InitializeProperties(ErrorStr,paramter1, parameter2...) Then
'            'something went wrong
'            MsgBox ErroStr
'            Set MYCLASS_var = Nothing
'        End If
'    End With
'    Set Create_MYCLASS = MYCLASS_var
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

' create an objext exporter
Public Function Create_cExportObjects(destpath As String, Optional strSourceDb As String = vbNullString) As cExportObjects
Dim oExpO As cExportObjects
    Set oExpO = New cExportObjects
    oExpO.InitializeProperties destpath, strSourceDb
    Set Create_cExportObjects = oExpO
End Function

Public Function Create_GIT_Rep_Tools(Optional GIT_Table As String = "USYS_GIT_Objects") As GIT_Rep_Tools
    Dim oGIT As GIT_Rep_Tools
    Set oGIT = New GIT_Rep_Tools
    If oGIT.InitializeProperties(GIT_Table) Then
        Set Create_GIT_Rep_Tools = oGIT
    Else
        Set Create_GIT_Rep_Tools = Nothing
    End If
End Function
