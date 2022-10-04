Attribute VB_Name = "mdl_Helpers"
'############################################################################################
'# Copyright (c) 2022 Thomas Möller                                                         #
'# MIT License  => https://github.com/team-moeller/better-access-barcode/blob/main/LICENSE  #
'# Version 1.05.13  published: 04.10.2022                                                   #
'############################################################################################

Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#End If

Public Function File2OLE(ByVal Table As String, ByVal PrimaryKeyFieldName As String, _
                         ByVal TargetFieldName As String, ByVal PrimaryKeyValue As Long, _
                         ByVal FileName As String, Optional ByVal InCurrentProjectPath As Boolean) As Long

'Prerequisit: Record with ID must already exist
'Call: File2OLE("USys_FileData","ID","FileData","1","Pivot.min.js",True)

    On Error GoTo Handle_Error

    Dim cnn As ADODB.Connection
    Dim strSQL As String
    Dim rst As ADODB.Recordset
    Dim FileID As Long
    Dim PathFilename As String
    Dim Buffer() As Byte
    Dim FileSize As Long

    If InCurrentProjectPath = True Then
        PathFilename = CurrentProject.Path & "\" & FileName
    Else
        PathFilename = FileName
    End If

    If Dir$(PathFilename) = vbNullString Then
        MsgBox "Thge file '" & PathFilename & "' does not exist."
        Exit Function
    End If

    strSQL = "SELECT " & TargetFieldName & ", FileName " & _
             "FROM " & Table & " " & _
             "WHERE " & PrimaryKeyFieldName & " = " & PrimaryKeyValue
    
    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
    FileID = FreeFile

    Open PathFilename For Binary Access Read Lock Read Write As FileID

    FileSize = FileLen(PathFilename)
    ReDim Buffer(FileSize)
    rst(TargetFieldName) = Null
    Get FileID, , Buffer
    rst(TargetFieldName).AppendChunk Buffer
    rst("FileName") = FileName
    rst.Update
    Close FileID
    File2OLE = True

Exit_Here:
    rst.Close
    Set rst = Nothing
    Set cnn = Nothing
    Close FileID
    Exit Function

Handle_Error:
    File2OLE = Err.Number
    Resume Exit_Here

End Function

Public Function SaveFileToDisk(ByVal FileName As String, ByVal Path As String) As Boolean
On Error GoTo Handle_Error

    'Declarations
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim FileID As Long
    Dim Buffer() As Byte
    Dim FileLen As Long
    Dim Success As Boolean

    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "Select FileData FROM USys_FileData Where FileName='" & FileName & "'", _
        cnn, adOpenDynamic, adLockOptimistic

    FileID = FreeFile
    FileLen = Nz(LenB(rst!FileData), 0)

    If FileLen > 0 Then
        ReDim Buffer(FileLen)
        MakeSureDirectoryPathExists (Path & "\")
        Open Path & "\" & FileName For Binary Access Write As FileID
        Buffer = rst!FileData.GetChunk(FileLen)
        Put FileID, , Buffer
        Close FileID
    End If
    Success = True

Exit_Here:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set cnn = Nothing
    SaveFileToDisk = Success
    Exit Function

Handle_Error:
    Select Case Err.Number
        Case 0
            Resume
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
            Resume Exit_Here
    End Select

End Function

Public Function IsFormOpen(ByVal strFormName As String) As Boolean

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        If Forms(strFormName).CurrentView <> 0 Then
            IsFormOpen = True
        End If
    End If
    
End Function

Public Sub ExportModules()
' This method is intended for the contributor who just wants to export the modules

    Call PrepareAndExportModules(False)

End Sub

Public Sub PrepareAndExportModules(Optional ByVal TagVersion As Boolean = True)
' This method is intended for the power user who wants to update the version numbers in the code headers

    'Declarations
    Dim Version As String
    Dim CodeLine As String
    Dim vbc As Object
    
    MakeSureDirectoryPathExists CurrentProject.Path & "\Modules\"
    Version = DLast("V_Number", "tbl_VersionHistory")
    CodeLine = "'# Version " & Version & "  published: " & Format$(Date, "dd.mm.yyyy") & "                                                   #"
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Then
            If TagVersion Then
                Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.InsertLines 4, CodeLine
                Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.DeleteLines 5, 1
            End If
    
            Application.VBE.ActiveVBProject.VBComponents(vbc.Name).Export CurrentProject.Path & "\Modules\" & vbc.Name & IIf(vbc.Type = 2, ".cls", ".bas")
        End If
    Next
    Application.DoCmd.RunCommand (acCmdCompileAndSaveAllModules)
    
    MsgBox "Export done", vbInformation, "Better Access PivotTable"

End Sub

Public Sub ImportModules()

    'Declarations
    Dim strFile As String
    Dim strModule As String
    Dim vbc As Object ' VBComponents
    Dim vbModule As Object
    Const ThisModuleName As String = "mdl_Helpers"
    
    Set vbc = Application.VBE.ActiveVBProject.VBComponents
    strFile = Dir$(CurrentProject.Path & "\Modules\")
    Do While Len(strFile) > 0
        On Error Resume Next
        strModule = Left$(strFile, InStr(strFile, ".") - 1)
        If strModule <> ThisModuleName Then
            Set vbModule = vbc.Item(strModule)
            If Not vbModule Is Nothing Then
                Application.VBE.ActiveVBProject.VBComponents.Remove vbModule
            End If
            On Error GoTo 0
            Application.VBE.ActiveVBProject.VBComponents.Import CurrentProject.Path & "\Modules\" & strFile
        End If
        strFile = Dir
    Loop
    Application.DoCmd.RunCommand (acCmdCompileAndSaveAllModules)
    
    MsgBox "Import done", vbInformation, "Better Access PivotTable"

End Sub

