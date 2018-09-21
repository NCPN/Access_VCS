Attribute VB_Name = "VCS_Reference"
Option Compare Database

Option Private Module
Option Explicit


' Import References from a CSV, true=SUCCESS
Public Function ImportReferences(ByVal obj_path As String) As Boolean
    Dim FSO As Object
    Dim InFile As Object
    Dim line As String
    Dim Item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim FileName As String
    Dim refName As String
    
    FileName = Dir$(obj_path & "references.csv")
    If Len(FileName) = 0 Then
        ImportReferences = False
        Exit Function
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(obj_path & FileName, iomode:=ForReading, Create:=False, Format:=TristateFalse)
    
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        Item = Split(line, ",")
        If UBound(Item) = 2 Then 'a ref with a guid
          GUID = Trim$(Item(0))
          Major = CLng(Item(1))
          Minor = CLng(Item(2))
          Application.References.AddFromGuid GUID, Major, Minor
        Else
          refName = Trim$(Item(0))
          Application.References.AddFromFile refName
        End If
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    ImportReferences = True
    Exit Function
    
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
    
End Function

' Export References to a CSV
Public Sub ExportReferences(ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim line As String
    Dim ref As Reference

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(obj_path & "references.csv", overwrite:=True, Unicode:=False)
    For Each ref In Application.References
        If ref.GUID <> vbNullString Then ' references of types mdb,accdb,mde etc don't have a GUID
            If Not ref.BuiltIn Then
                line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
                OutFile.WriteLine line
            End If
        Else
            line = ref.fullPath
            OutFile.WriteLine line
        End If
    Next
    OutFile.Close
End Sub
