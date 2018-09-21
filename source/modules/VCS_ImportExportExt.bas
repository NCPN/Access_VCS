Attribute VB_Name = "VCS_ImportExportExt"
Option Compare Database
Option Explicit

' =================================
' MODULE:       VCS_ImportExportExt
' Level:        VCS Development module
' Version:      1.00
'
' Description:  VCS import/export related functions & procedures for version control of
'               databases that are external to the VCS.accdb
'
' Source/date:  Bonnie Campbell, 5/3/2017
' Revisions:    BLC - 5/3/2017 - 1.00 - initial version
' =================================

' ===================================================================================
'  NOTE:
'       Functions and subroutines within this module are for development version control
'       purposes.
'
'       This module is necessary for calling VCS functions from a different database
'       and having VCS respond so that it imports/exports into/from that external db
'
' ===================================================================================

' ---------------------------------
' Declarations
' ---------------------------------
' NOTE: For external dbs Global Variables vs. Const are used
' ---------------------------------

' ---------
'  EXPORT
' ---------
' Usage: ExportAllSourceExternal()
Public g_INCLUDE_DB_TABLES As String   ' Lookup tables exported w/ source code (part of app vs. data)
Private Const ArchiveMyself As Boolean = False  'Causes the VCS_ code to be exported

' ---------
'  IMPORT
' ---------
' Usage: ImportAllSource
Private Const DebugOutput As Boolean = False

' ---------------------------------
'  Database Object Methods/Functions
' ---------------------------------


' ---------------------------------
' SUB:          ExportAllSourceExt
' Description:  Exports all source identified in the external db to `source` folder
'               includes: forms, reports, queries, macros, modules & lookup tables
' Assumptions:  Lookup tables are identified by setting g_INCLUDE_DB_TABLES for
'               those to include
' Parameters:   ExtDbFullPath - full path of database to export (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, May 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 5/3/2017 - initial version
' ---------------------------------
Public Sub ExportAllSourceExt(CallingDb As Database, ExtDbPath As String)
On Error GoTo Err_Handler

    Dim app As Access.Application
    Dim db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean

    Set db = CallingDb

    Set app = Access.Application

    CloseFormsReportsExt CallingDb, ExtDbPath
    'InitUsingUcs2

    With app
        source_path = ExtDbPath & "\extsource\"   'VCS_Dir.ProjectPath() & "source\"
        VCS_Dir.MkDirIfNotExist source_path
    
        Debug.Print
    
  '----------------
  '  Export Queries
  '----------------
        obj_path = source_path & "queries\"
        VCS_Dir.ClearTextFilesFromDir obj_path, "bas"
        Debug.Print VCS_String.PadRight("Exporting queries...", 24);
        
        obj_count = 0
        For Each qry In db.QueryDefs
            DoEvents
            If Left$(qry.Name, 1) <> "~" Then
                VCS_IE_Functions.ExportObject acQuery, qry.Name, obj_path & qry.Name & ".bas", UsingUcs2Ext(CallingDb) 'VCS_File.UsingUcs2
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print VCS_String.PadRight("Sanitizing...", 15);
        VCS_IE_Functions.SanitizeTextFiles obj_path, "bas"
        Debug.Print "[" & obj_count & "]"
    
  '----------------
  '  Export Forms, Reports, Macros, Modules
  '----------------
        
        For Each obj_type In Split( _
            "forms|Forms|" & acForm & "," & _
            "reports|Reports|" & acReport & "," & _
            "macros|Scripts|" & acMacro & "," & _
            "modules|Modules|" & acModule _
            , "," _
        )
            obj_type_split = Split(obj_type, "|")
            obj_type_label = obj_type_split(0)
            obj_type_name = obj_type_split(1)
            obj_type_num = val(obj_type_split(2))
            obj_path = source_path & obj_type_label & "\"
            obj_count = 0
            
            VCS_Dir.ClearTextFilesFromDir obj_path, "bas"
            
            Debug.Print VCS_String.PadRight("Exporting " & obj_type_label & "...", 24);
            For Each doc In db.Containers(obj_type_name).Documents
                DoEvents
                If (Left$(doc.Name, 1) <> "~") And _
                   (IsNotVCS(doc.Name) Or ArchiveMyself) Then
                    If obj_type_label = "modules" Then
                        ucs2 = False
                    Else
                        ucs2 = VCS_File.UsingUcs2
                    End If
                    VCS_IE_Functions.ExportObject obj_type_num, doc.Name, obj_path & doc.Name & ".bas", ucs2
                    
                    If obj_type_label = "reports" Then
                        VCS_Report.ExportPrintVars doc.Name, obj_path & doc.Name & ".pv"
                    End If
                    
                    obj_count = obj_count + 1
                End If
            Next
    
            Debug.Print VCS_String.PadRight("Sanitizing...", 15);
            
  '----------------
  '  Export Modules
  '----------------
            If obj_type_label <> "modules" Then
                VCS_IE_Functions.SanitizeTextFiles obj_path, "bas"
            End If
            Debug.Print "[" & obj_count & "]"
        Next
        
        VCS_Reference.ExportReferences source_path
    
  '----------------
  '  Export Tables
  '----------------
        obj_path = source_path & "tables\"
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        VCS_Dir.ClearTextFilesFromDir obj_path, "txt"
        
        Dim td As DAO.TableDef
        Dim tds As DAO.TableDefs
        Set tds = db.TableDefs
    
        obj_type_label = "tbldef"
        obj_type_name = "Table_Def"
        obj_type_num = acTable
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        obj_data_count = 0
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        
        'move these into Table and DataMacro modules?
        ' - We don't want to determin file extentions here - or obj_path either!
        VCS_Dir.ClearTextFilesFromDir obj_path, "sql"
        VCS_Dir.ClearTextFilesFromDir obj_path, "xml"
        VCS_Dir.ClearTextFilesFromDir obj_path, "LNKD"
  
  '----------------
  '  Export Lookup Tables
  '----------------
        Dim IncludeTablesCol As Collection
        Set IncludeTablesCol = StrSetToCol(g_INCLUDE_DB_TABLES, ",")
        
        Debug.Print VCS_String.PadRight("Exporting " & obj_type_label & "...", 24);
        
        For Each td In tds
            ' Exclude system & temporary tables
            If Left$(td.Name, 4) <> "MSys" And _
               Left$(td.Name, 1) <> "~" Then
                If Len(td.connect) = 0 Then ' this is not an external table
                    VCS_Table.ExportTableDef db, td, td.Name, obj_path
                    
                    'check for lookup tables
                    If g_INCLUDE_DB_TABLES = "*" Then
                        DoEvents
                        VCS_Table.ExportTableData CStr(td.Name), source_path & "tables\"
                        If Len(Dir$(source_path & "tables\" & td.Name & ".txt")) > 0 Then
                            obj_data_count = obj_data_count + 1
                        End If
                    
                    ElseIf (Len(Replace(g_INCLUDE_DB_TABLES, " ", vbNullString)) > 0) And _
                                        g_INCLUDE_DB_TABLES <> "*" Then
                        DoEvents
                        On Error GoTo Err_TableNotFound
                        If InCollection(IncludeTablesCol, td.Name) Then
                            VCS_Table.ExportTableData CStr(td.Name), source_path & "tables\"
                            obj_data_count = obj_data_count + 1
                        End If

Err_TableNotFound:
                        
                    'else don't export table data
                    End If
                Else
                    VCS_Table.ExportLinkedTable td.Name, obj_path
                End If
                
                obj_count = obj_count + 1
                
            End If
        Next
        Debug.Print "[" & obj_count & "]"
        If obj_data_count > 0 Then
          Debug.Print VCS_String.PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
        End If
        
  '----------------
  '  Relationships
  '----------------
        Debug.Print VCS_String.PadRight("Exporting Relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        VCS_Dir.MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
    
        VCS_Dir.ClearTextFilesFromDir obj_path, "txt"
    
        Dim aRelation As DAO.Relation
        
        For Each aRelation In CallingDb.Relations 'CurrentDb.Relations
            ' Exclude relations from system tables and inherited (linked) relations
            If Not (aRelation.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                    Or aRelation.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                    Or (aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                    DAO.RelationAttributeEnum.dbRelationInherited) Then
                VCS_Relation.ExportRelation aRelation, obj_path & aRelation.Name & ".txt"
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"
        
        Debug.Print "Done."
    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportAllSourceExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Sub

''returns true if named module is NOT part of the VCS code
Private Function IsNotVCS(ByVal Name As String) As Boolean
    If Name <> "VCS_ImportExport" And _
      Name <> "VCS_IE_Functions" And _
      Name <> "VCS_File" And _
      Name <> "VCS_Dir" And _
      Name <> "VCS_String" And _
      Name <> "VCS_Loader" And _
      Name <> "VCS_Table" And _
      Name <> "VCS_Reference" And _
      Name <> "VCS_DataMacro" And _
      Name <> "VCS_Report" And _
      Name <> "VCS_Relation" Then
        IsNotVCS = True
    Else
        IsNotVCS = False
    End If

End Function



' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSourceExt(CallingDb As Database, ExtDbPath As String)
On Error GoTo Err_Handler

    Dim FSO As Object
    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim FileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    Set FSO = CreateObject("Scripting.FileSystemObject")

    CloseFormsReportsExt CallingDb, ExtDbPath
    'InitUsingUcs2

    source_path = ExtDbPath & "source\" 'VCS_Dir.ProjectPath() & "source\"
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print
    
    If Not VCS_Reference.ImportReferences(source_path) Then
        Debug.Print "Info: no references file in " & source_path
        Debug.Print
    End If

    obj_path = source_path & "queries\"
    FileName = Dir$(obj_path & "*.bas")
    
    Dim tempFilePath As String
    tempFilePath = VCS_File.TempFile()
    
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            VCS_IE_Functions.ImportObject acQuery, obj_name, obj_path & FileName, VCS_File.UsingUcs2
            VCS_IE_Functions.ExportObject acQuery, obj_name, tempFilePath, VCS_File.UsingUcs2
            VCS_IE_Functions.ImportObject acQuery, obj_name, tempFilePath, VCS_File.UsingUcs2
            obj_count = obj_count + 1
            FileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    VCS_Dir.DelIfExist tempFilePath

    ' restore table definitions
    obj_path = source_path & "tbldef\"
    FileName = Dir$(obj_path & "*.sql")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.PadRight("Importing tabledefs...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            If DebugOutput Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            VCS_Table.ImportTableDef CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    ' restore linked tables - we must have access to the remote store to import these!
    FileName = Dir$(obj_path & "*.LNKD")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.PadRight("Importing Linked tabledefs...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            If DebugOutput Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            VCS_Table.ImportLinkedTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    
    ' NOW we may load data
    obj_path = source_path & "tables\"
    FileName = Dir$(obj_path & "*.txt")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.PadRight("Importing tables...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            VCS_Table.ImportTableData CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    'load Data Macros - not DRY!
    obj_path = source_path & "tbldef\"
    FileName = Dir$(obj_path & "*.xml")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.PadRight("Importing Data Macros...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            'VCS_Table.ImportTableData CStr(obj_name), obj_path
            VCS_DataMacro.ImportDataMacros obj_name, obj_path
            obj_count = obj_count + 1
            FileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    

        'import Data Macros
    

    For Each obj_type In Split( _
        "forms|" & acForm & "," & _
        "reports|" & acReport & "," & _
        "macros|" & acMacro & "," & _
        "modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_num = val(obj_type_split(1))
        obj_path = source_path & obj_type_label & "\"
         
            
        FileName = Dir$(obj_path & "*.bas")
        If Len(FileName) > 0 Then
            Debug.Print VCS_String.PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(FileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = VCS_File.UsingUcs2
                End If
                If IsNotVCS(obj_name) Then
                    VCS_IE_Functions.ImportObject obj_type_num, obj_name, obj_path & FileName, ucs2
                    obj_count = obj_count + 1
                Else
                    If ArchiveMyself Then
                            MsgBox "Module " & obj_name & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
                    End If
                End If
                FileName = Dir$()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    
    'import Print Variables
    Debug.Print VCS_String.PadRight("Importing Print Vars...", 24);
    obj_count = 0
    
    obj_path = source_path & "reports\"
    FileName = Dir$(obj_path & "*.pv")
    Do Until Len(FileName) = 0
        DoEvents
        obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
        VCS_Report.ImportPrintVars obj_name, obj_path & FileName
        obj_count = obj_count + 1
        FileName = Dir$()
    Loop
    Debug.Print "[" & obj_count & "]"
    
    'import relations
    Debug.Print VCS_String.PadRight("Importing Relations...", 24);
    obj_count = 0
    obj_path = source_path & "relations\"
    FileName = Dir$(obj_path & "*.txt")
    Do Until Len(FileName) = 0
        DoEvents
        VCS_Relation.ImportRelation obj_path & FileName
        obj_count = obj_count + 1
        FileName = Dir$()
    Loop
    Debug.Print "[" & obj_count & "]"
    DoEvents
    
    Debug.Print "Done."
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImportAllSourceExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Sub

' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProjectExt(CallingDb As Database, ExtDbPath As String)
On Error GoTo Err_Handler

    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr$(149) & " Tables" & vbCrLf & _
              Chr$(149) & " Forms" & vbCrLf & _
              Chr$(149) & " Macros" & vbCrLf & _
              Chr$(149) & " Modules" & vbCrLf & _
              Chr$(149) & " Queries" & vbCrLf & _
              Chr$(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CallingDb
    
    CloseFormsReportsExt CallingDb, ExtDbPath

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print
    
    Dim rel As DAO.Relation
    For Each rel In CurrentDb.Relations
        If Not (rel.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                rel.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
            CurrentDb.Relations.Delete (rel.Name)
        End If
    Next

    Dim dbObject As Object
    For Each dbObject In db.QueryDefs
        DoEvents
        If Left$(dbObject.Name, 1) <> "~" Then
'            Debug.Print dbObject.Name
            db.QueryDefs.Delete dbObject.Name
        End If
    Next
    
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If Left$(td.Name, 4) <> "MSys" And _
            Left$(td.Name, 1) <> "~" Then
            CurrentDb.TableDefs.Delete (td.Name)
        End If
    Next

    Dim ObjType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME As Byte = 0
    Const OTID As Byte = 1

    For Each ObjType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(ObjType, "|")
        DoEvents
        For Each doc In db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left$(doc.Name, 1) <> "~") And _
               (IsNotVCS(doc.Name)) Then
'                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.Name
            End If
        Next
    Next
    
    Debug.Print "================="
    Debug.Print "Importing Project"
    ImportAllSource
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImportProjectExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Sub

' Expose for use as function, can be called by query
Public Sub make()
    ImportProject
End Sub

Public Function Greetings(strName As String)
    MsgBox ("Hello, " & strName & "!"), vbInformation, "Greetings from" & Application.CurrentDb.Name
End Function

'===================================================================================================================================
'-----------------------------------------------------------'
' Helper Functions - these should be put in their own files '
'-----------------------------------------------------------'








' ---------------------------------
' SUB:          InCollectionExt
' Description:  Determine if an item or key is in a collection in the external database
' Assumptions:  -
' Parameters:   db - datbase to check (Database)
'               col - collection (Collection)
'               vItem - item to search for (variant)
'               vKey - key to search for (variant)
' Returns:      True or False depending on whether the item is in the collection (Boolean)
' Throws:       none
' References:   -
' Source/date:  VCS
' Adapted:      Bonnie Campbell, May 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 5/3/2017 - initial version
' ---------------------------------
Public Function InCollectionExt(db As Database, col As Collection, Optional vItem, Optional vKey) As Boolean
On Error GoTo Err_Handler

    Dim vColItem As Variant

    InCollectionExt = False

    If Not IsMissing(vKey) Then
        col.Item vKey

        '5 if not in collection, it is 91 if no collection exists
        If Err.Number <> 5 And Err.Number <> 91 Then
            InCollectionExt = True
        End If
    ElseIf Not IsMissing(vItem) Then
        For Each vColItem In col
            If vColItem = vItem Then
                InCollectionExt = True
                GoTo Exit_Handler
            End If
        Next vColItem
    End If

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - InCollectionExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          StrSetToColExt
' Description:  Converts string to column
' Assumptions:  -
' Parameters:   ExtDbFullPath - full path of database to export (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, May 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 5/3/2017 - initial version
' ---------------------------------
Public Function StrSetToColExt(ByVal strSet As String, ByVal delimiter As String) As Collection
 'throws errorsOn Error GoTo Err_Handler
'errno 457 - duplicate key (& item)

    Dim strSetArray() As String
    Dim col As Collection
    
    Set col = New Collection
    strSetArray = Split(strSet, delimiter)
    
    Dim Item As Variant
    For Each Item In strSetArray
        col.Add Item, Item
    Next
    
    Set StrSetToColExt = col

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - StrSetToColExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          CloseFormsReportsExt
' Description:  Closes forms & reports in an external access db
' Assumptions:  -
' Parameters:   Db - database to close forms/reports in (database)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, May 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 5/3/2017 - initial version
' ---------------------------------
Public Sub CloseFormsReportsExt(CallingDb As Database, ExtDbPath As String)
On Error GoTo Err_Handler

    Dim app As Access.Application
    Dim db As Database
    
    Set db = CallingDb
    Set app = Access.Application
    
    With app
        Do While .Forms.Count > 0
            .DoCmd.Close acForm, .Forms(0).Name
            DoEvents
        Loop
        Do While .Reports.Count > 0
            .DoCmd.Close acReport, .Reports(0).Name
            DoEvents
        Loop
    
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseFormsReportsExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Sub
' ---------------------------------
' SUB:          ExportAllSourceExt
' Description:  Exports all source identified in the external db to `source` folder
'               includes: forms, reports, queries, macros, modules & lookup tables
' Assumptions:  Lookup tables are identified by setting g_INCLUDE_DB_TABLES for
'               those to include
' Parameters:   ExtDbFullPath - full path of database to export (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, May 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 5/3/2017 - initial version
' ---------------------------------
Public Sub XX(CallingDb As Database, ExtDbPath As String)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportAllSourceExt[VCS_ImportExportExt])"
    End Select
    Resume Exit_Handler
End Sub

' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Public Function UsingUcs2Ext(CallingDb As Database) As Boolean
    Dim obj_name As String
    Dim obj_type As Variant
    Dim fn As Integer
    Dim bytes As String
    Dim obj_type_split() As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    
    If CallingDb.QueryDefs.Count > 0 Then
        obj_type_num = acQuery
        obj_name = CallingDb.QueryDefs(0).Name
    Else
        For Each obj_type In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
        )
            DoEvents
            obj_type_split = Split(obj_type, "|")
            obj_type_name = obj_type_split(0)
            obj_type_num = val(obj_type_split(1))
            If CallingDb.Containers(obj_type_name).Documents.Count > 0 Then
                obj_name = CallingDb.Containers(obj_type_name).Documents(0).Name
                Exit For
            End If
        Next
    End If

    If obj_name = vbNullString Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2Ext = True
        Exit Function
    End If

    Dim tempFileName As String
    tempFileName = VCS_File.TempFile()
    
    Application.SaveAsText obj_type_num, obj_name, tempFileName
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    If Asc(Mid$(bytes, 1, 1)) = &HFF And Asc(Mid$(bytes, 2, 1)) = &HFE Then
        UsingUcs2Ext = True
    Else
        UsingUcs2Ext = False
    End If
    Close fn
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFile (tempFileName)
End Function


