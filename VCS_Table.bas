Attribute VB_Name = "VCS_Table"
Option Compare Database
Option Explicit

' =================================
' MODULE:       VCS_Table
' Description:  Version control related functions & procedures for version control
' Source:
' Tim Abell
' https://github.com/timabell/msaccess-vcs-integration
' Armin Bra (deprecated)
' https://github.com/ArminBra/msaccess-vcs-integration
' http://stackoverflow.com/questions/187506/how-do-you-use-version-control-with-access-development
' Adapted:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 2/12/2015 - initial version
' =================================

' ===================================================================================
'  NOTE:
'  Backs up & regenerates components w/ functions using SaveAsText / LoadFromText
' ===================================================================================

'this is used in import table def
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' Structure to keep track of "on Update" and "on Delete" clauses
' Access does not in all cases execute such queries
Private Type structEnforce
    foreignTable As String
    foreignFields() As String
    table As String
    refFields() As String
    isUpdate As Boolean
End Type

' keeping "on Update" relations to be complemented after table creation
Private K() As structEnforce


Public Sub ExportLinkedTable(tbl_name As String, obj_path As String)
    On Error GoTo Err_LinkedTable:
    
    Dim tempFilePath As String
    
    tempFilePath = VCS_File.TempFile()
    
    Dim FSO, OutFile

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.MkDirIfNotExist obj_path
    
    Set OutFile = FSO.CreateTextFile(tempFilePath, True, True)
    
    OutFile.Write CurrentDb.tabledefs(tbl_name).name
    OutFile.Write vbCrLf
    OutFile.Write CurrentDb.tabledefs(tbl_name).Connect
    OutFile.Write vbCrLf
    OutFile.Write CurrentDb.tabledefs(tbl_name).SourceTableName
    OutFile.Write vbCrLf
    

    Dim db As Database
    Set db = CurrentDb
    Dim td As TableDef
    Set td = db.tabledefs(tbl_name)
    Dim idx As index
    
    For Each idx In td.Indexes
        If idx.Primary Then
            OutFile.Write Right(idx.Fields, Len(idx.Fields) - 1)
            OutFile.Write vbCrLf
        End If

    Next
    

Err_LinkedTable_Fin:
    On Error Resume Next
    OutFile.Close
    'save files as .odbc
    VCS_File.ConvertUcs2Utf8 tempFilePath, obj_path & tbl_name & ".LNKD"
    
    Exit Sub
    
Err_LinkedTable:

    OutFile.Close
    MsgBox Err.Description, vbCritical, "ERROR: EXPORT LINKED TABLE"
    Resume Err_LinkedTable_Fin:
End Sub


' Save a Table Definition as SQL statement
Public Sub ExportTableDef(db As Database, td As TableDef, TableName As String, directory As String)
    Dim fileName As String: fileName = directory & TableName & ".sql"
    Dim sql As String
    Dim fieldAttributeSql As String
    Dim idx As index
    Dim fi As Field
    Dim i As Integer
    Dim f As Field
    Dim rel As Relation
    Dim FSO, OutFile
    Dim ff As Object
    'Debug.Print tableName
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(fileName, True)
    sql = "CREATE TABLE " & strName(TableName) & " (" & vbCrLf
    For Each fi In td.Fields
        sql = sql & "  " & strName(fi.name) & " "
        If (fi.Attributes And dbAutoIncrField) Then
            sql = sql & "AUTOINCREMENT"
        Else
            sql = sql & strType(fi.Type) & " "
        End If
        Select Case fi.Type
            Case dbText, dbVarBinary
                sql = sql & "(" & fi.Size & ")"
            Case Else
        End Select
        For Each idx In td.Indexes
            fieldAttributeSql = ""
            If idx.Fields.count = 1 And idx.Fields(0).name = fi.name Then
                If idx.Primary Then fieldAttributeSql = fieldAttributeSql & " PRIMARY KEY "
                If idx.Unique Then fieldAttributeSql = fieldAttributeSql & " UNIQUE "
                If idx.Required Then fieldAttributeSql = fieldAttributeSql & " NOT NULL "
                If idx.Foreign Then
                    Set ff = idx.Fields
                    fieldAttributeSql = fieldAttributeSql & formatReferences(db, ff, TableName)
                End If
                If Len(fieldAttributeSql) > 0 Then fieldAttributeSql = " CONSTRAINT " & strName(idx.name) & fieldAttributeSql
            End If
        Next
        sql = sql + fieldAttributeSql
        sql = sql + "," & vbCrLf
    Next
    sql = Left(sql, Len(sql) - 3) ' strip off last comma and crlf
    
    Dim constraintSql As String
    For Each idx In td.Indexes
        If idx.Fields.count > 1 Then
            If Len(constraintSql) = 0 Then constraintSql = constraintSql & " CONSTRAINT "
            If idx.Primary Then constraintSql = constraintSql & formatConstraint("PRIMARY KEY", idx)
            If idx.Foreign Then constraintSql = formatConstraint("FOREIGN KEY", idx)
            If Len(constraintSql) > 0 Then
                sql = sql & "," & vbCrLf & "  " & constraintSql
                sql = sql & formatReferences(db, idx.Fields, TableName)
            End If
        End If
    Next
    sql = sql & vbCrLf & ")"

    'Debug.Print sql
    OutFile.WriteLine sql
    
    OutFile.Close
    
    'exort Data Macros
    VCS_DataMacro.ExportDataMacros TableName, directory
    
End Sub

Private Function formatReferences(db As Database, ff As Object, TableName As String)
    Dim rel As Relation
    Dim sql As String
    Dim f As Field
    For Each rel In db.Relations
        If (rel.foreignTable = TableName) Then
         If FieldsIdentical(ff, rel.Fields) Then
          sql = " REFERENCES "
          sql = sql & rel.table & " ("
          For Each f In rel.Fields
            sql = sql & strName(f.name) & ","
          Next
          sql = Left(sql, Len(sql) - 1) & ")"
          If rel.Attributes And dbRelationUpdateCascade Then
            sql = sql + " ON UPDATE CASCADE "
          End If
          If rel.Attributes And dbRelationDeleteCascade Then
            sql = sql + " ON DELETE CASCADE "
          End If
          Exit For
         End If
        End If
    Next
    formatReferences = sql
End Function

Private Function formatConstraint(keyw As String, idx As index) As String
    Dim sql As String
    Dim fi As Field
    
    sql = strName(idx.name) & " " & keyw & " ("
    For Each fi In idx.Fields
        sql = sql & strName(fi.name) & ", "
    Next
    sql = Left(sql, Len(sql) - 2) & ")" 'strip off last comma and close brackets
    
    'return value
    formatConstraint = sql
End Function

Private Function strName(s As String) As String
    strName = "[" & s & "]"
End Function

Private Function strType(i As Integer) As String
    Select Case i
    Case dbLongBinary
        strType = "LONGBINARY"
    Case dbBinary
        strType = "BINARY"
    'Case dbBit missing enum
    '    strType = "BIT"
    Case dbAutoIncrField
        strType = "COUNTER"
    Case dbCurrency
        strType = "CURRENCY"
    Case dbDate, dbTime
        strType = "DATETIME"
    Case dbGUID
        strType = "GUID"
    Case dbMemo
        strType = "LONGTEXT"
    Case dbDouble
        strType = "DOUBLE"
    Case dbSingle
        strType = "SINGLE"
    Case dbByte
        strType = "BYTE" 'was "UNSIGNED BYTE" which results in Run-time error 3292 Syntax error in field definition. on import
    Case dbInteger
        strType = "SHORT"
    Case dbLong
        strType = "LONG"
    Case dbNumeric
        strType = "NUMERIC"
    Case dbText
        strType = "VARCHAR"
    Case Else
        strType = "VARCHAR"
    End Select
End Function
Private Function FieldsIdentical(ff As Object, gg As Object) As Boolean
    Dim f As Field
    If ff.count <> gg.count Then
        FieldsIdentical = False
        Exit Function
    End If
    For Each f In ff
        If Not FieldInFields(f, gg) Then
        FieldsIdentical = False
        Exit Function
        End If
    Next
    FieldsIdentical = True
        
    
End Function

Private Function FieldInFields(fi As Field, ff As Fields) As Boolean
    Dim f As Field
    For Each f In ff
        If f.name = fi.name Then
            FieldInFields = True
            Exit Function
        End If
    Next
    FieldInFields = False
End Function

' Determine if a table or exists.
' based on sample code of support.microsoftcom
' ARGUMENTS:
'    TName: The name of a table or query.
'
' RETURNS: True (it exists) or False (it does not exist).
Private Function TableExists(TName As String) As Boolean
        Dim db As Database, Found As Boolean, test As String
        Const NAME_NOT_IN_COLLECTION = 3265

         ' Assume the table or query does not exist.
        Found = False
        Set db = CurrentDb()

         ' Trap for any errors.
        On Error Resume Next
         
         ' See if the name is in the Tables collection.
        test = db.tabledefs(TName).name
        If Err <> NAME_NOT_IN_COLLECTION Then Found = True

        ' Reset the error variable.
        Err = 0

        TableExists = Found

End Function

' Build SQL to export `tbl_name` sorted by each field from first to last
Private Function TableExportSql(tbl_name As String)
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)
    
    sb = VCS_String.Sb_Init()
    VCS_String.Sb_Append sb, "SELECT "
    count = 0
    For Each fieldObj In rs.Fields
        If count > 0 Then VCS_String.Sb_Append sb, ", "
        VCS_String.Sb_Append sb, "[" & fieldObj.name & "]"
        count = count + 1
    Next
    VCS_String.Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    count = 0
    For Each fieldObj In rs.Fields
        DoEvents
        If count > 0 Then VCS_String.Sb_Append sb, ", "
        VCS_String.Sb_Append sb, "[" & fieldObj.name & "]"
        count = count + 1
    Next

    TableExportSql = VCS_String.Sb_Get(sb)

End Function

' Export the lookup table `tblName` to `source\tables`.
Public Sub ExportTableData(tbl_name As String, obj_path As String)
    Dim FSO, OutFile
    Dim rs As Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim c As Long, Value As Variant
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset(TableExportSql(tbl_name))
    If rs.RecordCount = 0 Then
        Debug.Print "Error: Table " & tbl_name & "  empty"
        rs.Close
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.MkDirIfNotExist obj_path
    Dim tempFileName As String: tempFileName = VCS_File.TempFile()

    Set OutFile = FSO.CreateTextFile(tempFileName, True, True)

    c = 0
    For Each fieldObj In rs.Fields
        If c <> 0 Then OutFile.Write vbTab
        c = c + 1
        OutFile.Write fieldObj.name
    Next
    OutFile.Write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        c = 0
        For Each fieldObj In rs.Fields
            DoEvents
            If c <> 0 Then OutFile.Write vbTab
            c = c + 1
            Value = rs(fieldObj.name)
            If IsNull(Value) Then
                Value = ""
            Else
                Value = Replace(Value, "\", "\\")
                Value = Replace(Value, vbCrLf, "\n")
                Value = Replace(Value, vbCr, "\n")
                Value = Replace(Value, vbLf, "\n")
                Value = Replace(Value, vbTab, "\t")
            End If
            OutFile.Write Value
        Next
        OutFile.Write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close

    VCS_File.ConvertUcs2Utf8 tempFileName, obj_path & tbl_name & ".txt"
    FSO.DeleteFile tempFileName
End Sub
' Kill Table if Exists
Private Sub KillTable(tblName As String, db As Object)
    If TableExists(tblName) Then
        db.Execute "DROP TABLE [" & tblName & "]"
    End If
End Sub


Public Sub ImportLinkedTable(tblName As String, obj_path As String)
    Dim db As Database ' DAO.Database
    Dim FSO, InFile As Object
    
    Set db = CurrentDb
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFilePath As String
    tempFilePath = TempFile()
    
    ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFilePath, ForReading, False, TristateTrue)
    
    On Error GoTo err_notable:
    DoCmd.DeleteObject acTable, tblName
    
    GoTo err_notable_fin:
err_notable:
    Err.Clear
    Resume err_notable_fin:
err_notable_fin:
    On Error GoTo Err_CreateLinkedTable:
    
    Dim td As TableDef
    Set td = db.CreateTableDef(InFile.ReadLine())
    td.Connect = InFile.ReadLine()
    td.SourceTableName = InFile.ReadLine()
    db.tabledefs.Append td
    
    GoTo Err_CreateLinkedTable_Fin:
    
Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin:
Err_CreateLinkedTable_Fin:

    'this will throw errors if a primary key already exists or the table is linked to an access database table
    'will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:
    
    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim Field As Variant
    Dim sql As String
    sql = "CREATE INDEX __uniqueindex ON " & td.name & " ("
    
    For Each Field In Split(Fields, ";+")
        sql = sql & "[" & Field & "]" & ","
    Next
    'remove extraneous comma
    sql = Left(sql, Len(sql) - 1)
    
    sql = sql & ") WITH PRIMARY"
    CurrentDb.Execute sql
    
Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
    
   

End Sub





' Import Table Definition
Public Sub ImportTableDef(tblName As String, directory As String)
    Dim filePath As String: filePath = directory & tblName & ".sql"
    Dim db As Object ' DAO.Database
    Dim FSO, InFile As Object
    Dim buf As String
    Dim p As Integer
    Dim p1 As Integer
    Dim strFields() As String
    Dim strRef As String
    Dim strMsg As String
    Dim strForeignKeys() As String
    Dim s
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tempFileName As String: tempFileName = VCS_File.TempFile()

    n = -1
    Set FSO = CreateObject("Scripting.FileSystemObject")
    VCS_File.ConvertUtf8Ucs2 filePath, tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set db = CurrentDb
    KillTable tblName, db
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = buf & InFile.ReadLine()
    Loop
    
    ' The following block is needed because "on update" actions may cause problems
    For Each s In Split("UPDATE|DELETE", "|")
    p = InStr(buf, "ON " & s & " CASCADE")
    While p > 0
        n = n + 1
        ReDim Preserve K(n)
        K(n).table = tblName
        K(n).isUpdate = (s = "UPDATE")
        
        buf = Left(buf, p - 1) & Mid(buf, p + 18)
        p = InStrRev(buf, "REFERENCES", p)
        p1 = InStr(p, buf, "(")
        K(n).foreignFields = Split(VCS_String.SubString(p1, buf, "(", ")"), ",")
        K(n).foreignTable = Trim(Mid(buf, p + 10, p1 - p - 10))
        p = InStrRev(buf, "CONSTRAINT", p1)
        p1 = InStrRev(buf, "FOREIGN KEY", p1)
        If (p1 > 0) And (p > 0) And (p1 > p) Then
        ' multifield index
            K(n).refFields = Split(VCS_String.SubString(p1, buf, "(", ")"), ",")
        ElseIf p1 = 0 Then
        ' single field
        End If
        p = InStr(p, "ON " & s & " CASCADE", buf)
    Wend
    Next
    On Error Resume Next
    For i = 0 To n
        strMsg = K(i).table & " to " & K(i).foreignTable
        strMsg = strMsg & "(  "
        For j = 0 To UBound(K(i).refFields)
            strMsg = strMsg & K(i).refFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") to ("
        For j = 0 To UBound(K(i).foreignFields)
            strMsg = strMsg & K(i).foreignFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") Check "
        If K(i).isUpdate Then
            strMsg = strMsg & " on update cascade " & vbCrLf
        Else
            strMsg = strMsg & " on delete cascade " & vbCrLf
        End If
    Next
    On Error GoTo 0
    db.Execute buf
    InFile.Close
    If Len(strMsg) > 0 Then MsgBox strMsg, vbOKOnly, "Correct manually"
    
    'import Data Macros
    VCS_DataMacro.ImportDataMacros tblName, directory
    
End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub ImportTableData(tblName As String, obj_path As String)
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO, InFile As Object
    Dim c As Long, buf As String, Values() As String, Value As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFileName As String: tempFileName = VCS_File.TempFile()
    VCS_File.ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set db = CurrentDb

    db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim(buf)) > 0 Then
            Values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                Value = Values(c)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.name) = Value
                c = c + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
End Sub

