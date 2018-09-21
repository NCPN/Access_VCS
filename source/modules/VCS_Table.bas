Attribute VB_Name = "VCS_Table"
Option Compare Database

Option Private Module
Option Explicit

' --------------------------------
' Structures
' --------------------------------

' Structure to keep track of "on Update" and "on Delete" clauses
' Access does not in all cases execute such queries
Private Type structEnforce
    foreignTable As String
    foreignFields() As String
    table As String
    refFields() As String
    IsUpdate As Boolean
End Type

' keeping "on Update" relations to be complemented after table creation
Private K() As structEnforce


Public Sub ExportLinkedTable(ByVal tbl_name As String, ByVal obj_path As String)
    On Error GoTo Err_LinkedTable
    
    Dim tempFilePath As String
    
    tempFilePath = VCS_File.TempFile()
    
    Dim FSO As Object
    Dim OutFile As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.MkDirIfNotExist obj_path
    
    Set OutFile = FSO.CreateTextFile(tempFilePath, overwrite:=True, Unicode:=True)
    
    OutFile.Write CurrentDb.TableDefs(tbl_name).Name
    OutFile.Write vbCrLf
    
    If InStr(1, CurrentDb.TableDefs(tbl_name).connect, "DATABASE=" & CurrentProject.Path) Then
        'change to relatave path
        Dim connect() As String
        connect = Split(CurrentDb.TableDefs(tbl_name).connect, CurrentProject.Path)
        OutFile.Write connect(0) & "." & connect(1)
    Else
        OutFile.Write CurrentDb.TableDefs(tbl_name).connect
    End If
    
    OutFile.Write vbCrLf
    OutFile.Write CurrentDb.TableDefs(tbl_name).SourceTableName
    OutFile.Write vbCrLf
    
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    Set td = db.TableDefs(tbl_name)
    Dim idx As DAO.Index
    
    For Each idx In td.Indexes
        If idx.Primary Then
            OutFile.Write Right$(idx.Fields, Len(idx.Fields) - 1)
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
    Resume Err_LinkedTable_Fin
End Sub

' This requires Microsoft ADO Ext. 2.x for DLL and Security
' See reference: https://social.msdn.microsoft.com/Forums/office/en-US/883087ba-2c25-4571-bd3c-706061466a11/how-can-i-programmatically-access-scale-property-of-a-decimal-data-type-field?forum=accessdev
Private Function formatDecimal(ByVal Tablename As String, ByVal fieldName As String) As String

    Dim cnn As ADODB.Connection
    Dim cat As ADOX.Catalog
    Dim col As ADOX.Column
    
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    

    Set cnn = CurrentProject.Connection
    Set cat.ActiveConnection = cnn

    Set col = cat.Tables(Tablename).Columns(fieldName)

    formatDecimal = "(" & col.Precision & ", " & col.NumericScale & ")"

    Set col = Nothing
    Set cat = Nothing
    Set cnn = Nothing

End Function

' Save a Table Definition as SQL statement
Public Sub ExportTableDef(db As DAO.Database, td As DAO.TableDef, ByVal Tablename As String, _
                          ByVal directory As String)
    Dim FileName As String
    FileName = directory & Tablename & ".sql"
    Dim SQL As String
    Dim fieldAttributeSql As String
    Dim idx As DAO.Index
    Dim fi As DAO.field
    Dim FSO As Object
    Dim OutFile As Object
    Dim ff As Object
    'Debug.Print tableName
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(FileName, overwrite:=True, Unicode:=False)
    SQL = "CREATE TABLE " & strName(Tablename) & " (" & vbCrLf
    For Each fi In td.Fields
        SQL = SQL & "  " & strName(fi.Name) & " "
        If (fi.Attributes And dbAutoIncrField) Then
            SQL = SQL & "AUTOINCREMENT"
        Else
            SQL = SQL & strType(fi.Type) & " "
        End If
        Select Case fi.Type
            Case dbText, dbVarBinary
                SQL = SQL & "(" & fi.Size & ")"
            Case dbDecimal
                SQL = SQL & formatDecimal(Tablename, fi.Name)
            Case Else
        End Select
        For Each idx In td.Indexes
            fieldAttributeSql = vbNullString
            If idx.Fields.Count = 1 And idx.Fields(0).Name = fi.Name Then
                If idx.Primary Then fieldAttributeSql = fieldAttributeSql & " PRIMARY KEY "
                If idx.Unique Then fieldAttributeSql = fieldAttributeSql & " UNIQUE "
                If idx.Required Then fieldAttributeSql = fieldAttributeSql & " NOT NULL "
                If idx.Foreign Then
                    Set ff = idx.Fields
                    fieldAttributeSql = fieldAttributeSql & formatReferences(db, ff, Tablename)
                End If
                If Len(fieldAttributeSql) > 0 Then fieldAttributeSql = " CONSTRAINT " & strName(idx.Name) & fieldAttributeSql
            End If
            SQL = SQL & fieldAttributeSql
        Next
        SQL = SQL & "," & vbCrLf
    Next
    SQL = Left$(SQL, Len(SQL) - 3) ' strip off last comma and crlf
    
    Dim constraintSql As String
    For Each idx In td.Indexes
        If idx.Fields.Count > 1 Then
            If Len(constraintSql) = 0 Then constraintSql = constraintSql & " CONSTRAINT "
            If idx.Primary Then constraintSql = constraintSql & formatConstraint("PRIMARY KEY", idx)
            If Not idx.Foreign Then
                If Len(constraintSql) > 0 Then
                    SQL = SQL & "," & vbCrLf & "  " & constraintSql
                    SQL = SQL & formatReferences(db, idx.Fields, Tablename)
                End If
            End If
        End If
    Next
    SQL = SQL & vbCrLf & ")"

    'Debug.Print sql
    OutFile.WriteLine SQL
    
    OutFile.Close
    
    'exort Data Macros
    VCS_DataMacro.ExportDataMacros Tablename, directory
    
End Sub

Private Function formatReferences(db As DAO.Database, ff As Object, _
                                  ByVal Tablename As String) As String

    Dim rel As DAO.Relation
    Dim SQL As String
    Dim f As DAO.field
    
    For Each rel In db.Relations
        If (rel.foreignTable = Tablename) Then
         If FieldsIdentical(ff, rel.Fields) Then
          SQL = " REFERENCES "
          SQL = SQL & strName(rel.table) & " ("
          For Each f In rel.Fields
            SQL = SQL & strName(f.Name) & ","
          Next
          SQL = Left$(SQL, Len(SQL) - 1) & ")"
          If rel.Attributes And dbRelationUpdateCascade Then
            SQL = SQL + " ON UPDATE CASCADE "
          End If
          If rel.Attributes And dbRelationDeleteCascade Then
            SQL = SQL + " ON DELETE CASCADE "
          End If
          Exit For
         End If
        End If
    Next
    
    formatReferences = SQL
End Function

Private Function formatConstraint(ByVal keyw As String, ByVal idx As DAO.Index) As String
    Dim SQL As String
    Dim fi As DAO.field
    
    SQL = strName(idx.Name) & " " & keyw & " ("
    For Each fi In idx.Fields
        SQL = SQL & strName(fi.Name) & ", "
    Next
    SQL = Left$(SQL, Len(SQL) - 2) & ")" 'strip off last comma and close brackets
    
    'return value
    formatConstraint = SQL
End Function

Private Function strName(ByVal s As String) As String
    strName = "[" & s & "]"
End Function

Private Function strType(ByVal i As Integer) As String
    Select Case i
    Case dbLongBinary
        strType = "LONGBINARY"
    Case dbBinary
        strType = "BINARY"
    Case dbBoolean
        strType = "BIT"
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
        strType = "BYTE"
    Case dbInteger
        strType = "SHORT"
    Case dbLong
        strType = "LONG"
    Case dbNumeric
        strType = "NUMERIC"
    Case dbText
        strType = "VARCHAR"
    Case dbDecimal
        strType = "DECIMAL"
    Case Else
        strType = "VARCHAR"
    End Select
End Function

Private Function FieldsIdentical(ff As Object, gg As Object) As Boolean
    Dim f As DAO.field
    If ff.Count <> gg.Count Then
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

Private Function FieldInFields(fi As DAO.field, ff As DAO.Fields) As Boolean
    Dim f As DAO.field
    For Each f In ff
        If f.Name = fi.Name Then
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
Private Function TableExists(ByVal TName As String) As Boolean
    Dim db As DAO.Database
    Dim found As Boolean
    Dim Test As String
    
    Const NAME_NOT_IN_COLLECTION As Integer = 3265
    
     ' Assume the table or query does not exist.
    found = False
    Set db = CurrentDb()
    
     ' Trap for any errors.
    On Error Resume Next
     
     ' See if the name is in the Tables collection.
    Test = db.TableDefs(TName).Name
    If Err.Number <> NAME_NOT_IN_COLLECTION Then found = True
    
    ' Reset the error variable.
    Err = 0
    
    TableExists = found
End Function

' Build SQL to export `tbl_name` sorted by each field from first to last
Private Function TableExportSql(ByVal tbl_name As String) As String
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, Count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)
    
    sb = VCS_String.Sb_Init()
    VCS_String.Sb_Append sb, "SELECT "
    
    Count = 0
    For Each fieldObj In rs.Fields
        If Count > 0 Then VCS_String.Sb_Append sb, ", "
        VCS_String.Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next
    
    VCS_String.Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    
    Count = 0
    For Each fieldObj In rs.Fields
        DoEvents
        If Count > 0 Then VCS_String.Sb_Append sb, ", "
        VCS_String.Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next

    TableExportSql = VCS_String.Sb_Get(sb)
End Function

' Export the lookup table `tblName` to `source\tables`.
Public Sub ExportTableData(ByVal tbl_name As String, ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim rs As DAO.Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim c As Long, value As Variant
    
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    
    Set rs = CurrentDb.OpenRecordset(TableExportSql(tbl_name))
    If rs.RecordCount = 0 Then
        'why is this an error? Debug.Print "Error: Table " & tbl_name & "  empty"
        rs.Close
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.MkDirIfNotExist obj_path
    Dim tempFileName As String
    tempFileName = VCS_File.TempFile()

    Set OutFile = FSO.CreateTextFile(tempFileName, overwrite:=True, Unicode:=True)

    c = 0
    For Each fieldObj In rs.Fields
        If c <> 0 Then OutFile.Write vbTab
        c = c + 1
        OutFile.Write fieldObj.Name
    Next
    OutFile.Write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        c = 0
        For Each fieldObj In rs.Fields
            DoEvents
            If c <> 0 Then OutFile.Write vbTab
            c = c + 1
            value = rs(fieldObj.Name)
            If IsNull(value) Then
                value = vbNullString
            Else
                value = Replace(value, "\", "\\")
                value = Replace(value, vbCrLf, "\n")
                value = Replace(value, vbCr, "\n")
                value = Replace(value, vbLf, "\n")
                value = Replace(value, vbTab, "\t")
            End If
            OutFile.Write value
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
Private Sub KillTable(ByVal tblName As String, db As Object)
    If TableExists(tblName) Then
        db.Execute "DROP TABLE [" & tblName & "]"
    End If
End Sub

Public Sub ImportLinkedTable(ByVal tblName As String, ByRef obj_path As String)
    Dim db As DAO.Database
    Dim FSO As Object
    Dim InFile As Object
    
    Set db = CurrentDb
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFilePath As String
    tempFilePath = VCS_File.TempFile()
    
    ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFilePath, iomode:=ForReading, Create:=False, Format:=TristateTrue)
    
    On Error GoTo err_notable:
    DoCmd.DeleteObject acTable, tblName
    
    GoTo err_notable_fin
    
err_notable:
    Err.Clear
    Resume err_notable_fin
    
err_notable_fin:
    On Error GoTo Err_CreateLinkedTable:
    
    Dim td As DAO.TableDef
    Set td = db.CreateTableDef(InFile.ReadLine())
    
    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then 'replace relative path with literal path
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & CurrentProject.Path & "\")
    End If
    td.connect = connect
    
    td.SourceTableName = InFile.ReadLine()
    db.TableDefs.Append td
    
    GoTo Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable_Fin:
    'this will throw errors if a primary key already exists or the table is linked to an access database table
    'will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:
    
    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim field As Variant
    Dim SQL As String
    SQL = "CREATE INDEX __uniqueindex ON " & td.Name & " ("
    
    For Each field In Split(Fields, ";+")
        SQL = SQL & "[" & field & "]" & ","
    Next
    'remove extraneous comma
    SQL = Left$(SQL, Len(SQL) - 1)
    
    SQL = SQL & ") WITH PRIMARY"
    CurrentDb.Execute SQL
    
Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
    
End Sub

' Import Table Definition
Public Sub ImportTableDef(ByVal tblName As String, ByVal directory As String)
    Dim filePath As String
    filePath = directory & tblName & ".sql"
    Dim db As Object ' DAO.Database
    Dim FSO As Object
    Dim InFile As Object
    Dim buf As String
    Dim p As Integer
    Dim P1 As Integer
    Dim strMsg As String
    Dim s As Variant
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tempFileName As String
    tempFileName = VCS_File.TempFile()

    n = -1
    Set FSO = CreateObject("Scripting.FileSystemObject")
    VCS_File.ConvertUtf8Ucs2 filePath, tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, iomode:=ForReading, Create:=False, Format:=TristateTrue)
    Set db = CurrentDb
    KillTable tblName, db
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = buf & InFile.ReadLine()
    Loop
    
    ' The following block is needed because "on update" actions may cause problems
    For Each s In Split("UPDATE|DELETE", "|")
      p = InStr(buf, "ON " & s & " CASCADE")
      Do While p > 0
          n = n + 1
          ReDim Preserve K(n)
          K(n).table = tblName
          K(n).IsUpdate = (s = "UPDATE")
          
          buf = Left$(buf, p - 1) & Mid$(buf, p + 18)
          p = InStrRev(buf, "REFERENCES", p)
          P1 = InStr(p, buf, "(")
          K(n).foreignFields = Split(VCS_String.SubString(P1, buf, "(", ")"), ",")
          K(n).foreignTable = Trim$(Mid$(buf, p + 10, P1 - p - 10))
          p = InStrRev(buf, "CONSTRAINT", P1)
          P1 = InStrRev(buf, "FOREIGN KEY", P1)
          If (P1 > 0) And (p > 0) And (P1 > p) Then
          ' multifield index
              K(n).refFields = Split(VCS_String.SubString(P1, buf, "(", ")"), ",")
          ElseIf P1 = 0 Then
          ' single field
          End If
          p = InStr(p, "ON " & s & " CASCADE", buf)
      Loop
    Next
    On Error Resume Next
    For i = 0 To n
        strMsg = K(i).table & " to " & K(i).foreignTable
        strMsg = strMsg & "(  "
        For j = 0 To UBound(K(i).refFields)
            strMsg = strMsg & K(i).refFields(j) & ", "
        Next j
        strMsg = Left$(strMsg, Len(strMsg) - 2) & ") to ("
        For j = 0 To UBound(K(i).foreignFields)
            strMsg = strMsg & K(i).foreignFields(j) & ", "
        Next j
        strMsg = Left$(strMsg, Len(strMsg) - 2) & ") Check "
        If K(i).IsUpdate Then
            strMsg = strMsg & " on update cascade " & vbCrLf
        Else
            strMsg = strMsg & " on delete cascade " & vbCrLf
        End If
    Next
    On Error GoTo 0
    db.Execute buf
    InFile.Close
    If Len(strMsg) > 0 Then MsgBox strMsg, vbOKOnly, "Correct manually"
        
End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub ImportTableData(ByVal tblName As String, ByVal obj_path As String)
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO As Object
    Dim InFile As Object
    Dim c As Long, buf As String, values() As String, value As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFileName As String
    tempFileName = VCS_File.TempFile()
    VCS_File.ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, iomode:=ForReading, Create:=False, Format:=TristateTrue)
    Set db = CurrentDb

    db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim$(buf)) > 0 Then
            values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                value = values(c)
                If Len(value) = 0 Then
                    value = Null
                Else
                    value = Replace(value, "\t", vbTab)
                    value = Replace(value, "\n", vbCrLf)
                    value = Replace(value, "\\", "\")
                End If
                rs(fieldObj.Name) = value
                c = c + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
End Sub

