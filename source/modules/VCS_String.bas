Attribute VB_Name = "VCS_String"
Option Compare Database

Option Private Module
Option Explicit


'--------------------
' String Functions: String Builder,String Padding (right only), Substrings
'--------------------

' String builder: Init
Public Function Sb_Init() As String()
    Dim x(-1 To -1) As String
    Sb_Init = x
End Function

' String builder: Clear
Public Sub Sb_Clear(ByRef sb() As String)
    ReDim Sb_Init(-1 To -1)
End Sub

' String builder: Append
Public Sub Sb_Append(ByRef sb() As String, ByVal value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = value
End Sub

' String builder: Get value
Public Function Sb_Get(ByRef sb() As String) As String
    Sb_Get = Join(sb, "")
End Function


' Pad a string on the right to make it `count` characters long.
Public Function PadRight(ByVal value As String, ByVal Count As Integer) As String
    PadRight = value
    If Len(value) < Count Then
        PadRight = PadRight & Space$(Count - Len(value))
    End If
End Function

' returns substring between e.g. "(" and ")", internal brackets ar skippped
Public Function SubString(ByVal p As Integer, ByVal s As String, ByVal startsWith As String, _
                          ByVal endsWith As String) As String
    Dim start As Integer
    Dim cursor As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim level As Integer
    
    start = InStr(p, s, startsWith)
    level = 1
    P1 = InStr(start + 1, s, startsWith)
    P2 = InStr(start + 1, s, endsWith)
    
    Do While level > 0
        If P1 > P2 And P2 > 0 Then
            cursor = P2
            level = level - 1
        ElseIf P2 > P1 And P1 > 0 Then
            cursor = P1
            level = level + 1
        ElseIf P2 > 0 And P1 = 0 Then
            cursor = P2
            level = level - 1
        ElseIf P1 > 0 And P1 = 0 Then
            cursor = P1
            level = level + 1
        ElseIf P1 = 0 And P2 = 0 Then
            SubString = vbNullString
            Exit Function
        End If
        P1 = InStr(cursor + 1, s, startsWith)
        P2 = InStr(cursor + 1, s, endsWith)
    Loop
    
    SubString = Mid$(s, start + 1, cursor - start - 1)
End Function

