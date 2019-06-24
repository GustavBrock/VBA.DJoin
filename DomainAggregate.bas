Attribute VB_Name = "DomainAggregate"
Option Explicit

' Returns the joined (concatenated) values from a field of records having the same key.
' The joined values are stored in a collection which speeds up browsing a query or form
' as all joined values will be retrieved once only from the table or query.
' Null values and zero-length strings are ignored.
'
' If no values are found, Null is returned.
'
' The default separator of the joined values is a space.
' Optionally, any other separator can be specified.
'
' Syntax is held close to that of the native domain functions, DLookup, DCount, etc.
'
' Typical usage in a select query using a table (or query) as source:
'
'   Select
'       KeyField,
'       DJoin("[ValueField]", "[Table]", "[KeyField] = " & [KeyField] & "") As Values
'   From
'       Table
'   Group By
'       KeyField
'
' The source can also be an SQL Select string:
'
'   Select
'       KeyField,
'       DJoin("[ValueField]", "Select ValueField From SomeTable Order By SomeField", "[KeyField] = " & [KeyField] & "") As Values
'   From
'       Table
'   Group By
'       KeyField
'
' To clear the collection (cache), call DJoin with no arguments:
'
'   DJoin
'
' Requires:
'   CollectValues
'
' 2019-06-24, Cactus Data ApS, Gustav Brock
'
Public Function DJoin( _
    Optional ByVal Expression As String, _
    Optional ByVal Domain As String, _
    Optional ByVal Criteria As String, _
    Optional ByVal Delimiter As String = " ") _
    As Variant
    
    ' Expected error codes to accept.
    Const CannotAddKey      As Long = 457
    Const CannotReadKey     As Long = 5
    ' SQL.
    Const SqlMask           As String = "Select {0} From {1} {2}"
    Const SqlLead           As String = "Select "
    Const SubMask           As String = "({0}) As T"
    Const FilterMask        As String = "Where {0}"
    
    Static Values   As New Collection
    
    Dim Records     As DAO.Recordset
    Dim Sql         As String
    Dim SqlSub      As String
    Dim Filter      As String
    Dim Result      As Variant
    
    On Error GoTo Err_DJoin
    
    If Expression = "" Then
        ' Erase the collection of keys.
        Set Values = Nothing
        Result = Null
    Else
        ' Get the values.
        ' This will fail if the current criteria hasn't been added
        ' leaving Result empty.
        Result = Values.Item(Criteria)
        '
        If IsEmpty(Result) Then
            ' The current criteria hasn't been added to the collection.
            ' Build SQL to lookup values.
            If InStr(1, LTrim(Domain), SqlLead, vbTextCompare) = 1 Then
                ' Domain is an SQL expression.
                SqlSub = Replace(SubMask, "{0}", Domain)
            Else
                ' Domain is a table or query name.
                SqlSub = Domain
            End If
            If Trim(Criteria) <> "" Then
                ' Build Where clause.
                Filter = Replace(FilterMask, "{0}", Criteria)
            End If
            ' Build final SQL.
            Sql = Replace(Replace(Replace(SqlMask, "{0}", Expression), "{1}", SqlSub), "{2}", Filter)
            
            ' Look up the values to join.
            Set Records = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
            CollectValues Records, Delimiter, Result
            ' Add the key and its joined values to the collection.
            Values.Add Result, Criteria
        End If
    End If
    
    ' Return the joined values (or Null if none was found).
    DJoin = Result
    
Exit_DJoin:
    Exit Function
    
Err_DJoin:
    Select Case Err
        Case CannotAddKey
            ' Key is present, thus cannot be added again.
            Resume Next
        Case CannotReadKey
            ' Key is not present, thus cannot be read.
            Resume Next
        Case Else
            ' Some other error. Ignore.
            Resume Exit_DJoin
    End Select
    
End Function

' To be called from DJoin.
'
' Joins the content of the first field of a recordset to one string
' with a space as delimiter or an optional delimiter, returned by
' reference in parameter Result.
'
' 2019-06-11, Cactus Data ApS, Gustav Brock
'
Private Sub CollectValues( _
    ByRef Records As DAO.Recordset, _
    ByVal Delimiter As String, _
    ByRef Result As Variant)
    
    Dim SubRecords  As DAO.Recordset
    
    Dim Value       As Variant

    If Records.RecordCount > 0 Then
        While Not Records.EOF
            Value = Records.Fields(0).Value
            If Records.Fields(0).IsComplex Then
                ' Multi-value field (or attachment field).
                Set SubRecords = Records.Fields(0).Value
                CollectValues SubRecords, Delimiter, Result
            ElseIf Nz(Value) = "" Then
                ' Ignore Null values and zero-length strings.
            ElseIf IsEmpty(Result) Then
                ' First value found.
                Result = Value
            Else
                ' Join subsequent values.
                Result = Result & Delimiter & Value
            End If
            Records.MoveNext
        Wend
    Else
        ' No records found with the current criteria.
        Result = Null
    End If
    Records.Close

End Sub

