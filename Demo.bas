Attribute VB_Name = "Demo"
Option Explicit

Public Function SpeedTest()

    Dim Records     As DAO.Recordset
    
    Dim TimeStart   As Single
    Dim TimeEnd     As Single
    Dim Value       As Variant
    
    TimeStart = Timer
    Set Records = CurrentDb.OpenRecordset("Join", dbOpenSnapshot)
    While Not Records.EOF
        Value = Records.Fields(1).Value
        Records.MoveNext
    Wend
    TimeEnd = Timer
    Debug.Print "Join", Records.RecordCount, Format(TimeEnd - TimeStart, "0.00")
    TimeStart = Timer
    Set Records = CurrentDb.OpenRecordset("Join", dbOpenSnapshot)
    While Not Records.EOF
        Value = Records.Fields(1).Value
        Records.MoveNext
    Wend
    TimeEnd = Timer
    Debug.Print "Join", Records.RecordCount, Format(TimeEnd - TimeStart, "0.00")

    ' Clear cache.
    DJoin
    
    TimeStart = Timer
    Set Records = CurrentDb.OpenRecordset("Concat", dbOpenSnapshot)
    While Not Records.EOF
        Value = Records.Fields(1).Value
        Records.MoveNext
    Wend
    TimeEnd = Timer
    Debug.Print "ConcatRelated", Records.RecordCount, Format(TimeEnd - TimeStart, "0.00")
    TimeStart = Timer
    Set Records = CurrentDb.OpenRecordset("Concat", dbOpenSnapshot)
    While Not Records.EOF
        Value = Records.Fields(1).Value
        Records.MoveNext
    Wend
    TimeEnd = Timer
    Debug.Print "ConcatRelated", Records.RecordCount, Format(TimeEnd - TimeStart, "0.00")
    
    ' Clear cache.
    DJoin
    
End Function

