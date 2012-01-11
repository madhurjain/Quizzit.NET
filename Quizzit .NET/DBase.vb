Module DBase
    Dim MyConnection As New ADODB.Connection
    Public MyRecordset As New ADODB.Recordset
    Public ScoreSet As New ADODB.Recordset
    Dim ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\..\Resources\Database\QuizDb.mdb;Persist Security Info=False"

    Public Sub OpenConnection()
        If MyConnection.State = 0 Then
            MyConnection.Open(ConnString)
        End If
    End Sub
    Public Sub Recordset(ByVal Query As String)
        If MyRecordset.State = 0 Then
            MyRecordset.Open(Query, MyConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        End If
    End Sub

    Public Sub OpenScoreSet(ByVal Query As String)
        If ScoreSet.State = 0 Then
            ScoreSet.Open(Query, MyConnection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        End If
    End Sub
    Public Sub CloseScoreSet()
        If ScoreSet.State = 1 Then
            ScoreSet.Close()
            ScoreSet.ActiveConnection = Nothing
        End If
    End Sub
    Public Sub CloseRecordset()
        If MyRecordset.State = 1 Then
            MyRecordset.Close()
            MyRecordset.ActiveConnection = Nothing
        End If
    End Sub
    Public Sub CloseConnection()
        If MyConnection.State = 1 Then
            MyConnection.Close()
        End If
    End Sub
End Module
