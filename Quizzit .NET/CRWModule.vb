Module CRWModule
    Dim FS As System.IO.FileStream
    Dim SR As System.IO.TextReader
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    Dim y As Integer
    Dim Temp(150) As Integer
    Public TxtArray(10, 10) As Char


    Public Sub ReadFile()
        Try
            FS = New System.IO.FileStream(ResourcePath + "\CrossWordz_files\" + MainRound + ".txt", IO.FileMode.Open)
            SR = New System.IO.StreamReader(FS)
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        For i = 0 To FS.Length - 1
            Temp(i) = SR.Read
        Next
        For i = 0 To FS.Length - 1
            If Temp(i) <> 13 And Temp(i) <> 10 Then
                TxtArray(x, y) = Chr(Temp(i))
                y = y + 1
                If y = 10 Then
                    y = 0
                    x = x + 1
                End If
            End If
        Next
        x = 0
        y = 0
        SR.Close()
        FS.Close()
    End Sub
End Module
