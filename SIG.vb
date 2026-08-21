Public Class SIG
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pares As New List(Of Integer)
        For i = 0 To 5
            If (i Mod 2 = 0) Then
                pares.Add(i)
            End If
        Next
        For j = 0 To pares.Count - 1
            MsgBox(pares(j))
        Next

    End Sub
End Class