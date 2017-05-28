Public Class Form2
    Private numero = 0

    Public Function cridarCP(ds2 As DataSet) As System.Windows.Forms.DialogResult
        Dim CPAuxiliar

        If numero = 0 Then
            For index As Integer = 0 To ds2.Tables("Ciutats").Rows.Count - 1
                CPAuxiliar = ds2.Tables("Ciutats").Rows(index)("CP").ToString()
                ListBox1.Items.Add(CPAuxiliar)
            Next
        End If
        numero = numero + 1

        Return ShowDialog()
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
End Class