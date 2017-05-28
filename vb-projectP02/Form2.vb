Public Class Form2

    'Carrega els codis postals
    Public Function load(ds2 As DataSet) As DialogResult
        ListBox1.Items.Clear()

        For index As Integer = 0 To ds2.Tables("Ciutats").Rows.Count - 1
            ListBox1.Items.Add(ds2.Tables("Ciutats").Rows(index)("CP").ToString())
        Next
        Return ShowDialog()
    End Function

    'Asigna el codi postal
    'seleccionat
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form1.contacte_cp.Text = ListBox1.SelectedItem
        Me.Close()
    End Sub

    'Tanca el formulari
    'sense realitzar ninguna operacio
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class