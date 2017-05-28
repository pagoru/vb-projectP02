Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Form1

    Private contactes As New List(Of Contacte)
    Private currentContacte As Contacte
    Private currentIndexContacte As Integer

    Private connection As SqlConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connection = New SqlConnection()
        connection.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=CIUTATS;Trusted_Connection=True;"

        Try
            connection.Open()
        Catch ex As Exception
            MsgBox("No es pot connectar al servidor sql")
        End Try

        If connection.State = ConnectionState.Closed Then
            Close()
            Return
        End If

        LoadContacts()
        Me.Enabled = True
    End Sub

    Private Sub LoadContacts()

        Dim com As SqlCommand = New SqlCommand()
        Dim read As SqlDataReader

        com.Connection = connection
        com.CommandText = "SELECT * FROM CONTACTE co LEFT JOIN CIUTAT ci ON co.cp = ci.cp ORDER BY CAST(co.CODI AS INT) ASC"

        read = com.ExecuteReader

        While read.Read()
            contactes.Add(New Contacte(read(0).ToString, read(1).ToString, read(2).ToString, read(3).ToString, Integer.Parse(read(4).ToString), read(5).ToString, Decimal.Parse(read(6).ToString), read(8).ToString))
        End While

        read.Close()

        If contactes.Count = 0 Then
            Return
        End If
        currentContacte = contactes.First
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()

    End Sub

    Private Sub ClearContacte()
        Me.contacte_codi.Text = ""
        Me.contacte_nom.Text = ""
        Me.contacte_cp.Text = ""
        Me.contacte_telefon.Text = ""
        Me.contacte_categoria.Text = ""
        Me.contacte_email.Text = ""
        Me.contacte_riscmaxim.Text = ""

        Me.contacte_ciutat.Text = ""
    End Sub

    Private Sub LoadConacte()
        Me.contacte_codi.Text = currentContacte.codi
        Me.contacte_nom.Text = currentContacte.nom
        Me.contacte_cp.Text = currentContacte.cp
        Me.contacte_telefon.Text = currentContacte.telefon
        Me.contacte_categoria.Text = currentContacte.categoria
        Me.contacte_email.Text = currentContacte.email
        Me.contacte_riscmaxim.Text = currentContacte.riscmaxim

        Me.contacte_ciutat.Text = currentContacte.ciutat

        If currentIndexContacte = 0 Then
            Me.Button1.Enabled = False
            Me.Button3.Enabled = False
        Else
            Me.Button1.Enabled = True
            Me.Button3.Enabled = True
        End If
        If currentIndexContacte = contactes.Count - 1 Or contactes.Count = 0 Then
            Me.Button2.Enabled = False
            Me.Button4.Enabled = False
        Else
            Me.Button2.Enabled = True
            Me.Button4.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        currentContacte = contactes.First
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        currentContacte = contactes(currentIndexContacte - 1)
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        currentContacte = contactes(currentIndexContacte + 1)
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        currentContacte = contactes.Last
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim result = MessageBox.Show("Vols guardar els canvis?", "Avis!!", MessageBoxButtons.YesNoCancel)

        If result = DialogResult.Yes Then
            SaveChanges()
        ElseIf result = DialogResult.No Then

        ElseIf result = DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub DeleteCurrentContacte()
        If currentContacte Is Nothing Then
            Return
        End If

        Dim com As SqlCommand = New SqlCommand()

        com.Connection = connection
        com.CommandText = $"DELETE CONTACTE WHERE CODI='{currentContacte.codi}'"

        com.ExecuteNonQuery()

        ClearContacte()
        contactes.RemoveAt(currentIndexContacte)
        If contactes.Count = 0 Then
            Return
        End If
        currentContacte = contactes.First
        LoadConacte()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DeleteCurrentContacte()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        SaveChanges()
    End Sub

    Private Sub SaveChanges()
        Dim com As SqlCommand = New SqlCommand()

        com.Connection = connection

        For Each c As Contacte In contactes

            Dim cp As String = $"'{c.cp}'"
            If cp IsNot Nothing Then
                Dim com2 As SqlCommand = New SqlCommand()
                Dim reader2 As SqlDataReader
                com2.Connection = connection
                com2.CommandText = $"SELECT COUNT(*) FROM ciutat WHERE CP='{c.cp}'"
                reader2 = com2.ExecuteReader

                reader2.Read()
                If reader2(0).ToString = "0" Then
                    cp = "null"
                End If
                reader2.Close()
            End If



            Dim email As String = ""
            If (IsEmail(c.email)) Then
                email = c.email
            End If

            com.CommandText = $"UPDATE CONTACTE SET nom='{c.nom}', cp={cp}, telefon='{c.telefon}', categoria='{c.categoria}', email='{c.email}', riscmaxim='{c.riscmaxim}' WHERE codi='{c.codi}'"
            com.ExecuteNonQuery()

        Next

    End Sub

    Private Sub contacte_codi_TextChanged(sender As Object, e As EventArgs) Handles contacte_codi.TextChanged
        currentContacte.codi = contacte_codi.Text
    End Sub

    Private Sub contacte_nom_TextChanged(sender As Object, e As EventArgs) Handles contacte_nom.TextChanged
        currentContacte.nom = contacte_nom.Text
    End Sub

    Private Sub contacte_cp_TextChanged(sender As Object, e As EventArgs) Handles contacte_cp.TextChanged
        Try
            currentContacte.cp = Integer.Parse(contacte_cp.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub contacte_telefon_TextChanged(sender As Object, e As EventArgs) Handles contacte_telefon.TextChanged
        currentContacte.telefon = contacte_telefon.Text
    End Sub

    Private Sub contacte_categoria_TextChanged(sender As Object, e As EventArgs) Handles contacte_categoria.TextChanged
        Try
            currentContacte.categoria = Integer.Parse(contacte_categoria.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub contacte_email_TextChanged(sender As Object, e As EventArgs) Handles contacte_email.TextChanged
        currentContacte.email = contacte_email.Text
    End Sub

    Private Sub contacte_riscmaxim_TextChanged(sender As Object, e As EventArgs) Handles contacte_riscmaxim.TextChanged
        Try
            currentContacte.riscmaxim = Decimal.Parse(contacte_riscmaxim.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim com As SqlCommand = New SqlCommand()
        Dim reader As SqlDataReader
        com.Connection = connection
        com.CommandText = $"SELECT codi + 1 FROM contacte ORDER BY CAST(codi AS INT) DESC"
        reader = com.ExecuteReader

        Dim codi As String = "1"
        If reader.Read() Then
            codi = reader(0).ToString()
        End If

        reader.Close()

        com.CommandText = $"INSERT INTO contacte VALUES('{codi}', '', null, '', '0', '', '0')"
        com.ExecuteNonQuery()

        contactes.Add(New Contacte(codi))
        ClearContacte()
        currentContacte = contactes.Last
        currentIndexContacte = contactes.IndexOf(currentContacte)
        LoadConacte()
    End Sub

    Function IsEmail(ByVal email As String) As Boolean
        If email Is Nothing Then
            Return False
        End If
        Static emailExpression As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")

        Return emailExpression.IsMatch(email)
    End Function
End Class
