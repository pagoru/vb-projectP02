Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Form1

    Private con As SqlConnection

    Private ds, ds2 As DataSet
    Private ada, ada2 As SqlDataAdapter
    Private registreActual
    Private dv As DataView

    Private banderaDreta = False, banderaEsquerra = False, banderaGuardat = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnPersistencia.Enabled = False
        con = New SqlConnection
        con.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=CIUTATS;Trusted_Connection=true;"
        Try
            con.Open()
        Catch ex As Exception
        End Try

        If con.State = ConnectionState.Closed Then
            Me.Close()
        End If
        ds = New DataSet()
        ds2 = New DataSet()

        ada = New SqlDataAdapter("select * from CONTACTE", con)
        ada2 = New SqlDataAdapter("select CP from CIUTAT", con)

        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(ada)
        Dim cmBase2 As SqlCommandBuilder = New SqlCommandBuilder(ada2)

        ada.Fill(ds, "Contactes")
        ada2.Fill(ds2, "Ciutats")

        Dim pk(1) As DataColumn
        pk(0) = ds.Tables("Contactes").Columns("CODI")
        ds.Tables("Contactes").PrimaryKey = pk

        registreActual = 0
        mostrarRegistreActual()
    End Sub

    Private Sub mostrarRegistreActual()

        If ds.Tables("contactes").Rows.Count - 1 Then
            Try
                contacte_codi.Text = ds.Tables("Contactes").Rows(registreActual)("CODI").ToString()
                contacte_nom.Text = ds.Tables("Contactes").Rows(registreActual)("NOM").ToString()
                contacte_cp.Text = ds.Tables("Contactes").Rows(registreActual)("CP").ToString()
                contacte_telefon.Text = ds.Tables("Contactes").Rows(registreActual)("TELEFON").ToString()
                contacte_categoria.Text = ds.Tables("Contactes").Rows(registreActual)("CATEGORIA").ToString()
                contacte_email.Text = ds.Tables("Contactes").Rows(registreActual)("EMAIL").ToString()
                contacte_riscmaxim.Text = ds.Tables("Contactes").Rows(registreActual)("RISCMAXIM").ToString()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If registreActual > 0 Then
            Dim original As Integer = registreActual
            registreActual = registreActual - 1
            If registreActual > 0 Then
                While registreActual > ds.Tables("Contactes").Rows.Count - 1 And
                ds.Tables("Contactes").Rows(registreActual).RowState = DataRowState.Deleted
                    registreActual = registreActual - 1
                End While
            End If
            If registreActual > ds.Tables("Contactes").Rows.Count - 1 Then
                registreActual = original
            End If
            mostrarRegistreActual()

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If registreActual < ds.Tables("Contactes").Rows.Count - 1 Then
            Dim original As Integer = registreActual
            registreActual = registreActual + 1
            While registreActual < ds.Tables("Contactes").Rows.Count - 1 And
                ds.Tables("Contactes").Rows(registreActual).RowState = DataRowState.Deleted

                registreActual = registreActual + 1
            End While
            If registreActual > ds.Tables("Contactes").Rows.Count - 1 Then
                registreActual = original
            End If
            mostrarRegistreActual()

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        registreActual = 0
        While registreActual < ds.Tables("Contactes").Rows.Count - 1 And
            ds.Tables("Contactes").Rows(registreActual).RowState = DataRowState.Deleted
            registreActual = registreActual + 1
        End While

        mostrarRegistreActual()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        registreActual = ds.Tables("contactes").Rows.Count - 1
        While ds.Tables("Contactes").Rows(registreActual).RowState = DataRowState.Deleted
            registreActual = registreActual - 1
        End While
        mostrarRegistreActual()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        registreActual = 0
        Dim registreBuscat = 0
        Dim idContacte
        Dim idTractament
        Dim trobat = False

        If TextBox1.Text.Trim = String.Empty Then
            MsgBox("EL camp id no pot estar buid per buscar per id")
        Else
            idContacte = TextBox1.Text
            If registreActual < ds.Tables("Contactes").Rows.Count - 1 Then

                While registreActual < ds.Tables("Contactes").Rows.Count - 1

                    If ds.Tables("Contactes").Rows(registreActual).RowState = DataRowState.Deleted Then

                    Else
                        idTractament = ds.Tables("Contactes").Rows(registreActual)("CODI").ToString()
                        If idContacte = idTractament Then
                            registreBuscat = registreActual
                            trobat = True
                        End If
                    End If
                    registreActual = registreActual + 1
                End While
            End If
        End If

        registreActual = registreBuscat
        mostrarRegistreActual()
        If trobat = False Then
            MsgBox("Aquesta id no esta en el nostre sistema")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim dr As DataRow
        dr = ds.Tables("Contactes").NewRow

        dr("CODI") = "Camp obligatori"

        ds.Tables("Contactes").Rows.Add(dr)
        registreActual = ds.Tables("Contactes").Rows.Count - 1
        mostrarRegistreActual()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ds.Tables("Contactes").Rows(registreActual).Delete()
        Button1_Click(sender, e)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If guardarCanvisLocal() Then
            MsgBox("Canvis en la base de dades local")
        End If
    End Sub

    Private Function guardarCanvisLocal()

        Dim permitirFerCanvis = True

        permitirFerCanvis = ComprovacioNom()
        permitirFerCanvis = ComprovacioEmail()
        permitirFerCanvis = ComprovacionsRisc()
        permitirFerCanvis = ComprovacioCategoria()
        permitirFerCanvis = ComprovacioTelefon()
        permitirFerCanvis = ComprovacioCP()

        If permitirFerCanvis Then
            ds.Tables("Contactes").Rows(registreActual)("CODI") = contacte_codi.Text
            ds.Tables("Contactes").Rows(registreActual)("NOM") = contacte_nom.Text
            ds.Tables("Contactes").Rows(registreActual)("CP") = contacte_cp.Text
            ds.Tables("Contactes").Rows(registreActual)("TELEFON") = contacte_telefon.Text
            Try
                ds.Tables("Contactes").Rows(registreActual)("CATEGORIA") = contacte_categoria.Text
            Catch ex As Exception
                ds.Tables("Contactes").Rows(registreActual)("CATEGORIA") = 0
            End Try
            ds.Tables("Contactes").Rows(registreActual)("EMAIL") = contacte_email.Text
            ds.Tables("Contactes").Rows(registreActual)("RISCMAXIM") = contacte_riscmaxim.Text
            btnPersistencia.Enabled = True
        End If

        Return permitirFerCanvis
    End Function

    Public Function ComprovacioNom()
        Dim bandera = True
        If contacte_nom.Text.Trim = String.Empty Then
            MsgBox("El camp nom és obligatori!", MsgBoxStyle.OkOnly)
            bandera = False
        End If
        Return bandera
    End Function

    Public Function ComprovacioEmail()
        Dim bandera = True
        Dim email As String = contacte_email.Text

        If contacte_email.Text.Trim = String.Empty Then
        Else
            If Not FormatEmail(email) Then
                MsgBox("El format del email és incorrecte!", MsgBoxStyle.OkOnly)
                bandera = False
            End If
        End If
        Return bandera
    End Function

    Function FormatEmail(ByVal emailAddress As String) As Boolean
        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]" &
        "*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        Return emailAddressMatch.Success
    End Function

    Public Function ComprovacionsRisc()
        If contacte_riscmaxim.Text.Trim = String.Empty Then
            MsgBox("El camp risc màxim és obligatori!", MsgBoxStyle.OkOnly)
        Else
            If contacte_riscmaxim.Text < 0.0 Then
                MsgBox("El valor de risc màxim ha de ser més gran que 0.")
                Return False
            End If
        End If
        Return True
    End Function

    Public Function ComprovacioCategoria()
        If contacte_categoria.Text.Trim = String.Empty Then
        Else
            Try
                Dim categoria As Integer = contacte_categoria.Text
            Catch ex As Exception
                MsgBox("Categoria inorrecte")
                Return False
            End Try
        End If

        Return True
    End Function

    Public Function ComprovacioTelefon()
        If contacte_telefon.Text.Trim = String.Empty Then
        Else
            If IsNumeric(contacte_telefon.Text) Then
            Else
                MsgBox("El telefon ha de ser númeric")
                Return False
            End If
        End If

        Return True
    End Function

    Public Function ComprovacioCP()
        Dim buscar = False
        'No és obligatori el camp, en cas que hi hagui contingut mirem que el valor no sigui negatiu
        If contacte_cp.Text.Trim = String.Empty Then
        Else
            For index As Integer = 0 To ds2.Tables("Ciutats").Rows.Count - 1
                If contacte_cp.Text = ds2.Tables("Ciutats").Rows(index)("CP").ToString() Then
                    Exit For
                End If
            Next
            If buscar = False Then
                Return False
                MsgBox("Aquest codi postal no es valid")
            End If
        End If
        Return True
    End Function

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If btnPersistencia.Enabled = True Then
            If MessageBox.Show("Vols fer la persistencia dels canvis realitzats?", "Atenció", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                generarPersistencia()
            End If
        End If
    End Sub

    Private Sub generarPersistencia()
        Dim dt As DataTable
        Try
            dt = ds.Tables("Contactes").GetChanges()
            ada.Update(dt)
            ds.Tables("Contactes").AcceptChanges()
            btnPersistencia.Enabled = False
        Catch ex As Exception
            MsgBox("Error")
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Form2.cridarCP(ds2)
    End Sub

    Private Sub btnPersistencia_Click(sender As Object, e As EventArgs) Handles btnPersistencia.Click
        generarPersistencia()
        MsgBox("Les dades s'han registrat a la BD permanentment")
    End Sub

End Class
