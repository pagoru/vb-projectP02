Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Form1

    Private con As SqlConnection

    Private ds, ds2 As DataSet
    Private ada, ada2 As SqlDataAdapter
    Private actualContacte
    Private dv As DataView

    'Es carrega el formulari per primera vegada
    'Es conecta amb la base de dades i crea
    'els datasets necessaris i adaptadors de 
    'les dues taules de la base de dades
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

        ada = New SqlDataAdapter("select * from CONTACTE ORDER BY CAST(CODI AS INT) ASC", con)
        ada2 = New SqlDataAdapter("select CP from CIUTAT", con)

        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(ada)
        Dim cmBase2 As SqlCommandBuilder = New SqlCommandBuilder(ada2)

        ada.Fill(ds, "Contactes")
        ada2.Fill(ds2, "Ciutats")

        Dim pk(1) As DataColumn
        pk(0) = ds.Tables("Contactes").Columns("CODI")
        ds.Tables("Contactes").PrimaryKey = pk

        actualContacte = 0
        showActualContacte()
    End Sub

    'Mostra el contacte actual
    'mitjançant de forma idnividual
    'totes les textbox necessaries
    Private Sub showActualContacte()
        If ds.Tables("contactes").Rows.Count - 1 Then
            Try
                contacte_codi.Text = ds.Tables("Contactes").Rows(actualContacte)("CODI").ToString()
                contacte_nom.Text = ds.Tables("Contactes").Rows(actualContacte)("NOM").ToString()
                contacte_cp.Text = ds.Tables("Contactes").Rows(actualContacte)("CP").ToString()
                contacte_telefon.Text = ds.Tables("Contactes").Rows(actualContacte)("TELEFON").ToString()
                contacte_categoria.Text = ds.Tables("Contactes").Rows(actualContacte)("CATEGORIA").ToString()
                contacte_email.Text = ds.Tables("Contactes").Rows(actualContacte)("EMAIL").ToString()
                contacte_riscmaxim.Text = ds.Tables("Contactes").Rows(actualContacte)("RISCMAXIM").ToString()
            Catch ex As Exception

            End Try
        End If
    End Sub

    'Boto que mostra
    'l'anterior contacte de la llista
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If actualContacte > 0 Then
            Dim original As Integer = actualContacte
            actualContacte = actualContacte - 1
            If actualContacte > 0 Then
                While actualContacte > ds.Tables("Contactes").Rows.Count - 1 And
                ds.Tables("Contactes").Rows(actualContacte).RowState = DataRowState.Deleted
                    actualContacte = actualContacte - 1
                End While
            End If
            If actualContacte > ds.Tables("Contactes").Rows.Count - 1 Then
                actualContacte = original
            End If
            showActualContacte()
        End If
    End Sub

    'Boto que mostra
    'el seguent contacte de la llista
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If actualContacte < ds.Tables("Contactes").Rows.Count - 1 Then
            Dim original As Integer = actualContacte
            actualContacte += 1
            While actualContacte < ds.Tables("Contactes").Rows.Count - 1 And
                ds.Tables("Contactes").Rows(actualContacte).RowState = DataRowState.Deleted

                actualContacte += 1
            End While
            If actualContacte > ds.Tables("Contactes").Rows.Count - 1 Then
                actualContacte = original
            End If
            showActualContacte()
        End If
    End Sub

    'Boto que mostra
    'el primer contacte de la llista
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        actualContacte = 0
        While actualContacte < ds.Tables("Contactes").Rows.Count - 1 And
            ds.Tables("Contactes").Rows(actualContacte).RowState = DataRowState.Deleted
            actualContacte += 1
        End While

        showActualContacte()
    End Sub

    'Boto que mostra
    'l'ultim contacte de la llista
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        actualContacte = ds.Tables("contactes").Rows.Count - 1
        While ds.Tables("Contactes").Rows(actualContacte).RowState = DataRowState.Deleted
            actualContacte -= 1
        End While
        showActualContacte()
    End Sub

    'Botó que controla la
    'busca de ids en els contactes
    'actuals i mostra lactual
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        actualContacte = 0
        Dim finded = False
        Dim searched = 0
        Dim idContacte
        Dim id

        If TextBox1.Text.Trim = String.Empty Then
            MsgBox("EL camp id no pot estar buit")
        Else
            idContacte = TextBox1.Text
            If actualContacte < ds.Tables("Contactes").Rows.Count - 1 Then

                While actualContacte < ds.Tables("Contactes").Rows.Count - 1

                    If ds.Tables("Contactes").Rows(actualContacte).RowState = DataRowState.Deleted Then

                    Else
                        id = ds.Tables("Contactes").Rows(actualContacte)("CODI").ToString()
                        If idContacte = id Then
                            searched = actualContacte
                            finded = True
                        End If
                    End If
                    actualContacte = actualContacte + 1
                End While
            End If
        End If

        actualContacte = searched
        showActualContacte()
        If Not finded Then
            MsgBox("La id no existeix")
        End If
    End Sub

    'Botó que controla la 
    'creació de nous contactes
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim dr As DataRow
        dr = ds.Tables("Contactes").NewRow

        dr("CODI") = "Camp obligatori"

        ds.Tables("Contactes").Rows.Add(dr)
        actualContacte = ds.Tables("Contactes").Rows.Count - 1
        showActualContacte()
    End Sub

    'Botó que controla 
    'leliminació de contactes
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ds.Tables("Contactes").Rows(actualContacte).Delete()
        Button1_Click(sender, e)
    End Sub

    'Boto que guarda les dades
    'actuals
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If guardarCanvisLocal() Then
            MsgBox("Canvis en la base de dades local")
        End If
    End Sub

    'Guarda totes les dades
    'actuals de forma temporal en 
    'els datasets i comprova si
    'les dades son valides
    Private Function guardarCanvisLocal()

        Dim permitirFerCanvis = True

        permitirFerCanvis = isNameValid()
        permitirFerCanvis = isEmailValid()
        permitirFerCanvis = isCPValid()
        permitirFerCanvis = isRiscValid()
        permitirFerCanvis = isCategoriaValid()
        permitirFerCanvis = isTelefonValid()

        If permitirFerCanvis Then
            ds.Tables("Contactes").Rows(actualContacte)("CODI") = contacte_codi.Text
            ds.Tables("Contactes").Rows(actualContacte)("NOM") = contacte_nom.Text
            ds.Tables("Contactes").Rows(actualContacte)("CP") = contacte_cp.Text
            ds.Tables("Contactes").Rows(actualContacte)("TELEFON") = contacte_telefon.Text
            Try
                ds.Tables("Contactes").Rows(actualContacte)("CATEGORIA") = contacte_categoria.Text
            Catch ex As Exception
                ds.Tables("Contactes").Rows(actualContacte)("CATEGORIA") = 0
            End Try
            ds.Tables("Contactes").Rows(actualContacte)("EMAIL") = contacte_email.Text
            ds.Tables("Contactes").Rows(actualContacte)("RISCMAXIM") = contacte_riscmaxim.Text
            btnPersistencia.Enabled = True
        End If

        Return permitirFerCanvis
    End Function

    'Comprova si el nom es valid
    Public Function isNameValid()
        If contacte_nom.Text.Trim = String.Empty Then
            MsgBox("El camp nom és obligatori!", MsgBoxStyle.OkOnly)
            Return False
        End If
        Return True
    End Function

    'Comprova si el correu
    'electronic es valid
    Public Function isEmailValid()
        Dim email As String = contacte_email.Text

        If contacte_email.Text.Trim = String.Empty Then
        Else
            If Not FormatEmail(email) Then
                MsgBox("El format del email és incorrecte!", MsgBoxStyle.OkOnly)
                Return False
            End If
        End If
        Return True
    End Function

    'Comprova si el correu
    'electronic passat es
    'valid
    Function FormatEmail(ByVal emailAddress As String) As Boolean
        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]" &
        "*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        Return emailAddressMatch.Success
    End Function

    'Comprova si el riscmaxim
    'es valid
    Public Function isRiscValid()
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

    'Comprova si la categoria
    'es valida
    Public Function isCategoriaValid()
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

    'Comprova si el telefon
    'es valid
    Public Function isTelefonValid()
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

    'Comprova si el codif
    'postal es valid
    Public Function isCPValid()
        If contacte_cp.Text.Trim IsNot String.Empty Then
            For index As Integer = 0 To ds2.Tables("Ciutats").Rows.Count - 1
                If contacte_cp.Text = ds2.Tables("Ciutats").Rows(index)("CP").ToString() Then
                    Return True
                End If
            Next
        End If
        MsgBox("Aquest codi postal no es valid")
        Return False
    End Function

    'Comprova si abans de
    'tancar el formulari
    'sha realitzat la persistencia
    'a la base de dades
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If btnPersistencia.Enabled Then
            If MessageBox.Show("Vols fer la persistencia dels canvis realitzats?", "Atenció", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                generarPersistencia()
            End If
        End If
    End Sub

    'Realitza l'acció de generar 
    'persistencia de la taula de 
    'contactes actuals i guarda les dades
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

    'Boto que carrega el 
    'formulari de codis postals
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Form2.load(ds2)
    End Sub

    'Botó que realitza la
    'persistencia de les dades
    'actuals
    Private Sub btnPersistencia_Click(sender As Object, e As EventArgs) Handles btnPersistencia.Click
        generarPersistencia()
        MsgBox("Les dades s'han registrat a la BD permanentment")
    End Sub

End Class
