Public Class Contacte
    Public codi As String
    Public nom As String
    Public cp As String
    Public ciutat As String
    Public telefon As String
    Public categoria As Integer
    Public email As String
    Public riscmaxim As Decimal

    Public Sub New(codi As String)
        Me.codi = codi
    End Sub

    Public Sub New(codi As String, nom As String, cp As String, telefon As String, categoria As Integer, email As String, riscmaxim As Decimal, ciutat As String)
        Me.codi = codi
        Me.nom = nom
        Me.cp = cp
        Me.telefon = telefon
        Me.categoria = categoria
        Me.email = email
        Me.riscmaxim = riscmaxim

        Me.ciutat = ciutat
    End Sub

End Class
