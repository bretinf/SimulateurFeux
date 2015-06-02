Public Class Form1

    Private Property Etat As Boolean

    Private Property duree As Timer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub OvalShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Function Timer() As Timer
        Throw New NotImplementedException
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OvalShape1.Visible = True

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        With Me.BoutonClignotant1

            .Clignoter (Not .Clignotement_en_cours)

        End With

    End Sub

    Private Sub BoutonClignotant1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BoutonClignotant1.Click

    End Sub
End Class

