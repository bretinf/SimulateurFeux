Public Class BoutonClignotant

    Inherits System.Windows.Forms.PictureBox

    Friend WithEvents Timer1 As System.Windows.Forms.Timer

    Enum enum_couleur

        Bleu

        OrangeClair

        OrangeFoncé

        RougeClair

        RougeFoncé

        VertClair

        VertFoncé

    End Enum

Public Sub Set_Interval_clignotement_en_ms(Optional ByVal interval_en_ms As Integer = 1000) ', ByVal couleur As Color)

        Me.Timer1.Interval = interval_en_ms

    End Sub

    Public Sub Set_Couleur(ByVal Couleur As enum_couleur) ', ByVal couleur As Color)

        Dim key As String = ""

        Select Case Couleur

            Case enum_couleur.Bleu

                key = "Bleu"

            Case enum_couleur.OrangeClair

                key = "OrangeClair"

            Case enum_couleur.OrangeFoncé

                key = "OrangeFoncé"

            Case enum_couleur.RougeClair

                key = "RougeClair"

            Case enum_couleur.RougeFoncé

                key = "RougeFoncé"

            Case enum_couleur.VertClair

                key = "VertClair"

            Case enum_couleur.VertFoncé

                key = "VertFoncé"

        End Select

        Me.Image = Me.ImageList1.Images(key)

    End Sub

    'true pour démarrer et false pour arrêter

    Public Sub Clignoter(Optional ByVal Condition As Boolean = True)

        Me.Timer1.Enabled = Condition

        If Not Condition Then Me.Visible = False

    End Sub

    Function Clignotement_en_cours() As Boolean

        Return Me.Timer1.Enabled

    End Function

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Me.Visible = Not Me.Visible

    End Sub

End Class

