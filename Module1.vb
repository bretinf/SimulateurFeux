Option Explicit On
Module Clignote

    '---------------------------------------------------------------------------------------
    ' Auteur    : Didier FOURGEOT (myDearFriend!)
    '             www.mdf-xlpages.com
    ' Date      : 12/04/2010
    ' Sujet     : TextBox clignotant
    '---------------------------------------------------------------------------------------

    'En déclarant cette variable Temps en "public", elle devient accessible partout dans le projet
    Public Temps As Object

    Sub AffichUsf()
        Form1.Show()
    End Sub

    Sub Clign()
        'Programmation de l'évènement toutes les secondes
        Temps = Now + TimeValue("00:00:01")
        Application.OnTime(Temps, "Clign")
        'Affiche l'alerte ou la fait disparaître (alternativement)
        With UserForm1.LeTextBox
            .BackStyle = IIf(.BackStyle = fmBackStyleOpaque, fmBackStyleTransparent, fmBackStyleOpaque)
        End With
    End Sub

    Sub StopClign()
        On Error Resume Next
        'Stoppe la gestion de l'évènement OnTime
        Application.OnTime(Temps, "Clign", , False)
        On Error GoTo 0
        'Cache l'alerte
        UserForm1.LeTextBox.BackStyle = fmBackStyleOpaque
        Temps = Empty
    End Sub


End Module
