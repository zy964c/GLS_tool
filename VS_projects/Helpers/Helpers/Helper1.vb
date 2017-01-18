Module helper

    Sub Main()
        For Each s As String In My.Application.CommandLineArgs
            If s = "coord" Then
                coord1.coord()
            ElseIf s = "fn" Then
                fn1.fn()
            ElseIf s = "jd" Then
                jd2.jd()
            End If

        Next

    End Sub

End Module
