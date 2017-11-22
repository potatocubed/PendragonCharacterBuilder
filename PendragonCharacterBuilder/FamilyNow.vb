Public Module FamilyNow

    Function OldKnightGenerator(oldKnights As String(,), grandfatherName As String) As String(,)
        Dim x As Integer
        Dim s As String

        If oldKnights(0, 0) <> "" Then
            x = DiceRoller(1, 20)
            If x <= 15 Then
                s = $"{grandfatherName}'s younger brother. " & AliveAndMarried(True)
            Else
                s = $"{grandfatherName}'s illegitimate brother. " & AliveAndMarried(True)
            End If
            s = Replace(s, " married", " married to " & RandomName("female"))
            oldKnights(0, 2) = s
        End If

        OldKnightGenerator = oldKnights
    End Function

    Function MAKnightGenerator(maKnights As String(,), fatherName As String, motherName As String) As String(,)
        Dim x As Integer
        Dim s As String

        If maKnights(0, 0) <> "" Then
            For i = 0 To 3
                If maKnights(i, 0) = "" Then Exit For
                x = DiceRoller(1, 20)
                If x <= 12 Then
                    s = $"{fatherName}'s younger brother. " & AliveAndMarried(True)
                ElseIf x <= 16 Then
                    s = $"{motherName}'s brother. " & AliveAndMarried(True)
                ElseIf x <= 19 Then
                    s = $"{fatherName}'s illegitimate brother. " & AliveAndMarried(True)
                Else
                    s = $"{motherName}'s illegitimate brother. " & AliveAndMarried(True)
                End If
                s = Replace(s, " married", " married to " & RandomName("female"))
                maKnights(i, 2) = s
            Next
        End If

        MAKnightGenerator = maKnights
    End Function

    Function YoungKnightGenerator(youngKnights As String(,), charAge As Integer) As String(,)
        Dim x As Integer
        Dim s As String

        If youngKnights(0, 0) <> "" Then
            For i = 0 To 5
                If youngKnights(i, 0) = "" Then Exit For
                x = DiceRoller(1, 20)
                If charAge = 21 Then
                    x = DiceRoller(1, 4)
                    If x = 1 Then
                        s = "Your sister's husband. " & AliveAndMarried(True)
                    ElseIf x <= 3 Then
                        s = "Your first cousin (maternal). " & AliveAndMarried(True)
                    Else
                        s = "An illegitimate older brother. " & AliveAndMarried(True)
                    End If
                Else
                    x = DiceRoller(1, 20)
                    If x <= 8 Then
                        s = "Your younger brother. " & AliveAndMarried(True)
                    ElseIf x <= 14 Then
                        s = "Your first cousin (paternal). " & AliveAndMarried(True)
                    ElseIf x <= 15 Then
                        s = "Your sister's husband. " & AliveAndMarried(True)
                    ElseIf x <= 17 Then
                        s = "Your first cousin (maternal). " & AliveAndMarried(True)
                    ElseIf x <= 18 Then
                        s = "An illegitimate older brother. " & AliveAndMarried(True)
                    Else
                        s = "An illegitimate younger brother. " & AliveAndMarried(True)
                    End If
                End If
                s = Replace(s, " married", " married to " & RandomName("female"))
                youngKnights(i, 2) = s
            Next
        End If

        YoungKnightGenerator = youngKnights
    End Function

    Function AliveAndMarried(Optional definitelyAlive As Boolean = False, Optional mother As Boolean = False) As String
        Dim x As Integer
        Dim s As String

        x = DiceRoller(1, 20)
        If definitelyAlive Then x = x - 6
        If mother Then x = x - 3

        If x <= 7 Then
            s = "Alive and married. "
        ElseIf x <= 14 Then
            s = "Alive and unmarried. "
        ElseIf x <= 17 Then
            s = "Dead (was married). "
        Else
            s = "Dead (never married). "
        End If

        AliveAndMarried = s
    End Function
End Module
