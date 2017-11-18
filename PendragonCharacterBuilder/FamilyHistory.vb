Module FamilyHistory
    Dim charName As String
    Dim fatherName As String
    Dim grandfatherName As String
    Dim oldKnights As String(,)
    Dim maKnights As String(,)
    Dim youngKnights As String(,)
    Dim glory As Integer = 0
    Dim fatherGlory As Integer = 0
    Dim grandfatherGlory As Integer = ((DiceRoller(1, 20) * 100) + 1000) / 10 + DiceRoller(2, 20)
    Dim fatherAlive As Boolean = True
    Dim grandfatherAlive As Boolean = True
    Dim passionArray As New arraylist

    Public Sub Main()
        fatherName = "Bob"
        grandfatherName = "Grand-Bob"

        FamilyHistoryMaker("Steve", fatherName, grandfatherName, oldKnights, maKnights, youngKnights)
    End Sub

    Public Function FamilyHistoryMaker(charName_x As String, fatherName_x As String, grandfatherName_x As String,
                                       oldKnights_x As String(,), maKnights_x As String(,), youngKnights_x As String(,)) As String
        'Returns a multi-paragraph string covering the family history, with starting Glory at the very start.
        charName = charName_x
        fatherName = fatherName_x
        grandfatherName = grandfatherName_x
        oldKnights = oldKnights_x
        maKnights = maKnights_x
        youngKnights = youngKnights_x

        Dim s As String = ""
        Dim x As Integer
        Dim fhm As String = ""

        Dim familyDissident As Boolean = False
        Dim familyDissLeaders As Boolean = False

        fhm = fhm & Year439()
        If grandfatherAlive Then fhm = fhm & Year440()
        If grandfatherAlive Then fhm = fhm & Year441()
        If grandfatherAlive Then fhm = fhm & Year443()
        If grandfatherAlive Then fhm = fhm & Year444()
        If grandfatherAlive Then fhm = fhm & Year446()
        If grandfatherAlive Then
            s = Year447()
            If Left(s, 3) = "DIS" Then
                familyDissident = True
            End If
            s = Mid(s, 3)
            fhm = fhm & s
        End If
        If grandfatherAlive Then
            s = Year450(familyDissident)
            If Left(s, 3) = "REB" Then
                familyDissLeaders = True
            ElseIf Left(s, 3) = "DIS" Then
                familyDissident = True
            End If
            s = Mid(s, 3)
            fhm = fhm & s
        End If
        If grandfatherAlive Then fhm = fhm & Year451(familyDissident)
        If grandfatherAlive Then fhm = fhm & NothingYear("452-454")
        If grandfatherAlive Then fhm = fhm & NothingYear("455-456")

        FamilyHistoryMaker = CStr(glory) & "xx" & fhm
    End Function

    Function NothingYear(yearNo As String) As String
        Dim s As String = ""
        Dim x As Integer

        x = DiceRoller(1, 20)
        If x = 1 Then
            s = yearNo & ": " & grandfatherName & " met his end. " & MiscDeath() & vbNewLine & vbNewLine
            grandfatherAlive = False
        End If

        NothingYear = s
    End Function

    Function Year451(famDiss As Boolean) As String
        Dim s As String = $"451: {grandfatherName}"
        Dim x As Integer

        x = DiceRoller(1, 20)
        If x = 1 Then
            s = $"{s} met his end. " & MiscDeath()
            grandfatherAlive = False
        Else
            If famDiss Then
                s = $"{s} was sent to help Aetius in Gaul and fought at the Battle of Chalons"
                grandfatherGlory += (DiceRoller(1, 6) * 135)
                x = DiceRoller(1, 20)
                If x <= 2 Then
                    s = $"{s}, where he died gloriously."
                    grandfatherGlory += 1000
                    grandfatherAlive = False
                ElseIf x <= 5 Then
                    s = $"{s}, where he died."
                    grandfatherAlive = False
                Else
                    s = $"{s}."
                End If
            Else
                s = $"{s} served garrison duty and saw no combat."
            End If
        End If

        s = s & vbNewLine & vbNewLine
        Year451 = s
    End Function

    Function Year450(famDiss As Boolean) As String
        Dim s As String = ""
        Dim x As Integer

        Console.WriteLine()
        Console.WriteLine("Complaints about Vortigern's Saxon-loving ways become louder.")
        If famDiss Then
            Console.WriteLine("Does your family raise their voices further and stand at the front of this dissenting faction? [y or n]")
        Else
            Console.WriteLine("Does your family join them? [y or n]")
        End If

        Do
            s = Console.ReadLine()
            s = LCase(s)
            s = Left(s, 1)
            If s <> "y" And s <> "n" Then
                Console.WriteLine("Please enter y or n.")
                s = ""
            End If
        Loop While s = ""
        If famDiss Then
            If s = "y" Then
                s = "REB450: The family took a leadership role among those complaining about King Vortigern's Saxon-loving ways"
            Else
                s = "DIS450: The family continued to grumble about the influx of Saxons"
            End If
        Else
            If s = "y" Then
                s = "DIS450: The family's complaints about Saxon immigration had them marked as dissidents by King Vortigern"
            Else
                s = "LYL450: The family stayed loyal to their king, even as he flooded the country with foreigners"
            End If
        End If

        x = DiceRoller(1, 20)
        If x = 1 Then
            s = $"{s}. Unfortunately {grandfatherName} met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 5 Then
            s = $"{s}."
        Else
            s = $"{s}. {grandfatherName} attended the wedding of Vortigern and Rowena"
            grandfatherGlory += 25
            If x >= 19 Then
                s = $"{s} -- where he noticed that Rowena was already pregnant!"
                grandfatherGlory += 25
            Else
                s = $"{s}."
            End If
        End If

        s = s & vbNewLine & vbNewLine
        Year450 = s
    End Function

    Function Year447() As String
        Dim s As String = ""
        Dim x As String

        Console.WriteLine()
        Console.WriteLine("In the years 447-449, King Vortigern brings Saxons upon Saxons into Britain")
        Console.WriteLine("to help prosecute his (largely successful) war against the Picts.")
        Console.WriteLine("Many families complained about this immigration policy. Was yours one of them? [y or n]")

        Do
            s = Console.ReadLine()
            s = LCase(s)
            s = Left(s, 1)
            If s <> "y" And s <> "n" Then
                Console.WriteLine("Please enter y or n.")
                s = ""
            End If
        Loop While s = ""
        If s = "y" Then
            s = "DIS447-449: The family's complaints about Saxon immigration had them marked as dissidents by King Vortigern"
        End If
        If s = "n" Then
            s = "LYL447-449: The family's support of King Vortigern and his Saxon allies did not go unnoticed"
        End If

        x = DiceRoller(1, 20)
        If x = 1 Then
            s = $"{s}, and {grandfatherName} met his end. " & MiscDeath()
            grandfatherAlive = False
        Else
            s = $"{s}."
        End If

        s = s & vbNewLine & vbNewLine
        Year447 = s
    End Function

    Function Year446() As String
        Dim s As String
        Dim x As Integer

        s = $"446: {grandfatherName}"
        x = DiceRoller(1, 20)
        If x = 1 Then
            s = $"{s} met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 5 Then
            s = $"{s} served garrison duty and saw almost no combat."
        Else
            s = $"{s} fought at the Battle of Lincoln"
            grandfatherGlory += (DiceRoller(1, 6) * 60)
            x = DiceRoller(1, 2)
            If x = 2 Then passionArray.Add("PA/Hate (Picts)/" & (DiceRoller(1, 6) + 6))
            x = DiceRoller(1, 20)
            If x = 1 Then
                s = $"{s}, where he died gloriously."
                grandfatherGlory += 1000
                grandfatherAlive = False
            ElseIf x <= 3 Then
                s = $"{s}, where he died."
                grandfatherAlive = False
            Else
                s = $"{s}."
            End If
        End If

        s = s & vbNewLine & vbNewLine
        Year446 = s
    End Function

    Function Year444() As String
        Dim s As String
        Dim x As Integer

        s = $"444-445: {grandfatherName}"
        x = DiceRoller(1, 20)
        If x = 1 Then
            s = $"{s} met his end. " & MiscDeath()
            grandfatherAlive = False
        Else
            s = $"{s} served garrison duty and"
            x = DiceRoller(1, 19)
            If x = 1 Then
                s = $"{s} was killed by raiders."
                grandfatherAlive = False
                grandfatherGlory += 10
            ElseIf x <= 15 Then
                s = $"{s} defended well."
                grandfatherGlory += 10
            Else
                s = $"{s} saw basically no fighting."
            End If
        End If

        s = s & vbNewLine & vbNewLine
        Year444 = s
    End Function

    Function Year443() As String
        Dim s As String
        Dim x As Integer

        s = $"443: {grandfatherName}"
        x = DiceRoller(1, 20)

        If x = 1 Then
            s = $"{s} met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 10 Then
            s = $"{s} served garrison duty and was killed by raiders."
            grandfatherGlory += 20
            grandfatherAlive = False
        ElseIf x <= 18 Then
            s = $"{s} served garrison duty."
            grandfatherGlory += 20
        ElseIf x = 19 Then
            s = $"{s} was present at King Constans' murder but unable to prevent it."
            x = DiceRoller(1, 2)
            If x = 2 Then passionArray.Add("PA/Hate (Picts)/" & (DiceRoller(1, 6) + 6))
        Else
            s = $"{s} was present at King Constans' murder and died trying to save him."
            grandfatherGlory += 1000
            grandfatherAlive = False
            x = DiceRoller(1, 4)
            If x > 1 Then passionArray.Add("PA/Hate (Picts)/" & (DiceRoller(1, 6) + 6))
        End If

        s = s & vbNewLine & vbNewLine
        Year443 = s
    End Function

    Function Year441() As String
        Dim s As String
        Dim x As Integer

        s = $"441-442: {grandfatherName}"
        x = DiceRoller(1, 20)
        If x = 1 Then
            s = s & " met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 6 Then
            s = s & " served garrison duty and was killed by raiders."
            grandfatherGlory += 20
            grandfatherAlive = False
        Else
            s = $"{s} served garrison duty."
            grandfatherGlory += 20
        End If

        s = s & vbNewLine & vbNewLine
        Year441 = s
    End Function

    Function Year440() As String
        Dim s As String
        Dim x As Integer

        s = "440: "
        x = DiceRoller(1, 20)
        If x = 1 Then
            s = s & grandfatherName & " met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 9 Then
            s = s & grandfatherName & " stood garrison duty and was killed by Picts."
            grandfatherGlory += 20
            grandfatherAlive = False
        ElseIf x <= 18 Then
            s = s & grandfatherName & " stood garrison duty and fought Pictish raiders."
            grandfatherGlory += 10
        ElseIf x = 19 Then
            s = s & grandfatherName & " was present at the murder of King Constantin, but unable to protect his king."
            x = DiceRoller(1, 2)
            If x = 2 Then passionArray.Add("DT/Suspicious (Silchester Knights)/" & (DiceRoller(1, 6) + 6))
        Else
            s = s & grandfatherName & " was present at the murder of King Constantin and died trying to protect him."
            grandfatherGlory += 1000
            grandfatherAlive = False
            x = DiceRoller(1, 20)
            If x >= 4 Then passionArray.Add("DT/Suspicious (Silchester Knights)/" & (DiceRoller(1, 6) + 6))
        End If

        s = s & vbNewLine & vbNewLine
        Year440 = s
    End Function

    Function Year439() As String
        Dim s As String
        Dim x As Integer

        s = "439: A son is born to " & grandfatherName & "; your father " & fatherName & ". "
        x = DiceRoller(1, 20)
        If x = 1 Then
            s = s & "Unfortunately " & grandfatherName & " then died" & Mid(MiscDeath(), 8)
            grandfatherAlive = False
        ElseIf x <= 6 Then
            'Nothing.
        Else
            s = s & grandfatherName & " also fought in the Battle of Carlion"
            grandfatherGlory += (DiceRoller(1, 6) * 15)
            x = DiceRoller(1, 20)
            If x = 1 Then
                s = s & ", where he died gloriously."
                grandfatherGlory += 1000
                grandfatherAlive = False
            ElseIf x = 2 Then
                s = s & ", where he died."
                grandfatherAlive = False
            Else
                s = s & "."
            End If
            x = DiceRoller(1, 20)
            If x >= 16 Then
                passionArray.Add("PA/Hate (Irish)/" & CStr(DiceRoller(3, 6)))
            End If
        End If
        s = s & vbNewLine & vbNewLine

        Year439 = s
    End Function

    Function MiscDeath(Optional male As Boolean = True) As String
        Dim s As String
        Dim x As Integer

        x = DiceRoller(1, 20)
        If male Then
            s = "He died "
            Select Case x
                Case 1
                    s = s & "in battle, fighting a personal feud."
                Case 2
                    s = s & "in battle, fighting a personal feud."
                Case 3
                    s = s & "in battle, in a neighbouring land."
                Case 4
                    s = s & "in battle, in a neighbouring land."
                Case 5
                    s = s & "in battle, in a neighbouring land."
                Case 6
                    s = s & "in battle against foreign invaders."
                Case 7
                    s = s & "in battle against foreign invaders."
                Case 8
                    s = s & "in a hunting accident."
                Case 9
                    s = s & "in a hunting accident."
                Case 10
                    s = s & "in a miscellaneous accident."
                Case 11
                    s = s & "in a miscellaneous accident."
                Case 12
                    s = s & "of natural causes."
                Case 13
                    s = s & "of natural causes."
                Case 14
                    s = s & "of natural causes."
                Case 15
                    s = s & "of natural causes."
                Case 16
                    s = s & "of natural causes."
                Case 17
                    s = s & "(or disappeared) in unknown and mysterious circumstances."
                Case 18
                    s = s & "(or disappeared) in unknown and mysterious circumstances."
                Case 19
                    s = s & "(or disappeared) in unknown and mysterious circumstances."
                Case 20
                    s = s & "(or disappeared) in unknown and mysterious circumstances."
            End Select
        Else
            s = "She "
            Select Case x
                Case 1
                    s = s & "was killed by raiders."
                Case 2
                    s = s & "was kidnapped by raiders and died in captivity."
                Case 3
                    s = s & "died in a accident."
                Case 4
                    s = s & "died in a accident."
                Case 5
                    s = s & "died in a accident."
                Case 6
                    s = s & "died of pregnancy-related complications."
                Case 7
                    s = s & "died of pregnancy-related complications."
                Case 8
                    s = s & "died of pregnancy-related complications."
                Case 9
                    s = s & "died of pregnancy-related complications."
                Case 10
                    s = s & "died of pregnancy-related complications."
                Case 11
                    s = s & "died of pregnancy-related complications."
                Case 12
                    s = s & "died of natural causes."
                Case 13
                    s = s & "died of natural causes."
                Case 14
                    s = s & "died of natural causes."
                Case 15
                    s = s & "died of natural causes."
                Case 16
                    s = s & "died of natural causes."
                Case 17
                    s = s & "died of natural causes."
                Case 18
                    s = s & "died of natural causes."
                Case 19
                    s = s & "died (or disappeared) in unknown and mysterious circumstances."
                Case 20
                    s = s & "died (or disappeared) in unknown and mysterious circumstances."
            End Select
        End If

        MiscDeath = s
    End Function
End Module
