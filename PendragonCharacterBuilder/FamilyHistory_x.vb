Module FamilyHistory_x
    Dim charName As String
    Dim fatherName As String
    Dim motherName As String
    Dim grandfatherName As String
    Dim oldKnights As String(,)
    Dim maKnights As String(,)
    Dim youngKnights As String(,)
    Dim glory As Integer = 0
    Dim fatherGlory As Integer = 0
    Dim grandfatherGlory As Integer = ((DiceRoller(1, 20) * 100) + 1000) / 10 + DiceRoller(2, 20)
    Dim fatherAlive As Boolean = True
    Dim grandfatherAlive As Boolean = True
    Dim passionArray As New ArrayList
    Dim extraHeirlooms As Integer = 0
    Dim charBirthYear As Integer = 0

    Public Sub Main()
        fatherName = "Bob"
        grandfatherName = "Grand-Bob"

        FamilyHistoryMaker("Steve", fatherName, grandfatherName, oldKnights, maKnights, youngKnights, 464)
    End Sub

    Public Function FamilyHistoryMaker(charName_x As String, fatherName_x As String, grandfatherName_x As String,
                                       oldKnights_x As String(,), maKnights_x As String(,), youngKnights_x As String(,),
                                       charBirthYear_x As Integer) As String
        'Returns a multi-paragraph string covering the family history, with starting Glory at the very start.
        charName = charName_x
        fatherName = fatherName_x
        motherName = RandomName("female")
        grandfatherName = grandfatherName_x
        oldKnights = oldKnights_x
        maKnights = maKnights_x
        youngKnights = youngKnights_x
        charBirthYear = charBirthYear_x

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
            s = Mid(s, 4)
            fhm = fhm & s
        End If
        If grandfatherAlive Then
            s = Year450(familyDissident)
            If Left(s, 3) = "REB" Then
                familyDissLeaders = True
            ElseIf Left(s, 3) = "DIS" Then
                familyDissident = True
            End If
            s = Mid(s, 4)
            fhm = fhm & s
        End If
        If grandfatherAlive Then fhm = fhm & Year451(familyDissident)
        If grandfatherAlive Then fhm = fhm & GrandfatherNothingYear("452-454")
        If grandfatherAlive Then fhm = fhm & GrandfatherNothingYear("455-456")
        If grandfatherAlive Then fhm = fhm & Year457(familyDissident)
        If grandfatherAlive Then fhm = fhm & GrandfatherNothingYear("458-459")
        fhm = fhm & Year460()
        fhm = fhm & Year461()
        fhm = fhm & Year463()
        fhm = fhm & Year464()

        FamilyHistoryMaker = CStr(glory) & "xx" & fhm
        'And also return extraHeirlooms
    End Function

    Function GrandfatherNothingYear(yearNo As String) As String
        Dim s As String = ""
        Dim x As Integer = DiceRoller(1, 20)

        If x = 1 Then
            s = yearNo & ": " & grandfatherName & " met his end. " & MiscDeath() & vbNewLine & vbNewLine
            grandfatherAlive = False
        End If

        GrandfatherNothingYear = s
    End Function

    Function Year464() As String
        If charBirthYear = 464 Then
            Year464 = WeddingBirth("")
            Year464 = Replace(Year464, " He also", $"464: {fatherName}")
            Year464 = Year464 & vbNewLine & vbNewLine
        Else
            Year464 = ""
        End If
    End Function

    Function Year463() As String
        If grandfatherAlive Then
            Year463 = $"463: At the 'Night of Long Knives', {grandfatherName} is murdered. {fatherName} survived, one way or another."
            grandfatherAlive = False
        End If
        If charBirthYear = 463 Then
            Year463 = WeddingBirth(Year463)
        End If
        passionArray.Add("PA/Hate (Saxons)/" & (6 + DiceRoller(3, 6)))

        If passionArray.Count > 0 Then
            ConsolidatePassions()

            Console.WriteLine()
            Console.Write("When he died, your grandfather had " & passionArray.Count & " passion")
            If passionArray.Count > 1 Then Console.Write("s")
            Console.Write(".")
            Console.WriteLine()
            Console.WriteLine("For each passion, please choose if your father also had it by typing y or n.")

            Dim s As String
            Dim p As String
            Dim v As Integer
            Dim x As Integer
            Dim entry As String
            Dim i As Integer
            Dim stopValue As Integer

            s = ""

            stopValue = passionArray.Count - 1
            i = 0
            Do While i <= stopValue
                entry = passionArray(i)
                p = Mid(entry, 4)
                x = InStrRev(p, "/")
                v = CInt(Mid(p, x + 1))
                p = Left(p, x - 1)
                Do
                    Console.WriteLine(p & ": " & CStr(v) & " -- y or n?")
                    s = Console.ReadLine()
                    s = LCase(s)
                    s = Left(s, 1)
                    If s <> "y" And s <> "n" Then
                        s = ""
                        Console.WriteLine("Keep this trait, y or n?")
                    End If
                Loop While s = ""
                If s = "n" Then
                    passionArray.Remove(entry)
                    stopValue -= 1
                    i -= 1
                End If
                i += 1
            Loop
        End If
    End Function

    Sub ConsolidatePassions()
        Dim s As String
        Dim s2 As String
        Dim x As Integer
        Dim x2 As Integer
        Dim x3 As Integer
        Dim passValue As Integer
        Dim passValue2 As Integer
        Dim stopValue As Integer
        Dim i As Integer
        Dim j As Integer

        'passionArray.Add("PA/Hate (Saxons)/13")
        'passionArray.Add("PA/Hate (Saxons)/14")
        'passionArray.Add("PA/Hate (Saxons)/12")
        stopValue = passionArray.Count - 1

        i = 0
        Do While i <= stopValue
            s = passionArray(i)
            x2 = passionArray.IndexOf(s)
            s = Mid(s, 4)
            x = InStrRev(s, "/")
            passValue = Mid(s, x + 1)
            s = Left(s, x - 1)
            j = 0
            Do While j <= stopValue
                s2 = passionArray(j)
                x3 = passionArray.IndexOf(s2)
                If x2 <> x3 Then
                    s2 = Mid(s2, 4)
                    x = InStrRev(s2, "/")
                    passValue2 = Mid(s2, x + 1)
                    s2 = Left(s2, x - 1)
                    If s = s2 Then
                        If passValue >= passValue2 Then
                            passionArray.RemoveAt(x3)
                            stopValue -= 1
                        Else
                            passionArray.RemoveAt(x2)
                            stopValue -= 1
                            j = j - 1
                            i = i - 1
                        End If
                    End If
                End If
                j = j + 1
            Loop
            i = i + 1
        Loop
    End Sub

    Function Year461() As String
        Dim s As String = ""
        Dim x As Integer = DiceRoller(1, 20)

        If grandfatherAlive = True Then
            s = $"461-462: {fatherName} served garrison duty."
            If charBirthYear = 461 Or charBirthYear = 462 Then
                s = WeddingBirth(s)
            End If
            If x = 1 Then
                s = $"{s} Meanwhile, {grandfatherName} met his end. " & MiscDeath()
                s = s & vbNewLine & vbNewLine
                grandfatherAlive = False
            ElseIf x <= 5 Then
                s = ""
            Else
                s = $"{s} Meanwhile, {grandfatherName} fought at the Battle of Cambridge"
                grandfatherGlory += (DiceRoller(1, 6) * 30)
                x = DiceRoller(1, 20)
                If x = 1 Then
                    s = $"{s}, where he died gloriously in battle."
                    grandfatherAlive = False
                    grandfatherGlory += 1000
                ElseIf x <= 3 Then
                    s = $"{s}, where he died."
                    grandfatherAlive = False
                Else
                    s = $"{s}."
                    grandfatherGlory += 10
                End If
                s = s & vbNewLine & vbNewLine
            End If
        Else
            s = ""
        End If

        Year461 = s
    End Function

    Function WeddingBirth(s As String)
        Dim x As Integer = DiceRoller(1, 20)
        Dim x2 As Integer
        s = $"{s} He also married {motherName} (a widow) and fathered {charName}."

        If x <= 4 Then
            x2 = 25
        ElseIf x <= 7 Then
            x2 = 50
        ElseIf x <= 17 Then
            x2 = 100
        ElseIf x <= 19 Then
            x2 = 200
        Else
            x2 = 350
        End If

        fatherGlory += x2   'This will vanish if this is called earlier than Year460.
        WeddingBirth = s
    End Function

    Function Year460() As String
        Dim s As String = ""
        Dim x As Integer = DiceRoller(1, 20)

        s = $"460: {fatherName}, son of {grandfatherName} began his service as a knight in this year."
        fatherGlory = 1000 + Math.Round(grandfatherGlory / 10)

        If charBirthYear = 460 Then
            s = WeddingBirth(s)
        End If

        If x = 1 And grandfatherAlive = True Then
            s = $"{s} Unfortunately {grandfatherName} met his end. " & MiscDeath()
            grandfatherAlive = False
        End If

        s = s & vbNewLine & vbNewLine
        Year460 = s
    End Function

    Function Year457(famDiss As Boolean) As String
        Dim s As String = ""
        Dim x As Integer = DiceRoller(1, 20)

        If famDiss Then x = x + 10
        s = $"457: {grandfatherName}"

        If x = 1 Then
            s = $"{s} met his end. " & MiscDeath()
            grandfatherAlive = False
        ElseIf x <= 8 Then
            s = $"{s} managed to avoid every last shred of excitement this year."
        Else
            s = $"{s} fought at the Battle of Kent, on the side of the "
            If famDiss Then
                s = s & "rebels"
                x = DiceRoller(1, 20)
                grandfatherGlory += Math.Round((DiceRoller(1, 6) * 30 * 0.75))
                If x = 1 Then
                    s = $"{s}, where he died gloriously in battle."
                    grandfatherGlory += 1000
                    grandfatherAlive = False
                ElseIf x <= 5 Then
                    s = $"{s}, where he died in battle."
                    grandfatherAlive = False
                ElseIf x = 20 Then
                    s = $"{s}, where he defeated one of Vortigern's bodyguards but couldn't quite reach the king."
                    grandfatherGlory += 100
                Else
                    s = s & "."
                End If
            Else
                s = s & "loyalists"
                x = DiceRoller(1, 20)
                grandfatherGlory += (DiceRoller(1, 6) * 45)
                If x = 1 Then
                    s = $"{s}, where he died gloriously in battle."
                    grandfatherGlory += 1000
                    grandfatherAlive = False
                ElseIf x <= 3 Then
                    s = $"{s}, where he died in battle."
                    grandfatherAlive = False
                ElseIf x = 20 Then
                    s = $"{s}, where he defeated a rebel hero who was trying to fight through Vortigern's bodyguards. Later, the king publically praised him for valour and bestowed upon him many gifts."
                    grandfatherGlory += 125
                    extraHeirlooms += 2
                Else
                    s = s & "."
                End If
            End If
        End If

        s = s & vbNewLine & vbNewLine
        Year457 = s
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
        Console.WriteLine("In 450, complaints about Vortigern's Saxon-loving ways became louder.")
        If famDiss Then
            Console.WriteLine("Did your family raise their voices further and stand at the front of this dissenting faction? [y or n]")
        Else
            Console.WriteLine("Did your family join them? [y or n]")
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
End Module