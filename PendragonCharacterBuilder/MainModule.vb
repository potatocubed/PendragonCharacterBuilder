Module MainModule

    Sub Main()
        'The basic program will ask for a name and gender and assume that you're a Salisbury knight from the core book.
        'It won't assign your attributes or your various bonuses, just the standard stuff.
        'Then it'll output an XML file, then convert that into basic forum bbcode and plain text.
        'Everything generate-able from a random table will be generated from a table.

        Dim charName As String
        Dim charGender As String
        Dim tradWoman As Boolean = False
        Dim charAge As Integer
        Dim charYearBorn As Integer
        Dim homeland As String = "Salisbury"
        Dim culture As String = "Cymric"
        Dim charReligion As String = ""
        Dim charSonNumber As Integer = 1
        Dim charLeige As String = "Sir Roderick, Earl of Salisbury"
        Dim charClass As String = "squire"  'Might be different for women?
        Dim charManor As String = RandomHome()
        Dim charTraits As String(,)
        charTraits = InitialiseCharTraits()
        Dim charPassions As New ArrayList
        charPassions.Add("Loyalty (Lord)/15")
        charPassions.Add("Love (Family)/15")
        charPassions.Add("Hospitality/15")
        charPassions.Add("Honour/15")
        Dim charDirectedTraits As New ArrayList
        Dim charSIZ As Integer
        Dim charDEX As Integer
        Dim charSTR As Integer
        Dim charCON As Integer
        Dim charAPP As Integer
        Dim charFeatures As String
        Dim charSkills As String(,)
        Dim charGlory As Integer
        Dim fatherGlory As Integer
        Dim grandfatherGlory As Integer
        Dim pUncles As New ArrayList()
        Dim pAunts As New ArrayList()
        Dim brothers As New ArrayList()
        Dim sisters As New ArrayList()
        Dim mUncles As New ArrayList()
        Dim mAunts As New ArrayList()
        Dim cousins As New ArrayList()
        Dim familyHistory As String
        Dim charHorses As New ArrayList()
        charHorses.Add("Charger")
        charHorses.Add("Rouncy")
        charHorses.Add("Rouncy")
        charHorses.Add("Sumpter")
        Dim charSquire As String() = {RandomName(), "15", "First Aid", "6", "Battle", "1", "Horsemanship", "6", "xx", "5"}
        Dim charHeirlooms As New ArrayList   'Because you can have more than one!
        Dim charFamilyCharacteristic As String
        'Name, alive/dead, notes.
        Dim charOldKnights As String(,) = New String(0, 2) {}
        Dim charMAKnights As String(,) = New String(4, 2) {}
        Dim charYoungKnights As String(,) = New String(6, 2) {}
        Dim charLineageMen As Integer = DiceRoller(2, 6) + 5
        Dim charLevies As Integer = DiceRoller(5, 20)

        Dim fatherName As String = RandomName()
        Dim grandfatherName As String = RandomName()

        While fatherName = grandfatherName
            grandfatherName = RandomName()
        End While

        Dim motherName As String = RandomName("female")

        Dim fatherClass As String = "vassal knight"

        Dim xhim As String
        Dim xhis As String
        Dim xhe As String
        Dim skArray As New ArrayList()

        Dim x As Integer
        Dim x2 As Integer
        Dim x3 As Integer
        Dim spr As Integer
        Dim aspr As Integer
        Dim attMin As Integer
        Dim attMax As Integer
        Dim s As String
        Dim s2 As String
        Dim b As Boolean

        Dim here As String = My.Computer.FileSystem.CurrentDirectory
        'DEBUG
        here = "C:\Users\LonghurstC\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder"

        Dim sList As String = here & "\xml\pdcc_skill_list.xml"
        Dim hList As String = here & "\xml\pdcc_heirlooms.xml"

        Console.WriteLine("Welcome to the Pendragon character generator!")
        Console.WriteLine("This program will basically hammer through all the random tables in the")
        Console.WriteLine("Pendragon core book For you, but it won't make many actual *decisions*.")
        Console.WriteLine("So you'll have to dot the i's and cross the t's yourself.")
        Console.WriteLine()
        Console.WriteLine("To begin, is your character male [m] or female [f]?")
        charGender = ""
        Do
            charGender = Console.ReadLine()
            charGender = Left(charGender, 1)    'Getting just the first letter.
            charGender = LCase(charGender)
            If charGender <> "m" And charGender <> "f" Then
                charGender = ""
                Console.WriteLine("Please enter m or f.")
            End If
        Loop While charGender = ""

        If charGender = "m" Then
            charGender = "male"
            xhim = "him"
            xhis = "his"
            xhe = "he"
        Else
            charGender = "female"
            xhim = "her"
            xhis = "her"
            xhe = "she"
        End If

        Console.WriteLine()
        Console.WriteLine("You have chosen a " & charGender & " character. What is " & xhis & " name? ['random' for a random name]")
        charName = ""
        Do
            charName = Console.ReadLine()
            charName = StrConv(charName, VbStrConv.ProperCase)
            If charName = "Random" Then charName = RandomName(charGender)
        Loop While charName = ""

        Console.WriteLine()
        Console.WriteLine("And how old is " & charName & "? [A whole number between 21 and 26]")
        Console.WriteLine("(Older characters are closer to death but get some extra points and wotnot.)")
        Do
            x = Nothing
            s = Console.ReadLine()
            Try
                x = CInt(s)
            Catch ex As Exception
                x = Nothing
            End Try
            If x < 21 Or x > 26 Then x = Nothing
            If x = Nothing Then Console.WriteLine("Please enter a whole number between 21 and 26.")
        Loop While x = Nothing
        charAge = x
        charYearBorn = 485 - charAge

        Console.WriteLine()
        Console.WriteLine("Which religion do they follow?")
        Console.WriteLine("British Christian [b], Roman Christian [r], or Pagan [p]?")
        Do
            s = Nothing
            s = Console.ReadLine()
            s = LCase(s)
            s = Left(s, 1)
            Select Case s
                Case "b"
                    charReligion = "British Christian"
                    charTraits = TraitUpdate(charTraits, "Chaste", 13)
                    charTraits = TraitUpdate(charTraits, "Energetic", 13)
                    charTraits = TraitUpdate(charTraits, "Generous", 13)
                    charTraits = TraitUpdate(charTraits, "Modest", 13)
                    charTraits = TraitUpdate(charTraits, "Temperate", 13)
                Case "r"
                    charReligion = "Roman Christian"
                    charTraits = TraitUpdate(charTraits, "Chaste", 13)
                    charTraits = TraitUpdate(charTraits, "Forgiving", 13)
                    charTraits = TraitUpdate(charTraits, "Merciful", 13)
                    charTraits = TraitUpdate(charTraits, "Modest", 13)
                    charTraits = TraitUpdate(charTraits, "Temperate", 13)
                Case "p"
                    charReligion = "Pagan"
                    charTraits = TraitUpdate(charTraits, "Generous", 13)
                    charTraits = TraitUpdate(charTraits, "Energetic", 13)
                    charTraits = TraitUpdate(charTraits, "Honest", 13)
                    charTraits = TraitUpdate(charTraits, "Lustful", 13)
                    charTraits = TraitUpdate(charTraits, "Proud", 13)
                Case Else
                    s = Nothing
            End Select
            If s = Nothing Then Console.WriteLine("Please choose b, r, or p.")
        Loop While s = Nothing
        charTraits = TraitUpdate(charTraits, "Valorous", 15)

        Console.WriteLine()
        Console.WriteLine("And what's " & charName & "'s most famous trait?")
        Console.Write("(")
        For i = 0 To 12
            Console.Write(charTraits(0, i))
            If i < 12 Then Console.Write(", ")
        Next i
        Console.WriteLine()
        For i = 13 To 25
            Console.Write(charTraits(0, i))
            If i < 25 Then Console.Write(", ")
            If i = 25 Then Console.Write("?)")
        Next i
        Console.WriteLine()

        Do
            s = Nothing
            s = Console.ReadLine()
            s = StrConv(s, VbStrConv.ProperCase)
            For i = 0 To 25
                If charTraits(0, i) = s Then
                    Exit Do
                End If
            Next i
            s = Nothing
            Console.WriteLine("Please choose a trait which exists.")
        Loop While s = Nothing
        charTraits = TraitUpdate(charTraits, s, 16)

        If charGender = "female" Then
            Console.WriteLine()
            Console.WriteLine("Is she a female knight, or a more traditional Arthurian woman? [k or t]")
            Do
                s = Console.ReadLine()
                s = Left(s, 1)    'Getting just the first letter.
                s = LCase(s)
                If s <> "k" And s <> "t" Then
                    s = ""
                    Console.WriteLine("Please enter k for knight, or t for tradition.")
                End If
            Loop While s = ""
            If s = "k" Then
                charSkills = InitialiseCharSkills(sList)
            Else
                tradWoman = True
                charSkills = InitialiseCharSkills(sList, "female")
                charClass = "lady"
                charSquire(0) = RandomName("female")
                charSquire(6) = ""
                charSquire(9) = "6"
            End If
        Else
            charSkills = InitialiseCharSkills(sList)
        End If

        Select Case LCase(Left(charReligion, 1))
            Case "b"
                charSkills(0, 21) = Replace(charSkills(0, 21), "xx", "British Christianity")
            Case "r"
                charSkills(0, 21) = Replace(charSkills(0, 21), "xx", "Roman Christianity")
            Case "p"
                charSkills(0, 21) = Replace(charSkills(0, 21), "xx", "Pagan")
        End Select

        Console.WriteLine()
        Console.WriteLine("You've got 60 points to assign to the five attributes: SIZ, DEX, STR, CON, and APP.")
        Console.WriteLine("There are no takebacks because I'm a lazy coder, so work it out on paper first.")

        spr = 60
        aspr = 28

        For i = 1 To 5
            attMin = 5
            attMax = 18
            Select Case i
                Case 1
                    s = "SIZ"
                    attMin = 8
                    aspr = 20
                Case 2
                    s = "DEX"
                    aspr = 15
                Case 3
                    s = "STR"
                    aspr = 10
                Case 4
                    s = "CON"
                    aspr = 5
                Case 5
                    s = "APP"
                    If attMin < spr Then attMin = spr
                    If attMin >= attMax Then attMin = attMax
                    aspr = 0
            End Select
            If spr - aspr < attMax Then attMax = spr - aspr
            'If spr < attMax Then attMax = spr
            Console.WriteLine("What is " & charName & "'s " & s & "? [a whole number between " & attMin & " and " & attMax & "]")
            If i = 4 Then Console.WriteLine("(Remember that Cymric characters get +3 CON after you've chosen.)")
            Console.WriteLine("You have " & spr & " points remaining.")

            Do
                x = Nothing
                s = Console.ReadLine()
                Try
                    x = CInt(s)
                Catch ex As Exception
                    x = Nothing
                End Try
                If x < attMin Or x > attMax Then x = Nothing
                If x <> Nothing Then
                    spr = spr - x
                    Select Case i
                        Case 1
                            charSIZ = x
                        Case 2
                            charDEX = x
                        Case 3
                            charSTR = x
                        Case 4
                            charCON = x + 3
                            Console.WriteLine("(Cymric +3 bonus added!)")
                        Case 5
                            charAPP = x
                    End Select
                    Console.WriteLine()
                End If
                If x = Nothing Then Console.WriteLine("Please enter a number between " & attMin & " and " & attMax & ".")
            Loop While x = Nothing
        Next i

        Console.WriteLine()
        Console.WriteLine("Choose a knightly skill to be awesome at:")
        If tradWoman Then Console.WriteLine("(Yes, even as a traditional woman. The rules are a little odd here.)")
        skArray = SkillLister(sList, "true")
        PrintSkillList(skArray)
        s = Nothing
        Do
            s = Console.ReadLine()
            s = StrConv(s, VbStrConv.ProperCase)
            If Not skArray.Contains(s) Then
                s = Nothing
            Else
                For i = 0 To 39
                    If charSkills(0, i) = s Then charSkills(1, i) = 15
                Next i
            End If
            If s = Nothing Then Console.WriteLine("Please choose one of the knightly skills above.")
        Loop While s = Nothing

        skArray.Clear()
        If tradWoman Then
            'Women are allowed to be good at medicine, fashion, etc.
            skArray = SkillLister(sList, "", "", False)
        Else
            skArray = SkillLister(sList, "", False, False)
        End If
        x = skArray.IndexOf("Religion (xx)")
        skArray(x) = charSkills(0, 21)

        'Clear out the skills at 10+
        For i = 0 To 39
            If charSkills(1, i) >= 10 Then
                s = charSkills(0, i)
                skArray.Remove(s)
            End If
        Next

        Console.WriteLine()
        Console.WriteLine("And now choose three non-combat skills to be good at:")
        For j = 1 To 3
            PrintSkillList(skArray)
            s = Nothing
            Do
                s = Console.ReadLine
                s = StrConv(s, VbStrConv.ProperCase)
                If s = "Read" Then
                    s = charSkills(0, 19)
                ElseIf s = "Play" Then
                    s = charSkills(0, 18)
                ElseIf s = "Religion" Then
                    s = charSkills(0, 21)
                End If
                If Not skArray.Contains(s) Then
                    s = Nothing
                    Console.WriteLine("Please choose one of the skills above.")
                Else
                    skArray.Remove(s)
                    For k = 0 To 25
                        If charSkills(0, k) = s Then
                            charSkills(1, k) = 10
                            Exit For
                        End If
                    Next
                End If
            Loop While s = Nothing
        Next

        'Console.WriteLine()
        'Console.WriteLine("You get some more options for customising your character, ")
        'Console.WriteLine("but those are way too fiddly for me to bother with here")
        'Console.WriteLine("They'll be summarised on the character sheet output.")

        Heightening(charSkills, charTraits, charPassions, charSIZ, charDEX, charSTR, charCON, charAPP)

        Console.WriteLine()
        If tradWoman Then
            Console.WriteLine("Your lady-in-waiting's name is " & charSquire(0) & ". Choose a skill for her to be vaguely okay at:")
            Console.WriteLine("(The rules are a little vague about lady-in-waiting skills, so you'll")
            Console.WriteLine("probably want to run them past your GM after character generation Is done.)")
        Else
            Console.WriteLine("Your squire's name is " & charSquire(0) & ". Choose a skill for him to be vaguely okay at:")
        End If

        skArray.Clear()
        If tradWoman Then
            skArray = SkillLister(sList, "", "", False)
        Else
            skArray = SkillLister(sList, "", False, "")
        End If
        PrintSkillList(skArray)
        x = skArray.IndexOf("Religion (xx)")
        skArray(x) = charSkills(0, 21)

        s = Nothing
        Do
            s = Console.ReadLine
            s = StrConv(s, VbStrConv.ProperCase)
            If s = "Read" Then
                s = charSkills(0, 19)
            ElseIf s = "Play" Then
                s = charSkills(0, 18)
            ElseIf s = "Religion" Then
                s = charSkills(0, 21)
            End If
            If Not skArray.Contains(s) Then
                s = Nothing
                Console.WriteLine("Please choose one of the skills above.")
            Else
                charSquire(8) = s
            End If
        Loop While s = Nothing

        'STUFF goes here but it's entirely standard.
        Console.WriteLine()
        Console.Write("You have inherited something from your deceased father: ")
        If charReligion = "pagan" Then
            s = HeirloomGenerator(hList, True, True)
        Else
            s = HeirloomGenerator(hList, False, True)
        End If
        x = InStr(s, "//")
        If x > 0 Then
            charHeirlooms.Add(Left(s, x - 1))
            charHeirlooms.Add(Mid(s, x + 2))
        Else
            charHeirlooms.Add(s)
        End If

        'A quick bit to count extra horses.
        If InStr(s, "charger") Then charHorses.Add("Charger")
        If InStr(s, "courser") Then charHorses.Add("Courser")
        x2 = InStr(s, "rouncy")
        If x2 > 0 Then
            charHorses.Add("Rouncy")
            If InStr(x2 + 1, s, "rouncy") Then charHorses.Add("Rouncy")
        End If

        'And now back to your regularly-scheduled heirloom announcement.
        s = Replace(s, "//", " AND ")
        Console.Write(s & ".")
        Console.WriteLine()

        Console.WriteLine()
        Console.WriteLine("Finally, you get a heritable family characteristic!")
        Console.Write("The ")
        If charGender = "male" Then
            Console.Write("men")
        Else
            Console.Write("women")
        End If
        Console.Write(" of your line are all ")
        s = SpecialGiftGenerator(here, charGender)
        charFamilyCharacteristic = s
        Console.Write(s & ".")

        'And now a short bit to add the bonus you just got.
        'First check to see if you got one bonus or two.
        x = 0
        For i = 1 To Len(s) - 1
            s2 = Mid(s, i, 1)
            If s2 = "+" Then x += 1
        Next

        If x = 1 Then
            'If you got one, was it APP?
            If InStr(s, "APP") > 0 Then
                x = charAPP
                charAPP = charAPP + 10
                If charAPP > 18 Then charAPP = 18
            Else
                charSkills = SkillUpdater(s, charSkills)
            End If
        Else
            'Lucky you, you get two!
            x = InStr(s, "(")
            s = Mid(s, x)
            x = InStr(s, "and")
            s2 = "(" & Trim(Mid(s, x + 4))
            s = Left(s, x - 1) & ")"
            charSkills = SkillUpdater(s, charSkills)
            charSkills = SkillUpdater(s2, charSkills)
        End If

        x = 0
        If charAPP <= 6 Then
            x = 3
        ElseIf charAPP <= 9 Then
            x = 2
        ElseIf charAPP <= 12 Then
            x = 1
        ElseIf charAPP <= 16 Then
            x = 2
        Else
            x = 3
        End If

        'Distinctive features!
        'Right at the end because a woman's APP might go up at the family features stage.
        Console.WriteLine()
        If x = 1 Then
            Console.WriteLine("Thanks to " & xhis & " APP of " & charAPP & ", " & charName & " has " & x & " distinctive feature:")
        Else
            Console.WriteLine("Thanks to " & xhis & " APP of " & charAPP & ", " & charName & " has " & x & " distinctive features:")
        End If
        charFeatures = $"Something about {xhis} "

        Dim fArray As New ArrayList()
        fArray.Add("hair")
        fArray.Add("body")
        fArray.Add("facial expression")
        fArray.Add("speech")
        fArray.Add("facial feature")
        fArray.Add("limbs")

        For i = 1 To x
            x2 = DiceRoller(1, fArray.Count)
            charFeatures = charFeatures & (fArray(x2 - 1))
            fArray.RemoveAt(x2 - 1)
            If i = x Then
                charFeatures = charFeatures & "."
            ElseIf i = (x - 1) Then
                charFeatures = charFeatures & $" and {xhis} "
            Else
                charFeatures = charFeatures & $", {xhis} "
            End If
        Next i
        Console.WriteLine(charFeatures)
        Console.WriteLine()

        'Probably something about holdings here, which is just the same as home.

        x2 = 0
        s = " ("
        x = DiceRoller(1, 6) - 5
        If x = 1 Then
            charOldKnights(0, 0) = RandomName()
            charOldKnights(0, 1) = "alive"
            charOldKnights(0, 2) = ""
            s = s & "1 old, "
            x2 = x2 + x
        End If


        x = DiceRoller(1, 6) - 2
        If x > 0 Then
            For i = 0 To x - 1
                charMAKnights(i, 0) = RandomName()
                charMAKnights(i, 1) = "alive"
                charMAKnights(i, 2) = ""
            Next
            s = s & x & " middle-aged, "
            x2 = x2 + x
        End If

        x = DiceRoller(1, 6)
        For i = 0 To x - 1
            charYoungKnights(i, 0) = RandomName()
            charMAKnights(i, 1) = "alive"
            charMAKnights(i, 2) = ""
        Next
        If x2 > 0 Then
            s = s & "and " & x & " young)"
        Else
            s = s & "all young)"
        End If

        x2 = x2 + x
        Console.WriteLine("Your personal army consists of:")
        s2 = s
        s = ""

        If x2 > 1 Then
            s = x2 & " family knights " & s2
        Else
            s = "1 young family knight"
        End If
        If Not tradWoman Then
            s = s & ", plus yourself."
        Else
            s = s & "."
        End If
        Console.WriteLine(s)
        Console.WriteLine(charLineageMen & " lineage men.")
        Console.WriteLine(charLevies & " levies.")

        Console.WriteLine()
        Console.WriteLine(charName & "'s family history will now be generated and appended to their character sheet.")

        familyHistory = FamilyHistoryMaker(here, charName, fatherName, motherName, grandfatherName, charYearBorn)
        'number of extra heirlooms
        'Optional list of passions and directed traits.
        'Glories.
        'Line starting '439: ...' is the start of the actual history.

        Dim sr As New IO.StringReader(familyHistory)
        b = False
        s = sr.ReadLine
        Try
            x = CInt(s)
            For i = 1 To x
                If LCase(charReligion) = "pagan" Then b = True
                s = HeirloomGenerator(hList, b, True)
                If InStr(s, "charger") Then charHorses.Add("Charger")
                If InStr(s, "courser") Then charHorses.Add("Courser")
                x2 = InStr(s, "rouncy")
                If x2 > 0 Then
                    charHorses.Add("Rouncy")
                    If InStr(x2 + 1, s, "rouncy") Then charHorses.Add("Rouncy")
                End If
                x = InStr(s, "//")
                If x > 0 Then
                    charHeirlooms.Add(Left(s, x - 1))
                    charHeirlooms.Add(Mid(s, x + 2))
                Else
                    charHeirlooms.Add(s)
                End If
            Next
        Catch ex As Exception
            x = 0
        End Try
        s = sr.ReadLine()
        Do While InStr(Left(s, 10), "Glory") < 1
            If Left(s, 2) = "PA" Then
                charPassions.Add(Mid(s, 4))
            ElseIf Left(s, 2) = "DT" Then
                charDirectedTraits.Add(Mid(s, 4))
            End If
            s = sr.ReadLine
        Loop
        Do While Left(s, 3) <> "439" And Left(s, 3) <> "440"
            x = InStr(s, "/")
            s2 = Mid(s, x + 1)
            s = Left(s, x - 1)
            x = CInt(s2)

            Select Case s
                Case "cGlory"
                    charGlory = x
                Case "fGlory"
                    fatherGlory = x
                Case "gfGlory"
                    grandfatherGlory = x
            End Select

            s = sr.ReadLine
        Loop
        familyHistory = s & vbNewLine & sr.ReadToEnd
        sr.Dispose()

        charOldKnights = OldKnightGenerator(charOldKnights, grandfatherName)
        charMAKnights = MAKnightGenerator(charMAKnights, fatherName, motherName)
        charYoungKnights = YoungKnightGenerator(charYoungKnights, charAge)

        'Generating your family tree is OBNOXIOUS.
        'First, we determine if your mother remarried.
        Dim motherStatus As String
        motherStatus = AliveAndMarried(False, True)
        If InStr(motherStatus, "Dead ") > 0 Then
            motherStatus = "Deceased."
        ElseIf InStr(motherStatus, " married") > 0 Then
            motherStatus = Replace(motherStatus, " married", " married to " & RandomName("male"))
        End If

        'Now, we find out how many brothers and sisters you have. (Including pre-existing ones.)
        x2 = 0
        For i = 0 To 5
            If charYoungKnights(i, 0) = "" Then Exit For
            If InStr(charYoungKnights(i, 2), "brother") > 0 Then x2 += 1
        Next

        x3 = 0
        x = DiceRoller(1, 6)
        For i = 1 To x
            x = DiceRoller(1, 2)
            If x = 1 Then
                x3 += 1
            Else
                sisters.Add(RandomName("female"))
            End If
        Next

        x3 = x3 - x2
        If x3 > 0 Then
            For i = 1 To x3
                brothers.Add(RandomName)
            Next
        End If

        x2 = 0
        For i = 0 To 5
            If charYoungKnights(i, 0) = "" Then Exit For
            If InStr(charYoungKnights(i, 2), "sister") > 0 Then x2 += 1
        Next

        x2 = x2 - sisters.Count
        If x2 > 0 Then
            For i = 1 To x2
                sisters.Add(RandomName("female"))
            Next
        End If

        For i = 0 To 5
            If charYoungKnights(i, 0) = "" Then Exit For
            If InStr(charYoungKnights(i, 2), "RANDOMSISTER") > 0 Then
                x = DiceRoller(1, sisters.Count)
                charYoungKnights(i, 2) = Replace(charYoungKnights(i, 2), "RANDOMSISTER", sisters(x - 1))
            End If
        Next

        x2 = 0
        x3 = 0
        'Now for your father's siblings.
        For i = 0 To 3
            If charMAKnights(i, 0) = "" Then Exit For
            If InStr(charMAKnights(i, 2), $"{fatherName}") > 0 Then x2 += 1
        Next

        x3 = 0
        x = DiceRoller(1, 6)
        For i = 1 To x
            x = DiceRoller(1, 2)
            If x = 1 Then
                x3 += 1
            Else
                pAunts.Add(RandomName("female"))
            End If
        Next

        x3 = x3 - x2
        If x3 > 0 Then
            For i = 1 To x3
                pUncles.Add(RandomName)
            Next
        End If

        x2 = 0
        x3 = 0
        'And your mother's.
        For i = 0 To 3
            If charMAKnights(i, 0) = "" Then Exit For
            If InStr(charMAKnights(i, 2), $"{motherName}") > 0 Then x2 += 1
        Next

        x3 = 0
        x = DiceRoller(1, 6)
        For i = 1 To x
            x = DiceRoller(1, 2)
            If x = 1 Then
                x3 += 1
            Else
                mAunts.Add(RandomName("female"))
            End If
        Next

        x3 = x3 - x2
        If x3 > 0 Then
            For i = 1 To x3
                mUncles.Add(RandomName)
            Next
        End If

        'And now your pre-existing cousins.
        Dim al1 As New ArrayList
        al1.AddRange(pUncles)
        al1.AddRange(pAunts)
        Dim al2 As New ArrayList
        al2.AddRange(mUncles)
        al2.AddRange(mAunts)

        For i = 0 To 3
            If charMAKnights(i, 0) = "" Then Exit For
            If InStr(charMAKnights(i, 2), $"{motherName}") > 0 Then al2.Add(charMAKnights(i, 0))
            If InStr(charMAKnights(i, 2), $"{fatherName}") > 0 Then al1.Add(charMAKnights(i, 0))
        Next

        x2 = 0
        For i = 0 To 5
            Dim cName As String
            Dim cNotes As String
            cName = charYoungKnights(i, 0)
            If cName = "" Then Exit For
            cNotes = charYoungKnights(i, 2)
            If InStr(cNotes, "cousin (paternal)") > 0 Then
                x = DiceRoller(1, al1.Count)
                charYoungKnights(i, 2) += $" Son of {al1(x - 1)}."
            End If
            If InStr(cNotes, "cousin (maternal)") > 0 Then
                x = DiceRoller(1, al2.Count)
                charYoungKnights(i, 2) += $" Son of {al2(x - 1)}."
            End If
        Next

        MiscLifeDetails(here, pUncles, True)
        MiscLifeDetails(here, pAunts, False)
        MiscLifeDetails(here, mUncles, True)
        MiscLifeDetails(here, mAunts, False)
        MiscLifeDetails(here, sisters, False)
        MiscLifeDetails(here, brothers, True)
        MiscLifeDetails(here, cousins, True)    'The only cousins this system generates are male, so.

        'Commence the export stuff.
        'Dim charSheet As New Xml.XmlDocument

        Dim charSheet As New Xml.XmlDocument
        Dim cNode As Xml.XmlNode
        Dim cNode2 As Xml.XmlNode

        cNode = charSheet.CreateElement("pdcc_character")
        charSheet.AppendChild(cNode)
        cNode2 = charSheet.CreateElement("character")
        cNode.AppendChild(cNode2)
        cNode2 = charSheet.CreateElement("history")
        cNode.AppendChild(cNode2)
        cNode2 = charSheet.CreateElement("family")
        cNode.AppendChild(cNode2)

        ExportCharacter(charSheet, charName, charGender, tradWoman, charAge, homeland, culture, charReligion, fatherClass,
                        charSonNumber, charLeige, charClass, charManor, charTraits, charDirectedTraits, charPassions,
                        charSIZ, charDEX, charSTR, charCON, charAPP, charFeatures, charSkills, charGlory,
                        charSquire, charHorses, charHeirlooms, charOldKnights, charMAKnights, charYoungKnights,
                        charLineageMen, charLevies)

        ExportHistory(charSheet, familyHistory)

        ExportFamily(charSheet, fatherName, grandfatherName, fatherGlory, grandfatherGlory, motherName, motherStatus,
                     charOldKnights, charMAKnights, charYoungKnights, pUncles, pAunts, mUncles, mAunts, brothers, sisters, cousins)

        Dim bleah As String
        If tradWoman Then
            bleah = "Lady"
        Else
            bleah = "Sir"
        End If

        s = here & $"\{bleah} " & charName & ".xml"
        charSheet.Save(s)
        ProcessXMLOutput(s, here)

        Console.WriteLine()
        Console.WriteLine("Family and history all generated!")
        Console.WriteLine("Press any key to exit.")
        Console.ReadKey()

    End Sub

    Sub MiscLifeDetails(here As String, ByRef inList As ArrayList, Optional male As Boolean = True)
        If inList.Count = 0 Then Exit Sub

        For i = 0 To inList.Count - 1
            inList(i) = $"{inList(i)}. " & AliveAndMarried()
            If InStr(inList(i), "Dead ") Then inList(i) = inList(i) & MiscDeath(here, male)
        Next

    End Sub

    Sub Heightening(ByRef skills As String(,), ByRef traits As String(,), ByRef passions As ArrayList,
                    ByRef siz As Integer, ByRef dex As Integer, ByRef str As Integer, ByRef con As Integer,
                    ByRef app As Integer)
        Dim s As String = ""
        Dim s2 As String = ""
        Dim x As Integer
        Dim x2 As Integer
        Dim skArray As New ArrayList
        Dim SkArray2 As New ArrayList
        Dim gpArray As New ArrayList
        For i = 0 To 39
            If skills(1, i) <= 10 Then
                skArray.Add(skills(0, i))
            ElseIf skills(1, i) <= 14 Then
                SkArray2.Add(skills(0, i))
            End If
        Next
        Dim statNameArray As New ArrayList
        statNameArray.AddRange({"SIZ", "DEX", "STR", "CON", "APP"})
        Dim statArray = New Integer() {siz, dex, str, con, app}

        Console.WriteLine()
        Console.WriteLine("Now you get to choose four unique values to 'heighten'.")
        Console.WriteLine("This is a +5 bonus for skills and a +1 bonus for everything else.")
        Console.WriteLine("Skills are capped at 15, traits at 19, passions at 20, and stats at 18.")
        Console.WriteLine("Except CON, which can go up to 21.")
        Console.WriteLine()
        Console.WriteLine("Just like with ability scores, there are no takebacks because I am lazy.")
        Console.WriteLine()

        For i = 1 To 4
            Console.WriteLine("Type the name of the category you want to increase:")
            Console.WriteLine("Stats, Traits, Skills, or Passions.")

            Do
                s = Console.ReadLine()
                s = LCase(s)
                If s <> "stats" And s <> "traits" And s <> "skills" And s <> "passions" Then
                    s = ""
                    Console.WriteLine("Please enter Stats, Traits, Skills, or Passions.")
                End If
            Loop While s = ""

            Select Case s
                Case "stats"
                    Console.WriteLine("Choose a stat to heighten:")
                    For j = 0 To 4
                        If j = 3 Then x = 21 Else x = 18
                        If statArray(j) < x Then
                            Console.Write($"{statNameArray(j)} ({statArray(j)})")
                            If j < 4 Then
                                For k = j + 1 To 4
                                    If statArray(k) < 18 Then
                                        Console.Write(", ")
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                    Console.WriteLine()
                    Do
                        s = ""
                        s = Console.ReadLine()
                        s = LCase(s)
                        If Not statNameArray.Contains(UCase(s)) Then
                            s = ""
                            Console.WriteLine("Please choose one of the five stats: SIZ, DEX, STR, CON, or APP.")
                        Else
                            Select Case s
                                Case "siz"
                                    If siz < 18 Then
                                        siz += 1
                                        statArray(0) += 1
                                    Else
                                        s = ""
                                        Console.WriteLine("That stat is too high to be enhanced further.")
                                    End If
                                Case "dex"
                                    If dex < 18 Then
                                        dex += 1
                                        statArray(1) += 1
                                    Else
                                        s = ""
                                        Console.WriteLine("That stat is too high to be enhanced further.")
                                    End If
                                Case "str"
                                    If str < 18 Then
                                        str += 1
                                        statArray(2) += 1
                                    Else
                                        s = ""
                                        Console.WriteLine("That stat is too high to be enhanced further.")
                                    End If
                                Case "con"
                                    If con < 21 Then
                                        con += 1
                                        statArray(3) += 1
                                    Else
                                        s = ""
                                        Console.WriteLine("That stat is too high to be enhanced further.")
                                    End If
                                Case "app"
                                    If app < 18 Then
                                        app += 1
                                        statArray(4) += 1
                                    Else
                                        s = ""
                                        Console.WriteLine("That stat is too high to be enhanced further.")
                                    End If
                            End Select
                        End If
                    Loop While s = ""
                Case "traits"
                    Console.WriteLine("Choose a trait to heighten:")
                    PrintTraitList(traits, 19)
                    Do While s <> ""
                        s = "x"
                        s = Console.ReadLine()
                        If s = "" Then s = "x"
                        s = StrConv(s, vbProperCase)
                        For j = 0 To 25
                            If traits(0, j) = s And traits(1, j) < 19 Then
                                traits = TraitUpdate(traits, s, traits(1, j) + 1)
                                s = ""
                                Exit For
                            End If
                        Next
                        If s = "x" Then Console.WriteLine("Please choose one of the listed traits.")
                    Loop
                Case "skills"
                    Console.WriteLine("Choose a skill to heighten.")
                    Console.WriteLine("This first list are all less than or equal to 10. You'll get the full +5 bonus by heightening these:")
                    PrintSkillList(skArray)
                    If SkArray2.Count > 0 Then
                        Console.WriteLine("This second list are all between 11 and 14, so heightening these will bump them to maximum 15.")
                        PrintSkillList(SkArray2)
                    End If
                    Do While s <> ""
                        s = "x"
                        s = Console.ReadLine()
                        If s = "" Then s = "x"
                        s = StrConv(s, vbProperCase)
                        If Not (skArray.Contains(s) Or SkArray2.Contains(s)) Then s = "x"
                        For j = 0 To 39
                            If skills(0, j) = s Then
                                SkillUpdater($"(+5 {s})", skills, 15)
                                skArray.Clear()
                                SkArray2.Clear()
                                For k = 0 To 39
                                    If skills(1, k) <= 10 Then
                                        skArray.Add(skills(0, k))
                                    ElseIf skills(1, k) <= 14 Then
                                        SkArray2.Add(skills(0, k))
                                    End If
                                Next
                                s = ""
                                Exit For
                            End If
                        Next
                        If s = "x" Then Console.WriteLine("Please choose one of the listed skills.")
                    Loop
                Case "passions"
                    Console.WriteLine("Choose a passion to heighten:")
                    s = "x"
                    x2 = 0
                    Do While s <> ""
                        For Each p In passions
                            Try
                                x = Right(p, 2)
                                s2 = Left(p, Len(p) - 3)
                                gpArray.Add(s2)
                            Catch
                                Try
                                    x = Right(p, 1)
                                    s2 = Left(p, Len(p) - 2)
                                    gpArray.Add(s2)
                                Catch
                                    x = 0
                                    s2 = p
                                End Try
                            End Try

                            If x < 20 Then
                                If passions.IndexOf(p) <> 0 Then Console.Write(", ")
                                If x2 >= 4 Then
                                    Console.WriteLine()
                                    x2 = 0
                                Else
                                    x2 += 1
                                End If
                                Console.Write($"{s2} ({x})")
                            End If
                        Next
                        Console.WriteLine()

                        s = "x"
                        s = Console.ReadLine()
                        If s = "" Then s = "x"
                        s = StrConv(s, vbProperCase)
                        If gpArray.Contains(s) Then
                            s = s2 & "/" & x + 1
                            passions.Add(s)
                            s = ""
                        ElseIf passions.Contains(s) Then
                            passions.Add(s)
                            s = ""
                        Else
                            s = "x"
                        End If
                        ConsolidatePassions(passions, True)

                        If s = "x" Then Console.WriteLine("Please choose one of the listed passions.")
                    Loop
            End Select
        Next

    End Sub

    Sub PrintSkillList(skArray As ArrayList)
        Dim c As Integer
        c = 0
        For i = 0 To skArray.Count - 1
            Console.Write(skArray(i))
            c += 1
            If c < 4 And i <> skArray.Count - 1 Then
                Console.Write(", ")
            Else
                c = 0
                Console.WriteLine()
            End If
        Next
    End Sub

    Sub PrintTraitList(traits As String(,), Optional traitMax As Integer = 21)
        Dim c As Integer
        c = 0
        For i = 0 To 25
            If traits(1, i) < traitMax And traits(1, i) > 20 - traitMax Then
                'Don't display traits that are too high for the max OR their counterparts.
                'I suppose it wouldn't HURT to display the counterparts but then you'd be spending
                'points to undo your previous point spend and that's unlikely to ever come up.
                Console.Write($"{traits(0, i)} ({traits(1, i)})")
                c += 1
                If c < 4 And i <> 25 Then
                    Console.Write(", ")
                Else
                    c = 0
                    Console.WriteLine()
                End If
            End If
        Next
    End Sub

    Function SkillUpdater(inString As String, sArray As String(,), Optional limited As Integer = -1) As String(,)
        Dim s As String
        Dim s2 As String
        Dim x As Integer
        Dim x2 As Integer

        x = InStr(inString, "(")
        s = Mid(inString, x + 1)
        s = Left(s, Len(s) - 1)  's is now the bonus and skill

        x2 = InStr(s, " ")
        s2 = Trim(Left(s, x2))
        s2 = Mid(s2, 2)
        x = CInt(s2) 'x is now the bonus
        s = Trim(Mid(s, x2))    's is now the skill

        For i = 0 To 39
            If sArray(0, i) = s Then
                x2 = CInt(sArray(1, i))
                x2 = x2 + x
                If limited >= 0 Then
                    If x2 > limited Then x2 = limited
                End If
                sArray(1, i) = CStr(x2)
                Exit For
            End If
        Next
        SkillUpdater = sArray
    End Function

    Function InitialiseCharSkills(f As String, Optional gender As String = "male") As String(,)
        Dim skArray As New ArrayList
        Dim skXML As New Xml.XmlDocument
        Dim skNode As Xml.XmlElement
        Dim x As Integer
        Dim xp As String
        Dim sName As String

        skXML.Load(f)
        skArray = SkillLister(f)
        Dim a As String(,) = New String(1, skArray.Count - 1) {}
        For i = 0 To skArray.Count - 1
            sName = skArray(i)
            a(0, i) = sName
            xp = $"//skill[@name='{sName}']/start_value"
            skNode = skXML.SelectSingleNode(xp)
            x = skNode.GetAttribute(gender)
            a(1, i) = x
        Next

        InitialiseCharSkills = a
    End Function

    Function InitialiseCharPassions() As String(,)
        Dim a As String(,) = New String(1, 5) {}
        a(0, 0) = "Loyalty (Lord)"
        a(0, 1) = "Love (Family)"
        a(0, 2) = "Hospitality"
        a(0, 3) = "Honour"
        a(0, 4) = "Hate (Saxons)"

        a(0, 1) = CStr(15)
        a(0, 2) = CStr(15)
        a(0, 3) = CStr(15)
        a(0, 4) = CStr(15)
        a(0, 5) = "xx"

        InitialiseCharPassions = a
    End Function

    Function TraitUpdate(a As String(,), trait As String, value As Integer) As String(,)
        For i = 0 To 25
            If a(0, i) = trait Then
                a(1, i) = CStr(value)
                If i < 13 Then
                    a(1, i + 13) = CStr(20 - value)
                Else
                    a(1, i - 13) = CStr(20 - value)
                End If
                Exit For
            End If
        Next
        TraitUpdate = a
    End Function

    Function InitialiseCharTraits() As String(,)
        Dim a As String(,) = New String(1, 25) {}

        For i = 0 To 25
            a(1, i) = CStr(10)
        Next i

        a(0, 0) = "Chaste"
        a(0, 1) = "Energetic"
        a(0, 2) = "Forgiving"
        a(0, 3) = "Generous"
        a(0, 4) = "Honest"
        a(0, 5) = "Just"
        a(0, 6) = "Merciful"
        a(0, 7) = "Modest"
        a(0, 8) = "Prudent"
        a(0, 9) = "Spiritual"
        a(0, 10) = "Temperate"
        a(0, 11) = "Trusting"
        a(0, 12) = "Valorous"

        a(0, 13) = "Lustful"
        a(0, 14) = "Lazy"
        a(0, 15) = "Vengeful"
        a(0, 16) = "Selfish"
        a(0, 17) = "Deceitful"
        a(0, 18) = "Arbitrary"
        a(0, 19) = "Cruel"
        a(0, 20) = "Proud"
        a(0, 21) = "Reckless"
        a(0, 22) = "Worldly"
        a(0, 23) = "Indulgent"
        a(0, 24) = "Suspicious"
        a(0, 25) = "Cowardly"

        InitialiseCharTraits = a
    End Function

    Function RandomHome() As String
        Dim pArray As String() = {"Baverstock", "Berwick St. James", "Broughton", "Burcombe", "Cholderton", "Dinton", "Durnford", "Idmiston", "Laverstock", "Newton", "Newton Tony", "Pitton", "Shrewton", "Stapleford", "Steeple Langford", "Tisbury", "Winterbourne Gunnet", "Winterbourne Stoke", "Woodford", "Wylye"}
        Dim n As Integer = DiceRoller(1, UBound(pArray))
        RandomHome = pArray(n)
    End Function

    Function RandomName(Optional gender As String = "male") As String
        Dim male As Boolean = True
        If gender = "female" Then male = False

        Dim mNames As String
        mNames = "Adtherp, Alein, Aliduke, Annecians, Archade, Arnold, Arrouse, Bandelaine, Bellangere, Bellias, Berel, Bersules, Bliant, Breunis, Briant, Caulas, Chestelaine, Clegis, Cleremond, Dalan, Dinaunt, Driant, Ebel, Edward, Elias, Eliot, Emerause, Flannedrius, Florence, Floridas, Galardoun, Garnish, Gerin, Gauter, Gherard, Gilbert, Gilmere, Goneries, Gracian, Gumret, Guy, Gwinas, Harsouse, Harvis, Hebes, Hemison, Herawd, Heringdale, Herlews, Hermel, Hermind, Hervis, Hewgon, Idres, Jordans, Lardans, Leomie, Manasan, Maurel, Melion, Miles, Morganor, Morians, Moris, Nanowne, Nerovens, Pedivere, Pellandres, Pellogres, Perin, Phelot, Pillounes, Plaine, Plenorias, Sauseise, Selises, Selivant, Semond."
        Dim fNames As String
        fNames = "Ade, Alis, Arnive, Astrigis, Bene, Blancheflor, Carsenefide, Calire, Cunneware, Diane, Elidis, Enide, Elizabeth, Esclarmonde, Feimurgan, Felelolie, Felinete, Feunete, Florie, Gloris, Heliap, Iblis, Idain, Imane, Jeschute, Laufamour, Liaze, Lore, Loorette, Laudine, Malvis, Maugalie, Melior, Morchades, Obie, Obilot, Oruale, Repanse, Sangive, Tanree, Tryamour, Violette."

        Dim nArray As String()
        If male Then
            nArray = mNames.Split(", ")
        Else
            nArray = fNames.Split(", ")
        End If

        Dim n As Integer
        n = DiceRoller(1, UBound(nArray))

        RandomName = Trim(nArray(n))
    End Function

    Public Function DiceRoller(Optional quantity As Integer = 1, Optional sides As Integer = 20) As Integer
        DiceRoller = 0

        Static Generator As System.Random = New System.Random()

        For i = 1 To quantity
            DiceRoller = DiceRoller + Generator.Next(1, sides + 1)
            Generator.Next()
        Next i
    End Function
End Module
