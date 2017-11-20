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
        Dim charPassions As String(,)
        charPassions = InitialiseCharPassions()
        Dim charSIZ As Integer
        Dim charDEX As Integer
        Dim charSTR As Integer
        Dim charCON As Integer
        Dim charAPP As Integer
        Dim charDamage As String
        Dim charHealing As Integer
        Dim charMove As Integer
        Dim charHP As Integer
        Dim charUnconscious As Integer
        Dim charFeatures As String
        Dim charSkills As String(,)
        Dim charGlory As Integer
        Dim charHorses As String() = {"Charger #1", "Rouncy #1", "Rouncy #2", "Sumpter #1", "", ""}
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

        Dim fatherClass As String = "vassal knight"

        Dim xhim As String
        Dim xhis As String
        Dim xhe As String
        Dim skArray As New ArrayList()

        Dim x As Integer
        Dim x2 As Integer
        Dim spr As Integer
        Dim aspr As Integer
        Dim attMin As Integer
        Dim attMax As Integer
        Dim s As String
        Dim s2 As String

        Dim here As String = My.Computer.FileSystem.CurrentDirectory
        'DEBUG
        'here = "C:\Users\LonghurstC\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder"

        Dim sList As String = here & "\pdcc_skill_list.xml"
        Dim hList As String = here & "\pdcc_heirlooms.xml"

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

        charDamage = CStr(Math.Round((charSTR + charSIZ) / 6, 0)) & "d6"
        charHealing = Math.Round((charCON + charSTR) / 10)
        charMove = Math.Round((charSTR + charDEX) / 10)
        charHP = charCON + charSIZ
        charUnconscious = Math.Round(charHP / 4)

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
                For i = 0 To 31
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

        For i = 0 To 39
            'Skills which are at 10+ don't get improved at this step.
            s = charSkills(0, i)
            x = charSkills(1, i)
            If x >= 10 Then skArray.Remove(s)
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

        Console.WriteLine()
        Console.WriteLine("You get some more options for customising your character, ")
        Console.WriteLine("but those are way too fiddly for me to bother with here")
        Console.WriteLine("They'll be summarised on the character sheet output.")
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
        skArray.Clear()
        If InStr(s, "charger") Then skArray.Add("Charger #2")
        If InStr(s, "courser") Then skArray.Add("Courser #1")
        x2 = InStr(s, "rouncy")
        If x2 > 0 Then
            skArray.Add("Rouncy #3")
            If InStr(x2 + 1, s, "rouncy") Then skArray.Add("Rouncy #4")
        End If

        x2 = charHorses.Count
        If skArray.Count > 0 Then
            For i = 0 To skArray.Count - 1
                charHorses(4 + i) = skArray(i)
            Next
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
        s = SpecialGiftGenerator(charGender)
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
            charFeatures = charFeatures & (fArray(x2))
            fArray.RemoveAt(x2)
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
            For i = 1 To x
                charMAKnights(0, 0) = RandomName()
                charMAKnights(0, 1) = "alive"
                charMAKnights(0, 2) = ""
            Next
            s = s & x & " middle-aged, "
            x2 = x2 + x
        End If

        x = DiceRoller(1, 6)
        For i = 1 To x
            charYoungKnights(0, 0) = RandomName()
            charMAKnights(0, 1) = "alive"
            charMAKnights(0, 2) = ""
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

        s = Nothing
        Do
            s = Console.ReadLine()
        Loop While s Is Nothing

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

        For i = 0 To 31
            If sArray(0, i) = s Then
                x2 = CInt(sArray(1, i))
                x2 = x2 + x
                If limited >= 0 Then
                    If x2 > limited Then x = limited
                End If
                sArray(1, i) = CStr(x2)
                Exit For
            End If
        Next
        SkillUpdater = sArray
    End Function

    Function SpecialGiftGenerator(gender As String) As String
        Dim s As String = ""
        Dim x As Integer

        x = DiceRoller(1, 20)
        If gender = "male" Then
            Select Case x
                Case 1
                    s = "good with horses (+5 Horsemanship)"
                Case 2
                    s = "good with horses (+5 Horsemanship)"
                Case 3
                    s = "excellent singers (+10 Singing)"
                Case 4
                    s = "possessed of keen senses (+5 Awareness)"
                Case 5
                    s = "possessed of keen senses (+5 Awareness)"
                Case 6
                    s = "possessed of keen senses (+5 Awareness)"
                Case 7
                    s = "possessed of keen senses (+5 Awareness)"
                Case 8
                    s = "gifted at naturecraft (+5 Hunting)"
                Case 9
                    s = "light-footed (+10 Dancing)"
                Case 10
                    s = "natural healers (+5 First Aid)"
                Case 11
                    s = "naturally lovable (+10 Flirting)"
                Case 12
                    s = "good with faces (+10 Recognise)"
                Case 13
                    s = "remarkably deductive (+5 Intrigue)"
                Case 14
                    s = "like otters (+10 Swimming)"
                Case 15
                    s = "natural speakers (+10 Orate)"
                Case 16
                    s = "natural musicians (+15 Play (Harp))"
                Case 17
                    s = "good with words (+15 Compose)"
                Case 18
                    s = "handy with heraldry (+10 Heraldry)"
                Case 19
                    s = "good with birds (+15 Falconry)"
                Case 20
                    s = "clever gamblers (+10 Gaming)"
            End Select
        Else
            Select Case x
                Case 1
                    s = "beautiful (+10 APP)"
                Case 2
                    s = "beautiful (+10 APP)"
                Case 3
                    s = "beautiful (+10 APP)"
                Case 4
                    s = "beautiful (+10 APP)"
                Case 5
                    s = "beautiful (+10 APP)"
                Case 6
                    s = "natural healers (+5 First Aid and +5 Chirurgery)"
                Case 7
                    s = "natural healers (+5 First Aid and +5 Chirurgery)"
                Case 8
                    s = "natural healers (+5 First Aid and +5 Chirurgery)"
                Case 9
                    s = "natural healers (+5 First Aid and +5 Chirurgery)"
                Case 10
                    s = "natural healers (+5 First Aid and +5 Chirurgery)"
                Case 11
                    s = "good with animals (+5 Falconry and +5 Horsemanship)"
                Case 12
                    s = "good with animals (+5 Falconry and +5 Horsemanship)"
                Case 13
                    s = "good with animals (+5 Falconry and +5 Horsemanship)"
                Case 14
                    s = "good with animals (+5 Falconry and +5 Horsemanship)"
                Case 15
                    s = "good with animals (+5 Falconry and +5 Horsemanship)"
                Case 16
                    s = "excellent speakers (+5 Orate and +5 Singing)"
                Case 17
                    s = "excellent speakers (+5 Orate and +5 Singing)"
                Case 18
                    s = "nimble-fingered (+10 Industry)"
                Case 19
                    s = "diligent caretakers (+10 Stewardship)"
                Case 20
                    s = "diligent caretakers (+10 Stewardship)"
            End Select
        End If

        SpecialGiftGenerator = s
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
            a(0, 1) = x
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
            DiceRoller = DiceRoller + Generator.Next(1, sides)
        Next i
    End Function
End Module
