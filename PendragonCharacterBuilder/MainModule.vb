Module MainModule

    Sub Main()
        'The basic program will ask for a name and gender and assume that you're a Salisbury knight from the core book.
        'It won't assign your attributes or your various bonuses, just the standard stuff.
        'Then it'll output an XML file, then convert that into basic forum bbcode and plain text.
        'Everything generate-able from a random table will be generated from a table.

        Dim charName As String
        Dim charGender As String
        Dim charAge As Integer
        Dim charYearBorn As Integer
        Dim homeland As String = "Salisbury"
        Dim culture As String = "Cymric"
        Dim charReligion As String
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
        Dim charHeirloom As String
        Dim charFamilyCharacteristic As String

        Dim fatherName As String = RandomName()
        Dim grandfatherName As String = RandomName()

        While fatherName = grandfatherName
            grandfatherName = RandomName()
        End While

        Dim fatherClass As String = "vassal knight"

        Dim xhim As String
        Dim xhis As String
        Dim xhe As String

        Dim x As Integer
        Dim x2 As Integer
        Dim spr As Integer
        Dim aspr As Integer
        Dim attMin As Integer
        Dim attMax As Integer
        Dim s As String

        Console.WriteLine("Welcome to the Pendragon character generator!")
        Console.WriteLine("This program will basically hammer through all the random tables in the")
        Console.WriteLine("Pendragon core book For you, but it won't make any actual *decisions*.")
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
                charSkills = InitialiseCharSkills()
            Else
                charSkills = InitialiseCharSkills("female")
            End If
        Else
            charSkills = InitialiseCharSkills()
        End If

#Disable Warning BC42104 ' Variable is used before it has been assigned a value
        Select Case LCase(Left(charReligion, 1))
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
            Case "b"
                charSkills(0, 20) = Replace(charSkills(0, 20), "xx", "British Christianity")
            Case "r"
                charSkills(0, 20) = Replace(charSkills(0, 20), "xx", "Roman Christianity")
            Case "p"
                charSkills(0, 20) = Replace(charSkills(0, 20), "xx", "Pagan")
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

        If x = 1 Then
            Console.WriteLine("Thanks to " & xhis & " APP of " & charAPP & ", " & charName & " has " & x & " distinctive feature:")
        Else
            Console.WriteLine("Thanks to " & xhis & " APP of " & charAPP & ", " & charName & " has " & x & " distinctive features:")
        End If
        charFeatures = "Something about their "

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
                charFeatures = charFeatures & " and their "
            Else
                charFeatures = charFeatures & ", their "
            End If
        Next i
        Console.WriteLine(charFeatures)
        Console.WriteLine()

        Console.WriteLine("Choose a knightly skill to be awesome at:")
        Console.WriteLine("(Awareness, Courtesy, First Aid, Hunting; Battle, Dagger, Horsemanship, Lance, Spear, Sword)")

        Dim kSkills As String() = "Awareness,Courtesy,First Aid,Hunting,Battle,Dagger,Horsemanship,Lance,Spear,Sword".Split(",")
        s = Nothing
        Do
            s = Console.ReadLine()
            s = StrConv(s, VbStrConv.ProperCase)
            If Not kSkills.Contains(s) Then
                s = Nothing
            Else
                For i = 0 To 31
                    If charSkills(0, i) = s Then charSkills(1, i) = 15
                Next i
            End If
            If s = Nothing Then Console.WriteLine("Please choose one of the knightly skills above.")
        Loop While s = Nothing

        Dim skArray As New ArrayList()
        For i = 0 To 25
            s = charSkills(0, i)
            If s = "Chirurgery" Or s = "Fashion" Or s = "Industry" Then
                'Skip it
            ElseIf CInt(charSkills(1, i)) < 10 Then
                skArray.Add(charSkills(0, i))
            End If
        Next

        Console.WriteLine()
        Console.WriteLine("And now choose three non-combat skills to be good at:")
        For j = 1 To 3
            Console.WriteLine("Skill " & j & ":")
            For i = 0 To skArray.Count - 1
                Console.Write(skArray(i))
                If i = skArray.Count - 1 Then
                    Console.Write(".")
                Else
                    Console.Write(", ")
                End If
            Next
            Console.WriteLine()
            s = Nothing
            Do
                s = Console.ReadLine
                s = StrConv(s, VbStrConv.ProperCase)
                If s = "Read" Then
                    s = charSkills(0, 18)
                ElseIf s = "Play" Then
                    s = charSkills(0, 17)
                ElseIf s = "Religion" Then
                    s = charSkills(0, 20)
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
        Console.WriteLine("Your squire's name is " & charSquire(0) & ". Choose a skill for him to be vaguely okay at:")
        skArray.Clear()
        For i = 0 To 31
            s = charSkills(0, i)
            If s = "Chirurgery" Or s = "Fashion" Or s = "Industry" Or s = "First Aid" Or s = "Battle" Or s = "Horsemanship" Then
                'Skip it
            Else
                skArray.Add(charSkills(0, i))
            End If
        Next
        For i = 0 To skArray.Count - 1
            Console.Write(skArray(i))
            If i = skArray.Count - 1 Then
                Console.Write(".")
            Else
                Console.Write(", ")
            End If
        Next
        Console.WriteLine()
        s = Nothing
        Do
            s = Console.ReadLine
            s = StrConv(s, VbStrConv.ProperCase)
            If s = "Read" Then
                s = charSkills(0, 18)
            ElseIf s = "Play" Then
                s = charSkills(0, 17)
            ElseIf s = "Religion" Then
                s = charSkills(0, 20)
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
        s = HeirloomGenerator()
        charHeirloom = s

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
        Console.Write(s & ".")

        Console.WriteLine()
        Console.WriteLine("Finally, you get a heritable family characteristic!")
        Console.Write("The ")
        If charGender = "male" Then
            Console.Write("men")
        Else
            Console.Write("women")
        End If
        Console.Write(" of your line are all ")

    End Sub

    Function HeirloomGenerator(Optional pagan As Boolean = False, Optional recurse As Boolean = False) As String
        Dim s As String = ""
        Dim x As Integer

        x = DiceRoller(1, 20)
        Do While (pagan And x = 8) Or (recurse And x = 20)
            x = DiceRoller(1, 20)
        Loop
        Select Case x
            Case 1
                s = DiceRoller(3, 20) & "d in cash"
            Case 2
                s = (DiceRoller(3, 20) + 100) & "d in cash"
            Case 3
                s = (DiceRoller(3, 20) + 100) & "d in cash"
            Case 4
                s = "£1 in cash"
            Case 5
                s = "£1 in cash"
            Case 6
                s = "£1 in cash"
            Case 7
                s = "£" & DiceRoller(1, 6) & " in cash"
            Case 8
                s = "a sacred Christian relic -- "
                Select Case DiceRoller(1, 6)
                    Case 1
                        s = s & "a finger"
                    Case 2
                        s = s & "some tears"
                    Case 3
                        s = s & "some hair"
                    Case 4
                        s = s & "some hair"
                    Case 5
                        s = s & "a bone fragment"
                    Case 6
                        s = s & "some blood"
                End Select
            Case 9
                s = "an ancient bronze sword worth £2 (+1 to Sword skill, breaks in combat as if it wasn't a sword)"
        Case 10
                s = "a blessed lance worth 25d (+1 to Lance skill)"
            Case 11
                s = "a decorated saddle worth £1"
            Case 12
                s = "an engraved ring worth "
                Select Case DiceRoller(1, 3)
                    Case 1
                        s = s & "120d"
                    Case 2
                        s = s & "120d"
                    Case 3
                        s = s & "£2"
                End Select
            Case 13
                Select Case DiceRoller(1, 6)
                    Case 1
                        s = "a gold armband worth £8"
                    Case Else
                        s = "a silver armband worth £1"
                End Select
            Case 14
                s = "a valuable cloak from "
                Select Case DiceRoller(1, 6)
                    Case 1
                        s = s & "Byzantium"
                    Case 2
                        s = s & "Byzantium"
                    Case 3
                        s = s & "Germany"
                    Case 4
                        s = s & "Spain"
                    Case 5
                        s = s & "Spain"
                    Case 6
                        s = s & "Rome"
                End Select
                s = s & " worth £1"
            Case 15
                s = "a magic healing potion (!) which cures 1d6 damage, once"
            Case 16
                s = "an extra rouncy"
            Case 17
                s = "an extra rouncy"
            Case 18
                s = "a second charger"
            Case 19
                s = "a courser"
            Case 20
                s = HeirloomGenerator(pagan, True) & "; AND " & HeirloomGenerator(pagan, True)
        End Select

        HeirloomGenerator = s
    End Function

    Function InitialiseCharSkills(Optional gender As String = "male") As String(,)
        Dim a As String(,) = New String(1, 31) {}
        a(0, 0) = "Awareness"
        a(0, 1) = "Boating"
        a(0, 2) = "Chirurgery"
        a(0, 3) = "Compose"
        a(0, 4) = "Courtesy"
        a(0, 5) = "Dancing"
        a(0, 6) = "Faerie Lore"
        a(0, 7) = "Falconry"
        a(0, 8) = "Fashion"
        a(0, 9) = "First Aid"
        a(0, 10) = "Flirting"
        a(0, 11) = "Folk Lore"
        a(0, 12) = "Gaming"
        a(0, 13) = "Heraldry"
        a(0, 14) = "Industry"
        a(0, 15) = "Intrigue"
        a(0, 16) = "Orate"
        a(0, 17) = "Play (Harp)"
        a(0, 18) = "Read (Latin)"
        a(0, 19) = "Recognise"
        a(0, 20) = "Religion (xx)"
        a(0, 21) = "Romance"
        a(0, 22) = "Singing"
        a(0, 23) = "Stewardship"
        a(0, 24) = "Swimming"
        a(0, 25) = "Tourney"
        a(0, 26) = "Battle"
        a(0, 27) = "Horsemanship"
        a(0, 28) = "Sword"
        a(0, 29) = "Lance"
        a(0, 30) = "Spear"
        a(0, 31) = "Dagger"

        If gender = "male" Then
            a(1, 0) = "5"
            a(1, 1) = "1"
            a(1, 2) = "0"
            a(1, 3) = "1"
            a(1, 4) = "3"
            a(1, 5) = "2"
            a(1, 6) = "1"
            a(1, 7) = "3"
            a(1, 8) = "0"
            a(1, 9) = "10"
            a(1, 10) = "3"
            a(1, 11) = "2"
            a(1, 12) = "3"
            a(1, 13) = "3"
            a(1, 14) = "0"
            a(1, 15) = "3"
            a(1, 16) = "3"
            a(1, 17) = "3"
            a(1, 18) = "0"
            a(1, 19) = "3"
            a(1, 20) = "2"
            a(1, 21) = "2"
            a(1, 22) = "2"
            a(1, 23) = "2"
            a(1, 24) = "2"
            a(1, 25) = "2"
            a(1, 26) = "10"
            a(1, 27) = "10"
            a(1, 28) = "10"
            a(1, 29) = "10"
            a(1, 30) = "6"
            a(1, 31) = "5"
        Else
            a(1, 0) = "2"
            a(1, 1) = "0"
            a(1, 2) = "10"
            a(1, 3) = "1"
            a(1, 4) = "5"
            a(1, 5) = "2"
            a(1, 6) = "3"
            a(1, 7) = "2"
            a(1, 8) = "2"
            a(1, 9) = "10"
            a(1, 10) = "5"
            a(1, 11) = "2"
            a(1, 12) = "3"
            a(1, 13) = "1"
            a(1, 14) = "5"
            a(1, 15) = "2"
            a(1, 16) = "2"
            a(1, 17) = "3"
            a(1, 18) = "1"
            a(1, 19) = "2"
            a(1, 20) = "2"
            a(1, 21) = "2"
            a(1, 22) = "3"
            a(1, 23) = "5"
            a(1, 24) = "1"
            a(1, 25) = "1"
            a(1, 26) = "1"
            a(1, 27) = "3"
            a(1, 28) = "0"
            a(1, 29) = "0"
            a(1, 30) = "0"
            a(1, 31) = "5"
        End If

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

    Function DiceRoller(Optional quantity As Integer = 1, Optional sides As Integer = 20) As Integer
        DiceRoller = 0

        Static Generator As System.Random = New System.Random()

        For i = 1 To quantity
            DiceRoller = DiceRoller + Generator.Next(1, sides)
        Next i
    End Function
End Module
