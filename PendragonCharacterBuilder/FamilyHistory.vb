Module FamilyHistory
    Dim charName As String
    Dim charGlory As Integer
    Dim fatherName As String
    Dim motherName As String = RandomName("female")
    Dim grandfatherName As String
    Dim fatherGlory As Integer = 0
    Dim grandfatherGlory As Integer = (((DiceRoller(1, 20) * 100) + 1000) / 10) + DiceRoller(2, 20)
    Dim fatherAlive As Boolean = True
    Dim grandfatherAlive As Boolean = True
    Dim passionArray As New ArrayList
    Dim charBirthYear As Integer = 0
    Dim familyDissident As Boolean = False
    Dim familyRebel As Boolean = False

    Public Function FamilyHistoryMaker(here As String, charName_x As String, fatherName_x As String,
                                       motherName_x As String, grandfatherName_x As String, cby As Integer) As String
        'Returns a string with newline separators which fills in the whole family history.
        'The first line is all code, though, so we can also get stuff like starting Glory.
        Dim s As String = ""
        Dim fhm As String = ""
        Dim x As Integer
        Dim extraHeirlooms As Integer = 0

        charName = charName_x
        fatherName = fatherName_x
        grandfatherName = grandfatherName_x
        motherName = motherName_x
        charBirthYear = cby

        For i = 439 To 463
            'Look for the year file. If it exists, run it. If it doesn't exist, skip.
            'Check for birth year and add details as appropriate.
            If IO.File.Exists(here & "\xml\years\" & CStr(i) & ".xml") And grandfatherAlive Then
                fhm = fhm & YearGen(here, i, grandfatherGlory, grandfatherAlive)
            End If

            'DEBUG
            'grandfatherAlive = True

            'Here commences the big list of special cases.
            If grandfatherAlive Then
                If i = 447 Then
                    familyDissident = Dissidence(i, familyDissident)
                    If familyDissident Then
                        If Not InStr(fhm, "447") > 0 Then
                            fhm = fhm & "447-449: "
                        End If
                        fhm = fhm & "The family's grumbling about Vortigern's decisions got them labelled as dissidents." & vbNewLine & vbNewLine
                    End If
                End If

                If i = 450 Then
                    If familyDissident Then
                        If Dissidence(i, familyDissident) Then
                            If Not InStr(fhm, "450") > 0 Then
                                fhm = fhm & "450: "
                            End If
                            fhm = fhm & "The family became some of the lead instigators in the rebellion against Vortigern's rule." & vbNewLine & vbNewLine
                            familyRebel = True
                        End If
                    Else
                        If Not InStr(fhm, "450") > 0 Then
                            fhm = fhm & "450: "
                        End If
                        If Dissidence(i, familyDissident) Then
                            fhm = fhm & "The family's grumbling about Vortigern's decisions got them labelled as dissidents." & vbNewLine & vbNewLine
                            familyDissident = True
                        Else
                            fhm = fhm & "The family's staunch loyalty to Vortigern in this trying year did not go unnoticed!" & vbNewLine & vbNewLine
                            familyDissident = False
                        End If
                    End If
                End If
            End If
            If i = 457 And InStr(fhm, "showered him with gifts") Then extraHeirlooms = 2
            If i = 460 Then fatherGlory = (Math.Round(grandfatherGlory / 10)) + 1000
            If i = charBirthYear Then
                If Not InStr(fhm, CStr(i)) Then
                    fhm = fhm & CStr(i) & ": "
                End If
                fhm = fhm & $"In this year {fatherName} married {motherName}, a widow, and produced an heir: {charName}." & vbNewLine & vbNewLine
                x = DiceRoller(1, 20)
                If x <= 4 Then
                    fatherGlory += 25
                ElseIf x <= 7 Then
                    fatherGlory += 50
                ElseIf x <= 17 Then
                    fatherGlory += 100
                ElseIf x <= 19 Then
                    fatherGlory += 200
                Else
                    fatherGlory += 350
                End If
            End If
            If i = 463 Then PassionInheritor()
        Next

        For i = 464 To 484
            If IO.File.Exists(here & "\xml\years\" & CStr(i) & ".xml") And fatherAlive Then
                fhm = fhm & YearGen(here, i, fatherGlory, fatherAlive)
            End If

            If i = 477 Then i = i
        Next

        'Then the last battle of your father.
        If fatherAlive Then
            x = DiceRoller(1, 20)
            fhm = fhm & $"Almost immediately after Eburacum, {fatherName} travelled north with Uther and crushed the Saxons in a brutal surprise attack. "
            If x = 1 Then
                fhm = fhm & $"{fatherName} died there, but it was glorious."
                fatherGlory += 1000
            Else
                fhm = fhm & $"Unfortunately, {fatherName} died there."
            End If
            fhm = fhm & vbNewLine & vbNewLine
            fatherGlory += (90 * DiceRoller(1, 6))
        End If

        'Glory calculation.
        charGlory = Math.Round(fatherGlory / 10)

        'PassionConsolidation and adoption.
        PassionInheritor(True)

        fhm = Replace(fhm, "{fatherName}", fatherName)
        fhm = Replace(fhm, "{grandfatherName}", grandfatherName)

        fhm = $"cGlory/{charGlory}" & vbNewLine &
            $"fGlory/{fatherGlory}" & vbNewLine &
            $"gfGlory/{grandfatherGlory}" & vbNewLine & fhm

        If passionArray.Count > 0 Then
            For Each s In passionArray
                fhm = s & vbNewLine & fhm
            Next
        End If

        fhm = CStr(extraHeirlooms) & vbNewLine & fhm

        FamilyHistoryMaker = fhm
    End Function

    Sub PassionInheritor(Optional yourFather As Boolean = False)
        Dim s As String
        Dim s2 As String
        Dim x As Integer
        Dim value As Integer

        ConsolidatePassions()

        Console.WriteLine()
        If yourFather Then
            Console.WriteLine("You can inherit some, all, or none of the family traits.")
        Else
            Console.WriteLine("Your father can inherit some, all, or none of the family traits.")
        End If

        Console.WriteLine("One by one, please choose whether or not to keep [y] or drop [n] the trait.")

        Dim stopCount As Integer = passionArray.Count - 1
        For i = 0 To stopCount
            s = Mid(passionArray(i), 4)
            x = InStr(s, "/")
            value = CInt(Mid(s, x + 1))
            s = Left(s, x - 1)

            Console.WriteLine($"Keep ""{s}: {value}""?")
            s2 = ""
            Do
                s2 = Console.ReadLine
                s2 = Left(s2, 1)
                s2 = LCase(s2)
                If s2 = "k" Then s2 = "y"
                If s2 = "d" Then s2 = "n"
                If s2 <> "y" And s2 <> "n" Then
                    Console.WriteLine("Please choose y or n.")
                    s2 = ""
                End If
            Loop While s2 = ""
            If s2 = "n" Then
                passionArray.RemoveAt(i)
                i = i - 1
                stopCount = stopCount - 1
            End If
            If i >= stopCount Then Exit For
        Next

    End Sub

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

    Function Dissidence(y As Integer, famDiss As Boolean) As Boolean
        Dim s As String

        If y = 447 Then
            Console.WriteLine()
            Console.WriteLine("In 447 King Vortigern's habit of importing Saxons to fight his (admittedly successful)")
            Console.WriteLine("wars against the Picts aggravated many British families. Is yours among them? [y or n]")
        ElseIf y = 450 Then
            Console.WriteLine()
            Console.WriteLine("In 450 more and more families felt that King Vortigern favoured the Saxons over his")
            Console.Write("own people. ")
            If famDiss Then
                Console.Write("Did your family step up and lead the rebellion? [y or n]")
            Else
                Console.Write("Was your family among them? [y or n]")
            End If
            Console.WriteLine()
        End If
        s = ""
        Do
            s = Console.ReadLine()
            s = Left(s, 1)
            s = LCase(s)
            If s <> "y" And s <> "n" Then
                Console.WriteLine("Please enter y or n.")
                s = ""
            End If
        Loop While s = ""

        If s = "y" Then
            Dissidence = True
        Else
            Dissidence = False
        End If
    End Function

    Function YearGen(here As String, y As Integer, ByRef glory As Integer, ByRef lifedeath As Boolean) As String
        Dim yXML As New Xml.XmlDocument
        Dim yElem As Xml.XmlElement
        Dim yElem2 As Xml.XmlElement
        Dim yNode As Xml.XmlNode

        Dim yg As String = ""
        Dim s As String = ""
        Dim s2 As String = ""
        Dim s3 As String
        Dim x As Integer = DiceRoller(1, 20)
        Dim passionTracker As String = ""
        Dim g1 As Integer, g2 As Integer, g3 As Double
        Dim events As String = "events"

        yXML.Load(here & "\xml\years\" & CStr(y) & ".xml")
        yElem = yXML.SelectSingleNode("/year")
        s = yElem.GetAttribute("n")
        yg = $"{s}:"

        'Check for the branching years and apply branch modifiers.
        If yElem.GetAttribute("branch") = "true" Then
            'The branch options vary based on year.
            If y = 451 Then
                If familyDissident Then
                    events = "events[@branch='b']"
                Else
                    events = "events[@branch='a']"
                End If
            End If
            If y = 457 Then
                If familyRebel Then
                    events = "events[@branch='c']"
                ElseIf familyDissident Then
                    events = "events[@branch='a']"
                Else
                    events = "events[@branch='b']"
                End If
            End If
            If y = 460 Then
                If grandfatherAlive Then
                    events = "events[@branch='a']"
                Else
                    events = "events[@branch='b']"
                End If
            End If
        End If

        'If there's an events/text then that always displays.
        Try
            yNode = yXML.SelectSingleNode($"//{events}/text/node()")
            s = yNode.Value
        Catch ex As Exception
            s = ""
        End Try
        yg = $"{yg} {s}"

        'On to the year's event.
        yElem = yXML.SelectSingleNode($"//{events}/event[@dice_min <= {x} and @dice_max >= {x}]/text")
        yNode = yElem.SelectSingleNode("./text()")
        Try
            s = yNode.Value
        Catch ex As Exception
            'Do nothing
        End Try

        If yElem.GetAttribute("glory") <> "" Then
            glory += CInt(yElem.GetAttribute("glory"))
        End If
        If yElem.GetAttribute("dies") = "true" Then
            lifedeath = False
        End If
        If yElem.GetAttribute("passion") <> "" Then
            passionTracker = yElem.GetAttribute("passion")
            'I'm pretty certain you can't get more than one passion per year.
        End If
        If yElem.GetAttribute("battle") <> "" Then
            s2 = yElem.GetAttribute("battle")
            yElem2 = yXML.SelectSingleNode($"//battle[@name='{s2}']")

            'A battle!
            'Generate the glory.
            s3 = yElem2.GetAttribute("glory-die")
            g2 = yElem2.GetAttribute("glory-flat")
            g3 = yElem2.GetAttribute("glory-multiplier")
            g1 = DiceRoller(CInt(Left(s2, 1)), CInt(Mid(s2, 3)))
            glory += Math.Round(g1 * g2 * g3)

            'Then the battle event.
            x = DiceRoller(1, 20)
            yElem2 = yXML.SelectSingleNode($"//battle[@name='{s2}']/b-event[@dice_min <= {x} and @dice_max >= {x}]/text")
            yNode = yElem2.SelectSingleNode("./text()")
            Try
                s = s & yNode.Value
            Catch ex As Exception
                'Do nothing.
            End Try
            If yElem2.GetAttribute("glory") <> "" Then
                glory += CInt(yElem2.GetAttribute("glory"))
            End If
            If yElem2.GetAttribute("dies") = "true" Then
                lifedeath = False
            End If
            If yElem2.GetAttribute("passion") <> "" Then
                passionTracker = yElem2.GetAttribute("passion")
            End If
        End If
        'And now we check for passions.
        If passionTracker <> "" Then
            yElem = yXML.SelectSingleNode($"//passion[@id='{passionTracker}']")
            x = DiceRoller(1, 20)
            yElem2 = yXML.SelectSingleNode($"//passion[@id='{passionTracker}']/p-event[@dice_min <= {x} and @dice_max >= {x}]/text")
            Try
                yNode = yElem2.SelectSingleNode("./text()")
                s2 = yElem2.GetAttribute("value")
                g1 = DiceRoller(CInt(Left(s2, 1)), CInt(Mid(s2, 3)))
                If yElem2.GetAttribute("modifier") <> "" Then
                    g2 = CInt(yElem2.GetAttribute("modifier"))
                End If
                g3 = g1 + g2
                s2 = yNode.Value & CStr(g3)
                passionArray.Add(s2)
            Catch ex As Exception
                'Do nothing
            End Try
        End If

        yg = $"{yg} {s}"
        If InStr(yg, "MISCDEATH") > 0 Then yg = Replace(yg, "MISCDEATH", MiscDeath(here))
        If InStr(yg, "NOTHINGYEAR") > 0 Then
            yg = ""
        Else
            yg = yg & vbNewLine & vbNewLine
        End If
        YearGen = yg
    End Function
End Module
