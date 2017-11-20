Module FamilyHistory
    Dim charName As String
    Dim fatherName As String
    Dim grandfatherName As String
    Dim fatherGlory As Integer = 0
    Dim grandfatherGlory As Integer = (((DiceRoller(1, 20) * 100) + 1000) / 10) + DiceRoller(2, 20)
    Dim fatherAlive As Boolean = True
    Dim grandfatherAlive As Boolean = True
    Dim passionArray As New ArrayList
    Dim extraHeirlooms As Integer = 0
    Dim charBirthYear As Integer = 0

    Sub Main()
        'FamilyHistoryMaker("C:\Users\LonghurstC\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder",
        '                  "Steve", "Bob", "Grand-Bob", 464)

        FamilyHistoryMaker("C:\Users\Chris\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder",
                           "Steve", "Bob", "Grand-Bob", 464)
    End Sub

    Public Function FamilyHistoryMaker(here As String, charName_x As String, fatherName_x As String,
                                       grandfatherName_x As String, cby As Integer) As String
        'Returns a string with newline separators which fills in the whole family history.
        'The first line is all code, though, so we can also get stuff like starting Glory.
        Dim s As String = ""
        Dim fhm As String = ""
        Dim x As Integer

        charName = charName_x
        fatherName = fatherName_x
        grandfatherName = grandfatherName_x
        charBirthYear = cby

        For i = 439 To 463
            'Look for the year file. If it exists, run it. If it doesn't exist, skip.
            'Check for birth year and add details as appropriate.
            If IO.File.Exists(here & "\" & CStr(i) & ".xml") And grandfatherAlive Then
                fhm = fhm & YearGen(here, i, grandfatherGlory, grandfatherAlive)
            End If
        Next

        For i = 464 To 484

        Next

        fhm = Replace(fhm, "{fatherName}", fatherName)
        fhm = Replace(fhm, "{grandfatherName}", grandfatherName)

        FamilyHistoryMaker = fhm
    End Function

    Function YearGen(here As String, y As Integer, ByRef glory As Integer, ByRef lifedeath As Boolean) As String
        Dim yXML As New Xml.XmlDocument
        Dim yElem As Xml.XmlElement
        Dim yElem2 As Xml.XmlElement
        Dim yNode As Xml.XmlNode

        Dim yg As String = ""
        Dim s As String = ""
        Dim s2 As String = ""
        Dim x As Integer = DiceRoller(1, 20)
        Dim passionTracker As String = ""
        Dim g1 As Integer, g2 As Integer, g3 As Double

        yXML.Load(here & "\" & CStr(y) & ".xml")
        yElem = yXML.SelectSingleNode("/year")
        s = yElem.GetAttribute("n")
        yg = $"{s}:"

        'If there's an events/text then that always displays.
        Try
            yNode = yXML.SelectSingleNode("//events/text/node()")
            s = yNode.Value
        Catch ex As Exception
            s = ""
        End Try
        yg = $"{yg} {s}"

        'On to the year's event.
        yElem = yXML.SelectSingleNode($"//event[@dice_min <= {x} and @dice_max >= {x}]/text")
        yNode = yElem.SelectSingleNode("./text()")
        s = yNode.Value
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
            s2 = yElem2.GetAttribute("glory-die")
            g2 = yElem2.GetAttribute("glory-flat")
            g3 = yElem2.GetAttribute("glory-multiplier")
            g1 = DiceRoller(CInt(Left(s2, 1)), CInt(Mid(s2, 3)))
            glory += Math.Round(g1 * g2 * g3)

            'Then the battle event.
            x = DiceRoller(1, 20)
            yElem2 = yXML.SelectSingleNode($"//b-event[@dice_min <= {x} and @dice_max >= {x}]/text")
            yNode = yElem2.SelectSingleNode("./text()")
            s = s & yNode.Value
            If yElem2.GetAttribute("glory") <> "" Then
                glory += CInt(yElem.GetAttribute("glory"))
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
            yElem2 = yXML.SelectSingleNode($"//p-event[@dice_min <= {x} and @dice_max >= {x}]/text")
            yNode = yElem2.SelectSingleNode("./text()")
            Try
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
        If InStr(yg, "MISCDEATH") > 0 Then yg = Replace(yg, "MISCDEATH", MiscDeath())
        If InStr(yg, "NOTHINGYEAR") > 0 Then
            yg = ""
        Else
            yg = yg & vbNewLine & vbNewLine
        End If
        YearGen = yg
    End Function

    Public Function MiscDeath(Optional male As Boolean = True) As String
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
