Public Module ExportModule
    Sub ExportCharacter(ByRef charSheet As Xml.XmlDocument, name As String, gender As String, tradWoman As Boolean,
                        charAge As Integer, homeland As String, culture As String, charReligion As String, fatherClass As String,
                        charSonNumber As Integer, charLeige As String, charClass As String, charManor As String,
                        charTraits As String(,), charDirectedTraits As ArrayList, charPassions As ArrayList,
                        siz As Integer, dex As Integer, str As Integer, con As Integer, app As Integer,
                        features As String, skills As String(,), glory As Integer, squire As String(),
                        horses As ArrayList, heirlooms As ArrayList, oldKnights As String(,), maKnights As String(,),
                        youngKnights As String(,), men As Integer, levies As Integer)
        Dim cElem As Xml.XmlElement
        Dim cNode As Xml.XmlNode
        Dim cNode2 As Xml.XmlNode
        Dim cNode3 As Xml.XmlNode
        Dim cAtt As Xml.XmlAttribute

        Dim s As String
        Dim s2 As String
        Dim x As Integer

        'Personal Data
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("personal-data")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("personal-data")

        cNode = charSheet.CreateElement("name")
        cNode.AppendChild(charSheet.CreateTextNode(name))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("glory")
        cNode.AppendChild(charSheet.CreateTextNode(glory))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("age")
        cNode.AppendChild(charSheet.CreateTextNode(charAge))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("gender")
        cNode.AppendChild(charSheet.CreateTextNode(gender))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("child-number")
        cNode.AppendChild(charSheet.CreateTextNode(charSonNumber))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("father-class")
        cNode.AppendChild(charSheet.CreateTextNode(fatherClass))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("homeland")
        cNode.AppendChild(charSheet.CreateTextNode(homeland))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("culture")
        cNode.AppendChild(charSheet.CreateTextNode(culture))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("religion")
        cNode.AppendChild(charSheet.CreateTextNode(charReligion))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("lord")
        cNode.AppendChild(charSheet.CreateTextNode(charLeige))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("current-class")
        cNode.AppendChild(charSheet.CreateTextNode(charClass))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("current-home")
        cNode.AppendChild(charSheet.CreateTextNode(charManor))
        cElem.AppendChild(cNode)

        'Personality Traits
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("personality-traits")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("personality-traits")

        For i = 0 To 12
            cNode = charSheet.CreateElement("trait-pair")
            cNode2 = cElem.AppendChild(cNode)

            cNode = charSheet.CreateElement("trait")
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = charTraits(0, i)
            cNode.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = charTraits(1, i)
            cNode.Attributes.Append(cAtt)
            cNode2.AppendChild(cNode)

            cNode3 = charSheet.CreateElement("trait")
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = charTraits(0, i + 13)
            cNode3.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = charTraits(1, i + 13)
            cNode3.Attributes.Append(cAtt)
            cNode2.AppendChild(cNode3)
        Next

        If charDirectedTraits.Count > 0 Then
            For Each trait In charDirectedTraits
                cNode = charSheet.CreateElement("directed-trait")
                s = trait
                x = InStr(s, "/")
                s2 = Mid(s, x + 1)
                s = Left(s, x - 1)
                cAtt = charSheet.CreateAttribute("name")
                cAtt.Value = s
                cNode.Attributes.Append(cAtt)
                cAtt = charSheet.CreateAttribute("value")
                cAtt.Value = s2
                cNode.Attributes.Append(cAtt)
            Next
        End If

        'Passions
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("passions")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("passions")

        If charPassions.Count > 0 Then
            For Each passion In charPassions
                cNode = charSheet.CreateElement("passion")
                s = passion
                x = InStr(s, "/")
                s2 = Mid(s, x + 1)
                s = Left(s, x - 1)
                cAtt = charSheet.CreateAttribute("name")
                cAtt.Value = s
                cNode.Attributes.Append(cAtt)
                cAtt = charSheet.CreateAttribute("value")
                cAtt.Value = s2
                cNode.Attributes.Append(cAtt)
                cElem.AppendChild(cNode)
            Next
        End If

        'Attributes
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("attributes")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("attributes")

        cNode = charSheet.CreateElement("base-attributes")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("base-attributes")

        cNode = charSheet.CreateElement("size")
        cAtt = charSheet.CreateAttribute("short")
        cAtt.Value = "SIZ"
        cNode.Attributes.Append(cAtt)
        cNode.AppendChild(charSheet.CreateTextNode(siz))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("dexterity")
        cAtt = charSheet.CreateAttribute("short")
        cAtt.Value = "DEX"
        cNode.Attributes.Append(cAtt)
        cNode.AppendChild(charSheet.CreateTextNode(dex))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("strength")
        cAtt = charSheet.CreateAttribute("short")
        cAtt.Value = "STR"
        cNode.Attributes.Append(cAtt)
        cNode.AppendChild(charSheet.CreateTextNode(str))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("constitution")
        cAtt = charSheet.CreateAttribute("short")
        cAtt.Value = "CON"
        cNode.Attributes.Append(cAtt)
        cNode.AppendChild(charSheet.CreateTextNode(con))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("appearance")
        cAtt = charSheet.CreateAttribute("short")
        cAtt.Value = "APP"
        cNode.Attributes.Append(cAtt)
        cNode.AppendChild(charSheet.CreateTextNode(app))
        cElem.AppendChild(cNode)

        cElem = charSheet.SelectSingleNode("//character/attributes")
        cNode = charSheet.CreateElement("derived-attributes")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("derived-attributes")

        Dim charDamage As String = CStr(Math.Round((str + siz) / 6)) & "d6"
        Dim charHealing As Integer = Math.Round((con + str) / 10)
        Dim charMove As Integer = Math.Round((str + dex) / 10)
        Dim charHP As Integer = con + siz
        Dim charUnconscious As Integer = Math.Round(charHP / 4)

        cNode = charSheet.CreateElement("damage")
        cNode.AppendChild(charSheet.CreateTextNode(charDamage))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("healing")
        cNode.AppendChild(charSheet.CreateTextNode(charHealing))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("move")
        cNode.AppendChild(charSheet.CreateTextNode(charMove))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("HP")
        cNode.AppendChild(charSheet.CreateTextNode(charHP))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("unconscious")
        cNode.AppendChild(charSheet.CreateTextNode(charUnconscious))
        cElem.AppendChild(cNode)

        'Feature(s)
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("features")
        cNode.AppendChild(charSheet.CreateTextNode(features))
        cElem.AppendChild(cNode)

        'Skills
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("skills")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("skills")
        cNode = charSheet.CreateElement("non-combat")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("non-combat")

        For i = 0 To 33
            cNode = charSheet.CreateElement("skill")
            s = skills(0, i)
            s2 = skills(1, i)
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = s
            cNode.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = s2
            cNode.Attributes.Append(cAtt)
            cElem.AppendChild(cNode)
        Next

        cElem = charSheet.SelectSingleNode("//character/skills")
        cNode = charSheet.CreateElement("combat")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("combat")

        For i = 27 To 39
            cNode = charSheet.CreateElement("skill")
            s = skills(0, i)
            s2 = skills(1, i)
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = s
            cNode.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = s2
            cNode.Attributes.Append(cAtt)
            cElem.AppendChild(cNode)
        Next

        'Stuff
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("stuff")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("stuff")

        Dim stuffArray As New ArrayList
        If tradWoman Then
            stuffArray.Add("Sewing kit")
            stuffArray.Add("Fine clothing (worth £1)")
            stuffArray.Add("Simple jewellery")
            stuffArray.Add("Toilet articles")
            stuffArray.Add("A chest")
        Else
            stuffArray.Add("Chainmail")
            stuffArray.Add("Shield")
            stuffArray.Add("Spears (2)")
            stuffArray.Add("Sword")
            stuffArray.Add("Dagger")
            stuffArray.Add("Fine clothing (worth £1)")
            stuffArray.Add("Personal gear")
            stuffArray.Add("Travel gear")
            stuffArray.Add("War gear")
        End If

        stuffArray.AddRange(heirlooms)

        For Each item In stuffArray
            If Not (InStr(item, "charger") > 0 Or InStr(item, "rouncy") > 0 Or InStr(item, "courser") > 0) Then
                cNode = charSheet.CreateElement("item")
                cNode.AppendChild(charSheet.CreateTextNode(item))
                cElem.AppendChild(cNode)
            End If
        Next

        'Squire/Lady in Waiting
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("squire")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("squire")

        cNode = charSheet.CreateElement("name")
        cNode.AppendChild(charSheet.CreateTextNode(squire(0)))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("age")
        cNode.AppendChild(charSheet.CreateTextNode(squire(1)))
        cElem.AppendChild(cNode)

        For i = 2 To 8
            If squire(i) = "xx" Then Exit For
            cNode = charSheet.CreateElement("skill")
            s = squire(i)
            s2 = squire(i + 1)
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = s
            cNode.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = s2
            cNode.Attributes.Append(cAtt)
            cElem.AppendChild(cNode)
        Next

        'And horses
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("horses")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("horses")

        For Each h In horses
            cNode = charSheet.CreateElement("horse")
            cAtt = charSheet.CreateAttribute("type")
            cAtt.Value = h
            cNode.Attributes.Append(cAtt)
            cElem.AppendChild(cNode)
        Next

        'And your personal army.
        cElem = charSheet.SelectSingleNode("//character")
        cNode = charSheet.CreateElement("army")
        cElem.AppendChild(cNode)
        cElem = cElem.SelectSingleNode("army")

        x = 0
        If oldKnights(0, 0) <> "" Then x += 1
        For i = 0 To 3
            If maKnights(0, i) <> "" Then x += 1
        Next
        For i = 0 To 3
            If youngKnights(0, i) <> "" Then x += 1
        Next

        cNode = charSheet.CreateElement("family-knights")
        cNode.AppendChild(charSheet.CreateTextNode(CStr(x) & " plus yourself"))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("lineage-men")
        cNode.AppendChild(charSheet.CreateTextNode(CStr(men)))
        cElem.AppendChild(cNode)

        cNode = charSheet.CreateElement("levies")
        cNode.AppendChild(charSheet.CreateTextNode(CStr(levies)))
        cElem.AppendChild(cNode)
    End Sub

    Sub ExportHistory(ByRef sheet As Xml.XmlDocument, history As String)
        Dim cElem As Xml.XmlElement
        Dim cElem2 As Xml.XmlElement
        Dim cNode As Xml.XmlNode
        Dim passages As String()

        passages = history.Split(vbNewLine & vbNewLine)

        cElem = sheet.SelectSingleNode("//history")
        For Each p In passages
            p = Replace(p, vbNewLine, "")
            p = Replace(p, vbLf, "")
            p = Replace(p, "\n", "")
            If p <> "" Then
                cElem2 = sheet.CreateElement("p")
                cNode = sheet.CreateTextNode(p)
                cElem2.AppendChild(cNode)
                cElem.AppendChild(cElem2)
            End If
        Next
    End Sub

    Sub ExportFamily(ByRef sheet As Xml.XmlDocument, fatherName As String, grandfatherName As String, fatherGlory As Integer, grandfatherGlory As Integer,
                     motherName As String, motherStatus As String, oldKnights As String(,), maKnights As String(,), youngKnights As String(,), pUncles As ArrayList,
                     pAunts As ArrayList, mUncles As ArrayList, mAunts As ArrayList, brothers As ArrayList, sisters As ArrayList, cousins As ArrayList)

        Dim cElem As Xml.XmlElement
        Dim cElem2 As Xml.XmlElement
        Dim cNode As Xml.XmlNode
        Dim cAtt As Xml.XmlAttribute
        Dim s As String
        Dim x As Integer

        cElem = sheet.SelectSingleNode("//family")
        cElem2 = sheet.CreateElement("grandfather-generation")
        cElem.AppendChild(cElem2)
        cElem = sheet.SelectSingleNode("//family/grandfather-generation")
        cElem2 = sheet.CreateElement("person")
        cNode = sheet.CreateTextNode($"{grandfatherName} (your grandfather). Dead with {grandfatherGlory} glory.")
        cElem2.AppendChild(cNode)
        cElem.AppendChild(cElem2)

        If oldKnights(0, 0) <> "" Then
            cElem2 = sheet.CreateElement("person")
            cNode = sheet.CreateTextNode($"{oldKnights(0, 0)} (old knight). {oldKnights(0, 2)}")
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        End If

        cElem = sheet.SelectSingleNode("//family")
        cElem2 = sheet.CreateElement("father-generation")
        cElem.AppendChild(cElem2)
        cElem = sheet.SelectSingleNode("//family/father-generation")

        cElem2 = sheet.CreateElement("person")
        cAtt = sheet.CreateAttribute("relation")
        cAtt.Value = "father"
        cElem2.Attributes.Append(cAtt)
        cNode = sheet.CreateTextNode($"{fatherName} (your father). Dead with {fatherGlory} glory.")
        cElem2.AppendChild(cNode)
        cElem.AppendChild(cElem2)

        cElem2 = sheet.CreateElement("person")
        cAtt = sheet.CreateAttribute("relation")
        cAtt.Value = "mother"
        cElem2.Attributes.Append(cAtt)
        cNode = sheet.CreateTextNode($"{motherName} (your mother). {motherStatus}")
        cElem2.AppendChild(cNode)
        cElem.AppendChild(cElem2)

        For i = 0 To 3
            If maKnights(i, 0) <> "" Then
                cElem2 = sheet.CreateElement("person")
                cAtt = sheet.CreateAttribute("relation")
                If InStr(maKnights(i, 2), fatherName) Then
                    s = "paternal-uncle"
                Else
                    s = "maternal-uncle"
                End If
                cAtt.Value = s
                cElem2.Attributes.Append(cAtt)
                cNode = sheet.CreateTextNode($"{maKnights(i, 0)} (middle-aged knight, {s}). {maKnights(i, 2)}")
                cElem2.AppendChild(cNode)
                cElem.AppendChild(cElem2)
            End If
        Next

        For Each p In pUncles
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "paternal uncle"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        For Each p In mUncles
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "maternal uncle"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        For Each p In pAunts
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "paternal aunt"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        For Each p In mAunts
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "maternal aunt"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        cElem = sheet.SelectSingleNode("//family")
        cElem2 = sheet.CreateElement("same-generation")
        cElem.AppendChild(cElem2)
        cElem = sheet.SelectSingleNode("//family/same-generation")

        For i = 0 To 5
            If youngKnights(i, 0) <> "" Then
                cElem2 = sheet.CreateElement("person")
                cAtt = sheet.CreateAttribute("relation")
                If InStr(youngKnights(i, 2), "cousin") Then
                    s = "cousin"
                Else
                    s = "brother"
                End If
                cAtt.Value = s
                cElem2.Attributes.Append(cAtt)
                cNode = sheet.CreateTextNode($"{youngKnights(i, 0)} (young knight, {s}). {youngKnights(i, 2)}")
                cElem2.AppendChild(cNode)
                cElem.AppendChild(cElem2)
            End If
        Next

        For Each p In brothers
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "brother"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        For Each p In sisters
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "sister"
            cAtt.Value = s
            x = InStr(p, ".")
            p = p.ToString.Insert(x - 1, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

        For Each p In cousins
            p = Replace(p, "  ", " ")
            cElem2 = sheet.CreateElement("person")
            cAtt = sheet.CreateAttribute("relation")
            s = "cousin"
            cAtt.Value = s
            'x = InStr(p, ".")
            'p.ToString.Insert(x, $" ({s})")
            cElem2.Attributes.Append(cAtt)
            cNode = sheet.CreateTextNode(p)
            cElem2.AppendChild(cNode)
            cElem.AppendChild(cElem2)
        Next

    End Sub
End Module
