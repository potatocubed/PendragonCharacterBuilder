Public Module ExportModule
    Sub ExportCharacter(ByRef charSheet As Xml.XmlDocument, name As String, gender As String, tradWoman As Boolean,
                        charAge As Integer, homeland As String, culture As String, charReligion As String,
                        charSonNumber As Integer, charLeige As String, charClass As String, charManor As String,
                        charTraits As String(,), charDirectedTraits As ArrayList, charPassions As ArrayList,
                        siz As Integer, dex As Integer, str As Integer, con As Integer, app As Integer,
                        features As String, skills As String(,), glory As Integer, squire As String(),
                        horses As ArrayList, heirlooms As ArrayList)
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

        For Each h In heirlooms
            If InStr(h, "charger") > 0 Or InStr(h, "rouncy") > 0 Or InStr(h, "courser") > 0 Then
                'The horses are taken care of elsewhere.
                Try
                    heirlooms.Remove(h)
                Catch
                    'Do nothing, just in case this screws up.
                End Try
            End If
        Next

        stuffArray.AddRange(heirlooms)

        For Each item In stuffArray
            cNode = charSheet.CreateElement("item")
            cNode.AppendChild(charSheet.CreateTextNode(item))
            cElem.AppendChild(cNode)
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

        'And Horses
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
    End Sub

    Sub ExportHistory(ByRef sheet As Xml.XmlDocument)
        Dim cElem As Xml.XmlElement
        Dim cNode As Xml.XmlNode
        Dim cAtt As Xml.XmlAttribute

        Dim s As String
        Dim s2 As String
        Dim x As Integer

        cElem = sheet.SelectSingleNode("//history")
    End Sub

    Sub ExportFamily(ByRef sheet As Xml.XmlDocument)

    End Sub
End Module
