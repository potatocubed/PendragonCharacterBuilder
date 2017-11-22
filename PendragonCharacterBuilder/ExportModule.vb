Public Module ExportModule
    Sub ExportCharacter(ByRef charSheet As Xml.XmlDocument, name As String, gender As String, tradWoman As Boolean,
                        charAge As Integer, homeland As String, culture As String, charReligion As String,
                        charSonNumber As Integer, charLeige As String, charClass As String, charManor As String,
                        charTraits As String(,), charDirectedTraits As ArrayList, charPassions As ArrayList)
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
            cAtt.Value = charTraits(i, 0)
            cNode.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = charTraits(i, 1)
            cNode.Attributes.Append(cAtt)
            cNode2.AppendChild(cNode)

            cNode3 = charSheet.CreateElement("trait")
            cAtt = charSheet.CreateAttribute("name")
            cAtt.Value = charTraits(i + 13, 0)
            cNode3.Attributes.Append(cAtt)
            cAtt = charSheet.CreateAttribute("value")
            cAtt.Value = charTraits(i + 13, 1)
            cNode3.Attributes.Append(cAtt)
            cNode2.AppendChild(cNode3)
        Next

        If charDirectedTraits.Count > 0 Then
            For i = 0 To charDirectedTraits.Count - 1
                cNode = charSheet.CreateElement("directed-trait")
                s = charDirectedTraits(i)
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
    End Sub

    Sub ExportHistory(sheet As Xml.XmlDocument)

    End Sub

    Sub ExportFamily(sheet As Xml.XmlDocument)

    End Sub
End Module
