Public Module XMLInterpreters
    Function SkillLister(f As String, Optional knightly As String = "", Optional nonKnightly As String = "", Optional combat As String = "") As ArrayList
        'The various optional parameters should come in as 'true' or 'false' or nothing.
        knightly = LCase(knightly)
        nonKnightly = LCase(nonKnightly)
        combat = LCase(combat)

        Dim skArray As New ArrayList
        Dim skillDoc As New Xml.XmlDocument
        Dim skillList As Xml.XmlNodeList

        Dim xp As String
        xp = "//skill"

        If knightly <> "" Or nonKnightly <> "" Or combat <> "" Then
            xp = xp & "["

            If knightly <> "" Then
                'We care about the knightly status of skills.
                xp = $"{xp}@knightly='{knightly}'"
            End If

            If nonKnightly <> "" Then
                'We care about the non-knightly status of skills.
                If knightly <> "" Then xp = $"{xp} and "
                xp = $"{xp}@non-knightly='{nonKnightly}'"
            End If

            If combat <> "" Then
                'We care about combat vs non-combat skills.
                If knightly <> "" Or nonKnightly <> "" Then xp = $"{xp} and "
                xp = $"{xp}@combat='{combat}'"
            End If

            xp = xp & "]"
        End If

        skillDoc.Load(f)
        skillList = skillDoc.SelectNodes(xp)
        For Each skill As Xml.XmlElement In skillList
            skArray.Add(skill.GetAttribute("name"))
        Next

        SkillLister = skArray
    End Function

    Function HeirloomGenerator(f As String, Optional pagan As Boolean = False, Optional recurse As Boolean = False) As String
        Dim s As String = ""
        Dim s2 As String = ""
        Dim xp As String = "//heirloom"
        Dim x As Integer = DiceRoller(1, 20)

        Dim hXML As New Xml.XmlDocument
        Dim hElem As Xml.XmlElement
        Dim hElem2 As Xml.XmlElement
        Dim hNode As Xml.XmlNode

        Do While (pagan = True And x = 8) Or (recurse = False And x = 20)
            x = DiceRoller(1, 20)
        Loop

        If (recurse And x = 20) Then
            s = HeirloomGenerator(f, pagan, False) & "//" & HeirloomGenerator(f, pagan, False)
        Else
            xp = $"{xp}[@dice_min <= {x} and @dice_max >= {x}]"

            hXML.Load(f)
            hElem = hXML.SelectSingleNode(xp)
            Try
                hNode = hElem.SelectSingleNode("./text/node()[1]")
                s = hNode.Value
            Catch ex As Exception
                s = ""
            End Try

            Try
                hElem2 = hElem.SelectSingleNode("./sub-roll")
                x = DiceRoller(hElem2.GetAttribute("number"), hElem2.GetAttribute("sides"))
                hElem2 = hElem.SelectSingleNode($"./sub-roll/sub-outcome[@dice_min <= {x} and @dice_max >= {x}]")
                hNode = hElem2.SelectSingleNode("./node()")
                s = s & hNode.Value
                s = Replace(s, "SUBROLL", CStr(x))
            Catch ex As Exception
                s = s & ""
            End Try
        End If

        HeirloomGenerator = s
    End Function
End Module
