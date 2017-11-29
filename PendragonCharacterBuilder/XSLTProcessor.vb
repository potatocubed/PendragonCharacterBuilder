Module XSLTProcessor
    Public Sub ProcessXMLOutput(sheetLoc As String, exportDir As String)
        'Loop through every xslt in the 'transformations' subdirectory and apply it to the XML.
        Dim transformDir As String = exportDir & "\xml\transformations\"
        Dim xslt As New Xml.Xsl.XslCompiledTransform
        Dim c As New Collection
        Dim f As String
        Dim s As String = "whatever"
        Dim x As New Xml.XmlDocument
        Dim xElem As Xml.XmlElement
        Dim charName As String
        Dim n As Integer
        Dim nsManager As Xml.XmlNamespaceManager

        n = InStrRev(sheetLoc, "\")
        Try
            charName = Mid(sheetLoc, n + 1)
        Catch ex As Exception
            charName = sheetLoc
        End Try

        charName = Left(charName, Len(charName) - 4)

        f = Dir(transformDir)
        Do While f <> ""
            If Right(f, 3) = "xsl" Then c.Add(f)
            f = Dir()
        Loop

        For Each fl In c
            s = ""
            x.Load(transformDir & fl)
            nsManager = New Xml.XmlNamespaceManager(x.NameTable)
            nsManager.AddNamespace("xsl", x.DocumentElement.Attributes("xmlns:xsl").InnerText)
            xElem = x.SelectSingleNode("//xsl:param[@name='outputTag']", nsManager)
            s = xElem.InnerText
            charName = $"{charName} {s}.txt"
            xslt.Load(transformDir & fl)
            xslt.Transform(sheetLoc, exportDir & "\" & charName)
        Next
    End Sub

    Sub Main()
        'For testing.

        Dim x As String = "C:\Users\LonghurstC\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder\Lady Laudine.xml"
        Dim d As String = "C:\Users\LonghurstC\source\repos\PendragonCharacterBuilder\PendragonCharacterBuilder"

        ProcessXMLOutput(x, d)
    End Sub
End Module
