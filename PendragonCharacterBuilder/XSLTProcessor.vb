Module XSLTProcessor
    Public Sub ProcessXMLOutput(sheet As Xml.XmlDocument, exportDir As String)
        'Loop through every xslt in the 'transformations' subdirectory and apply it to the XML.
        Dim transformDir As String = exportDir & "\xml\transformations\"
    End Sub
End Module
