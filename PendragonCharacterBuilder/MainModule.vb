Module MainModule

    Sub Main()
        'The basic program will ask for a name and gender and assume that you're a Salisbury knight from the core book.
        'It won't assign your attributes or your various bonuses, just the standard stuff.
        'Then it'll output an XML file, then convert that into basic forum bbcode and plain text.
        'Everything generate-able from a random table will be generated from a table.

        Dim charName As String
        Dim charGender As String
        Dim charAge As String
        Dim fatherName As String
        Dim grandfatherName As String

        Console.WriteLine("Welcome to the Pendragon character generator!")
        Console.WriteLine("This program will basically hammer through all the random tables in the Pendragon core book for you, but it won't make any actual *decisions*.")
        Console.WriteLine("So you'll have to dot the i's and cross the t's yourself.")

        Console.WriteLine("To begin, is your character male [m] or female [f]?")
        charGender = ""
        While charGender = ""
            charGender = Console.Read()
            charGender = Left(charGender, 1)    'Getting just the first letter.
            charGender = LCase(charGender)
            If charGender <> "m" And charGender <> "f" Then
                Console.WriteLine("Please enter m or f.")
            End If
        End While
    End Sub

End Module
