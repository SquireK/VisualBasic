Module Module1

    Sub Main()

        Console.WriteLine("Please type something and press Enter.")
        Dim userValue As String

        userValue = Console.ReadLine()
        Console.WriteLine("You typed: " & userValue)
        Console.ReadLine()

        Console.WriteLine("Would you prefer what is behind door number 1, 2, or 3?")

        userValue = Console.ReadLine()

        If userValue = "1" Then
            Console.WriteLine("You won a new car!")
            Console.ReadLine()
        ElseIf userValue = "2" Then
            Console.WriteLine("You won a new boat!")
            Console.ReadLine()
        ElseIf userValue = "3" Then
            Console.WriteLine("You won a new cat!")
            Console.ReadLine()
        Else
            Console.WriteLine("Sorry, we did not understand your response. You Lose!")
            Console.ReadLine()
        End If

    End Sub

End Module
