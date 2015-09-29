Module Module1

    Sub Main()

        Dim userName As String
        Dim userNum1 As Double
        Dim userNum2 As Double
        Dim numTotal As Double

        Console.WriteLine("Hello and Welcome to a Visual Basic Console Application.") 'introduction
        Console.WriteLine("My name is V-Bot and we'll be doing math today.")
        Console.ReadLine()

        Console.Write("So, why don't you tell me your name?" & vbCrLf) 'get name

        userName = Console.ReadLine()

        Console.WriteLine(vbCrLf & "Hello and welcome " & userName & vbCrLf) 'read name
        Console.ReadLine()

        Console.WriteLine("Now I will do some math for you.") 'math intro
        Console.WriteLine("I want you to give me 2 numbers.")

        Console.Write("What is your first number?" & vbCrLf) 'first number
        userNum1 = Console.ReadLine()

        Console.Write("What is your second number?" & vbCrLf) 'second number
        userNum2 = Console.ReadLine()

        Console.WriteLine("Let's check if I got those correctly." & vbCrLf)
        Console.ReadLine()

        Console.WriteLine("Your first number was " & userNum1 & ".")
        Console.WriteLine("Your second number was " & userNum2 & ".")
        Console.ReadLine()

        'an IF statement can come here later

        Console.WriteLine("Good, now first I will add those numbers.") 'addition
        numTotal = userNum1 + userNum2
        Console.ReadLine()
        Console.WriteLine(userNum1 & " + " & userNum2 & " = " & numTotal)
        Console.ReadLine()

        Console.WriteLine("Next I will subtract the second number from the first number.") 'subtraction
        numTotal = userNum1 - userNum2
        Console.ReadLine()
        Console.WriteLine(userNum1 & " - " & userNum2 & " = " & numTotal)
        Console.ReadLine()

        Console.WriteLine("Now I will multiply them.") 'multiplication
        numTotal = userNum1 * userNum2
        Console.ReadLine()
        Console.WriteLine(userNum2 & " * " & userNum1 & " = " & numTotal)
        Console.ReadLine()

        Console.WriteLine("Finally, I shall divide number 2 from number 1.") 'division
        numTotal = userNum1 / userNum2
        Console.ReadLine()
        Console.WriteLine(userNum2 & " / " & userNum1 & " = " & numTotal)
        Console.ReadLine()

        Console.WriteLine("Now I will show you a couple extra calculations to show how smart I am.")
        Console.ReadLine()
        Console.WriteLine("Number 1 squared: " & userNum1 ^ 2) 'both numbers squared
        Console.WriteLine("Number 2 squared: " & userNum2 ^ 2)

        Console.WriteLine("Number 1 Modulus 2: " & userNum1 Mod userNum2) '1 modulus 2
        Console.WriteLine("Number 2 Modulus 1: " & userNum2 Mod userNum1) '2 modulus 1
        Console.ReadLine()

    End Sub

End Module
