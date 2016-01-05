Function checkString(ByVal statementToCheck)
        Dim a, b, c As Integer
        
        'the unallowed characters
        Dim badLetters As String = "abcdefghijklmnopqrstuvwxyz ;'""!@#$%^&*()_-+=,<>/?\|]}[{`~"
        
        'the allowed characters
        Dim goodLetters As String = "1234567890."
        
        'the arrays
        Dim stringToCheck() As Char
        Dim badLettersArray() As Char
        Dim goodLettersArray() As Char
        
        'the flag to indicate whether there is a bad character
        Dim badFlag As Boolean = False

        'creates the arrays and forces the input to a lowercase form
        stringToCheck = statementToCheck.Text.ToLower
        badLettersArray = badLetters
        goodLettersArray = goodLetters

        '=====the "wrapper" for loop=====
        'the for loop gets the text's length and examines each character
        For a = 0 To statementToCheck.TextLength - 1

            'these next 2 for loops check to see if the selected character is "bad"
            For b = 0 To badLetters.Length - 1

                'this if statement checks to see if the character at location (a) in the counter
                '   is equal to location (b) in the bad character counter
                If stringToCheck(a) = badLettersArray(b) Then

                    'this sets the flag to true
                    badFlag = True

                End If

            Next

        Next
        
        'this returns to the program with a true or false indicator
        If badFlag Then

            Return True

        ElseIf Not badFlag Then

            Return False

        End If
End Function
