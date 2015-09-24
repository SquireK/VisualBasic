Module Module1

  Sub Main()

    Dim myNewCar As New Car()
    
    myNewCar.make = "Oldsmobile"
    myNewCar.model = "Cutlas Supreme"
    myNewCar.year = 1986
    myNewCar.color = "Silver"
    
    'With myNewCar
    
    '  '.make = "Oldsmobile"
    '  '.model = "Cutlas Supreme"
    '  '.year = 1986
    '  '.color = "Silver"
    
    'End With
    
    Console.WriteLine("{0} - {1} - {2}", myNewCar.Make, myNewCar.Model, myNewCar.Year)
    Console.ReadLine()
  
  End Sub
  
  'Function determineMarketValue(ByVal myCar As car) As Double

  '  'Return 100.00

  'End Function

End Module
