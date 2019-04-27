Attribute VB_Name = "Module2"
Sub Stocks()
   
    Dim Begining_Year As Double
    Dim end_year As Double
    Dim year_difference As Double
    
    Begining_Year = 0
    end_year = 0
    year_difference = Begining_Year - end_year

    ' make ticker a vaiable
    Dim ticker As String

    'set a variable for holding the ticker
    Dim Ticker_Vol As Double
    ' having a hard time remembering why i set the ticker_vol to 0
    Ticker_Vol = 0

        Dim Ticker_Name As Integer
        Ticker_Name = 2
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through all tickers
    For i = 2 To LastRow

        ' check if it is the same ticker and if not then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Change ticker
        ticker = Cells(i, 1).Value

        'add to the ticker total
        Total = Total + Cells(i, 7).Value

    'print the Ticker name in the ticker name
    Range("I" & Ticker_Name).Value = ticker

    'print the Volume amount in column
    Range("J" & Ticker_Name).Value = Ticker_Vol

    'add one to ticker name
    Ticker_Name = Ticker_Name + 1

    'Reset Ticker vol total
    Ticker_Vol = 0


' if cell following the same ticker do this
    Else


'add vol total
Ticker_Vol = Ticker_Vol + Cells(i, 7).Value

End If

Next i
End Sub

