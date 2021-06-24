Attribute VB_Name = "Module1"
Sub stocktest()

'The ticker variable will hold the row that the ticker will be output.
Dim ticker As Integer

'setting the starting row to 2 as the first line is reserved for the column headers

'The Totalstock variable will hold the total value of the stock to be printed in the "total stock volume" column of the worksheet
Dim totalttock As Long

'Rather than go all the way to the bottom, this variable and the following formula will be used to calculate the final row.
Dim lastrow As Long

'These two will be used to grab the start of year <open> and the end of year <close> to calculate teh yearly change and percent change
Dim startyear As Double
Dim endyear As Double

'Defining variables for the #bonus section
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Variant
Dim increaseticker As String
Dim decreaseticker As String
Dim volumeticker As String


'This should move to the next worksheet
For Each ws In Worksheets

'Naming our new columns
Range("I1, O1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'These are exclusive to the #bonus section
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"


totalstock = 0


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
ticker = 2

greatestincrease = 0
greatestdecrease = 0
greatestvolume = 0

    'The way I've got this macro set up will skip the first one, so this will establish the start of the first year for the first ticker
startyear = Cells(2, 3).Value
    'r here means "row"
    For r = 2 To lastrow

        'If the ticker for the next line is not the same ticker as the current line...
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        'Set the value of the ticker line to the current ticker showing we've marked it...
        Cells(ticker, 9).Value = Cells(r, 1).Value

        'Add the current row to the total stock volume...
        totalstock = totalstock + Cells(r, 7).Value

        'Display the total stock volume for that ticker...
        Cells(ticker, 12).Value = totalstock

        'This part we will calculate the yearly change and the percent change
        endyear = Cells(r, 6).Value
        'calculating the total change value
        Cells(ticker, 10).Value = endyear - startyear
        
            'This nested IF checks if the value is greater or less than 0 to change the formatting to GREEN for positive and RED for negative
            If Cells(ticker, 10).Value > 0 Then
        
            Cells(ticker, 10).Interior.ColorIndex = 4
        
            ElseIf Cells(ticker, 10).Value < 0 Then
        
            Cells(ticker, 10).Interior.ColorIndex = 3
        
            End If

    
        'calculating the total change percentage
            'You can't divide by 0! If there's a zero at the start it returns an "N/A" instead.
            If startyear = 0 Then
                Cells(ticker, 11).Value = "N/A"
            Else
                Cells(ticker, 11).Value = ((endyear - startyear) / startyear) 'You have to divide by the start to get the % change so if it's 0 it doesn't work
                Cells(ticker, 11).NumberFormat = "0.00%"
            End If
    
        'Calculating if this is the greatest volume we've seen and storing that value
            If totalstock > greatestvolume Then
                greatestvolume = totalstock
                volumeticker = Cells(ticker, 9).Value
            End If
            
        'Calculating if this is the greatest % increase we've seen and storing that value
            If Cells(ticker, 11).Value <> "N/A" Then
                If Cells(ticker, 11).Value > greatestincrease Then
                    greatestincrease = Cells(ticker, 11).Value
                    increaseticker = Cells(ticker, 9).Value
                End If
            End If
        
        'calculating if this is the greatest % decrease we've seen
            If Cells(ticker, 11).Value <> "N/A" Then
                If Cells(ticker, 11).Value < greatestdecrease Then
                    greatestdecrease = Cells(ticker, 11).Value
                    decreaseticker = Cells(ticker, 9).Value
                End If
            End If
            
        'Reset the startyear value for the next ticker
        startyear = Cells(r + 1, 3).Value
    
        'Reset the total stock value to calculate for the next ticker
        totalstock = 0
    
        'Finally, go to the next line for the next ticker
        ticker = ticker + 1
    
        'Otherwise, if the next ticker is the same as the current ticker
        'All this one really needs to do is update the total stock. The first IF that checks if not equal is doing most of the work
        ElseIf Cells(r + 1, 1).Value = Cells(r, 1).Value Then
    
        totalstock = totalstock + Cells(r, 7).Value
    
    
        'Close the IF statement
        End If

    Next r

Range("O2").Value = increaseticker
Range("O3").Value = decreaseticker
Range("O4").Value = volumeticker
Range("P2").Value = greatestincrease
Range("P3").Value = greatestdecrease
Range("P4").Value = greatestvolume
Range("P2 : P3").NumberFormat = "0.00%"

'Why is this at the end instead of the beginning?
'Well, for some reason putting this at the beginning of the loop makes this not work correctly.
'To be honest, I'm not sure why. It works identically on every sheet but the last one when this is at the beginning
'When this is at the beginning it doesn't add the headers like it should. Weird right? Not sure why. But it works now!
ws.Activate

Next ws

End Sub
