Sub stocks()

Dim Summary_Row As Integer
Summary_Row = 2
Dim Ticker As String
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Opening_Value As Boolean
Opening_Value = True
Dim Starting_Stock As Double
Starting_Stock = 0
Dim Year_End_Stock As Double
Year_End_Stock = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0



Dim Greatest_Percentage_Increase As Double
Greatest_Percentage_Increase = 0
Dim Ticker_GPI As String
Dim Greatest_Percentage_Decrease As Double
Greatest_Percentage_Decrease = 0
Dim Ticker_GPD As String
Dim Greatest_Volume_Change As Double
Greatest_Volume_Change = 0
Dim Ticker_GVC As String


For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
'ws.Range("N2").Value = "Greatest % Increase"
'ws.Range("N3").Value = "Greatest % Decrease"
'ws.Range("N4").Value = "Greatest Total Volume"
'ws.Range("O1").Value = "Ticker"
'ws.Range("P1").Value = "Value"


    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value

            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            Year_End_Stock = Year_End_Stock + ws.Cells(i, 6)

            Yearly_Change = Year_End_Stock - Starting_Stock

            Percent_Change = (Yearly_Change / Starting_Stock) * 100

                If (Yearly_Change Or Starting_Stock) = 0 Then

                    Percent_Change = 0

                End If

            ws.Range("I" & Summary_Row).Value = Ticker

            ws.Range("J" & Summary_Row).Value = Yearly_Change

                If Yearly_Change > 0 Then

                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4

                Else

                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                
                End If
            
            ws.Range("K" & Summary_Row).Value = FormatPercent(Percent_Change)

            ws.Range("L" & Summary_Row).Value = Total_Stock_Volume
            
            Summary_Row = Summary_Row + 1
            
            Total_Stock_Volume = 0

            Starting_Stock = 0
            
            Year_End_Stock = 0

            Percent_Change = 0

            Opening_Value = True

        Else

            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            
            'I only want the action to happen once
           Do While Opening_Value = True

            Starting_Stock = Starting_Stock + ws.Cells(i, 3).Value

            Opening_Value = False

           Loop


        End If


    Next i



    
    
    'For j = 2 To 172
            
           ' If Cells(j + 1, 11).Value > Cells(j, 11).Value Then
           
              '  Greatest_Percentage_Increase = Cells(j + 1, 11).Value

               ' Ticker_GPI = cells(j+1, 9).Value
           
           'Else
           
            'Greatest_Percentage_Increase = Cells(j, 11).Value

            'Ticker_GPI = cells(j, 9).Value
          
          'End If

            'If Cells(j + 1, 11).Value < Cells(j, 11).Value Then
            
              '  Greatest_Percentage_Decrease = Cells(j + 1, 11).Value

             '   Ticker_GPD = cells(j+1,9).Value
            
            'Else
            
                'Greatest_Percentage_Decrease = Cells(j, 11).Value

                'Ticker_GPD = cells(j,9).Value
            
           ' End If
    'Next j
    
    
        'For k = 2 to 172

            'If Cells(k + 1, 12).Value > Cells(k, 12).Value Then
            
               ' Greatest_Volume_Change = Cells(k + 1, 12).Value

                'Ticker_GVC = cells(k + 1, 9)
            
            'Else
            
               ' Greatest_Volume_Change = Cells(k, 12).Value
                
                'Ticker_GVC = cells(k, 9)
            
           ' End If
    
            'Next k


'Range("O2").value = Ticker_GPI
'Range("O3").value = Ticker_GPD
'Range("O4").value = Ticker_GVC


'Range("P2").value = Greatest_Percentage_Increase
'Range("P3").value = Greatest_Percentage_Decrease
'Range("P4").value = Greatest_Volume_Change

Summary_Row = 2

Next ws


End Sub

