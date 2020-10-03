Attribute VB_Name = "Module1"
Sub Stock_Market_Summary()

'Set object variable
Dim CurrentWs As Worksheet
    
'Add header to each sheet
For Each CurrentWs In Worksheets
    CurrentWs.Range("I1").Value = "Ticker"
    CurrentWs.Range("J1").Value = "Yearly Change"
    CurrentWs.Range("K1").Value = "Percent Change"
    CurrentWs.Range("L1").Value = "Total Stock Volume"
    CurrentWs.Range("P1").Value = "Ticker"
    CurrentWs.Range("Q1").Value = "Value"
    
Next CurrentWs


For Each CurrentWs In Worksheets

  ' Set variables for calculations
    Dim Ticker_Name As String
    Ticker_Name = ""
    Dim Volume_Total As Double
    Volume_Total = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Yearly_Percent_Change As Double
    Yearly_Pecent_Change = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Open_Price As Double
    Open_Price = 0
   
    'Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Set row count for each sheet
    Dim Lastrow As Long
    Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

    'set Open_Price first location
    Open_Price = CurrentWs.Cells(2, 3).Value
  
    'Loop through all ticker
     For i = 2 To Lastrow

        'Check if we are still within the same ticker name, if it is not...
         If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
    
          'Set the ticker name
          Ticker_Name = CurrentWs.Cells(i, 1).Value
    
          'Add to the total ticker volume
           Volume_Total = Volume_Total + CurrentWs.Cells(i, 7).Value
          
          'Calculate yearly price change
           Close_Price = CurrentWs.Cells(i, 6).Value
           
           Yearly_Change = Close_Price - Open_Price
                    
           If Open_Price <= 0 Then
              Yearly_Percent_Change = 0
           Else
              Yearly_Percent_Change = (Yearly_Change / Open_Price) * 100
        
           End If
                        
           'add interior color
           If (Yearly_Change) > 0 Then
                CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = rgbGreen
                
           ElseIf Yearly_Price <= 0 Then
                CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = rgbRed
                                       
           End If
                                      
          'Print the Ticker Name, Total Ticker Volume, Yearly Price Change and Yearly Percent Change in the Summary Table
          CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
          CurrentWs.Range("L" & Summary_Table_Row).Value = Volume_Total
          CurrentWs.Range("J" & Summary_Table_Row).Value = Yearly_Change
          CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Percent_Change) & "%")
              
          'Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          'Set the next Open Price
          Open_Price = CurrentWs.Cells(i + 1, 3).Value
      
          'Reset values
          Volume_Percent_Change = 0
          Total_Ticker_Volume = 0
      
    'If the cell immediately following a row is the same Ticker Name...
    Else

        'Add to the Volume Total
        Volume_Total = Volume_Total + CurrentWs.Cells(i, 7).Value
      
    End If
  
  Next i

Next CurrentWs


For Each CurrentWs In Worksheets

'Declare variables for the greatest % summary
Dim a As Double             'greatest value
Dim b As Double             'ticker name
Dim Lookup_value As Double
Dim r As Long
Dim Lastrow_two As Long

a = CurrentWs.Cells(2, 11).Value
r = 1
c = ""
CurrentWs.Range("M2").Value = c
Lastrow_two = CurrentWs.Cells(Rows.Count, 11).End(xlUp).Row

'find the greatest % increase
For r = 2 To Lastrow_two

    b = CurrentWs.Cells(r + 1, 11).Value
    
    If a > b Then
        CurrentWs.Range("Q2").Value = a
        Lookup_value = a
         
        If r = 1 Then
           c = CurrentWs.Cells(r, 9).Value
             
        End If
        
    Else
     Lookup_value = b
     CurrentWs.Range("Q2").Value = b
     c = CurrentWs.Cells(r + 1, 9).Value
          
    End If
     
     a = Lookup_value
     CurrentWs.Range("Q2").Value = a
     CurrentWs.Range("P2").Value = c

Next r

'find the greatest % decrease
For r = 2 To Lastrow_two

    b = CurrentWs.Cells(r + 1, 11).Value
    
    If a < b Then
        CurrentWs.Range("Q3").Value = a
        Lookup_value = a
         
        If r = 1 Then
           c = CurrentWs.Cells(r, 9).Value
             
        End If
        
    Else
     Lookup_value = b
     CurrentWs.Range("Q3").Value = b
     c = CurrentWs.Cells(r + 1, 9).Value
          
    End If
     
     a = Lookup_value
     CurrentWs.Range("Q3").Value = a
     CurrentWs.Range("P3").Value = c

Next r
        
'find the greatest total volume
For r = 2 To Lastrow_two

    b = CurrentWs.Cells(r + 1, 12).Value
    
    If a > b Then
        CurrentWs.Range("Q4").Value = a
        Max_value = a
         
        If r = 1 Then
           c = CurrentWs.Cells(r, 9).Value
             
        End If
        
    Else
     Lookup_value = b
     CurrentWs.Range("Q4").Value = b
     c = CurrentWs.Cells(r + 1, 9).Value
          
    End If
     
     a = Lookup_value
     CurrentWs.Range("Q4").Value = a
     CurrentWs.Range("P4").Value = c

Next r
        
        'add summary title
        CurrentWs.Range("O2").Value = "Greatest % Increase"
        CurrentWs.Range("O3").Value = "Greatest % Decrease"
        CurrentWs.Range("O4").Value = "Greatest Total Volume"
                
        
Next CurrentWs

For Each CurrentWs In Worksheets
    With CurrentWs
    .Columns.AutoFit
    End With
    
Next CurrentWs

End Sub



