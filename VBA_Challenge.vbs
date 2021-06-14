Sub Stock()

For Each ws In Worksheets
'This action loops worksheets

Dim ticker As String
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Volume As Double
'Set volume to 0
Volume = 0
Dim Summary_Table As Integer
'Setting Varitables

Summary_Table = 2
'This action starts the summary table in row 2
Dim first As Long
'This statement gose along with the opneing_price in the lines to follow
first = 2
Dim percentage As Double




'Set headings to columns
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Price"
ws.Range("K1") = "Percentage"
ws.Range("L1") = "Volume"

'Grabs everything column A and loops it to the last working cell.
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'use for i to loop each row from first sheet to last
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'Everything under this statement is apart of the If loop
'If the present cell dose not equal the next cell then. 'These products use these cells which will go into the summary table
ticker = ws.Cells(i, 1).Value
'(Volume +) means add the next volume in the nxet row until the loop ends
Volume = Volume + ws.Cells(i, 7).Value
'This statement gose with first = 2. 'It offsets the rows by two, there is a loop for opening price in a line to follow, wtich is first = i + 1
Opening_Price = ws.Range("C" & first).Value

Closing_Price = ws.Cells(i, 6).Value

Yearly_Change = Closing_Price - Opening_Price
'If within the If to cancel out dividing by 0
If Opening_Price = 0 Then
percentage = 0
Else
'If not 0 then follow these rules
Opening_Price = ws.Range("C" & first).Value
'opening price = open price, + 2 which will start counting at the second row
percentage = Yearly_Change / Opening_Price

'Print statements in these columns
ws.Range("I" & Summary_Table).Value = ticker
ws.Range("L" & Summary_Table).Value = Volume
ws.Range("J" & Summary_Table).Value = Yearly_Change
ws.Range("K" & Summary_Table).Value = percentage
'change precentage column to precent
ws.Range("K" & Summary_Table).NumberFormat = "0.00%"

'Reset volume
Volume = 0

'End If for the second loop
End If
'Second If statement within the If for changing colors.
If Yearly_Change > 0 Then
ws.Range("J" & Summary_Table).Interior.Color = RGB(0, 255, 0)

Else
ws.Range("J" & Summary_Table).Interior.Color = RGB(255, 0, 0)
End If
'Summary Table + 1 will go to the next row once the condition is met
Summary_Table = Summary_Table + 1

'This loops the previous opening price
first = i + 1
Else
'last statement of first If, this will add up the volume until the condition is met
Volume = Volume + ws.Cells(i, 7).Value


End If

Next i
'input bonus
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'This is looking for the greatest incerase
For i = 2 To lastrow
If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
ws.Range("Q2").Value = ws.Cells(i, 11).Value
ws.Range("P2").Value = ws.Cells(i, 9).Value
End If

'This is looking for the greatest decrease
If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
ws.Range("Q3").Value = ws.Cells(i, 11).Value
ws.Range("P3").Value = ws.Cells(i, 9).Value
End If

'This is lookinh for the total volume
If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
ws.Range("Q4").Value = ws.Cells(i, 12).Value
ws.Range("P4").Value = ws.Cells(i, 9).Value
End If
'change precentage columns to precent
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"


Next i
Next ws

End Sub

