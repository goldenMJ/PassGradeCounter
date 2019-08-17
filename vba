Sub passCounter()

'create a variable to store pass counter
Dim passCounter As Integer

'initially set the passCounter to zero for each row
passCounter = 0

'loop through each row
For i = 2 To 13
    'while in each row, loop through each exam column
    For j = 4 To 8
    
        'if a column contains the word "pass"
        If (Cells(i, j).Value) = "Pass" Then
        
            'add 1 to passCounter
            passCounter = passCounter + 1
            
        End If
      Next j
      
            'Once we have iterated through each column in row i, print the value in the total column
            Cells(i, 9).Value = passCounter
            
    Next i

End Sub
