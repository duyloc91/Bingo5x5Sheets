Attribute VB_Name = "Module2"
Sub randomizeBingoWithoutDuplicates()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(1)
    
    ws.Activate
    
    For i = 0 To 49
            'first box
            For j = 1 To 5
                Cells(i * 6 + 2, j).Value = Int(20 * Rnd + 20 * (j - 1) + 1)
                
                firstRow = Cells(i * 6 + 2, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Do While randomVal = firstRow
                    randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Loop
                Cells(i * 6 + 3, j).Value = randomVal
                
                secondRow = Cells(i * 6 + 3, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow
                    randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Loop
                Cells(i * 6 + 4, j).Value = randomVal
                
                thirdRow = Cells(i * 6 + 4, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow Or randomVal = thirdRow
                    randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Loop
                Cells(i * 6 + 5, j).Value = randomVal
                
                
                fourRow = Cells(i * 6 + 5, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow Or randomVal = thirdRow Or randomVal = fourRow
                    randomVal = Int(20 * Rnd + 20 * (j - 1) + 1)
                Loop
                Cells(i * 6 + 6, j).Value = randomVal
            Next j
            
            
            'Second box
            For j = 7 To 11
                Cells(i * 6 + 2, j).Value = Int(20 * Rnd + 20 * (j - 7) + 1)
                
                firstRow = Cells(i * 6 + 2, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Do While randomVal = firstRow
                    randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Loop
                Cells(i * 6 + 3, j).Value = randomVal
                
                secondRow = Cells(i * 6 + 3, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow
                    randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Loop
                Cells(i * 6 + 4, j).Value = randomVal
                
                thirdRow = Cells(i * 6 + 4, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow Or randomVal = thirdRow
                    randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Loop
                Cells(i * 6 + 5, j).Value = randomVal
                
                
                fourRow = Cells(i * 6 + 5, j).Value
                randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Do While randomVal = firstRow Or randomVal = secondRow Or randomVal = thirdRow Or randomVal = fourRow
                    randomVal = Int(20 * Rnd + 20 * (j - 7) + 1)
                Loop
                Cells(i * 6 + 6, j).Value = randomVal
            Next j
            
    Next i

End Sub
