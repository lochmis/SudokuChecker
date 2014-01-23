Attribute VB_Name = "Module1"
'First to declare that a puzzle is complete. Should any empty cell or error occure, their value
'will change its value to False
Dim complete As Boolean
Dim error As Boolean
'folowing is for a switch to prevent messages
Dim answer As Integer



'Main function which will call all of the others to check the board
Sub checkBoard()
Attribute checkBoard.VB_ProcData.VB_Invoke_Func = "q\n14"

'we need to unlock a sheet for VBA to be able to make changes
ActiveSheet.Unprotect

complete = True
error = False
answer = 1


'subrutines of checking board to execute
isBoardComplete
checkRows
checkColumns
checkSubGrids
displayInformation

'and lock it again
ActiveSheet.Protect

End Sub
'Starting with checking if every cell have a number and if not, change background colour
Sub isBoardComplete()

error = False

Dim row As Integer
Dim col As Integer

'Using a loop to check each cell independentely
ActiveSheet.Range("B2:J10").Select
For col = 0 To Selection.Columns.Count - 1
    For row = 0 To Selection.Rows.Count - 1
        'If there is no value, background will be yellow
        If Selection.Offset(row, col).Range("A1").Value = 0 Then
            Selection.Offset(row, col).Range("A1").Interior.Color = RGB(255, 255, 102)
            'If there is a cell with no value, puzzle is not complete
            complete = False
        Else
            'If number was inputed, background collor changes to grey
            Selection.Offset(row, col).Range("A1").Interior.Color = RGB(235, 235, 235)
        End If
    Next row
Next col

End Sub
'Now to check each Row for duplicates within
Sub checkRows()

Dim row As Integer
Dim col As Integer
Dim base As Integer
Dim y As Integer

'Starting with simple loop within loop within loop :o)
ActiveSheet.Range("B2:J10").Select
For row = 0 To Selection.Rows.Count - 1
    For col = 0 To Selection.Columns.Count - 1
            'base is the input which we compare to other numbers if they are the same
            base = Selection.Offset(row, col).Range("A1").Value
        For y = 0 To Selection.Columns.Count - 1
            'If y equals to col, that would check the same cell against each other; therefore we
            'have to rule it out, also if base is 0 there is no point in checking
            'If there is a duplicate in the row, it will change their background to red
            If y <> col And base <> 0 And Selection.Offset(row, y).Range("A1").Value = base Then
                Selection.Offset(row, y).Range("A1").Interior.Color = RGB(255, 0, 0)
                Selection.Offset(row, col).Range("A1").Interior.Color = RGB(255, 0, 0)
                error = True
            End If
        Next y
    Next col
Next row

End Sub
'The same way we check for duplicates in columns
Sub checkColumns()

Dim row As Integer
Dim col As Integer
Dim base As Integer
Dim x As Integer

ActiveSheet.Range("B2:J10").Select
For col = 0 To Selection.Columns.Count - 1
    For row = 0 To Selection.Rows.Count - 1
        base = Selection.Offset(row, col).Range("A1").Value
        For x = 0 To Selection.Rows.Count - 1
            If x <> row And base <> 0 And Selection.Offset(x, col).Range("A1").Value = base Then
                Selection.Offset(x, col).Range("A1").Interior.Color = RGB(255, 0, 0)
                Selection.Offset(row, col).Range("A1").Interior.Color = RGB(255, 0, 0)
                error = True
            End If
        Next x
    Next row
Next col

End Sub

Sub checkSubGrids()

Dim grid As Integer
Dim row As Integer
Dim col As Integer
Dim base As Integer
Dim y As Integer
Dim x As Integer
Dim skipcell As Boolean

'Here we predefine the sub grids and use loop to switch between them
For grid = 1 To 9

    If grid = 1 Then
        ActiveSheet.Range("B2:D4").Select
    ElseIf grid = 2 Then
        ActiveSheet.Range("E2:G4").Select
    ElseIf grid = 3 Then
        ActiveSheet.Range("H2:J4").Select
    ElseIf grid = 4 Then
        ActiveSheet.Range("B5:D7").Select
    ElseIf grid = 5 Then
        ActiveSheet.Range("E5:G7").Select
    ElseIf grid = 6 Then
        ActiveSheet.Range("H5:J7").Select
    ElseIf grid = 7 Then
        ActiveSheet.Range("B8:D10").Select
    ElseIf grid = 8 Then
        ActiveSheet.Range("E8:G10").Select
    ElseIf grid = 9 Then
        ActiveSheet.Range("H8:J10").Select
    End If
    
    'Here we use basicly the same method like in checkColumn and Row but here we have to check rows and columns in subgrid
    For row = 0 To Selection.Rows.Count - 1
        For col = 0 To Selection.Columns.Count - 1
            base = Selection.Offset(row, col).Range("A1").Value
            For x = 0 To Selection.Rows.Count - 1
                For y = 0 To Selection.Columns.Count - 1
                    'we have to make sure that we are not comparing the same cell...
                    skipcell = False
                    If x = row And y = col Then
                        skipcell = True
                    End If
                    
                    If skipcell = False And base <> 0 And Selection.Offset(x, y).Range("A1").Value = base Then
                        Selection.Offset(x, y).Range("A1").Interior.Color = RGB(255, 0, 0)
                        Selection.Offset(row, col).Range("A1").Interior.Color = RGB(255, 0, 0)
                        error = True
                        
                    End If
                Next y
            Next x
        Next col
    Next row
Next grid

End Sub

Sub displayInformation()

'this little peace of code is just that after function, board will not end up selected
ActiveSheet.Range("A1").Select

'This switch is to show correct answer depending on which function war run
Select Case answer
    Case 1
        If error = True Then
            MsgBox "There seem to be some duplicate numbers in row, column and/or subgrid. Numbers has been highlighted"

        ElseIf complete = True Then
            ActiveSheet.Range("B2:J10").Interior.Color = RGB(198, 239, 206)
            MsgBox "Congratulation, You've solved this sudoku!"

        Else
            MsgBox "So far everything is in order, keep up the good work!"
        
        End If
        
    Case 2
        If complete = True Then
            ActiveSheet.Range("B2:J10").Interior.Color = RGB(235, 235, 235)
            MsgBox "Board is completely filled"

        Else
            MsgBox "Board is not complete. Empty cells are marked yellow"
        
        End If
        
    Case 3
        If error = True Then
            MsgBox "There seem to be some duplicate numbers in row(s). Numbers has been highlighted"
    
        Else
            MsgBox "So far all rows are in order, keep up the good work!"
        
        End If
        
    Case 4
        If error = True Then
            MsgBox "There seem to be some duplicate numbers in column(s). Numbers has been highlighted"
    
        Else
            MsgBox "So far all columns are in order, keep up the good work!"
        
        End If
        
    Case 5
        If error = True Then
            MsgBox "There seem to be some duplicate numbers in sub grid(s). Numbers has been highlighted"
    
        Else
            MsgBox "So far all sub grids are in order, keep up the good work!"
        
        End If
        
End Select

End Sub

'Now this might not be the best way to do this, but it is simple enough
Sub restartBoard()

'This piece of code I got from following website: http://www.ozgrid.com/forum/showthread.php?t=60440
     
    If MsgBox("Do you really want to restore board to its original setting?", vbYesNo, "Selection") = vbNo Then
        Exit Sub
    
    Else

        'There is basic board on the page where user would normaly not go (let's hope)
        'In case we want to restart the board, we copy the original one and replace the "used one"
        'Therefore we have a fresh start

        ActiveSheet.Unprotect
        ActiveSheet.Range("A100:I108").Select
        Selection.Copy
        ActiveSheet.Range("B2:J10").Select
        ActiveSheet.Paste
        ActiveSheet.Protect
        Application.CutCopyMode = True
        ActiveSheet.Range("A1").Select
    
        MsgBox "Board restored to normal"
        
    End If
    
End Sub
'Following subrutines are here for each button
Sub buttonComplete()

ActiveSheet.Unprotect

answer = 2
isBoardComplete
displayInformation

ActiveSheet.Protect

End Sub

Sub buttonRow()

ActiveSheet.Unprotect

answer = 3
'isBoardComplete is here to reset the background colours, so only errors from current function will be highlighted
isBoardComplete
checkRows
displayInformation

ActiveSheet.Protect

End Sub

Sub buttonColumn()

ActiveSheet.Unprotect

answer = 4
isBoardComplete
checkColumns
displayInformation

ActiveSheet.Protect

End Sub

Sub buttonGrid()

ActiveSheet.Unprotect

answer = 5
isBoardComplete
checkSubGrids
displayInformation

ActiveSheet.Protect

End Sub



