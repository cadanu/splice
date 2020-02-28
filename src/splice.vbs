Option Explicit

'macro that evaluates, cleans and concatenates the contents
'of an array of cells.
Sub splice()

'variable declarations
Dim row: row = 4
Dim cols(7)
Dim count
Dim index
Dim stren
Dim garbage

'loops to the end of file
While row <= 3936
    
    count = 0
    index = 0
    
    'input array values according to 'row' value
    cols(0) = Cells(row, 86)
    cols(1) = Cells(row, 87)
    cols(2) = Cells(row, 88)
    cols(3) = Cells(row, 89)
    cols(4) = Cells(row, 90)
    cols(5) = Cells(row, 91)
    cols(6) = Cells(row, 92)
    cols(7) = Cells(row, 93)
    
    'assigns empty value to duplicate cell values
    For ia = 0 To 7
        'index = i
        For ib = 0 To 7
            If ia <> ib Then
                If cols(ia) = cols(ib) Then
                    cols(ib) = ""
                End If
            End If
        Next
    Next
    
    'creates string from 'cols' array
    For Each k In cols
    
        'k = index
        'if loop will make index value empty if it matches 'k'
        'If k <> index Then
        '    If k = index Then
        '        k = ""
        '    End If
        'End If
        'disposes of empty indexes
        If k = "" Then
            garbage = garbage & k
        'assigns index value to 'stren' variable
        Else
            stren = stren & k & ";"
        End If
    Next
    
    'places value in 'stren' in selected cell
    ActiveCell = stren
    
    'changes selected cell (down 1)
    ActiveCell.Offset(1, 0).Select
    
    'increments row for array input at top
    row = row + 1
    
    'empties variables
    stren = ""
    garbage = ""

Wend

    'ActiveCell.Offset(1,0)
    'ActiveCell = stren


    'MsgBox "Hello World"
    'Cells(4, 96) = 25
    'MsgBox stren

End Sub
