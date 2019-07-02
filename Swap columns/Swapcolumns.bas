Sub Scaldolewejnormal()
'moves both colums data to left column (with ", " separator)

Dim selectionrow As String
Dim startingrow As Integer
Dim endingrow As Integer
Dim startingcolumn As String
Dim endingcolumn As String

selectionrow = Selection.Address

startingrow = Mid(selectionrow, 4, 2)
endingrow = Mid(selectionrow, 10, 2)
startingcolumn = Mid(selectionrow, 2, 1)
endingcolumn = Mid(selectionrow, 8, 1)

For i = startingrow To endingrow

Range(startingcolumn & i).Value = Range(startingcolumn & i).Value & ", " & Range(endingcolumn & i)
Range(endingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldolewejreverse()
'swap columns and moves both colums data to left column (with ", " separator)
Dim selectionrow As String
Dim startingrow As Integer
Dim endingrow As Integer
Dim startingcolumn As String
Dim endingcolumn As String

selectionrow = Selection.Address

startingrow = Mid(selectionrow, 4, 2)
endingrow = Mid(selectionrow, 10, 2)
startingcolumn = Mid(selectionrow, 2, 1)
endingcolumn = Mid(selectionrow, 8, 1)

For i = startingrow To endingrow

Range(startingcolumn & i).Value = Range(endingcolumn & i) & ", " & Range(startingcolumn & i).Value
Range(endingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldoprawejnormal()
'moves both colums data to right column (with ", " separator)
Dim selectionrow As String
Dim startingrow As Integer
Dim endingrow As Integer
Dim startingcolumn As String
Dim endingcolumn As String

selectionrow = Selection.Address

startingrow = Mid(selectionrow, 4, 2)
endingrow = Mid(selectionrow, 10, 2)
startingcolumn = Mid(selectionrow, 2, 1)
endingcolumn = Mid(selectionrow, 8, 1)

For i = startingrow To endingrow

Range(endingcolumn & i).Value = Range(startingcolumn & i).Value & ", " & Range(endingcolumn & i)
Range(startingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldoprawejreverse()
'swap columns and moves both colums data to right column (with ", " separator)
Dim selectionrow As String
Dim startingrow As Integer
Dim endingrow As Integer
Dim startingcolumn As String
Dim endingcolumn As String

selectionrow = Selection.Address

startingrow = Mid(selectionrow, 4, 2)
endingrow = Mid(selectionrow, 10, 2)
startingcolumn = Mid(selectionrow, 2, 1)
endingcolumn = Mid(selectionrow, 8, 1)

For i = startingrow To endingrow

Range(endingcolumn & i).Value = Range(endingcolumn & i) & ", " & Range(startingcolumn & i).Value
Range(startingcolumn & i).Value = ""
Next i

End Sub


Sub reversenumbername()
'swap places between two columns of selection

Dim selectionrow As String
Dim startingrow As Integer
Dim endingrow As Integer
Dim startingcolumn As String
Dim endingcolumn As String
Dim stringlewy As String
Dim stringprawy As String

selectionrow = Selection.Address

startingrow = Mid(selectionrow, 4, 2)
endingrow = Mid(selectionrow, 10, 2)
startingcolumn = Mid(selectionrow, 2, 1)
endingcolumn = Mid(selectionrow, 8, 1)

For i = startingrow To endingrow
stringlewy = Range(startingcolumn & i).Value
stringprawy = Range(endingcolumn & i).Value

Range(startingcolumn & i).Value = stringprawy
Range(endingcolumn & i).Value = stringlewy

Next i

End Sub