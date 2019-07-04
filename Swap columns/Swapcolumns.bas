Public startingrow As Integer
Public endingrow As Integer
Public startingcolumn As String
Public endingcolumn As String

Sub readselectionrange()
'function to get range of selection and save it as usable variables for subs below
Dim selectionaddress As String
Dim selectionaddressplit As Variant

Dim firstcellletter As String
Dim secondcelletter As String
Dim firstcellnumber As Integer
Dim secondcellnumber As Integer

selectionaddress = Selection.Address                   'gets selection range
selectionaddress = Replace(selectionaddress, "$", "")  'delete $ from address

selectionaddressplit = Split(selectionaddress, ":")    'split range of starting cell [0] and ending cell [1]

firstcellletter = selectionaddressplit(0)              'define to use in for function
secondcellletter = selectionaddressplit(1)             'define to use in for function

For i = 0 To 9
    firstcellletter = Replace(firstcellletter, i, "")   'delete numbers from string
    secondcellletter = Replace(secondcellletter, i, "") 'delete numbers from string
Next i
startingcolumn = firstcellletter                        'assing extracted column letter to variable used in other subs
endingcolumn = secondcellletter                         'assing extracted column letter to variable used in other subs

startingrow = Replace(selectionaddressplit(0), firstcellletter, "") 'assing extracted row number to variable used in other subs
endingrow = Replace(selectionaddressplit(1), secondcellletter, "")  'assing extracted row number to variable used in other subs

End Sub


Sub Scaldolewejnormal()
'moves both colums data to left column (with ", " separator)

Call readselectionrange

For i = startingrow To endingrow

Range(startingcolumn & i).Value = Range(startingcolumn & i).Value & ", " & Range(endingcolumn & i)
Range(endingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldolewejreverse()
'swap columns and moves both colums data to left column (with ", " separator)
Call readselectionrange

For i = startingrow To endingrow

Range(startingcolumn & i).Value = Range(endingcolumn & i) & ", " & Range(startingcolumn & i).Value
Range(endingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldoprawejnormal()
'moves both colums data to right column (with ", " separator)
Call readselectionrange

For i = startingrow To endingrow

Range(endingcolumn & i).Value = Range(startingcolumn & i).Value & ", " & Range(endingcolumn & i)
Range(startingcolumn & i).Value = ""
Next i

End Sub

Sub Scaldoprawejreverse()
'swap columns and moves both colums data to right column (with ", " separator)
Call readselectionrange

For i = startingrow To endingrow

Range(endingcolumn & i).Value = Range(endingcolumn & i) & ", " & Range(startingcolumn & i).Value
Range(startingcolumn & i).Value = ""
Next i

End Sub


Sub reversenumbername()
'swap places between two columns of selection

Call readselectionrange

For i = startingrow To endingrow
stringlewy = Range(startingcolumn & i).Value
stringprawy = Range(endingcolumn & i).Value

Range(startingcolumn & i).Value = stringprawy
Range(endingcolumn & i).Value = stringlewy

Next i

End Sub
