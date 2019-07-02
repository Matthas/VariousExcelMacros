Attribute VB_Name = "AutoTripDistance"
Sub Bomfinderstart()

Dim postcoderange As String
Dim contractorrange As String
Dim i As Integer
Dim BOM As String
Dim bomdata As String
Dim countnumerrange As String
Dim contractorlistrange As String
Dim n As Integer
Dim m As Integer
Dim postcode As String
Dim contractor As String
Dim listnumber As Integer


n = 0
m = 0
contractorrange = "E5"  'range gdzie wybierany jest contractor
postcoderange = "D5"    'range gdzie wpisywany jest kod pocztowy
countnumberrange = "O1" 'range gdzie jest zliczana ilosc contractorow
bomdata = "BOMdata"  'nazwa arkusza gdzie przechowywane sa dane o contractorach
BOM = "BOM" 'nazwa arkusza w ktorym zostaja wyswietlone wyniki
contractorlistrange = "I" 'kolumna w ktorej jest lista contractorow
listnumber = Worksheets(bomdata).Range("F1") 'tutaj wpisac adres komorki gdzie znajduje sie zliczenie ilosci pozycji



If Worksheets(BOM).Range(postcoderange).Value = "" Then
    MsgBox ("Wpisz kod pocztowy")
Else
    If Worksheets(BOM).Range(contractorrange).Value = "" Then
        MsgBox ("Wpisz contractora")
        Else
        
        
            
        
            For i = 2 To Worksheets(bomdata).Range(countnumberrange).Value   'from first to last unique contractor
                If Worksheets(BOM).Range(contractorrange) = Worksheets(bomdata).Range(contractorlistrange & i) Then  'if there is contractor match
                    contractor = Worksheets(BOM).Range(contractorrange).Value         'save contractor name
                    postcode = Worksheets(BOM).Range(postcoderange).Value             'save postcode
                    Call Bomfinderdistance(contractor, postcode, i, contractorlistrange, listnumber, bomdata, BOM, postcoderange)     'send all data to next sub (clarity reasons)
                    m = 1
                    Exit For
                Else
                n = 1
                End If
            
            Next i
            If m = 1 Then
            'all good, operation Finished successfully
            Else
                If n = 1 Then
                    MsgBox ("Nie ma takiego contractora")
                End If
            End If
        End If
End If
    
           
End Sub

Sub Bomfinderdistance(contractor As String, postcode As String, contractorrownumber As Integer, contractorlistrange As String, listnumber As Integer, bomdata As String, _
BOM As String, postcoderange As String)

Dim rangestartrow As Integer
Dim rangestartcolumn As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim minval As Integer
Dim minvalrange As String
Dim contractorlistcolumn As String
Dim rangestartcolumnoffset As String
Dim myformula As String
Dim bestcontractorrange As String

contractorlistcolumn = "A" 'column where contractor names are located; default should be A
rangestartrow = 6
rangestartcolumn = "H"
rangeendcolumn = "J"
rangestartcolumnoffset = "I"
bestcontractorrange = "C9"

Worksheets(BOM).Range(rangestartcolumn & rangestartrow & ":" & "J" & (rangestartrow + 50)).Clear   'clear cells from starting cell H6 to J50 (in default scenario)


For k = 2 To listnumber
    If contractor = Worksheets(bomdata).Range(contractorlistcolumn & k) Then
        Exit For
    End If

Next k


For j = 2 To listnumber
    If Worksheets(bomdata).Range(contractorlistcolumn & j).Value = contractor Then
        For i = 0 To (Worksheets(bomdata).Range(contractorlistrange & contractorrownumber).Offset(0, 1).Value - 1) 'loop from 0 to (load number of BOM locations for this contractor)
            'create list of matching BOMs for selected contractor
            Worksheets(BOM).Range(rangestartcolumn & (rangestartrow + i)).Value = Worksheets(bomdata).Range(contractorlistcolumn & (k + i)).Offset(0, 1).Value & " - " & _
            Worksheets(bomdata).Range(contractorlistcolumn & (k + i)).Offset(0, 2).Value
            'Name - location
            'same as above but with postcodes
            Worksheets(BOM).Range(rangestartcolumn & (rangestartrow + i)).Offset(0, 1).Value = Worksheets(bomdata).Range(contractorlistcolumn & (k + i)).Offset(0, 3).Value
            '
            'insert formula to calculate distance
           
           Worksheets(BOM).Range(rangestartcolumn & (rangestartrow + i)).Offset(0, 2).formula = "=TripDistance(" & rangestartcolumnoffset & rangestartrow + i & "," & postcoderange & ")"
            
        
        Next i
    End If
Next j

minval = 9999
For i = 0 To (Worksheets(bomdata).Range(contractorlistrange & contractorrownumber).Offset(0, 1).Value - 1)

    If Worksheets(BOM).Range(rangeendcolumn & rangestartrow + i).Value < minval Then
        minval = Worksheets(BOM).Range(rangeendcolumn & rangestartrow + i).Value
        minvalrange = rangeendcolumn & (rangestartrow + i)
End If
Next i

Worksheets(BOM).Range(bestcontractorrange).Value = Worksheets(BOM).Range(minvalrange).Offset(0, -2).Value

End Sub


