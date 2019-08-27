Sub AnI()
'===========zliczanie ulosci uzytkownikow
userscount = 0
For i = 20 To 200
    If Range("A" & i).Value = "" Then   'jesli puste pole to idz dalej
    Else
        userscount = userscount + 1  'zliaczanei ilosci uzytkownikow
    End If
'nastepna komorka
Next i

'==========zliczanie ilosci kompow
pccount = 0
For i = 0 To 500
    If Range("I3").Offset(0, i).Value = "" Then 'jesli puste pole to idz dalej.  Offset bo tak latwiej chodzic po kolumnach.
             'zliczamy po wierszu 3 bo w wierszu drugim sa puste pole, ktore zepsuly by obliczenia
    Else
        pccount = pccount + 1       'zliczanie ilosci kompow
    End If
'nastepna komorka
Next i


Rows("2").MergeCells = False        'usuniecie scalenia komorek

'wypelnienie pustych pol inicjalami ludzi

For i = 0 To (pccount - 1)  'jeden mniej bo tak wyjdzie ze podwoi jedno nazwisko na samym koncu listy
    If Range("I2").Offset(0, i).Value = "" Then 'jezeli puste pole zobacz pole obok, jezeli nie jest puste przekopiuj wartosc z pola obok
        Range("I2").Offset(0, i).Value = Range("I2").Offset(0, i - 1).Value 'przepisz wartosc z komorki z lewej strony
    Else
        'jezeli komorka nie jest pusta, to pomin
    End If
'nastepna komorka
Next i

Range("A20:A" & (usercount + 20)).Interior.Color = xlNone

MsgBox (userscount & " + " & pccount)
For j = 20 To (20 + userscount) ' +20 bo zaczynamy od wiersza 20
    For i = 1 To pccount
        If Range("I2").Offset(0, i - 1).Value = Range("A" & j).Value Then  'jezeli nazwa kompa zgadza sie z aktualnie przeszukana nazwa uzytkownika to
            'sprawdz czy uzywa kompa
            If Range("I3").Offset(0, i - 1).Value > 0 Then 'sprawdz czy wartosc jest wieksza od zera (wiersz 3)
                'jezeli jest aktywny to wpisz w kolumnie C active i usun kolor z kolumny A
                Range("C" & j).Value = "Active"
                Range("A" & j).Interior.Color = xlNone
            Else
                'jezeli nie to podswietl na zolto
                If Range("C" & j).Value = "Active" Then 'sprawdz czy juz nie byl aktywny w innej kolumnie
                Else 'jak nie byl to koloruj
                    Range("A" & j).Interior.Color = RGB(255, 255, 0) 'zolty
                End If
            End If
        End If
    Next i
Next j

For j = 20 To (20 + userscount) ' +20 bo zaczynamy od wiersza 20
    For i = 1 To pccount
        If Range("I2").Offset(0, i - 1).Value = Range("A" & j).Value Then  'jezeli nazwa kompa zgadza sie z aktualnie przeszukana nazwa uzytkownika to
            'sprawdz czy uzywa kompa
            If Range("B" & j).Value = "x" Then
                If Range("I3").Offset(0, i - 1).Value > 0 Then 'sprawdz czy wartosc jest wieksza od zera (wiersz 3)
                    'jezeli jest aktywny to wpisz w kolumnie C active i usun kolor z kolumny A
                    Range("C" & j).Value = "intruder"
                    Range("B" & j).Interior.Color = RGB(255, 0, 0)
                Else
                    'jezeli nie to podswietl na zolto
                    If Range("C" & j).Value = "intruder" Then 'sprawdz czy juz nie byl aktywny w innej kolumnie
                    Else 'jak nie byl to koloruj
                        Range("B" & j).Interior.Color = xlNone 'gumka
                    End If
                End If
            End If
        End If
    Next i
Next j



End Sub