Attribute VB_Name = "MThee"
Option Explicit

Sub UpdateRekenbladThee()

    Dim iZetsystemen As Integer, i As Integer, rngInput As Range, c As Range, rngTheePakket As Range, clsThee As Rekenblad
    Dim iRow, dctThee As Dictionary
    
    
    If Range("_ptr.H.2eZetJN").Value = "Ja" Then iZetsystemen = 2 Else iZetsystemen = 1
    Set rngTheePakket = ThisWorkbook.Worksheets("Tabellen2").ListObjects("tblTheePakket").DataBodyRange
    Set dctThee = New Dictionary
    
    For i = 1 To iZetsystemen
        Set rngInput = Range("_rng." & i & "Z.InputThee")
        For Each c In rngInput
            If c.Value <> 0 Then
                With rngTheePakket
                    For iRow = 1 To .Rows.Count
                        If .Cells(iRow, 1) = c.Offset(, -1) Then
                            Set clsThee = New Rekenblad
                            clsThee.ArtikelNr = .Cells(iRow, 3)
                            clsThee.RegelType = "Thee | " & i
                            clsThee.Omschrijving = .Cells(iRow, 4)
                            clsThee.Drinks = c.Offset(, 1) * .Cells(iRow, 5)
                            clsThee.Korting1 = Range("_ptr.affactuurkortingThee").Value
                            If dctThee.Exists(clsThee.Omschrijving & i) Then
                                dctThee.Item(clsThee.Omschrijving & i).Drinks = dctThee.Item(clsThee.Omschrijving & i).Drinks + clsThee.Drinks
                            Else
                                dctThee.Add clsThee.Omschrijving & i, clsThee
                            End If
                        End If
                    Next iRow
                End With
            End If
        Next c
    Next i
    
    RefreshRekenbladRange "_rng.rb.Thee", DictionaryToCollection(dctThee)
    
End Sub
