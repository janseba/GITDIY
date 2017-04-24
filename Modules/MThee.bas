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
                Set clsThee = New Rekenblad
                clsThee.ArtikelNr = GetTheeArtikelNr(rngTheePakket, c.Offset(, -1))
                clsThee.RegelType = "Thee | " & i
                clsThee.Omschrijving = c.Offset(, -1)
                clsThee.Drinks = c.Offset(, 1)
                clsThee.Korting1 = Range("_ptr.affactuurkortingThee").Value
                If dctThee.Exists(clsThee.Omschrijving & i) Then
                    dctThee.Item(clsThee.Omschrijving & i).Drinks = dctThee.Item(clsThee.Omschrijving & i).Drinks + clsThee.Drinks
                Else
                    dctThee.Add clsThee.Omschrijving & i, clsThee
                End If
            End If
        Next c
    Next i
    
    RefreshRekenbladRange "_rng.rb.Thee", DictionaryToCollection(dctThee)
    
End Sub

Function GetTheeArtikelNr(ByRef rngTheePakket As Range, ByVal sPakket As String) As String

    Dim iRow As Integer, maxPerc As Double, artNr As String
    
    maxPerc = 0
    With rngTheePakket
        For iRow = 1 To .Rows.Count
            If .Cells(iRow, 1) = sPakket Then
                If .Cells(iRow, 5) > maxPerc Then
                    maxPerc = .Cells(iRow, 5)
                    artNr = .Cells(iRow, 3)
                End If
            End If
        Next iRow
    End With
    
    GetTheeArtikelNr = artNr
End Function
