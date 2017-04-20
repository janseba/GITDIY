Attribute VB_Name = "MOutput2"
Option Explicit
Sub GenerateOutput2()
    SetPerformance True
    UpdateLocatieOverzicht2
    UpdateApparatuurOverzicht2
    UpdateServiceOverzicht2
    UpdateEenmaligeKosten2
    'UpdateOperatingOverzicht
    UpdateBonusTabel2
    'UpdateIngKortingen1
    'UpdateIngKortingen2
    UpdateBonusstaffel2 1
    UpdateBonusstaffel2 2
    UpdateContractgegevens2
    UpdateRetouren2
    UpdateCoderingen2
    UpdateBijlagen2
    UpdateHoodlocatieBonus2
    UpdateBonusConditie2
    UpdateMalusConditie2
    UpdateOpbouwOfferteContract2
    UpdateCheckDeal2 OK:=IIf(ThisWorkbook.Names("_ptr.Dealoordeel").RefersToRange.Value = 2, True, False)
    UpdateKoffie2
    UpdateServies
    UpdateServiesBudget
    UpdatePos
    SetPerformance False
End Sub
Sub UpdateLocatieOverzicht2()
    Dim aantalLocaties As Integer, rngLocaties As Range, rngLocatieOverzicht As Range, i As Integer
    aantalLocaties = Range("_ptr.AantalLocaties")
    Set rngLocaties = Range("_rng.Locaties")
    Set rngLocatieOverzicht = Range("_out.Locatieoverzicht")
    If rngLocatieOverzicht.Rows.Count > 1 Then rngLocatieOverzicht.Offset(1).Resize(rngLocatieOverzicht.Rows.Count - 1).ClearContents
    If rngLocatieOverzicht.Rows.Count - 1 < aantalLocaties Then '-1 ivm headerrow
        InsertRowsInRange aantalLocaties - (rngLocatieOverzicht.Rows.Count - 1), "_out.Locatieoverzicht"
    ElseIf rngLocatieOverzicht.Rows.Count - 1 > aantalLocaties Then
        DeleteRowsFromRange rngLocatieOverzicht.Rows.Count - 1 - aantalLocaties, "_out.Locatieoverzicht"
    End If
    For i = 1 To aantalLocaties
        rngLocatieOverzicht(i + 1, 1) = rngLocaties.Cells(i, 1) 'Klantnummer
        rngLocatieOverzicht(i + 1, 2) = rngLocaties.Cells(i, 2) 'Klantnaam
        rngLocatieOverzicht(i + 1, 3) = rngLocaties.Cells(i, 5) 'Adres
        rngLocatieOverzicht(i + 1, 4) = rngLocaties.Cells(i, 8) 'Plaats
        rngLocatieOverzicht(i + 1, 5) = GetAantalMachines2(rngLocaties.Cells(i, 1)) 'AantalMachines
        rngLocatieOverzicht(i + 1, 6) = rngLocaties.Cells(i, 12) 'Weeknummer
        rngLocatieOverzicht(i + 1, 7) = rngLocaties.Cells(i, 11) 'MaandagWeeknummer
        rngLocatieOverzicht(i + 1, 8) = rngLocaties.Cells(i, 13) 'Contactpersoon
        rngLocatieOverzicht(i + 1, 9) = rngLocaties.Cells(i, 14) 'TelefoonnummerCp
        rngLocatieOverzicht(i + 1, 10) = rngLocaties.Cells(i, 15) 'Openingstijden
        rngLocatieOverzicht(i + 1, 11) = rngLocaties.Cells(i, 16) 'Operating
    Next i
    
End Sub
Function GetAantalMachines2(ByVal KlantNummer As String) As String
    Dim rngLocatieDetails As Range, row As Integer, dctMachines As Dictionary, result As String
    Set rngLocatieDetails = Range("_rng.locatiedetails")
    Set dctMachines = New Dictionary
    With rngLocatieDetails
        For row = 1 To .Rows.Count
            If .Cells(row, 1) = KlantNummer Then
                If Not dctMachines.Exists(.Cells(row, 7).Value) Then
                    dctMachines.Add .Cells(row, 7).Value, 1
                Else
                    dctMachines.Item(.Cells(row, 7).Value) = dctMachines.Item(.Cells(row, 7).Value) + 1
                End If
            End If
        Next row
    End With
    
    Dim k As Variant
    For Each k In dctMachines.Keys
        result = result & "|" & k & " (" & dctMachines.Item(k) & ")"
    Next k
    If dctMachines.Count > 0 Then GetAantalMachines2 = Right(result, Len(result) - 1) Else GetAantalMachines2 = "0"
End Function
Sub UpdateApparatuurOverzicht2()
    Dim machine As Integer, inputRange As Range, apparatuur As ApparatuurOverzicht, dctApparatuur As Dictionary, iContractType As Integer
    Dim j As Integer, zetSysteem As Integer, contractType As String, regelNr As Variant, r As Variant, p As Integer, iBrutoPrijs As Integer
    Dim iKorting As Integer
    Set dctApparatuur = New Dictionary
    contractType = Range("_ptr.Contracttype").Value
    If contractType = "Koop" Then
        iBrutoPrijs = 3: iKorting = 4
    ElseIf contractType = "Huur" Then
        iBrutoPrijs = 9: iKorting = 10
    ElseIf contractType = "Bruikleen" Then
        iBrutoPrijs = 12: iKorting = 13
    End If
    regelNr = Array(1, 2, 5, 6, 7, 8, 9)
    'rng.1Z.app.koop.M1.input
    For zetSysteem = 1 To 2
        For machine = 1 To 3
            Set inputRange = Names("_rng." & zetSysteem & "Z.M" & machine).RefersToRange
            With inputRange
                If .Cells(1, 2) > 0 Then 'Er is een machine met aantal > 0
                    p = 1
                    For Each r In regelNr
                        If .Cells(r, 2) > 0 Then 'Aantal > 0
                            Set apparatuur = New ApparatuurOverzicht
                            apparatuur.RegelNrTb = 200 + ((zetSysteem - 1) * 30 + machine * 10) + p
                            apparatuur.MachinetypeAccessoires = .Cells(r, 1) 'MachineType
                            apparatuur.Aantal = .Cells(r, 2) 'Aantal
                            apparatuur.Koopprijs = .Cells(r, iBrutoPrijs) / apparatuur.Aantal 'BrutoPrijs
                            apparatuur.KortingKoop = .Cells(r, iKorting) 'Korting%
                            dctApparatuur.Add apparatuur.RegelNrTb, apparatuur
                            p = p + 1
                        End If
                    Next r
                End If
            End With
        Next machine
    Next zetSysteem
    AdjustRange dctApparatuur, "_out.ApparatuurOverzicht"
    
    Dim rngApparatuurOverzicht As Range, k As Variant
    Set rngApparatuurOverzicht = Names("_out.ApparatuurOverzicht").RefersToRange
    machine = 1
    With rngApparatuurOverzicht
        For Each k In dctApparatuur
            machine = machine + 1
            .Cells(machine, 1) = dctApparatuur(k).RegelNrTb
            .Cells(machine, 2) = dctApparatuur(k).MachinetypeAccessoires
            .Cells(machine, 3) = dctApparatuur(k).Aantal
            .Cells(machine, 4) = dctApparatuur(k).Koopprijs
            .Cells(machine, 5) = dctApparatuur(k).KortingKoop
        Next k
    End With
End Sub

Sub UpdateServiceOverzicht2()
    Dim machine As Integer, zetSysteem As Integer, contractType As String, inputRange As Range, service As ServiceOverzicht
    Dim dctService As Dictionary, iServices As Integer, iBrutoPrijs As Integer, iKorting As Integer
    Set dctService = New Dictionary
    contractType = Range("_ptr.Contracttype").Value
    If contractType = "Koop" Then
        iBrutoPrijs = 6: iKorting = 7
    ElseIf contractType = "Huur" Then
        iBrutoPrijs = 9: iKorting = 10
    ElseIf contractType = "Bruikleen" Then
        iBrutoPrijs = 12: iKorting = 13
    End If
    
    Dim iRow As Long
    For zetSysteem = 1 To 2
        For machine = 1 To 3
            iServices = 0
            Set inputRange = Names("_rng." & zetSysteem & "Z.M" & machine).RefersToRange
            'BSP (moet alleen berekend worden bij koop en bruikleen, zit in huurprijs)
            If contractType <> "Huur" Then
                With inputRange
                    For iRow = 1 To .Rows.Count
                        If .Cells(iRow, iBrutoPrijs) <> 0 And InStr("Onderzetkast,Aanvullende servicemodule,eenmalig", .Cells(iRow, 1).Offset(, -1)) = 0 Then
                            Set service = New ServiceOverzicht
                            iServices = iServices + 1
                            service.RegelNrTbl = 300 + ((zetSysteem - 1) * 30 + machine * 10) + iServices
                            If iRow = 4 Then service.MachinetypeAccessoires = .Cells(iRow, 1) Else service.MachinetypeAccessoires = "BSP " & .Cells(iRow, 1)
                            service.Aantal = .Cells(iRow, 2)
                            service.BSPPrijsKoop = .Cells(iRow, iBrutoPrijs) / service.Aantal
                            service.KortingKoop = .Cells(iRow, iKorting)
                            dctService.Add service.RegelNrTbl, service
                        End If
                    Next iRow
                End With
            End If
            'Aanvullende service
            With inputRange
                If .Cells(4, iBrutoPrijs) > 0 Then
                    Set service = New ServiceOverzicht
                    iServices = iServices + 1
                    service.RegelNrTbl = 300 + ((zetSysteem - 1) * 30 + machine * 10) + iServices
                    service.MachinetypeAccessoires = .Cells(4, 1)
                    service.Aantal = .Cells(4, 2)
                    service.BSPPrijsKoop = .Cells(4, iBrutoPrijs)
                    service.KortingKoop = .Cells(4, iKorting)
                    dctService.Add service.RegelNrTbl, service
                End If
            End With
        Next machine
    Next zetSysteem
    AdjustRange dctService, "_out.ServiceOverzicht"
    
    Dim rngserviceOverzicht As Range, k As Variant
    Set rngserviceOverzicht = Names("_out.ServiceOverzicht").RefersToRange
    machine = 1
    With rngserviceOverzicht
        For Each k In dctService
            machine = machine + 1
            .Cells(machine, 1) = dctService(k).RegelNrTbl
            .Cells(machine, 2) = dctService(k).MachinetypeAccessoires
            .Cells(machine, 3) = dctService(k).Aantal
            .Cells(machine, 4) = dctService(k).BSPPrijsKoop
            .Cells(machine, 5) = dctService(k).KortingKoop
        Next k
    End With
End Sub

Sub UpdateEenmaligeKosten2()
    Dim machine As Integer, zetSysteem As Integer, contractType As String, inputRange As Range, Kosten As EenmaligeKosten, dctKosten As Dictionary
    Dim i As Integer, regels As Variant, regel As Variant, sRecyclingbijdrage As String, dblRecBijdrage As Double, iBrutoPrijs As Integer
    Dim iKorting As Integer, vArtikel As Variant
    Set dctKosten = New Dictionary
    contractType = Range("_ptr.Contracttype").Value
    
    If contractType = "Koop" Then
        iBrutoPrijs = 3: iKorting = 4
    ElseIf contractType = "Huur" Then
        iBrutoPrijs = 9: iKorting = 10
    ElseIf contractType = "Bruikleen" Then
        iBrutoPrijs = 12: iKorting = 13
    End If
    
    regels = Array(1, 3, 5, 6, 7, 8, 9, 10, 11, 12)
    
    i = 0
    ' Get recyclingbijdrage kolom 25 (omschrijving) en 26 (bedrag)
    For zetSysteem = 1 To 2
        For machine = 1 To 3
            Set inputRange = Names("_rng." & zetSysteem & "Z.M" & machine).RefersToRange
            i = 0
            For Each regel In regels
                With inputRange
                    If InStr("|3|10|11|", "|" & regel & "|") = 0 Then '1 en 5 t/m 9 betreft recyclingbijdrage
                        If contractType = "Koop" Then 'recyclingbijdrage alleen meenemen bij koop
                            vArtikel = .Cells(regel, 1).Offset(, -3)
                            If Not IsError(vArtikel) Then
                                sRecyclingbijdrage = Lookup(Range("tblMachines[Omschrijving Recyclingbijdrage]"), .Cells(regel, 1).Offset(, -3), Range("tblMachines[Calc Nr]"))
                                dblRecBijdrage = Lookup(Range("tblMachines[RecBijdrage]"), .Cells(regel, 1).Offset(, -3), Range("tblMachines[Calc Nr]"))
                            End If
                            If dblRecBijdrage > 0 And Not IsError(vArtikel) Then
                                i = i + 1
                                Set Kosten = New EenmaligeKosten
                                Kosten.RegelNrTbl = 400 + ((zetSysteem - 1) * 30 + machine * 10) + i
                                Kosten.Omschrijving = sRecyclingbijdrage
                                Kosten.Aantal = .Cells(regel, 2)
                                Kosten.PrijsPerStuk = dblRecBijdrage
                                Kosten.Korting = 0
                                dctKosten.Add Kosten.RegelNrTbl, Kosten
                            End If
                        End If
                    Else '3, 11 en 12 zijn koop items en horen ook bij eenmalige kosten thuis
                        If .Cells(regel, iBrutoPrijs) > 0 And .Cells(regel, 2) > 0 Then
                            i = i + 1
                            Set Kosten = New EenmaligeKosten
                            Kosten.RegelNrTbl = 400 + ((zetSysteem - 1) * 30 + machine * 10) + i
                            Kosten.Omschrijving = .Cells(regel, 1)
                            Kosten.Aantal = .Cells(regel, 2)
                            Kosten.PrijsPerStuk = .Cells(regel, iBrutoPrijs) / Kosten.Aantal
                            Kosten.Korting = .Cells(regel, iKorting)
                            dctKosten.Add Kosten.RegelNrTbl, Kosten
                        End If
                    End If
                End With
            Next regel
        Next machine
    Next zetSysteem
    
    ' Get aflever- en installatiekosten
    Set inputRange = Range("_rng.Installatiekosten")
    Set Kosten = New EenmaligeKosten
    With inputRange
        Kosten.RegelNrTbl = 471
        Kosten.Omschrijving = "Aflever- en installatiekosten"
        Kosten.Aantal = .Cells(1, 1)
        Kosten.PrijsPerStuk = .Cells(2, 1)
        Kosten.Korting = Range("_ptr.KortingInstallatie")
    End With
    dctKosten.Add Kosten.RegelNrTbl, Kosten
    
    AdjustRange dctKosten, "_out.EenmaligeKosten"
    
    Dim rngEenmaligeKosten As Range, k As Variant
    Set rngEenmaligeKosten = Range("_out.EenmaligeKosten")
    i = 1
    With rngEenmaligeKosten
        For Each k In dctKosten
            i = i + 1
            .Cells(i, 1) = dctKosten(k).RegelNrTbl
            .Cells(i, 2) = dctKosten(k).Omschrijving
            .Cells(i, 3) = dctKosten(k).Aantal
            .Cells(i, 4) = dctKosten(k).PrijsPerStuk
            .Cells(i, 5) = dctKosten(k).Korting
        Next k
    End With
    
End Sub

Sub UpdateBonusTabel2()
    Dim i As Integer, rngVertaaltabel As Range, j As Integer
    Set rngVertaaltabel = Names("tbl.VertaaltabelIngredienten").RefersToRange
    
    Dim index As Integer, rngFound As Range, bonusRegel As BonusTabel, aantalThee As Integer, zetSysteem As Integer
    Dim dctBonusTabel As Dictionary, c As Range, rngThee As Range, rngCS As Range
    Set dctBonusTabel = New Dictionary
    
    ' Get thee
    aantalThee = 0
    Set rngThee = Range("_rng.rb.Thee")
    With rngThee
    For i = 1 To .Rows.Count
        If .Cells(i, 4) > 0 And Range("_ptr.BonusThee").Value > 0 Then
            aantalThee = aantalThee + 1
            Set bonusRegel = New BonusTabel
            bonusRegel.RegelNrTbl = 6100 + aantalThee
            bonusRegel.ArtikelOmschrijving = .Cells(i, 3)
            bonusRegel.bonus = Range("_ptr.BonusThee").Value
            bonusRegel.Afname = .Cells(i, 1).Offset(, -4)
            dctBonusTabel.Add bonusRegel.RegelNrTbl, bonusRegel
        End If
    Next i
    End With
    
    ' Get CS
    Set rngCS = Range("_rng.rb.CS")
    With rngCS
        For i = 1 To .Rows.Count
            If .Cells(i, 4) > 0 And Range("_ptr.BonusCS").Value > 0 Then
               j = j + 1
               Set bonusRegel = New BonusTabel
               bonusRegel.RegelNrTbl = 6200 + j
               bonusRegel.ArtikelOmschrijving = .Cells(i, 3)
               bonusRegel.bonus = Range("_ptr.BonusCS").Value
                   bonusRegel.Afname = .Cells(i, 1).Offset(, -2)
               dctBonusTabel.Add bonusRegel.RegelNrTbl, bonusRegel
            End If
        Next i
    End With
    
    AdjustRange dctBonusTabel, "_out.BonusTabel"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.BonusTabel").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctBonusTabel
            i = i + 1
            .Cells(i, 1) = dctBonusTabel(k).RegelNrTbl
            .Cells(i, 2) = dctBonusTabel(k).ArtikelOmschrijving
            .Cells(i, 3) = dctBonusTabel(k).bonus
            .Cells(i, 4) = dctBonusTabel(k).Afname
        Next k
    End With
    
End Sub
Sub UpdateBonusstaffel2(ByVal zetSysteem As Integer)
    Dim rngBonus As Range, bonusJN As String, bonusBedrag As Double, i As Integer, dctBonusTabel As Dictionary
    Dim bonusRegel As New BonusregelKoffie, sEenheid As String
    
    bonusJN = Range("_ptr." & zetSysteem & "Z.BonusKoffieJN").Value
    bonusBedrag = Range("_ptr." & zetSysteem & "Z.BonusBasistredeEuro").Value
    sEenheid = Range("_ptr." & zetSysteem & "Z.Eenheid")
    
    Set dctBonusTabel = New Dictionary
    Set rngBonus = Range("_rng." & zetSysteem & "Z.Bonustabel")
    
    If bonusJN = "Ja" And bonusBedrag > 0 Then
        Set bonusRegel = New BonusregelKoffie
        'Check of 1e en 2e regel allebei 0 zijn
        If rngBonus.Cells(1, 3) = 0 And rngBonus.Cells(2, 3) = 0 Then 'als beide lijnen 0 zijn dan mogen ze worden samengevoegd.
            'schuif tabel 1 rij naar beneden (om eerste regel te skippen) en maak hem 1 rij kleiner
            Set rngBonus = rngBonus.Offset(1).Resize(rngBonus.Rows.Count - 1)
        End If
        With rngBonus
            For i = 1 To rngBonus.Rows.Count
                Set bonusRegel = New BonusregelKoffie
                bonusRegel.RegelNrTbl = (zetSysteem + 8) * 100 + i
                If i = 1 Then '1e regel
                    bonusRegel.AfgenomenVolumePerJaar = "Tot " & Format(.Cells(i, 2), "#,##0") & " " & sEenheid
                ElseIf i = .Rows.Count Then 'laatste regel
                    bonusRegel.AfgenomenVolumePerJaar = "Vanaf " & Format(.Cells(i, 1), "#,##0") & " " & sEenheid
                Else 'tussenliggende regels
                    bonusRegel.AfgenomenVolumePerJaar = "Van " & Format(.Cells(i, 1), "#,##0") & " tot " & Format(.Cells(i, 2), "#,##0") & " " & sEenheid
                End If
                bonusRegel.Bedrag = .Cells(i, 3)
                bonusRegel.Eenheid = sEenheid
                dctBonusTabel.Add bonusRegel.RegelNrTbl, bonusRegel
            Next i
        End With
    ElseIf bonusJN <> "Ja" And Range("_ptr." & zetSysteem & "Z.MalusJN").Value = "Ja" Then 'geen bonus wel malus
        Set bonusRegel = New BonusregelKoffie
        With rngBonus
            Set bonusRegel = New BonusregelKoffie
            bonusRegel.RegelNrTbl = (zetSysteem + 8) * 100 + 1
            bonusRegel.AfgenomenVolumePerJaar = "Tot " & Format(.Cells(1, 2), "#,##0") & " " & sEenheid
            bonusRegel.Bedrag = .Cells(1, 3)
            bonusRegel.Eenheid = sEenheid
            dctBonusTabel.Add bonusRegel.RegelNrTbl, bonusRegel
        End With
    End If
    AdjustRange dctBonusTabel, "_out.Bonusstaffel" & zetSysteem
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.Bonusstaffel" & zetSysteem).RefersToRange
    i = 1
    With rngOutput
        For Each k In dctBonusTabel
            i = i + 1
            .Cells(i, 1) = dctBonusTabel(k).RegelNrTbl
            .Cells(i, 2) = dctBonusTabel(k).AfgenomenVolumePerJaar
            .Cells(i, 3) = dctBonusTabel(k).Bedrag
            .Cells(i, 4) = dctBonusTabel(k).Eenheid
        Next k
    End With

End Sub


Sub UpdateContractgegevens2()
    Dim outputRange As Range
    Set outputRange = Range("_out.ContractGegevens")
    
    With outputRange
        .Cells(2, 1) = 1101 'RegelNrTbl
        .Cells(2, 2) = "Klantnaam"
        .Cells(2, 3) = Range("_ptr.Klantnaam").Value
        .Cells(3, 1) = 1102
        .Cells(3, 2) = "Voorstelnummer"
        .Cells(3, 3) = Format(Range("_ptr.Calculatiedatum"), "yy-mmdd") & "." & Range("_ptr.ReferentienummerSAP")
        .Cells(4, 1) = 1103
        .Cells(4, 2) = "Datum"
        .Cells(4, 3) = Range("_ptr.Calculatiedatum").Value
        .Cells(5, 1) = 1104
        .Cells(5, 2) = "Referentienummer"
        .Cells(5, 3) = Range("_ptr.ReferentienummerSAP").Value
    End With

End Sub
Sub UpdateRetouren2()
    Dim rngKlantnummer As Range, c As Range, retour As RetourRegel, i As Integer, j As Integer, dctRetouren As Dictionary
    Set rngKlantnummer = Range("_rng.Retouren").Resize(, 1)
    Set dctRetouren = New Dictionary
    i = 0: j = 0
    For Each c In rngKlantnummer
        i = i + 1
        If Not IsEmpty(c) Then
            j = j + 1
            Set retour = New RetourRegel
            retour.RegelNrTb = 1200 + j
            retour.MachineType = c.Offset(, 2)
            retour.MachineNummer = c.Offset(, 3)
            retour.KlantNummer = c.Value
            dctRetouren.Add retour.RegelNrTb, retour
        End If
    Next c
    
    AdjustRange dctRetouren, "_out.Retouren"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.Retouren").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctRetouren
            i = i + 1
            .Cells(i, 1) = dctRetouren(k).RegelNrTb
            .Cells(i, 2) = dctRetouren(k).MachineType
            .Cells(i, 3) = dctRetouren(k).MachineNummer
            .Cells(i, 4) = dctRetouren(k).KlantNummer
        Next k
    End With
End Sub

Sub UpdateCoderingen2()
    Dim i As Integer, rngOmschrijvingVeld As Range, c As Range, codering As CoderingRegel, dctCodering As Dictionary
    Set rngOmschrijvingVeld = Range("tblCoderingen[OmschrijvingVeld]")
    Set dctCodering = New Dictionary
    i = 0
    For Each c In rngOmschrijvingVeld
        i = i + 1
        Set codering = New CoderingRegel
        codering.RegelNrTbl = 1300 + i
        codering.OmschrijvingVeld = c.Value
        codering.KomtVoorInTekstblok = c.Offset(, 1).Value
        codering.Code = c.Offset(, 2).Value
        If codering.RegelNrTbl = 1321 Then
            codering.Waarde = GetKoffieType2(1)
        ElseIf codering.RegelNrTbl = 1322 Then
            codering.Waarde = GetKoffieType2(2)
        Else
            codering.Waarde = CStr(c.Offset(, 3).Value)
        End If
        dctCodering.Add codering.RegelNrTbl, codering
    Next c
    
    AdjustRange dctCodering, "_out.Codering"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.Codering").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctCodering
            i = i + 1
            .Cells(i, 1) = dctCodering(k).RegelNrTbl
            .Cells(i, 2) = dctCodering(k).OmschrijvingVeld
            .Cells(i, 3) = dctCodering(k).KomtVoorInTekstblok
            .Cells(i, 4) = dctCodering(k).Code
            If k = 1323 Then 'Geldigheid offerte
                .Cells(i, 5) = "'" & Format(dctCodering(k).Waarde, "dd-mm-yy")
            Else
                .Cells(i, 5) = dctCodering(k).Waarde
            End If
        Next k
    End With
        
End Sub
Function GetKoffieType2(ByVal zetSysteem As Integer) As String
    Dim rngKoffieKg As Range, c As Range, i As Integer, dctTypes As Dictionary, sKoffieType As String
    Set rngKoffieKg = Range("_rng." & zetSysteem & "Z.KoffieAantal")
    Set dctTypes = New Dictionary
    For Each c In rngKoffieKg
        i = i + 1
        If c.Value > 0 Then
            sKoffieType = Lookup(Range("tblKoffie[Type]"), c.Offset(, -3), Range("tblKoffie[CalculatieNr]"))
            If Not dctTypes.Exists(sKoffieType) Then dctTypes.Add sKoffieType, sKoffieType
        End If
    Next c
    
    Dim k As Variant, result As String
    For Each k In dctTypes.Keys
        result = result & " en " & k
    Next k
    result = Mid(result, 5, 999)
    GetKoffieType2 = result
End Function
Sub UpdateBijlagen2()
    Dim rngBijlagen As Range, i As Integer, rngOutput As Range
    Set rngBijlagen = Range("tblBijlagen")
    Set rngOutput = Range("_out.Bijlagen")
    
    With rngOutput
        For i = 2 To rngBijlagen.Rows.Count + 1
            .Cells(i, 3) = rngBijlagen.Cells(i - 1, 2)
        Next i
    End With
    
End Sub
Sub UpdateHoodlocatieBonus2()
    Dim rngOutput As Range
    
    Set rngOutput = Range("_out.HoofdlocatieBonus")
    
    If rngOutput.Rows.Count > 1 Then rngOutput.Offset(1).Resize(rngOutput.Rows.Count - 1).ClearContents
    
    If Range("_ptr.BonusJN").Value = 1 Then
        If rngOutput.Rows.Count < 2 Then InsertRowsInRange 1, "_out.HoofdlocatieBonus"
        With rngOutput
            .Cells(2, 1) = 1501
            .Cells(2, 2) = Range("_ptr.Bonuslocatie.KlantNr")
            .Cells(2, 3) = Range("_ptr.Bonuslocatie.KlantNaam")
            .Cells(2, 4) = Range("_ptr.Bonuslocatie.KlantPlaats")
        End With
    Else
        If rngOutput.Rows.Count > 1 Then DeleteRowsFromRange 1, "_out.HoofdlocatieBonus"
    End If
End Sub

Sub UpdateBonusConditie2()
    Dim rngOutput As Range
    Set rngOutput = Range("_out.BonusConditie")
    With rngOutput
        If Range("_ptr.1Z.BonusKoffieJN").Value = "Ja" Then .Cells(2, 2) = 1 Else .Cells(2, 2) = 0
        If Range("_ptr.2Z.BonusKoffieJN").Value = "Ja" Then .Cells(3, 2) = 1 Else .Cells(3, 2) = 0
    End With
End Sub
Sub UpdateMalusConditie2()
    Dim rngOutput As Range
    Set rngOutput = Range("_out.MalusCondities")
    With rngOutput
        If Range("_ptr.1Z.MalusJN").Value = "Ja" Then .Cells(2, 2) = 1 Else .Cells(2, 2) = 0
        If Range("_ptr.2Z.MalusJN").Value = "Ja" Then .Cells(3, 2) = 1 Else .Cells(3, 2) = 0
    End With
End Sub
Sub UpdateCheckDeal2(ByVal OK As Boolean)
    Dim rngOuptut As Range
    Set rngOuptut = Range("_out.CheckDeal")
    If OK Then
        rngOuptut.Cells(2, 2) = 1
    Else
        rngOuptut.Cells(2, 2) = 0
    End If
End Sub
Sub UpdateKoffie2()
    Dim rngKoffieKg As Range, rngKoffieOmschrijving As Range, c As Range, i As Integer, dctKoffie As Dictionary, zetSysteem As Integer
    Set dctKoffie = New Dictionary
    For zetSysteem = 1 To 2
        Set rngKoffieKg = Range("_rng." & zetSysteem & "Z.KoffieAantal")
        Set rngKoffieOmschrijving = rngKoffieKg.Offset(, -2)
        i = 0
        For Each c In rngKoffieKg
            i = i + 1
            If c.Value > 0 Then
                If Not dctKoffie.Exists(rngKoffieOmschrijving.Cells(i).Value) Then
                    dctKoffie.Add rngKoffieOmschrijving.Cells(i).Value, c.Value
                Else
                    dctKoffie.Item(rngKoffieOmschrijving.Cells(i).Value) = dctKoffie.Item(rngKoffieOmschrijving.Cells(i).Value) + c.Value
                End If
            End If
        Next c
    Next zetSysteem
    
    AdjustRange dctKoffie, "_out.Koffie"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.Koffie").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctKoffie
            i = i + 1
            .Cells(i, 1) = k
            .Cells(i, 2) = dctKoffie(k)
        Next k
    End With
    
End Sub
Sub UpdateServies()
    Dim rngServies As Range, clsServies As Servies, i As Integer, dctServies As Dictionary
    
    Set rngServies = Range("_rng.rb.Servies")
    Set dctServies = New Dictionary
    
    With rngServies
        For i = 1 To .Rows.Count
            If .Cells(i, 4) > 0 Then
                Set clsServies = New Servies
                clsServies.RegelNrTbl = 1600 + i
                clsServies.Servies = .Cells(i, 3)
                clsServies.Aantal = .Cells(i, 4)
                If .Cells(i, 5) <> 0 Then clsServies.AantalColli = .Cells(i, 4) / .Cells(i, 5)
                dctServies.Add clsServies.RegelNrTbl, clsServies
            End If
        Next i
    End With
    
    AdjustRange dctServies, "_out.Servies"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.Servies").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctServies
            i = i + 1
            .Cells(i, 1) = k
            .Cells(i, 2) = dctServies(k).Servies
            .Cells(i, 3) = dctServies(k).Aantal
            .Cells(i, 4) = dctServies(k).AantalColli
        Next k
    End With
    
End Sub
Sub UpdateServiesBudget()
    Dim clsServiesBudget As ServiesBudget, i As Integer, dctServiesBudget As Dictionary
    
    Set dctServiesBudget = New Dictionary
    
    If Range("_ptr.VervangingServiesJN") = "Ja" Then
        For i = 1 To Range("_ptr.Looptijd") - 1
            Set clsServiesBudget = New ServiesBudget
            clsServiesBudget.RegelNrTbl = 1700 + i
            clsServiesBudget.Jaar = i + 1
            clsServiesBudget.Budget = Application.WorksheetFunction.Round(Range("_ptr.rb.VervangingServiesVAP").Value, -2) / (Range("_ptr.Looptijd") - 1)
            dctServiesBudget.Add clsServiesBudget.RegelNrTbl, clsServiesBudget
        Next i
    End If
        
    AdjustRange dctServiesBudget, "_out.serviesbudget"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.serviesbudget").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctServiesBudget
            i = i + 1
            .Cells(i, 1) = k
            .Cells(i, 2) = dctServiesBudget(k).Jaar
            .Cells(i, 3) = dctServiesBudget(k).Budget
        Next k
    End With
End Sub

Sub UpdatePos()
    Dim rngPOS As Range, clsPOS As POS, i As Integer, dctPOS As Dictionary
    
    Set rngPOS = Range("_rng.InputPOS")
    Set dctPOS = New Dictionary
    
    With rngPOS
        For i = 1 To .Rows.Count
            If .Cells(i, 1) > 0 Then
                Set clsPOS = New POS
                clsPOS.RegelNrTbl = 1800 + i
                clsPOS.PosMateriaal = .Cells(i, 0)
                clsPOS.Aantal = .Cells(i, 1)
                dctPOS.Add clsPOS.RegelNrTbl, clsPOS
            End If
        Next i
    End With
    
    AdjustRange dctPOS, "_out.POS"
    
    Dim k As Variant, rngOutput As Range
    Set rngOutput = Names("_out.POS").RefersToRange
    i = 1
    With rngOutput
        For Each k In dctPOS
            i = i + 1
            .Cells(i, 1) = k
            .Cells(i, 2) = dctPOS(k).PosMateriaal
            .Cells(i, 3) = dctPOS(k).Aantal
        Next k
    End With
    
End Sub
Sub UpdateOpbouwOfferteContract2()
    Dim rngOutput As Range, rngInput As Range
    Set rngInput = Range("tblTekstblokken")
    Set rngOutput = Range("_out.OpbouwOfferteContract")
    
    If rngOutput.Rows.Count > 1 Then rngOutput.Offset(1).Resize(rngOutput.Rows.Count - 1).ClearContents
    If rngOutput.Rows.Count - 1 < rngInput.Rows.Count Then '-1 ivm headerrow
        InsertRowsInRange rngInput.Rows.Count - (rngOutput.Rows.Count - 1), "_out.OpbouwOfferteContract"
    ElseIf rngOutput.Rows.Count - 1 > rngInput.Rows.Count Then
        DeleteRowsFromRange rngOutput.Rows.Count - 1 - rngInput.Rows.Count, "_out.OpbouwOfferteContract"
    End If
    rngOutput.Offset(1).Resize(rngOutput.Rows.Count - 1, 2).Value = rngInput.Offset(, 1).Resize(, 2).Value
End Sub

