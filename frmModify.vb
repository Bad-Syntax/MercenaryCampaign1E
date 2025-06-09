Public Class frmModify
    Private CurrentSquad As cSquad
    Private TempCash As Long

    Sub frmCreate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            lstBase.SelectedIndex = 1 'Medium Mech
            lstQuality.SelectedIndex = 0 'New
            lstExperience.SelectedIndex = 1 'Regular

            TempCash = Mercenary.Cash_On_Hand
            Refresh_Units()
            Refresh_Label()
            Me.Show()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Public Sub ContinueMerc()
        cmdFinished_Click(Nothing, Nothing)
    End Sub
    Private Sub cmdFinished_Click(sender As Object, e As EventArgs) Handles cmdFinished.Click
        Try
            'Find monthly costs
            Dim TempMonthly As Int64 = (Mercenary.Salaries + Mercenary.SquadCount * 500) * (Mercenary.Overhead / 100) + Mercenary.Maintenance_Cost

            'Now Save this unit
            'TODO:  Check for overwrites
            Dim OutFile As String = "LeaderData" & vbTab & Leader.Save & vbCrLf
            OutFile &= "UnitName" & vbTab & Mercenary.UnitName & vbCrLf
            'OutFile &= "InitialUP" & vbTab & Mercenary.InitialUP & vbCrLf
            'OutFile &= "CurrentUP" & vbTab & Mercenary.CurrentUP & vbCrLf
            OutFile &= "Cash_On_Hand" & vbTab & TempCash & vbCrLf
            OutFile &= "Supplies_On_Hand" & vbTab & Mercenary.Supplies_On_Hand & vbCrLf
            OutFile &= "SupplyID" & vbTab & Mercenary.SupplyID & vbCrLf
            OutFile &= "Consumables" & vbTab & Mercenary.Consumables & vbCrLf
            OutFile &= "OverallSkill" & vbTab & Mercenary.OverallSkill & vbCrLf
            OutFile &= "Overhead" & vbTab & Mercenary.Overhead & vbCrLf
            OutFile &= "MoraleMod" & vbTab & Mercenary.MoraleMod & vbCrLf
            OutFile &= "ReputationMod" & vbTab & Mercenary.ReputationMod & vbCrLf
            OutFile &= "SquadsStart" & vbTab & Mercenary.SquadsStart & vbCrLf
            OutFile &= "Dishonorable" & vbTab & Mercenary.Dishonorable.ToString & vbCrLf
            For Each S As cSupplyOrder In Mercenary.Supply_Orders
                OutFile &= "Supply_Order" & vbTab & S.OrderID & vbTab & S.Supplies & vbTab & S.Delay & vbTab & S.Lost.ToString & vbCrLf
            Next

            Dim I As Int32 = 0 'Index, no real reason for it
            For Each R As cRegiment In Mercenary.Regiments.Values
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                OutFile &= I & vbTab & L.Name & vbTab & "" & vbTab & C.Name & vbTab & "" & vbTab & B.Name & vbTab & "" & vbTab & R.Name & vbTab & "" & vbTab & S.Rank & vbTab & S.Pilot & vbTab & S.Quality.ToString & vbTab & S.UnitType.ToString & vbTab & S.State.ToString & vbCrLf
                                I += 1
                            Next
                        Next
                    Next
                Next
            Next

            Mercenary.DragoonRating()

            'Update minus the differences
            Mercenary.Cash_On_Hand = TempCash

            'Dim Filename As String = Mercenary.UnitName.ToString & "_" & Format(Now.Hour, "##00").ToString & Format(Now.Minute, "##00").ToString & Format(Now.Day, "##00").ToString & Format(Now.Month, "##00").ToString & Format(Now.Year, "####0000").ToString & ".txt"
            'If System.IO.File.Exists("Saves\" & Filename & ".txt") Then
            'If MsgBox("Overwrite existing file?", MsgBoxStyle.YesNo, "Overwrite") = MsgBoxResult.No Then
            'Exit Sub
            'End If
            'End If
            'System.IO.File.WriteAllText("Saves\" & Filename, OutFile)

            Me.Hide()
            frmOperations.Enabled = True
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub lstBase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstQuality.SelectedIndexChanged, lstExperience.SelectedIndexChanged, lstBase.SelectedIndexChanged
        If lstBase.SelectedIndex >= 9 And lstBase.SelectedIndex <= 21 Then lstQuality.Visible = False Else lstQuality.Visible = True
        Refresh_Base_Unit_Label()
    End Sub
    Private Sub chkHireling_CheckedChanged(sender As Object, e As EventArgs) Handles chkHireling.CheckedChanged
        Refresh_Base_Unit_Label()
    End Sub

    Public Sub Refresh_Base_Unit_Label()
        Try
            'Debug.Print("Refresh_Base_Unit_Label")
            GetCurrentSquad()
            Dim TempText As String = ""
            TempText &= "Unit Type: " & CurrentSquad.UnitType.ToString & vbCrLf
            TempText &= "Cost: €" & Format(CurrentSquad.Cost * 10000, "###,##0") & vbCrLf
            TempText &= "Reputation: " & Format(CurrentSquad.Reputation / 12, "###,##0") & vbCrLf
            TempText &= "Morale: " & Format(CurrentSquad.Morale, "###,##0.00") & vbCrLf
            If CurrentSquad.Hireling Then
                TempText &= "Salary: €" & Format(CurrentSquad.Salary, "###,##0") & " (Random until saved)" & vbCrLf
            Else
                TempText &= "Salary: €" & Format(CurrentSquad.Salary, "###,##0") & vbCrLf
            End If
            TempText &= "Support Generated: " & CurrentSquad.Support_Generated & vbCrLf
            TempText &= "Support Required: " & CurrentSquad.Support_Required & vbCrLf
            TempText &= "Support Cost: €" & Format(Math.Max(0, CurrentSquad.Support_Required - CurrentSquad.Support_Generated) * 5000, "###,##0") & vbCrLf
            TempText &= "Consumables: €500" & vbCrLf
            TempText &= "Aerospace: " & CurrentSquad.Aerospace & vbCrLf
            TempText &= "Maneuver: " & CurrentSquad.Maneuver & vbCrLf
            TempText &= "Combat: " & CurrentSquad.Combat & vbCrLf
            labSquadTemplate.Text = TempText
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub Refresh_Units()
        Try
            lstBattalions.Items.Clear()
            lstCompanies.Items.Clear()
            lstLances.Items.Clear()
            lstSquads.Items.Clear()
            Dim OldIndex As Int32 = Math.Max(0, lstRegiments.SelectedIndex)
            With lstRegiments
                .BeginUpdate()
                .Items.Clear()
                For Each R As String In Mercenary.Regiments.Keys
                    .Items.Add(R)
                Next
                If .Items.Count > 0 Then .SelectedIndex = Math.Min(.Items.Count - 1, OldIndex)
                .EndUpdate()
            End With
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub Refresh_Label()
        Try
            Dim TotalExperience As Int32 = 0
            Dim TempText As String = ""
            With Mercenary
                .Consumables = .SquadCount * 500

                TempText &= "Mercenary Unit Stats: " & .UnitName & vbCrLf
                TempText &= "Leader Name: " & Leader.Name & vbCrLf
                TempText &= "Leader Rank: " & Leader.Rank
                If Leader.Rank = "" Then
                    Select Case .SquadCount
                        Case <= 6 : TempText &= "Lieutenant"
                        Case <= 12 : TempText &= "Captain"
                        Case <= 40 : TempText &= "Major"
                        Case <= 156 : TempText &= "Colonel"
                        Case Else : TempText &= "General"
                    End Select
                End If
                TempText &= vbCrLf
                TempText &= "Leader Leadership: " & Leader.Skills(eSkill.Leadership) & vbCrLf
                TempText &= "Leader Tactics: " & Leader.Skills(eSkill.Tactics) & vbCrLf
                TempText &= "Leader Diplomacy: " & Leader.Skills(eSkill.Diplomacy) & vbCrLf
                TempText &= "Leader Land Management: " & Leader.Skills(eSkill.Land_Management) & vbCrLf
                TempText &= vbCrLf

                TempText &= "Morale: " & Format(.ForceMorale, "##0.00") & "  Reputation: " & Format(.Reputation / 12, "###,##0") & vbCrLf
                TempText &= "Support Generated: " & Format(.Support_Generated, "###,##0") & "  Support Required: " & Format(.Support_Required, "###,##0") & vbCrLf
                TempText &= vbCrLf

                TempText &= "Total Squads: " & .SquadCount & "  Combat: " & .SquadCount - .SupportSquads & "  Support: " & .SupportSquads

                TempText &= vbCrLf

                Dim TempOverhead As Int64 = (.Salaries + .Consumables) * (.Overhead / 100)
                TempText &= "Overhead (" & .Overhead & "%) = €" & Format(TempOverhead, "###,##0") & vbCrLf
                Dim TempSalary As Int64
                For Each R As cRegiment In Mercenary.Regiments.Values
                    TempSalary += R.Salary
                Next
                TempText &= "Salary: €" & Format(TempSalary, "###,##0") & "  Consumables: €" & Format(.Consumables, "###,##0") & vbCrLf

                Dim TempSupportCosts As Int64 = 0
                If .Support_Required > .Support_Generated Then
                    TempSupportCosts = (.Support_Required - .Support_Generated) * 5000
                End If
                TempText &= "Support Costs: €" & Format(TempSupportCosts, "###,##0") & vbCrLf

                TempText &= "Cash On Hand: €" & Format(TempCash, "###,##0") & vbCrLf


                For Each R As cRegiment In Mercenary.Regiments.Values
                    If R.PillagingCivilians > 0 Then TempText &= R.Name & "  Pillaging Civilians For: " & R.PillagingCivilians & "/20" & vbCrLf
                    If R.PlayingItSafe > 0 Then TempText &= R.Name & "  Playing It Safe Value: " & R.PlayingItSafe & vbCrLf
                Next

                TempText &= vbCrLf

                If .SquadCount > 0 Then TempText &= "Carrying Capacity Used/Total:" & vbCrLf
                If .Get_Capacity_Used.Mechs > 0 Or .Get_Capacity_Available.Mechs > 0 Then TempText &= "   Mechs: " & .Get_Capacity_Used.Mechs & "/" & .Get_Capacity_Available.Mechs & vbCrLf
                If .Get_Capacity_Used.LAM > 0 Or .Get_Capacity_Available.LAM > 0 Then TempText &= "   Land-Air Mechs: " & .Get_Capacity_Used.LAM & "/" & .Get_Capacity_Available.LAM & vbCrLf
                If .Get_Capacity_Used.Aerospace > 0 Or .Get_Capacity_Available.Aerospace > 0 Then TempText &= "   Aerospace Fighters: " & .Get_Capacity_Used.Aerospace & "/" & .Get_Capacity_Available.Aerospace & vbCrLf
                If .Get_Capacity_Used.Armor_Light > 0 Or .Get_Capacity_Available.Armor_Light > 0 Then TempText &= "   Light Armor: " & .Get_Capacity_Used.Armor_Light & "/" & .Get_Capacity_Available.Armor_Light & vbCrLf
                If .Get_Capacity_Used.Armor_Heavy > 0 Or .Get_Capacity_Available.Armor_Heavy > 0 Then TempText &= "   Heavy Armor: " & .Get_Capacity_Used.Armor_Heavy & "/" & .Get_Capacity_Available.Armor_Heavy & vbCrLf
                If .Get_Capacity_Used.Infantry_Platoons > 0 Or .Get_Capacity_Available.Infantry_Platoons > 0 Then TempText &= "   Infantry Platoons: " & .Get_Capacity_Used.Infantry_Platoons & "/" & .Get_Capacity_Available.Infantry_Platoons & vbCrLf
                If .Get_Capacity_Used.Aircraft > 0 Or .Get_Capacity_Available.Aircraft > 0 Then TempText &= "   Conventional Aircraft: " & .Get_Capacity_Used.Aircraft & "/" & .Get_Capacity_Available.Aircraft & vbCrLf
                If .Get_Capacity_Used.Dropships > 0 Or .Get_Capacity_Available.Dropships > 0 Then TempText &= "   Dropships: " & .Get_Capacity_Used.Dropships & "/" & .Get_Capacity_Available.Dropships & vbCrLf
                TempText &= vbCrLf

                TempText &= "Aerospace Value: " & .Aerospace & vbCrLf
                TempText &= "Maneuver Value: " & .Maneuver & vbCrLf
                TempText &= "Combat Value: " & .Combat & vbCrLf

                'If lstRegiments.SelectedIndex > -1 Then
                '   Dim TempVal() As String = Split(lstRegiments.SelectedItem, " - ")
                '   ReDim Preserve TempVal(3)
                '   TempText &= vbCrLf
                '   TempText &= "Element Type: " & TempVal(0) & vbCrLf
                '  TempText &= "Element Quality: " & TempVal(1) & vbCrLf
                '   If TempVal(2) = "Salvaged" Then TempText &= "Element Type: Salvaged Condition" & vbCrLf
                'End If

            End With
            labCBills.Text = "C-Bills Available: " & TempCash
            labData.Text = TempText
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Public Sub Refresh_Labels_Formation()
        Try
            Dim TempLabel As String = ""
            If lstRegiments.SelectedIndex > -1 Then
                With Mercenary.Regiments(lstRegiments.SelectedItem)
                    TempLabel &= "Cost: €" & Format(.Cost * 10000, "###,##0") & "  Sqds: " & .SquadCount & "  ComSqd: " & .SquadCount & "  SptSqd: " & .SupportSquads & vbCrLf
                    TempLabel &= "Aero: " & .Aerospace & "  Man: " & .Maneuver & "  Com: " & .Combat & "  Mor: " & Format(Math.Round(.TotalMorale / (.SquadCount - .SupportSquads), 2), "###,##0.00") & "  Rep: " & .Reputation & vbCrLf
                    TempLabel &= "SP Gen: " & Format(.Support_Generated, "###,##0") & "  SP Req: " & Format(.Support_Required, "###,##0") & " SP Def: " & Format(.Support_Generated - .Support_Required, "###,##0") & "  Maint: €" & Format(.Maintenance_Cost, "###,##0") & vbCrLf
                    Dim ConCost As Int64 = .SquadCount * 500
                    Dim OvCost As Int64 = (Mercenary.Overhead / 100) * (.Salary + ConCost)
                    Dim TotCost As Int64 = ConCost + OvCost + .Maintenance_Cost
                    TempLabel &= "Salary: €" & Format(.Salary, "###,##0") & "  Supply: €" & Format(ConCost, "###,##0") & "  Overhead(" & Mercenary.Overhead & "%): €" & Format(OvCost, "###,##0") & vbCrLf
                    TempLabel &= "Total Cost/Month: €" & Format(TotCost, "###,##0")
                End With
                labRegiment.Text = TempLabel

                If lstBattalions.SelectedIndex > -1 Then
                    TempLabel = ""
                    With Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem)
                        TempLabel &= "Cost: €" & Format(.Cost * 10000, "###,##0") & "  Sqds: " & .SquadCount & "  ComSqd: " & .SquadCount & "  SptSqd: " & .SupportSquads & vbCrLf
                        TempLabel &= "Aero: " & .Aerospace & "  Man: " & .Maneuver & "  Com: " & .Combat & "  Mor: " & Format(Math.Round(.TotalMorale / (.SquadCount - .SupportSquads), 2), "###,##0.00") & "  Rep: " & .Reputation & vbCrLf
                        TempLabel &= "SP Gen: " & Format(.Support_Generated, "###,##0") & "  SP Req: " & Format(.Support_Required, "###,##0") & " SP Def: " & Format(.Support_Generated - .Support_Required, "###,##0") & "  Maint: €" & Format(.Maintenance_Cost, "###,##0") & vbCrLf
                        Dim ConCost As Int64 = .SquadCount * 500
                        Dim OvCost As Int64 = (Mercenary.Overhead / 100) * (.Salary + ConCost)
                        Dim TotCost As Int64 = ConCost + OvCost + .Maintenance_Cost
                        TempLabel &= "Salary: €" & Format(.Salary, "###,##0") & "  Supply: €" & Format(ConCost, "###,##0") & "  Overhead(" & Mercenary.Overhead & "%): €" & Format(OvCost, "###,##0") & vbCrLf
                        TempLabel &= "Total Cost/Month: €" & Format(TotCost, "###,##0")
                    End With
                    labBattalion.Text = TempLabel
                    If lstCompanies.SelectedIndex > -1 Then
                        TempLabel = ""
                        With Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem)
                            TempLabel &= "Cost: €" & Format(.Cost * 10000, "###,##0") & "  Sqds: " & .SquadCount & "  ComSqd: " & .SquadCount & "  SptSqd: " & .SupportSquads & vbCrLf
                            TempLabel &= "Aero: " & .Aerospace & "  Man: " & .Maneuver & "  Com: " & .Combat & "  Mor: " & Format(Math.Round(.TotalMorale / (.SquadCount - .SupportSquads), 2), "###,##0.00") & "  Rep: " & .Reputation & vbCrLf
                            TempLabel &= "SP Gen: " & Format(.Support_Generated, "###,##0") & "  SP Req: " & Format(.Support_Required, "###,##0") & " SP Def: " & Format(.Support_Generated - .Support_Required, "###,##0") & "  Maint: €" & Format(.Maintenance_Cost, "###,##0") & vbCrLf
                            Dim ConCost As Int64 = .SquadCount * 500
                            Dim OvCost As Int64 = (Mercenary.Overhead / 100) * (.Salary + ConCost)
                            Dim TotCost As Int64 = ConCost + OvCost + .Maintenance_Cost
                            TempLabel &= "Salary: €" & Format(.Salary, "###,##0") & "  Supply: €" & Format(ConCost, "###,##0") & "  Overhead(" & Mercenary.Overhead & "%): €" & Format(OvCost, "###,##0") & vbCrLf
                            TempLabel &= "Total Cost/Month: €" & Format(TotCost, "###,##0")
                        End With
                        labCompany.Text = TempLabel
                        If lstLances.SelectedIndex > -1 Then
                            TempLabel = ""
                            With Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem)
                                TempLabel &= "Cost: €" & Format(.Cost * 10000, "###,##0") & "  Sqds: " & .SquadCount & "  ComSqd: " & .SquadCount & "  SptSqd: " & .SupportSquads & vbCrLf
                                TempLabel &= "Aero: " & .Aerospace & "  Man: " & .Maneuver & "  Com: " & .Combat & "  Mor: " & Format(Math.Round(.TotalMorale / (.SquadCount - .SupportSquads), 2), "###,##0.00") & "  Rep: " & .Reputation & vbCrLf
                                TempLabel &= "SP Gen: " & Format(.Support_Generated, "###,##0") & "  SP Req: " & Format(.Support_Required, "###,##0") & " SP Def: " & Format(.Support_Generated - .Support_Required, "###,##0") & "  Maint: €" & Format(.Maintenance_Cost, "###,##0") & vbCrLf
                                Dim ConCost As Int64 = .SquadCount * 500
                                Dim OvCost As Int64 = (Mercenary.Overhead / 100) * (.Salary + ConCost)
                                Dim TotCost As Int64 = ConCost + OvCost + .Maintenance_Cost
                                TempLabel &= "Salary: €" & Format(.Salary, "###,##0") & "  Supply: €" & Format(ConCost, "###,##0") & "  Overhead(" & Mercenary.Overhead & "%): €" & Format(OvCost, "###,##0") & vbCrLf
                                TempLabel &= "Total Cost/Month: €" & Format(TotCost, "###,##0")
                            End With
                            labLance.Text = TempLabel
                            If lstSquads.SelectedIndex > -1 Then
                                TempLabel = ""
                                Dim TempSquad() As String = Split(lstSquads.SelectedItem, Space(100))
                                With Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads(TempSquad(1))
                                    TempLabel &= "Cost: €" & Format(.Cost * 10000, "###,##0") & vbCrLf
                                    TempLabel &= "Element: " & .Element & "  Type: " & .UnitType.ToString & vbCrLf
                                    TempLabel &= "Aero: " & .Aerospace & "  Man: " & .Maneuver & "  Com: " & .Combat & "  Mor: " & Format(Math.Round(.Morale, 2), "###,##0.00") & "  Rep: " & .Reputation & vbCrLf
                                    TempLabel &= "SP Gen: " & Format(.Support_Generated, "###,##0") & "  SP Req: " & Format(.Support_Required, "###,##0") & " SP Def: " & Format(.Support_Generated - .Support_Required, "###,##0") & "  Maint: €" & Format(.Maintenance_Cost, "###,##0") & vbCrLf
                                    Dim ConCost As Int64 = .SquadCount * 500
                                    Dim OvCost As Int64 = (Mercenary.Overhead / 100) * (.Salary + ConCost)
                                    Dim TotCost As Int64 = ConCost + OvCost + .Maintenance_Cost
                                    TempLabel &= "Salary: €" & Format(.Salary, "###,##0") & "  Supply: €" & Format(ConCost, "###,##0") & "  Overhead(" & Mercenary.Overhead & "%): €" & Format(OvCost, "###,##0") & vbCrLf
                                    TempLabel &= "Total Cost/Month: €" & Format(TotCost, "###,##0")
                                End With
                                labSquad.Text = TempLabel
                            End If
                        End If
                    End If

                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Function GetUnitString(Unit As cSquad) As String
        Try
            Dim TempString As String = ""
            If Unit.Hireling Then TempString &= "*"
            Select Case Unit.UnitType
                Case eUnitType.Mech_Light : TempString &= "Light Mech"
                Case eUnitType.Mech_Medium : TempString &= "Medium Mech"
                Case eUnitType.Mech_Heavy : TempString &= "Heavy Mech"
                Case eUnitType.Mech_Assault : TempString &= "Assault Mech"
                Case eUnitType.Fighter_Light : TempString &= "Light Fighter"
                Case eUnitType.Fighter_Medium : TempString &= "Medium Fighter"
                Case eUnitType.Fighter_Heavy : TempString &= "Heavy Fighter"
                Case eUnitType.LAM_Light : TempString &= "Light LAM"
                Case eUnitType.LAM_Medium : TempString &= "Medium LAM"
                Case eUnitType.Infantry_Regular : TempString &= "Regular Infantry"
                Case eUnitType.Infantry_Motorized : TempString &= "Motorized Infantry"
                Case eUnitType.Infantry_Jump : TempString &= "Jump Infantry"
                Case eUnitType.Infantry_Scout : TempString &= "Scout"
                Case eUnitType.Armor_Light : TempString &= "Light Armor"
                Case eUnitType.Armor_Heavy : TempString &= "Heavy Armor"
                Case eUnitType.Artillery : TempString &= "Artillery"
                Case eUnitType.Aircraft : TempString &= "Aircraft"
                Case eUnitType.Support : TempString &= "Support"
                Case eUnitType.Infantry_Regular_Airmobile : TempString &= "Airmobile Regular Infantry"
                Case eUnitType.Infantry_Motorized_Airmobile : TempString &= "Airmobile Motorized Infantry"
                Case eUnitType.Infantry_Jump_Airmobile : TempString &= "Airmobile Jump Infantry"
                Case eUnitType.Infantry_Scout_Airmobile : TempString = "Airmobile Scout"

                Case eUnitType.Leopard_Dropship : TempString &= "Leopard Dropship"
                Case eUnitType.Union_Dropship : TempString &= "Union Dropship"
                Case eUnitType.Overlord_Dropship : TempString &= "Overlord Dropship"
                Case eUnitType.Fury_Dropship : TempString &= "Fury Dropship"
                Case eUnitType.Gazelle_Dropship : TempString &= "Gazelle Dropship"
                Case eUnitType.Seeker_Dropship : TempString &= "Seeker Dropship"
                Case eUnitType.Triumph_Dropship : TempString &= "Triumph Dropship"
                Case eUnitType.Condor_Dropship : TempString &= "Condor Dropship"
                Case eUnitType.Excaliber_Dropship : TempString &= "Excalibur Dropship"
                Case eUnitType.Scout_Jumpship : TempString &= "Scout Jumpship"
                Case eUnitType.Invader_Jumpship : TempString &= "Invader Jumpship"
                Case eUnitType.Monolith_Jumpship : TempString &= "Monolith Jumpship"
                Case eUnitType.Star_Lord_Jumpship : TempString &= "Star Lord Jumpship"
                Case eUnitType.Merchant_Jumpship : TempString &= "Merchant Jumpship"

                Case eUnitType.Avenger_Dropship : TempString &= "Avenger Dropship"
                Case eUnitType.Leopard_CV_Dropship : TempString &= "Leopard CV Dropship"
                Case eUnitType.Intruder_Dropship : TempString &= "Intruder Dropship"
                Case eUnitType.Buccaneer_Dropship : TempString &= "Buccaneer Dropship"
                Case eUnitType.Achilles_Dropship : TempString &= "Achilles Dropship"
                Case eUnitType.Monarch_Dropship : TempString &= "Monarch Dropship"
                Case eUnitType.Fortress_Dropship : TempString &= "Fortress Dropship"
                Case eUnitType.Vengeance_Dropship : TempString &= "Vengeance Dropship"
                Case eUnitType.Mule_Dropship : TempString &= "Mule Dropship"
                Case eUnitType.Mammoth_Dropship : TempString &= "Mammoth Dropship"
                Case eUnitType.Behemoth_Dropship : TempString &= "Behemoth Dropship"
            End Select
            If Unit.Experience = eExperience.Green Then
                TempString &= " - Green"
            ElseIf Unit.Experience = eExperience.Regular Then
                TempString &= " - Regular"
            ElseIf Unit.Experience = eExperience.Veteran Then
                TempString &= " - Veteran"
            ElseIf Unit.Experience = eExperience.Elite Then
                TempString &= " - Elite"
            End If
            If Unit.Quality = eQuality.Salvaged Then
                TempString &= " - Salvaged"
            ElseIf Unit.Quality = eQuality.Destroyed Then
                TempString &= " - Destroyed"
            End If
            Return TempString
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Function
    Private Sub GetCurrentSquad(Optional ExistingUnit As String = "")
        Try
            Debug.Print("GetCurrentSquad(" & ExistingUnit & ")")
            If ExistingUnit <> "" Then
                Dim TempVal() As String = Split("" & lstSquads.SelectedItem, Space(100))
                If TempVal.GetUpperBound(0) > 0 Then
                    CurrentSquad = Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads(TempVal(1))
                Else
                    CurrentSquad = New cSquad()
                    If Mid(ExistingUnit, 1, 1) = "*" Then
                        CurrentSquad.Hireling = True
                        ExistingUnit = Mid(ExistingUnit, 2, Len(ExistingUnit)).Trim
                    End If
                    TempVal = Split(ExistingUnit, " - ")
                    ReDim Preserve TempVal(3)
                    If TempVal(1) = "Elite" Then
                        CurrentSquad.Experience = eExperience.Elite
                    ElseIf TempVal(1) = "Veteran" Then
                        CurrentSquad.Experience = eExperience.Veteran
                    ElseIf TempVal(1) = "Regular" Then
                        CurrentSquad.Experience = eExperience.Regular
                    ElseIf TempVal(1) = "Green" Then
                        CurrentSquad.Experience = eExperience.Green
                    End If
                    If ("" & TempVal(2)) = "Salvaged" Then
                        CurrentSquad.Quality = eQuality.Salvaged
                    ElseIf ("" & TempVal(2)) = "Destroyed" Then
                        CurrentSquad.Quality = eQuality.Destroyed
                    Else
                        CurrentSquad.Quality = eQuality.LikeNew
                    End If

                    Select Case TempVal(0)
                        Case "Light Mech" : CurrentSquad.UnitType = eUnitType.Mech_Light
                        Case "Medium Mech" : CurrentSquad.UnitType = eUnitType.Mech_Medium
                        Case "Heavy Mech" : CurrentSquad.UnitType = eUnitType.Mech_Heavy
                        Case "Assault Mech" : CurrentSquad.UnitType = eUnitType.Mech_Assault
                        Case "Light Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Light
                        Case "Medium Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Medium
                        Case "Heavy Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Heavy
                        Case "Light LAM" : CurrentSquad.UnitType = eUnitType.LAM_Light
                        Case "Medium LAM" : CurrentSquad.UnitType = eUnitType.LAM_Medium
                        Case "Regular Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Regular
                        Case "Motorized Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Motorized
                        Case "Jump Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Jump
                        Case "Scout" : CurrentSquad.UnitType = eUnitType.Infantry_Scout
                        Case "Light Armor" : CurrentSquad.UnitType = eUnitType.Armor_Light
                        Case "Heavy Armor" : CurrentSquad.UnitType = eUnitType.Armor_Heavy
                        Case "Artillery" : CurrentSquad.UnitType = eUnitType.Artillery
                        Case "Aircraft" : CurrentSquad.UnitType = eUnitType.Aircraft
                        Case "Support" : CurrentSquad.UnitType = eUnitType.Support
                        Case "Airmobile Regular Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Regular_Airmobile
                        Case "Airmobile Motorized Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Motorized_Airmobile
                        Case "Airmobile Jump Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Jump_Airmobile
                        Case "Airmobile Scout" : CurrentSquad.UnitType = eUnitType.Infantry_Scout_Airmobile

                        Case "Leopard Dropship" : CurrentSquad.UnitType = eUnitType.Leopard_Dropship
                        Case "Union Dropship" : CurrentSquad.UnitType = eUnitType.Union_Dropship
                        Case "Overlord Dropship" : CurrentSquad.UnitType = eUnitType.Overlord_Dropship
                        Case "Fury Dropship" : CurrentSquad.UnitType = eUnitType.Fury_Dropship
                        Case "Gazelle Dropship" : CurrentSquad.UnitType = eUnitType.Gazelle_Dropship
                        Case "Seeker Dropship" : CurrentSquad.UnitType = eUnitType.Seeker_Dropship
                        Case "Triumph Dropship" : CurrentSquad.UnitType = eUnitType.Triumph_Dropship
                        Case "Condor Dropship" : CurrentSquad.UnitType = eUnitType.Condor_Dropship
                        Case "Excaliber Dropship" : CurrentSquad.UnitType = eUnitType.Excaliber_Dropship

                        Case "Scout Jumpship" : CurrentSquad.UnitType = eUnitType.Scout_Jumpship
                        Case "Invader Jumpship" : CurrentSquad.UnitType = eUnitType.Invader_Jumpship
                        Case "Monolith Jumpship" : CurrentSquad.UnitType = eUnitType.Monolith_Jumpship
                        Case "Star Lord Jumpship" : CurrentSquad.UnitType = eUnitType.Star_Lord_Jumpship
                        Case "Merchant Jumpship" : CurrentSquad.UnitType = eUnitType.Merchant_Jumpship

                        Case "Avenger Dropship" : CurrentSquad.UnitType = eUnitType.Avenger_Dropship
                        Case "Leopard CV Dropship" : CurrentSquad.UnitType = eUnitType.Leopard_CV_Dropship
                        Case "Intruder Dropship" : CurrentSquad.UnitType = eUnitType.Intruder_Dropship
                        Case "Buccaneer Dropship" : CurrentSquad.UnitType = eUnitType.Buccaneer_Dropship
                        Case "Achilles Dropship" : CurrentSquad.UnitType = eUnitType.Achilles_Dropship
                        Case "Monarch Dropship" : CurrentSquad.UnitType = eUnitType.Monarch_Dropship
                        Case "Fortress Dropship" : CurrentSquad.UnitType = eUnitType.Fortress_Dropship
                        Case "Vengeance Dropship" : CurrentSquad.UnitType = eUnitType.Vengeance_Dropship
                        Case "Mule Dropship" : CurrentSquad.UnitType = eUnitType.Mule_Dropship
                        Case "Mammoth Dropship" : CurrentSquad.UnitType = eUnitType.Mammoth_Dropship
                        Case "Behemoth Dropship" : CurrentSquad.UnitType = eUnitType.Behemoth_Dropship
                    End Select
                End If

            Else
                CurrentSquad = New cSquad()
                If chkHireling.Checked Then CurrentSquad.Hireling = True
                CurrentSquad.Experience = eExperience.Regular
                CurrentSquad.Quality = eQuality.LikeNew
                If lstQuality.Visible Then
                    If lstQuality.SelectedIndex = 0 Then
                        CurrentSquad.Quality = eQuality.LikeNew
                    ElseIf lstQuality.SelectedIndex = 2 Then
                        CurrentSquad.Quality = eQuality.Destroyed
                    Else
                        CurrentSquad.Quality = eQuality.Salvaged
                    End If
                End If
                If lstExperience.Visible Then CurrentSquad.Experience = lstExperience.SelectedIndex + 1
                If lstBase.SelectedIndex < 0 Then lstBase.SelectedIndex = 0
                Select Case lstBase.SelectedItem.ToString
                    Case "Light Mech" : CurrentSquad.UnitType = eUnitType.Mech_Light
                    Case "Medium Mech" : CurrentSquad.UnitType = eUnitType.Mech_Medium
                    Case "Heavy Mech" : CurrentSquad.UnitType = eUnitType.Mech_Heavy
                    Case "Assault Mech" : CurrentSquad.UnitType = eUnitType.Mech_Assault
                    Case "Light Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Light
                    Case "Medium Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Medium
                    Case "Heavy Fighter" : CurrentSquad.UnitType = eUnitType.Fighter_Heavy
                    Case "Light LAM" : CurrentSquad.UnitType = eUnitType.LAM_Light
                    Case "Medium LAM" : CurrentSquad.UnitType = eUnitType.LAM_Medium
                    Case "Regular Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Regular
                    Case "Motorized Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Motorized
                    Case "Jump Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Jump
                    Case "Scout" : CurrentSquad.UnitType = eUnitType.Infantry_Scout
                    Case "Light Armor" : CurrentSquad.UnitType = eUnitType.Armor_Light
                    Case "Heavy Armor" : CurrentSquad.UnitType = eUnitType.Armor_Heavy
                    Case "Artillery" : CurrentSquad.UnitType = eUnitType.Artillery
                    Case "Aircraft" : CurrentSquad.UnitType = eUnitType.Aircraft
                    Case "Support" : CurrentSquad.UnitType = eUnitType.Support
                    Case "Airmobile Regular Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Regular_Airmobile
                    Case "Airmobile Motorized Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Motorized_Airmobile
                    Case "Airmobile Jump Infantry" : CurrentSquad.UnitType = eUnitType.Infantry_Jump_Airmobile
                    Case "Airmobile Scout" : CurrentSquad.UnitType = eUnitType.Infantry_Scout_Airmobile

                    Case "Leopard Dropship" : CurrentSquad.UnitType = eUnitType.Leopard_Dropship
                    Case "Union Dropship" : CurrentSquad.UnitType = eUnitType.Union_Dropship
                    Case "Overlord Dropship" : CurrentSquad.UnitType = eUnitType.Overlord_Dropship
                    Case "Fury Dropship" : CurrentSquad.UnitType = eUnitType.Fury_Dropship
                    Case "Gazelle Dropship" : CurrentSquad.UnitType = eUnitType.Gazelle_Dropship
                    Case "Seeker Dropship" : CurrentSquad.UnitType = eUnitType.Seeker_Dropship
                    Case "Triumph Dropship" : CurrentSquad.UnitType = eUnitType.Triumph_Dropship
                    Case "Condor Dropship" : CurrentSquad.UnitType = eUnitType.Condor_Dropship
                    Case "Excaliber Dropship" : CurrentSquad.UnitType = eUnitType.Excaliber_Dropship

                    Case "Scout Jumpship" : CurrentSquad.UnitType = eUnitType.Scout_Jumpship
                    Case "Invader Jumpship" : CurrentSquad.UnitType = eUnitType.Invader_Jumpship
                    Case "Monolith Jumpship" : CurrentSquad.UnitType = eUnitType.Monolith_Jumpship
                    Case "Star Lord Jumpship" : CurrentSquad.UnitType = eUnitType.Star_Lord_Jumpship
                    Case "Merchant Jumpship" : CurrentSquad.UnitType = eUnitType.Merchant_Jumpship

                    Case "Avenger Dropship" : CurrentSquad.UnitType = eUnitType.Avenger_Dropship
                    Case "Leopard CV Dropship" : CurrentSquad.UnitType = eUnitType.Leopard_CV_Dropship
                    Case "Intruder Dropship" : CurrentSquad.UnitType = eUnitType.Intruder_Dropship
                    Case "Buccaneer Dropship" : CurrentSquad.UnitType = eUnitType.Buccaneer_Dropship
                    Case "Achilles Dropship" : CurrentSquad.UnitType = eUnitType.Achilles_Dropship
                    Case "Monarch Dropship" : CurrentSquad.UnitType = eUnitType.Monarch_Dropship
                    Case "Fortress Dropship" : CurrentSquad.UnitType = eUnitType.Fortress_Dropship
                    Case "Vengeance Dropship" : CurrentSquad.UnitType = eUnitType.Vengeance_Dropship
                    Case "Mule Dropship" : CurrentSquad.UnitType = eUnitType.Mule_Dropship
                    Case "Mammoth Dropship" : CurrentSquad.UnitType = eUnitType.Mammoth_Dropship
                    Case "Behemoth Dropship" : CurrentSquad.UnitType = eUnitType.Behemoth_Dropship

                End Select
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub


#Region "Selecting any of the formation list boxes"
    Private Sub lstRegiments_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRegiments.SelectedIndexChanged
        Try
            If lstRegiments.SelectedIndex < 0 Then
                labRegiment.Text = ""
                Exit Sub
            End If

            With lstBattalions
                .BeginUpdate()
                .Items.Clear()
                For Each B As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions.Keys
                    .Items.Add(B)
                Next
                If .Items.Count > 0 Then .SelectedIndex = 0
                .EndUpdate()
            End With

            Refresh_Labels_Formation()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub lstBattalions_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBattalions.SelectedIndexChanged
        Try
            If lstBattalions.SelectedIndex < 0 Then
                labBattalion.Text = ""
                Exit Sub
            End If
            With lstCompanies
                .BeginUpdate()
                .Items.Clear()
                For Each C As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies.Keys
                    .Items.Add(C)
                Next
                If .Items.Count > 0 Then .SelectedIndex = 0
                .EndUpdate()
            End With
            Refresh_Labels_Formation()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub lstCompanies_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCompanies.SelectedIndexChanged
        Try
            If lstCompanies.SelectedIndex < 0 Then Exit Sub
            With lstLances
                .BeginUpdate()
                .Items.Clear()
                For Each L As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances.Keys
                    .Items.Add(L)
                Next
                If .Items.Count > 0 Then .SelectedIndex = 0
                .EndUpdate()
            End With
            Refresh_Labels_Formation()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub lstLances_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstLances.SelectedIndexChanged
        If lstLances.SelectedIndex < 0 Then
            labLance.Text = ""
            Exit Sub
        End If
        Refresh_Squads()
    End Sub
    Private Sub lstSquads_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSquads.SelectedIndexChanged
        If lstSquads.SelectedIndex < 0 Then
            labSquad.Text = ""
            Exit Sub
        End If
        Refresh_Labels_Formation()

    End Sub
#End Region



#Region "Simple add/remove formation buttons"
    Private Sub cmdRegimentsAdd_Click(sender As Object, e As EventArgs) Handles cmdRegimentsAdd.Click
        Try
            Dim Name As String = InputBox("Enter Regiment Name: ", "Regiment Name", "")
            If Name = "" Then Exit Sub
            Dim TempUnit As New cRegiment(Name)
            If Mercenary.Regiments.ContainsKey(Name) Then
                MsgBox("You already have that unit!  Names must be unique!", vbOKOnly)
                Exit Sub
            End If
            Mercenary.Regiments.Add(Name, TempUnit)
            Refresh_Regiments()
            lstRegiments.SelectedIndex = lstRegiments.Items.Count - 1
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdBattalionsAdd_Click(sender As Object, e As EventArgs) Handles cmdBattalionsAdd.Click
        Try
            Dim Name As String = InputBox("Enter Battalion Name: ", "Battalion Name", "")
            If Name = "" Then Exit Sub
            Dim TempUnit As New cBattalion(Name)
            If Mercenary.Regiments(lstRegiments.SelectedItem).Battalions.ContainsKey(Name) Then
                MsgBox("You already have that unit!  Names must be unique!", vbOKOnly)
                Exit Sub
            End If
            Mercenary.Regiments(lstRegiments.SelectedItem).Battalions.Add(Name, TempUnit)
            Refresh_Battalions()
            lstBattalions.SelectedIndex = lstBattalions.Items.Count - 1
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdCompaniesAdd_Click(sender As Object, e As EventArgs) Handles cmdCompaniesAdd.Click
        Try
            Dim Name As String = InputBox("Enter Company Name: ", "Company Name", "")
            If Name = "" Then Exit Sub
            Dim TempUnit As New cCompany(Name)
            If Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies.ContainsKey(Name) Then
                MsgBox("You already have that unit!  Names must be unique!", vbOKOnly)
                Exit Sub
            End If
            Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies.Add(Name, TempUnit)
            Refresh_Companies()
            lstCompanies.SelectedIndex = lstCompanies.Items.Count - 1
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdLancesAdd_Click(sender As Object, e As EventArgs) Handles cmdLancesAdd.Click
        Try
            Dim Name As String = InputBox("Enter Lance Name: ", "Lance Name", "")
            If Name = "" Then Exit Sub
            Dim TempUnit As New cLance(Name)
            If Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances.ContainsKey(Name) Then
                MsgBox("You already have that unit!  Names must be unique!", vbOKOnly)
                Exit Sub
            End If
            Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances.Add(Name, TempUnit)
            Refresh_Lances()
            lstLances.SelectedIndex = lstLances.Items.Count - 1
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdSquadsAdd_Click(sender As Object, e As EventArgs) Handles cmdSquadsAdd.Click
        Try
            If TempCash < CurrentSquad.Cost * 10000 Then
                MsgBox("You do not have enough C-Bills!", vbOKOnly, "Not enough funds!")
                CurrentSquad = Nothing
            Else
                If lstLances.SelectedIndex = -1 Then
                    MsgBox("You need at least 1 lance to add units!")
                    Exit Sub
                End If
                If lstCompanies.SelectedIndex = -1 Then
                    MsgBox("You need at least 1 company to add units!")
                    Exit Sub
                End If
                If lstBattalions.SelectedIndex = -1 Then
                    MsgBox("You need at least 1 battalion to add units!")
                    Exit Sub
                End If
                If lstRegiments.SelectedIndex = -1 Then
                    MsgBox("You need at least 1 regiment to add units!")
                    Exit Sub
                End If
                GetCurrentSquad()

                Dim ShipFound As Boolean = False
                If CurrentSquad.UnitType.ToString.ToLower = "jumpship" Or CurrentSquad.UnitType.ToString.ToLower = "dropship" Then
                    If (CurrentSquad.Quality = eQuality.LikeNew And Availability(CurrentSquad.UnitType) = False) Or (CurrentSquad.Quality = eQuality.Salvaged And Availability_S(CurrentSquad.UnitType) = False) Or (CurrentSquad.Quality = eQuality.Destroyed And Availability_D(CurrentSquad.UnitType) = False) Then
                        MsgBox("You already tried to find this vessel this month, and none could be found.  Try again after unit creation!", vbOKOnly)
                        Exit Sub
                    End If
                    For Each DR As Int32 In Available_Leads()
                        If CurrentSquad.UnitType.ToString.ToLower = "jumpship" Then
                            If Availability_Jumpship(CurrentSquad.UnitType, CurrentSquad.Quality) Then
                                ShipFound = True
                                Exit For
                            End If
                        ElseIf CurrentSquad.UnitType.ToString.ToLower = "dropship" Then
                            If Availability_Dropship(CurrentSquad.UnitType, CurrentSquad.Quality) Then
                                ShipFound = True
                                Exit For
                            End If
                        End If
                    Next
                    If ShipFound = False Then
                        If CurrentSquad.Quality = eQuality.LikeNew Then Availability(CurrentSquad.UnitType) = False
                        If CurrentSquad.Quality = eQuality.Salvaged Then Availability_S(CurrentSquad.UnitType) = False
                        If CurrentSquad.Quality = eQuality.Destroyed Then Availability_D(CurrentSquad.UnitType) = False
                        MsgBox("The vessel could not be found anywhere this month!")
                        Exit Sub
                    End If
                End If

                CurrentSquad.ID = Guid.NewGuid.ToString
                CurrentSquad.Pilot = GetName()
                CurrentSquad.Element = GetUnitString(CurrentSquad)
                Debug.Print("Adding " & CurrentSquad.Element & "/" & CurrentSquad.ID)
                Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads.Add(CurrentSquad.ID, CurrentSquad)
                Refresh_Squads()
                Refresh_Labels_Formation()
                lstSquads.SelectedIndex = lstSquads.Items.Count - 1
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub


    Private Sub Refresh_Regiments()
        Dim OldIndex As Int32 = lstRegiments.SelectedIndex
        With lstRegiments
            .BeginUpdate()
            For Each R As String In Mercenary.Regiments.Keys
                .Items.Add(R)
            Next
            .EndUpdate()
        End With
        lstRegiments.SelectedIndex = Math.Min(lstRegiments.Items.Count - 1, OldIndex)
        Refresh_Labels_Formation()
    End Sub
    Private Sub Refresh_Battalions()
        Dim OldIndex As Int32 = lstBattalions.SelectedIndex
        With lstBattalions
            .BeginUpdate()
            For Each B As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions.Keys
                .Items.Add(B)
            Next
            .EndUpdate()
        End With
        lstBattalions.SelectedIndex = Math.Min(lstBattalions.Items.Count - 1, OldIndex)
        Refresh_Labels_Formation()
    End Sub
    Private Sub Refresh_Companies()
        Dim OldIndex As Int32 = lstCompanies.SelectedIndex
        With lstCompanies
            .BeginUpdate()
            For Each C As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies.Keys
                .Items.Add(C)
            Next
            .EndUpdate()
        End With
        lstCompanies.SelectedIndex = Math.Min(lstCompanies.Items.Count - 1, OldIndex)
        Refresh_Labels_Formation()
    End Sub
    Private Sub Refresh_Lances()
        Dim OldIndex As Int32 = lstLances.SelectedIndex
        With lstLances
            .BeginUpdate()
            For Each L As String In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances.Keys
                .Items.Add(L)
            Next
            .EndUpdate()
        End With
        lstLances.SelectedIndex = Math.Min(lstLances.Items.Count - 1, OldIndex)
        Refresh_Labels_Formation()
    End Sub

    Private Sub Refresh_Squads()
        'Debug.Print("Refresh_Squads")
        Dim OldIndex As Int32 = lstSquads.SelectedIndex
        With lstSquads
            .BeginUpdate()
            .Items.Clear()
            For Each S As cSquad In Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads.Values
                .Items.Add(GetUnitString(S) & Space(100) & S.ID)
            Next
            .EndUpdate()
        End With
        lstSquads.SelectedIndex = Math.Min(lstSquads.Items.Count - 1, OldIndex)
        'If lstSquads.Items.Count = 0 Then labSquad.Text = ""
        Refresh_Labels_Formation()
        Refresh_Label()
    End Sub

    Private Sub cmdRegimentsRemove_Click(sender As Object, e As EventArgs) Handles cmdRegimentsRemove.Click
        Try
            If lstRegiments.Items.Count = 1 Then
                MsgBox("You must have at least one regiment defined", vbOKOnly)
                Exit Sub
            End If
            Dim OldIndex As Int32 = lstRegiments.SelectedIndex
            If lstRegiments.SelectedIndex >= 0 Then
                Mercenary.Regiments.Remove(lstRegiments.SelectedItem)
            End If
            lstRegiments.SelectedIndex = Math.Min(lstRegiments.Items.Count - 1, OldIndex)
            Refresh_Regiments()
            Refresh_Battalions()
            Refresh_Companies()
            Refresh_Lances()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdBattalionsRemove_Click(sender As Object, e As EventArgs) Handles cmdBattalionsRemove.Click
        Try
            If lstBattalions.Items.Count = 1 Then
                MsgBox("You must have at least one battalion defined", vbOKOnly)
                Exit Sub
            End If
            Dim OldIndex As Int32 = lstBattalions.SelectedIndex
            If lstBattalions.SelectedIndex >= 0 Then
                Mercenary.Regiments(lstRegiments.SelectedItem).Battalions.Remove(lstBattalions.SelectedItem)
            End If
            lstBattalions.SelectedIndex = Math.Min(lstBattalions.Items.Count - 1, OldIndex)
            Refresh_Battalions()
            Refresh_Companies()
            Refresh_Lances()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdCompaniesRemove_Click(sender As Object, e As EventArgs) Handles cmdCompaniesRemove.Click
        Try
            If lstCompanies.Items.Count = 1 Then
                MsgBox("You must have at least one company defined", vbOKOnly)
                Exit Sub
            End If
            Dim OldIndex As Int32 = lstCompanies.SelectedIndex
            If lstCompanies.SelectedIndex >= 0 Then
                Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies.Remove(lstCompanies.SelectedItem)
            End If
            lstCompanies.SelectedIndex = Math.Min(lstCompanies.Items.Count - 1, OldIndex)
            Refresh_Companies()
            Refresh_Lances()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdLancesRemove_Click(sender As Object, e As EventArgs) Handles cmdLancesRemove.Click
        Try
            If lstLances.Items.Count = 1 Then
                MsgBox("You must have at least one lance defined", vbOKOnly)
                Exit Sub
            End If
            Dim OldIndex As Int32 = lstLances.SelectedIndex
            If lstLances.SelectedIndex >= 0 Then
                Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances.Remove(lstLances.SelectedItem)
            End If
            lstLances.SelectedIndex = Math.Min(lstLances.Items.Count - 1, OldIndex)
            Refresh_Lances()
            Refresh_Squads()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdSquadsRemove_Click(sender As Object, e As EventArgs) Handles cmdSquadsRemove.Click
        Try
            If lstSquads.SelectedIndex > -1 Then
                Dim OldIndex As Int32 = lstSquads.SelectedIndex
                Dim TempVal() As String = Split(lstSquads.SelectedItem.ToString, Space(100))
                Debug.Print("Removing " & TempVal(0) & " - " & TempVal(1))
                Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads.Remove(TempVal(1))

                If Mercenary.Regiments(lstRegiments.SelectedItem).Battalions(lstBattalions.SelectedItem).Companies(lstCompanies.SelectedItem).Lances(lstLances.SelectedItem).Squads.Count = 0 Then
                    lstSquads.Items.Clear()
                    labSquad.Text = ""
                Else
                    lstSquads.SelectedIndex = Math.Min(lstSquads.Items.Count - 1, OldIndex)
                End If
            End If
            Refresh_Squads()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

#End Region

End Class

