Public Module mCompaignOperations
    Public Function AerospaceOperations(FS As Single, ES As Single) As eAerospaceOperationsResult
        Try
            DebugEvent("Aerospace Operations: ")
            DebugEvent("  Friendly Aerospace: " & FS)
            DebugEvent("  Enemy Aerospace: " & ES)
            Dim TempResult As eAerospaceOperationsResult = eAerospaceOperationsResult.A
            Dim DieRoll As Int32 = D(2, 6)
            DebugEvent("Rolled " & DieRoll & " in Aerospace Operations phase.")
            Select Case FS / ES
                Case >= 3 : TempResult = Choose(DieRoll - 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                Case >= 2.5 : TempResult = Choose(DieRoll - 1, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1)
                Case >= 2 : TempResult = Choose(DieRoll - 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1)
                Case >= 1.5 : TempResult = Choose(DieRoll - 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1)
                Case >= 1 : TempResult = Choose(DieRoll - 1, -1, -1, 0, 0, 0, 0, 0, 0, 0, 1, 1)
                Case >= 1 / 1.5 : TempResult = Choose(DieRoll - 1, -1, -1, -1, -1, 0, 0, 0, 0, 0, 0, 0)
                Case >= 0.5 : TempResult = Choose(DieRoll - 1, -1, -1, -1, -1, -1, -1, 0, 0, 0, 0, 0)
                Case >= 1 / 2.5 : TempResult = Choose(DieRoll - 1, -1, -1, -1, -1, -1, -1, -1, -1, 0, 0, 0)
                Case Else : TempResult = Choose(DieRoll - 1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 0) '1/3
            End Select
            Return TempResult
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Function
    Public Function ManeuverOperations(FS As Single, ES As Single, FriendlyTactics As Int32, EnemyTactics As Int32) As eManeuverOperationsResult
        Try
            DebugEvent("Maneuver Operations: ")
            DebugEvent("  Friendly Maneuver: " & FS & " with a tactics skill of " & FriendlyTactics)
            DebugEvent("  Enemy Maneuver: " & ES & " with a tactics skill of " & EnemyTactics)

            Dim DieRoll As Int32 = D(2, 6) - 1
            DebugEvent("Rolled " & DieRoll & " in Maneuver Operations phase.")
            DieRoll += FriendlyTactics - EnemyTactics
            If DieRoll < 2 Then DieRoll = 2
            If DieRoll > 12 Then DieRoll = 12
            DieRoll -= 1
            DebugEvent("  This roll was modified to " & DieRoll)
            Dim TempResult As eManeuverOperationsResult = eManeuverOperationsResult.Att
            Select Case FS / ES
                Case >= 3 : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Con, eManeuverOperationsResult.DecF, eManeuverOperationsResult.DecF, eManeuverOperationsResult.DecF)
                Case >= 2.5 : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Con, eManeuverOperationsResult.Con, eManeuverOperationsResult.DecF, eManeuverOperationsResult.DecF)
                Case >= 2 : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Con, eManeuverOperationsResult.DecF, eManeuverOperationsResult.DecF)
                    TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.Con, eManeuverOperationsResult.Con, eManeuverOperationsResult.DecF)
                    TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Con, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.Con, eManeuverOperationsResult.DecF)
                Case >= 1 / 1.5 : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Con, eManeuverOperationsResult.Con, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Att, eManeuverOperationsResult.DecF)
                    TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Con, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.Sk, eManeuverOperationsResult.DecF)
                Case >= 1 / 2.5 : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Con, eManeuverOperationsResult.Con, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.Bat, eManeuverOperationsResult.DecF)
                Case Else : TempResult = Choose(DieRoll, eManeuverOperationsResult.DecE, eManeuverOperationsResult.DecE, eManeuverOperationsResult.DecE, eManeuverOperationsResult.Con, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Att, eManeuverOperationsResult.Sk, eManeuverOperationsResult.Bat, eManeuverOperationsResult.NP, eManeuverOperationsResult.DecF) '1/3
            End Select
            Return TempResult
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Function
    Public Sub CombatOperations(FriendlySquads As Int32, EnemySquads As Int32, DecisiveLoserDidNotAcceptHonor As Boolean, FriendlyAS As Int64, EnemyAS As Int64)
        Try
            DebugEvent("Combat Operations: ")
            DebugEvent("  Friendly Squads: " & FriendlySquads & " with " & FriendlyAS & " Aerospace")
            DebugEvent("  Enemy Squads: " & EnemySquads & " with " & EnemyAS & " Aerospace")

            Dim AttackerResult As Int32 = 0
            Dim DefenderResult As Int32 = 0
            Dim CombatResult As Int32 = D(2, 6) - 1
            DebugEvent("Rolled " & CombatResult & " in Combat Operations phase.")
            Dim TempCol As Int32 = 0
            Select Case (Friendly.Combat + FriendlyAS) / (Enemy.Combat + EnemyAS)
                Case >= 3 : TempCol = 1
                Case >= 2.5 : TempCol = 2
                Case >= 2 : TempCol = 3
                Case >= 1.5 : TempCol = 4
                Case >= 1 : TempCol = 5
                Case >= 1 / 1.5 : TempCol = 6
                Case >= 0.5 : TempCol = 7
                Case >= 1 / 2.5 : TempCol = 8
                Case Else : TempCol = 9 '1/3
            End Select
            If DecisiveLoserDidNotAcceptHonor Then TempCol += 1
            DebugEvent("  Column Used: " & TempCol)

            Select Case TempCol
                Case 1
                    AttackerResult = Choose(CombatResult, 7, 5, 4, 3, 2, 2, 1, 1, 0, 0, 0)
                    DefenderResult = Choose(CombatResult, 3, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9)
                Case 2
                    AttackerResult = Choose(CombatResult, 7, 5, 5, 4, 3, 2, 2, 1, 1, 0, 0)
                    DefenderResult = Choose(CombatResult, 3, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9)
                Case 3
                    AttackerResult = Choose(CombatResult, 7, 6, 5, 5, 4, 3, 2, 2, 1, 1, 0)
                    DefenderResult = Choose(CombatResult, 2, 3, 4, 5, 5, 6, 6, 7, 7, 8, 9)
                Case 4
                    AttackerResult = Choose(CombatResult, 8, 7, 6, 5, 5, 4, 3, 2, 2, 1, 1)
                    DefenderResult = Choose(CombatResult, 2, 2, 3, 4, 5, 5, 6, 6, 7, 7, 8)
                Case 5
                    AttackerResult = Choose(CombatResult, 8, 7, 6, 6, 5, 5, 4, 3, 2, 2, 1)
                    DefenderResult = Choose(CombatResult, 1, 2, 2, 3, 4, 5, 5, 6, 6, 7, 8)
                Case 6
                    AttackerResult = Choose(CombatResult, 8, 7, 7, 6, 6, 5, 5, 4, 3, 2, 2)
                    DefenderResult = Choose(CombatResult, 1, 1, 2, 2, 3, 4, 5, 5, 6, 7, 8)
                Case 7
                    AttackerResult = Choose(CombatResult, 9, 8, 7, 7, 6, 6, 5, 5, 4, 3, 2)
                    DefenderResult = Choose(CombatResult, 0, 1, 1, 2, 2, 3, 4, 5, 5, 6, 7)
                Case 8
                    AttackerResult = Choose(CombatResult, 9, 8, 8, 7, 7, 6, 6, 5, 5, 4, 3)
                    DefenderResult = Choose(CombatResult, 0, 0, 1, 1, 2, 2, 3, 4, 5, 5, 7)
                Case 9
                    AttackerResult = Choose(CombatResult, 9, 9, 8, 8, 7, 7, 6, 6, 5, 5, 3)
                    DefenderResult = Choose(CombatResult, 0, 0, 0, 1, 1, 2, 2, 3, 4, 5, 7)
            End Select

            DebugEvent("  Resolving Friendly Combat Operations:")
            ResolveCombatOperation(Friendly, AttackerResult, FriendlySquads)
            DebugEvent("  Resolving Enemy Combat Operations:")
            ResolveCombatOperation(Enemy, DefenderResult, EnemySquads)
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Public Sub ResolveCombatOperation(Unit As cForce, CombatResults As Int32, EngagedSquads As Int32)
        Try
            Dim Salvaged As Int32 = 0 'Number per company, use 1/3 for squad, 2/3 for lance/platoon, 3 for bn, 9 for regiment or per regiment
            Dim Destroyed As Int32 = 0 'Number per company, use 1/3 for squad, 2/3 for lance/platoon, 3 for bn, 9 for regiment or per regiment
            Select Case CombatResults
                Case 0 'Win, Can Salvage
                    Unit.ReputationMod += 5
                    Salvaged = D(1, 6) / 3
                    Destroyed = D(1, 6) / 3
                Case 1 'Win, Can Salvage
                    Unit.ReputationMod += 3
                    Salvaged = D(1, 6) / 2
                    Destroyed = D(1, 6) / 2
                Case 2 'Win, Can Salvage
                    Unit.ReputationMod += 2
                    Salvaged = D(1, 6) / 2
                    Destroyed = D(1, 6) / 2
                Case 3 'Draw, Can Salvage
                    Unit.ReputationMod += 1
                    Salvaged = D(1, 6)
                    Destroyed = D(1, 6)
                Case 4 'Draw, Can Salvage
                    Unit.ReputationMod += 0
                    Salvaged = D(1, 6)
                    Destroyed = D(1, 6)
                Case 5 'Draw
                    Unit.ReputationMod -= 0
                    Salvaged = D(1, 6) + 1
                    Destroyed = D(1, 6) + 1
                Case 6 'Draw
                    Unit.ReputationMod -= 1
                    Salvaged = D(1, 6) + 2
                    Destroyed = D(1, 6) + 2
                Case 7 'Loss
                    Unit.ReputationMod -= 2
                    Salvaged = D(1, 6) + 2
                    Destroyed = D(1, 6) + 2
                Case 8 'Loss
                    Unit.ReputationMod -= 3
                    Salvaged = D(1, 6) + 3
                    Destroyed = D(1, 6) + 3
                Case 9 'Loss
                    Unit.ReputationMod -= 5
                    Salvaged = D(2, 6)
                    Destroyed = D(2, 6)
            End Select

            'Now multiply by size
            'TODO: Option to make this more precise
            Select Case EngagedSquads
                Case 1 : Salvaged *= 1 / 3 : Destroyed *= 1 / 3
                Case 4 : Salvaged *= 2 / 3 : Destroyed *= 2 / 3
                Case 12
                Case 36 : Salvaged *= 3 : Destroyed *= 3
                Case 108 : Salvaged *= 9 : Destroyed *= 9
                Case > 108
                    Salvaged *= 9 * Math.Round(EngagedSquads / 108, MidpointRounding.AwayFromZero)
                    Destroyed *= 9 * Math.Round(EngagedSquads / 108, MidpointRounding.AwayFromZero)
            End Select

            DebugEvent("  Combat Results:")
            DebugEvent("    Reputation Modifier: " & Unit.ReputationMod)
            DebugEvent("    Engaged Squads: " & EngagedSquads)
            DebugEvent("    Salvaged: " & Salvaged)
            DebugEvent("    Destroyed: " & Destroyed)


            'Here is where I randomly determine units, up to the squad count, that are impacted

            'Add all the regiments, battalions, companies, and lances to a list
            Dim List_Regiments As New List(Of cRegiment)
            Dim List_Battalions As New List(Of cBattalion)
            Dim List_Companies As New List(Of cCompany)
            Dim List_Lances As New List(Of cLance)
            Dim List_Squads As New List(Of cLance)
            For Each R As cRegiment In Unit.Regiments.Values
                List_Regiments.Add(R)
            Next
            For Each R As cRegiment In Mercenary.Regiments.Values
                List_Regiments.Add(R)
            Next

            'Now add all battalions to a list
            Dim TotalSquadCount As Int64 = 0
            For Each R As cRegiment In List_Regiments
                For Each B As cBattalion In R.Battalions.Values
                    List_Battalions.Add(B)
                    For Each C As cCompany In B.Companies.Values
                        List_Companies.Add(C)
                        For Each L As cLance In C.Lances.Values
                            List_Lances.Add(L)
                            TotalSquadCount += L.Squads.Count
                        Next
                    Next
                Next
            Next

            DebugEvent("  Engaged Squads:" & EngagedSquads)
            DebugEvent("  Total Squads:" & TotalSquadCount)


            Dim Regiments_Engaged As New List(Of cRegiment)
            Dim Regiments_Battalions As New List(Of cBattalion)
            Dim Regiments_Companies As New List(Of cCompany)
            Dim Regiments_Lances As New List(Of cLance)
            Dim Regiments_Squads As New List(Of cLance)
            Do
                Dim TempRoll As Int32 = 0
                If EngagedSquads > 108 Then 'Regiment(s)
                    Dim DieRoll As Int32 = D(1, List_Regiments.Count)
                    Regiments_Engaged.Add(List_Regiments(DieRoll))
                    EngagedSquads -= List_Regiments(DieRoll).SquadCount
                    List_Regiments.Remove(List_Regiments(DieRoll))
                    For Each B As cBattalion In List_Regiments(DieRoll).Battalions.Values
                        List_Battalions.Remove(B)
                        For Each C As cCompany In B.Companies.Values
                            List_Companies.Remove(C)
                            For Each L As cLance In C.Lances.Values
                                List_Lances.Remove(L)
                            Next
                        Next
                    Next
                ElseIf EngagedSquads > 36 Then 'Regiment
                    Dim DieRoll As Int32 = D(1, List_Battalions.Count)
                    Regiments_Battalions.Add(List_Battalions(DieRoll))
                    EngagedSquads -= List_Battalions(DieRoll).SquadCount
                    List_Battalions.Remove(List_Battalions(DieRoll))
                    For Each C As cCompany In List_Battalions(DieRoll).Companies.Values
                        List_Companies.Remove(C)
                        For Each L As cLance In C.Lances.Values
                            List_Lances.Remove(L)
                        Next
                    Next
                ElseIf EngagedSquads > 12 Then 'Battalion
                    Dim DieRoll As Int32 = D(1, List_Companies.Count)
                    Regiments_Companies.Add(List_Companies(DieRoll))
                    EngagedSquads -= List_Companies(DieRoll).SquadCount
                    List_Companies.Remove(List_Companies(DieRoll))
                    For Each L As cLance In List_Companies(DieRoll).Lances.Values
                        List_Lances.Remove(L)
                    Next
                ElseIf EngagedSquads > 4 Then 'Lance
                    Dim DieRoll As Int32 = D(1, List_Lances.Count)
                    Regiments_Lances.Add(List_Lances(DieRoll))
                    EngagedSquads -= List_Lances(DieRoll).SquadCount
                    List_Lances.Remove(List_Lances(DieRoll))
                ElseIf EngagedSquads <= 4 Then 'Squads
                    Dim DieRoll As Int32 = D(1, List_Squads.Count)
                    Regiments_Lances.Add(List_Squads(DieRoll))
                    EngagedSquads -= List_Squads(DieRoll).SquadCount
                    Regiments_Squads.Remove(List_Squads(DieRoll))
                End If
                If EngagedSquads <= 0 Then Exit Do
            Loop

            DebugEvent("  Engaged Units:")
            DebugEvent("    Regiments: " & Regiments_Engaged.Count)
            DebugEvent("    Battalions: " & Regiments_Battalions.Count)
            DebugEvent("    Companies: " & Regiments_Companies.Count)
            DebugEvent("    Lances: " & Regiments_Lances.Count)

        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Public Sub CampaignOutcome(Unit As cForce, Mission As eMission, DecisiveWin As Boolean, DecisiveLoss As Boolean, DishonorableFriendly As Boolean, DishonorableEnemy As Boolean)
        Try
            Dim DieRoll As Int32 = D(1, 6)
            DebugEvent("Campaign Outcome:")
            DebugEvent("  Dieroll: " & DieRoll & " + " & Leader.Skills(eSkill.Tactics))
            DieRoll += Leader.Skills(eSkill.Tactics)
            If DecisiveWin Then
                DieRoll += 6
                DebugEvent("    +6 for Decisive Victory")
            End If
            If DecisiveLoss Then
                DieRoll -= 6
                DebugEvent("    -6 for Decisive Loss")
            End If
            DebugEvent("    Final Roll:  " & DieRoll) 'Range is between -5 and 18
            Dim Result As Int32 = 0
            DieRoll += 6
            DieRoll = Math.Min(DieRoll, 24)
            Select Case Mission
                Case eMission.Cadre_Duty : Result = Choose(DieRoll, 12, 12, 12, 11, 11, 10, 10, 10, 9, 9, 9, 8, 8, 7, 7, 7, 6, 6, 6, 5, 5, 4, 4, 4)
                Case eMission.Garrison_Duty : Result = Choose(DieRoll, 16, 15, 14, 13, 12, 11, 11, 10, 10, 9, 9, 8, 8, 7, 7, 6, 6, 5, 5, 5, 4, 4, 3, 2)
                Case eMission.Security_Duty : Result = Choose(DieRoll, 16, 15, 14, 13, 12, 11, 11, 10, 10, 9, 8, 8, 8, 8, 7, 7, 6, 6, 5, 4, 3, 2, 1, 0)
                Case eMission.Offensive_Campaign : Result = Choose(DieRoll, 16, 15, 14, 14, 13, 13, 12, 12, 11, 10, 9, 8, 8, 7, 6, 5, 4, 4, 3, 3, 2, 2, 1, 0)
                Case eMission.Defensive_Campaign : Result = Choose(DieRoll, 16, 15, 15, 14, 14, 13, 13, 12, 12, 11, 10, 9, 8, 7, 6, 5, 5, 4, 3, 3, 2, 2, 1, 1)
                Case eMission.Planetary_Assault : Result = Choose(DieRoll, 16, 16, 15, 15, 14, 14, 13, 13, 12, 12, 10, 9, 8, 7, 6, 5, 4, 4, 3, 2, 1, 1, 0, 0)
                Case eMission.Relief_Duty : Result = Choose(DieRoll, 16, 16, 15, 15, 14, 14, 13, 13, 12, 12, 11, 10, 9, 8, 7, 6, 5, 4, 4, 3, 2, 1, 1, 0)
                Case eMission.Riot_Duty : Result = Choose(DieRoll, 15, 15, 14, 14, 13, 13, 13, 12, 12, 12, 11, 10, 9, 8, 7, 6, 5, 5, 4, 4, 3, 3, 2, 2)
                Case eMission.Siege_Duty : Result = Choose(DieRoll, 16, 15, 14, 13, 12, 11, 11, 10, 10, 9, 9, 8, 8, 7, 7, 6, 6, 5, 5, 4, 3, 2, 1, 0)
                Case eMission.Recon_Raid : Result = Choose(DieRoll, 16, 16, 15, 14, 13, 12, 12, 11, 10, 9, 9, 8, 8, 7, 7, 6, 5, 4, 4, 3, 2, 1, 0, 0)
                Case eMission.Objective_Raid : Result = Choose(DieRoll, 16, 15, 14, 14, 13, 12, 12, 11, 10, 9, 9, 8, 8, 7, 7, 6, 5, 4, 4, 3, 2, 2, 1, 0)
                Case eMission.Divisionary_Raid : Result = Choose(DieRoll, 16, 16, 15, 15, 14, 14, 13, 13, 12, 12, 11, 11, 10, 10, 9, 8, 7, 6, 5, 4, 4, 3, 2, 1)
                Case eMission.Guerrilla_Warfare : Result = Choose(DieRoll, 16, 15, 14, 14, 13, 13, 13, 12, 12, 12, 11, 11, 10, 10, 9, 9, 8, 8, 7, 6, 5, 4, 4, 3)
            End Select

            Dim Danger As Int32 = 0 'All players and important NPCs must roll this number or higher, adding morale.  If unsuccessful roll against BOD to avoid 6d6 Damage
            Dim Salvaged As Int32 = 0 'Must roll or higher on each sizemult, adding morale, to have 1d6*sizemult reduced to salvage
            Dim Destroyed As Int32 = 0 'Must roll or higher on each sizemult, adding morale, to have 1d6*sizemult reduced to salvage
            Dim Plunder As Int32 = 0 'Roll 2d6 or less, subtracting morale from roll, to get plunder.  Plunder is 2d6*10000SizeMult
            Dim MoraleMod As Int32 = 0
            Dim ReputationMod As Int32 = 0

            Select Case Result
                Case 0 'A
                    ReputationMod += 6
                    Danger = 5
                    Salvaged = 6
                    Destroyed = 5
                    Plunder = 10
                    MoraleMod += 1'Won
                Case 1 'B
                    ReputationMod += 5
                    Danger = 5
                    Salvaged = 6
                    Destroyed = 5
                    Plunder = 10
                    MoraleMod += 1'Won
                Case 2 'C
                    ReputationMod += 4
                    Danger = 5
                    Salvaged = 6
                    Destroyed = 5
                    Plunder = 9
                    MoraleMod += 1'Won
                Case 3 'D
                    ReputationMod += 3
                    Danger = 6
                    Salvaged = 7
                    Destroyed = 6
                    Plunder = 9
                    MoraleMod += 1'Won
                Case 4 'E
                    ReputationMod += 2
                    Danger = 6
                    Salvaged = 7
                    Destroyed = 6
                    Plunder = 8
                    MoraleMod += 1'Won
                Case 5 'F
                    ReputationMod += 1
                    Danger = 6
                    Salvaged = 7
                    Destroyed = 6
                    Plunder = 7
                Case 6 'G
                    Danger = 7
                    Salvaged = 8
                    Destroyed = 7
                    Plunder = 7
                Case 7 'H
                    Danger = 7
                    Salvaged = 8
                    Destroyed = 7
                    Plunder = 6
                Case 8 'I
                    Danger = 7
                    Salvaged = 8
                    Destroyed = 7
                    Plunder = 6
                Case 9 'J
                    Danger = 7
                    Salvaged = 8
                    Destroyed = 7
                    Plunder = 5
                Case 10 'K
                    Danger = 7
                    Salvaged = 8
                    Destroyed = 7
                    Plunder = 5
                Case 11 'L
                    ReputationMod -= 1
                    Danger = 8
                    Salvaged = 9
                    Destroyed = 8
                    Plunder = 5
                Case 12 'M
                    ReputationMod -= 2
                    Danger = 8
                    Salvaged = 9
                    Destroyed = 8
                    Plunder = 4
                    MoraleMod -= 1'Loss
                Case 13 'N
                    ReputationMod -= 3
                    Danger = 8
                    Salvaged = 9
                    Destroyed = 8
                    Plunder = 4
                    MoraleMod -= 1'Loss
                Case 14 'O
                    ReputationMod -= 4
                    Danger = 9
                    Salvaged = 10
                    Destroyed = 9
                    Plunder = 4
                    MoraleMod -= 1'Loss
                Case 15 'P
                    ReputationMod -= 5
                    Danger = 9
                    Salvaged = 10
                    Destroyed = 9
                    Plunder = 4
                    MoraleMod -= 1'Loss
                Case 16 'Q
                    ReputationMod -= 6
                    Danger = 9
                    Salvaged = 10
                    Destroyed = 9
                    Plunder = 4
                    MoraleMod -= 1 'Loss
            End Select
            DebugEvent("  Final Result:")
            DebugEvent("    Current Reputation Modifier: " & ReputationMod)
            DebugEvent("    Current Morale Modifier: " & MoraleMod)
            DebugEvent("    Danger: " & Danger)
            DebugEvent("    Salvaged: " & Salvaged)
            DebugEvent("    Destroyed: " & Destroyed)
            DebugEvent("    Plunder: " & Plunder)


            Dim TotalPlunder As Int64 = 0
            For Each R As cRegiment In Forces_Friendly
                Dim PlayingItSafeMod As Int32 = R.PlayingItSafe

                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        DebugEvent("      Company: " & C.Name & " Status: ")

                        Dim SalvagedUnits As Int32 = 0
                        Dim DestroyedUnits As Int32 = 0
                        Dim PlunderedUnits As Int32 = 0 'TODO:  No idea what to do with this #, maybe get their cost???

                        If D(2, 6) - (C.TotalMorale / (C.SquadCount - C.SupportSquads)) - MoraleMod < Plunder Then
                            TotalPlunder += D(2, 6) - PlayingItSafeMod
                        End If


                        If D(2, 6) + (C.TotalMorale / (C.SquadCount - C.SupportSquads)) + MoraleMod - PlayingItSafeMod < Destroyed Then
                            DestroyedUnits = D(1, 6)
                        End If

                        PlunderedUnits = Math.Max(PlunderedUnits, 0)
                        DestroyedUnits = Math.Max(DestroyedUnits, 0)
                        Do
                            Dim Total As Int32 = 0
                            For Each L As cLance In C.Lances.Values
                                For Each S As cSquad In L.Squads.Values
                                    If S.State = eSquadState.Active And S.Quality <> eQuality.Destroyed Then Total += 1
                                    If D(1, 100) / 100 >= (DestroyedUnits / L.SquadCount) Then
                                        If S.State = eSquadState.Active And S.Quality <> eQuality.Destroyed Then
                                            S.Quality = eQuality.Destroyed
                                            DestroyedUnits -= 1
                                            DebugEvent("      Squad Destroyed:" & S.UnitType.ToString)
                                            If DestroyedUnits <= 0 Then Exit Do
                                        End If
                                    End If
                                Next
                            Next
                            If DestroyedUnits >= Total Then Exit Do 'Not enough units left
                        Loop

                        If D(2, 6) + (C.TotalMorale / (C.SquadCount - C.SupportSquads)) + MoraleMod - PlayingItSafeMod < Salvaged Then
                            SalvagedUnits = D(1, 6)
                        End If
                        SalvagedUnits = Math.Max(SalvagedUnits, 0)
                        Do
                            Dim Total As Int32 = 0
                            For Each L As cLance In C.Lances.Values
                                For Each S As cSquad In L.Squads.Values
                                    If S.State = eSquadState.Active And S.Quality <> eQuality.Destroyed And S.Quality <> eQuality.Salvaged Then Total += 1
                                    If D(1, 100) / 100 >= (SalvagedUnits / L.SquadCount) Then
                                        If S.State = eSquadState.Active And S.Quality <> eQuality.Destroyed And S.Quality <> eQuality.Salvaged Then
                                            S.Quality = eQuality.Salvaged
                                            SalvagedUnits -= 1
                                            DebugEvent("      Squad Salvaged:" & S.UnitType.ToString)
                                            If SalvagedUnits <= 0 Then Exit Do
                                        End If
                                    End If
                                Next
                            Next
                            If SalvagedUnits >= Total Then Exit Do 'Not enough units left
                        Loop


                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If D(2, 6) + S.Morale + MoraleMod - PlayingItSafeMod < Danger Then
                                    DebugEvent("      Squad Must Save vs BOD to avoid 6d6 HTK:" & S.UnitType.ToString)
                                    'TODO:  Must save vs BOD to avoid taking 6d6 HTK
                                End If
                            Next
                        Next

                    Next
                Next
            Next
            TotalPlunder *= 10000
            MsgBox("  You plundered €" & Format(TotalPlunder, "###,##0") & " worth of equipment, ransomed prisoners, and so forth.", MsgBoxStyle.OkOnly) 'TODO:  This is being added TWICE!
            Mercenary.Cash_On_Hand += TotalPlunder
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

End Module
