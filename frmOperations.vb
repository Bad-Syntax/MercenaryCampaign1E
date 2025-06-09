Public Class frmOperations
    Dim LastMajorEvent As eEvent = eEvent.None

    Private Sub frmOperations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RefreshLabels()
    End Sub
    Public Sub RefreshLabels()
        Try
            Dim TempText As String = ""
            TempText = "Unit Status:" & vbCrLf
            Dim TempSquads As Int32 = Mercenary.SquadCount
            Dim TempSupport As Int32 = Mercenary.SupportSquads
            TempText &= "Squads (Combat/Support):  " & TempSquads & "/" & TempSupport & vbCrLf
            TempText &= "Morale:  " & Format(Mercenary.ForceMorale, "0.00") & vbCrLf '   Morale (Reduction due to overhead)
            TempText &= "Reputation:  " & Format(Math.Round(Mercenary.Reputation / 12, 0), "##0") & vbCrLf
            TempText &= "Unit Cost:  €" & Format(Mercenary.Cost, "###,##0") & vbCrLf
            TempText &= "Overhead (" & Mercenary.Overhead & "%):  €" & Format((Mercenary.SquadCount * 500 + Mercenary.Salaries) * (Mercenary.Overhead / 100), "###,##0") & vbCrLf
            TempText &= "Cash on Hand:  €" & Format(Mercenary.Cash_On_Hand, "###,##0") & vbCrLf
            TempText &= "Salary Cost:  €" & Format(Mercenary.Salaries, "###,##0") & vbCrLf
            TempText &= "Maintenance Cost:  €" & Format(Mercenary.Maintenance_Cost, "###,##0") & vbCrLf
            TempText &= "Supply Cost:  €" & Format(Mercenary.SquadCount * 500, "###,##0") & vbCrLf
            TempText &= "Supply Orders Pending:  " & Mercenary.Supplies_Pending & " supply units in " & Mercenary.Supply_Orders.Count & " Orders" & vbCrLf
            TempText &= "Supply On Hand:  " & Format(Mercenary.Supplies_On_Hand, "###,##0") & vbCrLf
            TempText &= "Aerospace Value:  " & Format(Mercenary.Aerospace, "###,##0") & vbCrLf 'Reduction due to lack of maintenance
            TempText &= "Maneuver Value:  " & Format(Mercenary.Maneuver, "###,##0") & vbCrLf 'Reduction due to lack of maintenance
            TempText &= "Combat Value:  " & Format(Mercenary.Combat, "###,##0") & vbCrLf 'Reduction due to lack of maintenance
            TempText &= "Dragoon Rating:  " & Mercenary.DragoonRating.ToString & vbCrLf
            labUnitStatus.Text = TempText

            If CurrentContract Is Nothing Then
                TempText = "No Contract"
                cmdFindContract.Enabled = True
            Else
                If CurrentContract.Active Then
                    TempText = "Contract Status:  " & IIf(CurrentContract.Length_Remaining > 0, "Active", "Completed") & vbCrLf
                    TempText &= "Mission:  " & CurrentContract.Mission.ToString & vbCrLf
                    TempText &= "Length:  " & CurrentContract.Length & " Months" & vbCrLf
                    TempText &= "Months Remaining:  " & CurrentContract.Length_Remaining & vbCrLf
                    TempText &= "Remuneration:  €" & Format(CurrentContract.Remuneration * Mercenary.SquadCount * 4, "###,##0") & vbCrLf
                    'TODO:  Advance and Fees
                    TempText &= "Guarantee:  " & CurrentContract.Guarantees.ToString & vbCrLf
                    TempText &= "Command Rights:  " & CurrentContract.Command.ToString & vbCrLf
                    TempText &= "Transport Fees:  " & CurrentContract.TransportFee * 100 & "%" & vbCrLf
                    TempText &= "Supply Fees:  €" & Format(CurrentContract.SupplyFee * Mercenary.SquadCount * 500, "###,##0") & " Per Month" & vbCrLf
                    TempText &= "Salvage Rights:  " & CurrentContract.Salvage.ToString & vbCrLf
                Else
                    TempText = "No Contract"
                End If
            End If

            labContract.Text = TempText


            TempText = "Friendly Forces:" & vbCrLf
            For Each R As cRegiment In Forces_Friendly
                DebugEvent(ShowRegiment(R), False)

                If R.Name = "No Regiment" Then
                    For Each B As cBattalion In R.Battalions.Values
                        If B.Name = "No Battalion" Then
                            For Each C As cCompany In B.Companies.Values
                                If C.Name = "No Company" Then
                                    For Each L As cLance In C.Lances.Values
                                        Dim Strength As Int32 = Math.Round(L.CountByState(eSquadState.Active) / L.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                                        Dim Wounded As Int32 = Math.Round(L.CountByState(eSquadState.Wounded) / L.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                                        TempText &= L.Name & " (" & Strength & "%/" & Wounded & "%)" & vbCrLf 'Add % Strength / % Wounded
                                    Next
                                Else
                                    Dim Strength As Int32 = Math.Round(C.CountByState(eSquadState.Active) / C.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                                    Dim Wounded As Int32 = Math.Round(C.CountByState(eSquadState.Wounded) / C.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                                    TempText &= C.Name & " (" & Strength & "%/" & Wounded & "%)" & vbCrLf 'Add % Strength / % Wounded
                                End If
                            Next
                        Else
                            Dim Strength As Int32 = Math.Round(B.CountByState(eSquadState.Active) / B.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                            Dim Wounded As Int32 = Math.Round(B.CountByState(eSquadState.Wounded) / B.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                            TempText &= B.Name & " (" & Strength & "%/" & Wounded & "%)" & vbCrLf 'Add % Strength / % Wounded
                        End If
                    Next
                Else
                    Dim Strength As Int32 = Math.Round(R.CountByState(eSquadState.Active) / R.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                    Dim Wounded As Int32 = Math.Round(R.CountByState(eSquadState.Wounded) / R.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                    TempText &= R.Name & " (" & Strength & "%/" & Wounded & "%)" & vbCrLf 'Add % Strength / % Wounded
                End If
            Next
            TempText &= vbCrLf & "Enemy Forces:" & vbCrLf
            For Each R As cRegiment In Forces_Enemy
                Dim Strength As Int32 = Math.Round(R.CountByState(eSquadState.Active) / R.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                Dim Wounded As Int32 = Math.Round(R.CountByState(eSquadState.Wounded) / R.OriginalSize * 100, 0, MidpointRounding.AwayFromZero)
                TempText &= R.Name & " (" & Strength & "%/" & Wounded & "%)" & vbCrLf
                DebugEvent(ShowRegiment(R), False)
            Next
            labForces.Text = TempText

            Me.Text = "Mercenary Operations For " & CurrentTurn.ToLongDateString

            'Now update my txt box
            txtEvents.Text = Join(EventList.ToArray, vbCrLf)
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

    Private Sub cmdSupplies_Click(sender As Object, e As EventArgs) Handles cmdSupplies.Click
        Try
            Dim SupplyAmount As Int64 = "0" & InputBox("Enter the amount of supplies you wish to order", "Order Supplies", Mercenary.SquadCount)
            If SupplyAmount * 500 > Mercenary.Cash_On_Hand Then
                MsgBox("You do not have that much money!", MsgBoxStyle.Critical, "Insufficient Funds!")
                Exit Sub
            ElseIf SupplyAmount = 0 Then
                MsgBox("So you do NOT want supplies then, got it!", MsgBoxStyle.Critical, "I want NOTHING!")
                Exit Sub
            ElseIf SupplyAmount > 0 Then
                DebugEvent(SupplyAmount * 500 & " C-Bills worth of supplies ordered", True)
                Mercenary.Cash_On_Hand -= SupplyAmount * 500
                Mercenary.SupplyID += 1
                Dim SupplyOrder As New cSupplyOrder
                SupplyOrder.OrderID = Mercenary.SupplyID
                SupplyOrder.Supplies = SupplyAmount
                SupplyOrder.Delay = D(1, 6) - 1
                If D(1, 6) = 6 Then SupplyOrder.Lost = True Else SupplyOrder.Lost = False
                Mercenary.Supply_Orders.Add(SupplyOrder)
                RefreshLabels()
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

    Private Sub RecoverWounded()
        Try
            'If CurrentContract.Active Then
            For Each R As cRegiment In Forces_Friendly
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If S.State = eSquadState.Wounded Then
                                    DebugEvent(S.UnitType.ToString & " recovered from wounds received.", True)
                                    S.State = eSquadState.Active
                                End If
                            Next
                        Next
                    Next
                Next
            Next
            'End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Pillage()
        Try
            'Pillage the Civilians.  Normally should only have 1 value, but I wanted to track it per company
            Dim Pillage As Int64 = 0
            Dim PillageMax As Int64 = 0
            Dim ForcesPillaging As Int64 = 0
            For Each R As cRegiment In Mercenary.Regiments.Values
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        If R.PillagingCivilians <> 0 Or B.PillagingCivilians <> 0 Or C.PillagingCivilians <> 0 Then
                            Pillage += C.PillagingCivilians
                            PillageMax = Math.Max(Pillage, PillageMax)
                            ForcesPillaging += C.SquadCount
                        End If
                    Next
                Next
            Next
            If Pillage = 0 Then Exit Sub
            MsgBox("Your troops pillaged civilians and extorted " & Format(Pillage * 10000, "###,###") & " C-Bills.", vbOKOnly, "Raping the Horses and Riding off on the Women!")

            Mercenary.Cash_On_Hand += Pillage * 10000
            If D(3, 6) <= PillageMax Then
                MsgBox("However, those civilians REALLY did not like that, and a response was provoked!", MsgBoxStyle.OkOnly, "Locals are Pissed!") 'TODO:  Rebellion?  Formal Complaints?  Use ForcesPillaging for that level
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub UpdateSupplyOrders()
        Try
            'Get supplies
            For Each Order As cSupplyOrder In Mercenary.Supply_Orders
                Order.Delay -= 1
                If Order.Delay <= 1 Then
                    If Order.Lost Then
                        MsgBox("Order #" & Order.OrderID & " was lost in transit and the €" & Order.Supplies * 500 & " worth of supplies was lost!", MsgBoxStyle.OkOnly, "Supplies Lost!")
                    Else
                        MsgBox("Order #" & Order.OrderID & " arrived!  Adding " & Order.Supplies & " to supply stores!", MsgBoxStyle.OkOnly, "Supplies Arrived")
                        Mercenary.Supplies_On_Hand += Order.Supplies
                    End If
                End If
            Next
            'Clear received orders
            For I As Int32 = Mercenary.Supply_Orders.Count - 1 To 0 Step -1
                If Mercenary.Supply_Orders(I).Delay <= 1 Or Mercenary.Supply_Orders(I).Lost Then Mercenary.Supply_Orders.RemoveAt(I)
            Next
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub CheckForMutiny()
        Try
            Dim Mutiny As Boolean = False
            For Each R As cRegiment In Mercenary.Regiments.Values
                If D(2, 6) < R.BackPayMonths Then
                    Mutiny = True
                    DebugEvent("Mutiny occurs as salaries have not been paid!  " & R.BackPayMonths & " Months behind", True)
                End If
                If D(2, 6) < R.SupplyMonthsMissed Then
                    Mutiny = True
                    DebugEvent("Mutiny occurs as unit is not fully supplied!  " & R.SupplyMonthsMissed & " Months missed", True)
                End If
            Next
            For Each R As cRegiment In Mercenary.Regiments.Values
                R.MoraleMod -= R.BackPayMonths
                If R.SupplyMonthsMissed > 0 Then R.MoraleMod -= 1
                If Mercenary.Overhead < 5 Then
                    Mercenary.MoraleMod -= 5
                    If D(2, 6) < (6 + R.MonthsWithoutOverhead) Then
                        Mutiny = True
                        DebugEvent("Mutiny occurs in " & R.Name & " as overhead has not been provided!  " & R.MonthsWithoutOverhead & " Months missed", True)
                    Else
                        DebugEvent("Morale reduced by " & R.MonthsWithoutOverhead & " because overhead has not been provided!  " & R.MonthsWithoutOverhead & " Months missed", True)
                    End If
                End If
                R.MoraleMod -= R.MonthsWithoutOverhead
            Next
            If Mutiny Then Mercenary.DoMutiny()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

#Region "Event Processing"
    Private Sub Process_Event_None()
        DebugEvent("No events this month.", True)
        Try
            Dim CanCampaign As Boolean = False

            'Part of larger force
            If (Forces_Friendly.Count - MercForceSize) > MercForceSize Then
                CanCampaign = True
                DebugEvent("Campaign allowed as friendly forces outnumber mercenaries.", True)
            End If

            'Alone, or part of larger force under house/integrated command rights
            If MercForceSize <= (Forces_Friendly.Count - MercForceSize) And (CurrentContract.Command = eCommandRights.House Or CurrentContract.Command = eCommandRights.Integrated) Then
                CanCampaign = True
                DebugEvent("Campaign allowed as friendly forces are not larger than the mercs, and command rights allow it.", True)
            End If

            If Forces_Enemy.Count = 0 Then
                DebugEvent("There are no enemies present, so a campaign will not happen.", True)
                CanCampaign = False 'Gotta have baddies to fight!
            End If

            If CanCampaign And MercForceSize <= (Forces_Friendly.Count - MercForceSize) And CurrentContract.Command = eCommandRights.Independent Then
                If MsgBox("There are enemy forces present." & vbCrLf & "Independent command rights allow you to execute a military campaign." & vbCrLf & "Would you like to initiate a campaign?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Campaign = True
                End If
            Else
                If D(2, 6) - 2 <= Forces_Friendly.Count Then
                    DebugEvent("Campaign allowed as 2d6 roll was less than friendly forces size.", True)
                    Campaign = True
                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Process_Event_Civil_Disturbance()
        MsgBox("Civil Disturbance!  Local riots and small-scale agitation take place!", MsgBoxStyle.OkOnly, "Civil Disturbance")
        If D(2, 6) <= 3 Then
            MsgBox("Mercenary Unit fights 1 regiment") 'TODO:  Fix
        End If
    End Sub
    Private Sub Process_Event_Sporadic_Uprisings()
        MsgBox("Sporadic Uprisings!  Substantial civil uprisings sweep the planet!", MsgBoxStyle.OkOnly, "Sporadic Uprisings")
        If D(2, 6) <= 8 Then
            MsgBox("Mercenary Unit fights " & D(1, 6) & " regiment(s)") 'TODO:  Fix
        End If
    End Sub
    Private Sub Process_Event_Rebellion()
        Try
            MsgBox("Rebellion!  Local populate erupts in widespread resistance to government.  Mission changed to Riot Duty!", MsgBoxStyle.OkOnly, "Rebellion")
            CurrentContract.Mission = eMission.Riot_Duty
            Dim TempVal As Int32 = EnemyCommitment(CurrentContract.Mission) + EnemyCommitment(CurrentContract.Mission)
            Forces_Enemy.Clear()
            Enemy = New cForce(eForceType.Enemy)
            For I As Int32 = 1 To TempVal
                Enemy.AddRegiment()
            Next
            For Each R As cRegiment In Enemy.Regiments.Values
                Forces_Enemy.Add(R)
            Next
            If Enemy.LaunchCampaign() Then
                DebugEvent("Enemy rebellion forces a campaign", True)
                Campaign = True
            End If
            RefreshLabels()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Process_Event_Betrayal() 'Use table on p64
        Try
            'MsgBox("Betrayal!!", MsgBoxStyle.OkOnly, "Betrayal")
            Dim DieRoll As Int32 = D(2, 6)
            Dim TempResult As Int32 = 0
            Select Case Leader.Faction
                Case eFaction.Davion : TempResult = Choose(D(2, 6) - 1, 5, 0, 4, 4, 2, 6, 6, 2, 4, 3, 1)
                Case eFaction.Kurita : TempResult = Choose(D(2, 6) - 1, 2, 4, 0, 3, 5, 6, 6, 5, 1, 4, 3)
                Case eFaction.Liao : TempResult = Choose(D(2, 6) - 1, 2, 0, 5, 3, 3, 6, 6, 3, 4, 1, 3)
                Case eFaction.Steiner : TempResult = Choose(D(2, 6) - 1, 1, 0, 4, 3, 3, 6, 6, 3, 3, 5, 2)
                Case eFaction.Marik : TempResult = Choose(D(2, 6) - 1, 5, 0, 4, 2, 4, 6, 6, 4, 3, 1, 5)
                Case eFaction.ComStar : TempResult = Choose(D(2, 6) - 1, 1, 0, 4, 6, 6, 6, 6, 6, 4, 3, 5)
                Case eFaction.Other : TempResult = Choose(D(2, 6) - 1, 5, 1, 0, 4, 4, 6, 6, 4, 3, 3, 2)
            End Select
            Select Case TempResult
                Case 0 : MsgBox("Logistics Problem, -" & D(2, 6) - 2 * 10 & "%", vbOKOnly, "Betrayal!")'A
                Case 1 : MsgBox("Unit is abandoned in the field by house transports, if applicable", vbOKOnly, "Betrayal!")'B
                Case 2 : MsgBox("Mission actually diversionary raid" & If(D(2, 6) >= 10, " and unit abandoned in the field by house transports, if applicable", ""), vbOKOnly, "Betrayal!")'C
                Case 3 : MsgBox("Invalid contract, employer will not honor unless outcome A/B/C.  If Comstar is an intermediary, the matter submitted to them for arbitration", vbOKOnly, "Betrayal!")'D
                Case 4 : MsgBox("False intelligence on enemy strength", vbOKOnly, "Betrayal!")'E
                Case 5 : MsgBox("Employer will furnish logistics/transport at highest possible price, even if negotiated.  Employer will lend money to help, but deduct twice the value from final payments", vbOKOnly, "Betrayal!")'F
                Case 6 : MsgBox("Just rumors, no actual betrayal", vbOKOnly, "Betrayal!") 'G
            End Select
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Process_Event_Treachery()
        MsgBox("Treachery!", MsgBoxStyle.OkOnly, "Treachery")
        Debug.Print("TODO:  Treachery")
    End Sub
    Private Sub Process_Event_Logistics_Failure()
        Try
            Dim Loss As Single = 1 - ((D(2, 6) - 2) / 10)
            Select Case D(1, 3)
                Case 1 : MsgBox("Logistics Failure!  Bureaucratic snarls lost our supplies!  " & Loss * 100 & "% of Supplies Lost!", MsgBoxStyle.OkOnly, "Logistics Failure")
                Case 2 : MsgBox("Logistics Failure!  Enemy raided our supplies!  " & Loss * 100 & "% of Supplies Lost!", MsgBoxStyle.OkOnly, "Logistics Failure")
                Case 3 : MsgBox("Logistics Failure!  Supply train captured!  " & Loss * 100 & "% of Supplies Lost!", MsgBoxStyle.OkOnly, "Logistics Failure")
            End Select
            Mercenary.Supplies_On_Hand -= Loss * Mercenary.Supplies_On_Hand
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Process_Event_Reinforcements()
        Try
            Dim OldRegCount As Int32 = Friendly.Regiments.Count + MercForceSize
            For I As Int32 = 1 To FriendlyCommitment(CurrentContract.Mission)
                Dim TempUnit As New cRegiment("")
                TempUnit.ForceComposition = GetRegiment_Friendly(CurrentContract.Mission)
                TempUnit.Experience = GetGenericQuality()
                TempUnit.LeaderSkill = GetTacticalSkill(TempUnit.ForceComposition)
                FillSquads(TempUnit)
                Dim TempName As String = ""
                Select Case TempUnit.ForceComposition
                    Case eForceComposition.Aerospace
                        Friendly.SubUnitList_Aerospace += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Aerospace) & " Aerospace Wing"
                    Case eForceComposition.Air
                        Friendly.SubUnitList_Air += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Air) & " Air Regiment"
                    Case eForceComposition.Infantry
                        Friendly.SubUnitList_Infantry += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Infantry) & " Infantry Regiment"
                    Case eForceComposition.Militia
                        Friendly.SubUnitList_Militia += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Militia) & " Militia Regiment"
                    Case eForceComposition.Armor
                        Friendly.SubUnitList_Armor += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Armor) & " Armor Regiment"
                    Case eForceComposition.Mech
                        Friendly.SubUnitList_Mech += 1
                        TempName = GetNumberString(Friendly.SubUnitList_Mech) & " Mech Regiment"
                End Select
                TempUnit.Name = TempName
                GetOriginalSize(TempUnit)
                DebugEvent("Unit reinforced by " & TempUnit.Name, True)
                Friendly.Regiments.Add(TempName, TempUnit)
                Forces_Friendly.Add(TempUnit)
            Next
            MsgBox("Friendly Reinforcements!!" & vbCrLf & Friendly.Regiments.Count - OldRegCount & " Regiments Joined the Fight!", MsgBoxStyle.OkOnly, "Reinforcements")
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Process_Event_Attrition()
        Try
            Dim KIA As Int32 = D(2, 6) - 2 '% Gone, forever
            Dim WIA As Int32 = (D(2, 6) - 2) * 5 '% wounded, return next month
            Select Case D(1, 3)
                Case 1 : MsgBox("Attrition!  Natural Disaster devastated our troops!  " & KIA & "% Killed and " & WIA & "% Wounded!", MsgBoxStyle.OkOnly, "Attrition")
                Case 2 : MsgBox("Attrition!  Disease devastated our troops!  " & KIA & "% Killed and " & WIA & "% Wounded!", MsgBoxStyle.OkOnly, "Attrition")
                Case 3 : MsgBox("Attrition!  STD Outbreak devastated our troops!  " & KIA & "% Killed and " & WIA & "% Wounded!", MsgBoxStyle.OkOnly, "Attrition")
            End Select

            For Each R As cRegiment In Forces_Friendly
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If D(1, 100) <= KIA Then
                                    S.State = eSquadState.Killed
                                    DebugEvent("Attrition! " & S.Pilot & " KIA, " & L.Name & ", " & C.Name & ", " & B.Name & ", " & R.Name, True)
                                ElseIf D(1, 100) <= WIA Then
                                    S.State = eSquadState.Wounded
                                    DebugEvent("Attrition! " & S.Pilot & " WIA, " & L.Name & ", " & C.Name & ", " & B.Name & ", " & R.Name, True)
                                End If
                            Next
                        Next
                    Next
                Next
            Next

            For Each R As cRegiment In Forces_Enemy
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If D(1, 100) <= KIA Then
                                    S.State = eSquadState.Killed
                                    DebugEvent("Attrition! " & S.Pilot & " KIA, " & L.Name & ", " & C.Name & ", " & B.Name & ", " & R.Name, True)
                                ElseIf D(1, 100) <= WIA Then
                                    S.State = eSquadState.Wounded
                                    DebugEvent("Attrition! " & S.Pilot & " WIA, " & L.Name & ", " & C.Name & ", " & B.Name & ", " & R.Name, True)
                                End If
                            Next
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Private Sub Process_Event_Internal_Dissension()
        MsgBox("Internal Dissension" & vbCrLf & "At the gamemaster's discretion, internal dissension can take two very different forms. Normally, this event refers to internal discontent with House politics, expressing itself as planetary revolt, an assassination attempt against a Duke or Warlord, or as a full-scale power bid by a disaffected ruling House member (like Michael Hasek-Davion, Theodore Kurita, or Frederick Steiner). Alternatively, gamemasters can interpret this event as an outbreak of dissension within the players' own unit, which could result in a mutiny or a permanent split-up of the unit.", MsgBoxStyle.OkOnly, "Internal Dissension")
    End Sub
    Private Sub Process_Event_Armistice()
        MsgBox("Armistice" & vbCrLf & "Peace is temporarily declared along one or more of the fronts of the Succession Wars as a result of an agreement between two or more of the five Houses. This armistice may cover an area as large as an entire border between Houses or as small as a single planet, and can be for either a set or indefinite time limit. Gamemasters should note that although few, if any, parties will violate an armistice agreement by initiating a full-scale assault in the armistice zone, the agreement will normally not stop the opponents from continuing raiding activities or military buildups in the affected area.", MsgBoxStyle.OkOnly, "Armistice")
    End Sub
    Private Sub Process_Event_Change_of_Allegiance()
        MsgBox("The unit stationed with the players' mere unit changes allegiance from one Successor House to another (or in some cases, from the existing House regime to a rebel faction). As a result of this change of loyalties, player characters may gain or lose large amounts of land, spare parts, or cold, hard cash. If the players' unit is stationed alone, roll again.", MsgBoxStyle.OkOnly, "Allegiance Change")
    End Sub
    Private Sub Process_Event_ComStar_Activity()
        If LastMajorEvent = eEvent.ComStar_Activity Then
            MsgBox("Comstar Activity" & vbCrLf & "Fullscale communications interdict against one Successor House's territory (as punishment for some gross transgression against the facilities under ComStar's sacred trust).", MsgBoxStyle.OkOnly, "ComStar Activity")
        Else
            MsgBox("Comstar Activity" & vbCrLf & "A significant development affecting the relationship between ComStar and one or more of the ruling Houses occurs. Possible events might include the opening or closing of a major relay station, a ComStar request for 'Mech unit volunteers to assist its Explorer Corps teams on a mission, or a call for aid from a relay station that has been attacked or struck by a natural disaster. ", MsgBoxStyle.OkOnly, "ComStar Activity")
            LastMajorEvent = eEvent.ComStar_Activity
        End If
    End Sub
    Private Sub Process_Event_Periphery_Contact()
        MsgBox("Periphery Contact." & vbCrLf & "Interaction between the players' unit (or their current House employer) and inhabitants of the Periphery occurs. The many possibilities include the arrival of free traders from the Periphery, House-sponsored exploratory missions to Periphery worlds believed to be uninhabited, raids on House worlds by Bandit King forces (or vice versa), or even the establishment of long-term relationships between the two areas (like the one between Kyalla Centralia of the Magistracy of Canopus and Catherine Humphreys of the Principality of Andurien described in MechWarrior.", MsgBoxStyle.OkOnly, "Periphery Contact")
    End Sub
    Private Sub Process_Event_Major_Campaign()
        MsgBox("Major Campaign" & vbCrLf & "One or more Houses begins a major offensive against a bordering neighbor. Such a campaign can consist of either a general assault on a number of enemy-held worlds along a given front, seeking to make inroads wherever the enemy is weakest, or a series of 'steppingstone' invasions aimed at the ultimate conquest of a single key objective. In either case, the campaign should involve a minimum of 10-20 regiments of 'Mechs per side, and occupy the primary attention of all the forces stationed on either side of the campaign front.", MsgBoxStyle.OkOnly, "Major Campaign")
    End Sub
    Private Sub Process_Event_Technological_Advance()
        MsgBox("Technological Advance" & vbCrLf & "Techs from one of the Successor State Houses succeed in recovering one of the lost technologies of the Star League years, or they uncover an entirely new product or process. This advance is most likely to be related to military matters, and especially to technologies dealing with 'Mech, AeroSpace Fighter, or JumpShip construction or repair. At the gamemaster's discretion, the advance may be applied to an entirely different aspect of society (medicine, communications, agriculture, and so on). If this result recurs several times within a period of a few Months, the gamemaster should increase the significance of the discovery made." & "The most logical source of research developments like these is the NAIS in New Avalon or the copycat universities recently built by House Kurita and Marik. Of course, no matter where the initial discovery is made, the other Houses will soon have their own spies (or maybe even a full-scale invasion force, as in the case of the Battle for Hoff) on the scene to try to gain the secret for themselves." & vbCrLf, MsgBoxStyle.OkOnly, "Technological Advance")
    End Sub
    Private Sub Process_Event_Star_League_Facility()
        MsgBox("Star League Facility" & vbCrLf & "A major Star League facility (storehouse, administrative headquarters, military or naval base) is discovered on a Successor State world. As in the case of a Technological Advance, rumors of the discovery will travel quickly through the Inner Sphere, drawing spies or military units from other Houses to investigate or attempt to seize the facility.", MsgBoxStyle.OkOnly, "Star League Facility")
    End Sub
    Private Sub Process_Event_Fall_of_Major_World()
        MsgBox("Fall of Major World" & vbCrLf & "A planet with valuable resources or strategic significance changes hands from one House to another. Such an event will usually have repercussions (retaliatory raids, unit transfers, and further assaults) up and down both sides of the border where the change has occurred.", MsgBoxStyle.OkOnly, "Fall of Major World")
    End Sub
    Private Sub Process_Event_Death_of_Major_Personage()
        MsgBox("Death of Major Personage" & vbCrLf & "An important Successor State figure, ranging from a major military leader or a planetary ruling Duke on up to a royal family member. heir, or even a War, No fee too high! lord, dies. The death can occur as a result of combat, illness, or a successful assassination attempt. The possible reprcussions can be as simple as the transfer of power to a newly-promoted military leader or noble, or they can be as disruptive as a planetary rebellion or full-scale civil war. At the gamemaster's discretion, this event can be treated instead as an attempt on a major personage's life, with the player characters having the opportunity to either perform or prevent the attempt.", MsgBoxStyle.OkOnly, "Death of Major Personage")
    End Sub
    Private Sub Process_Event_Change_in_House_House_Relationship()
        MsgBox("Change in House/House Relations" & vbCrLf & "A significant change in the relationship between two or more Successor Houses occurs. This event can be either a sweeping change such as the formation or breaking of an alliance, or a subtle incident whose public effect is small but that will subtly improve or erode an existing relationship over the course of time." & vbCrLf & "Except as otherwise noted, the Special Events step is followed by the Friendly Campaign Operations step. In cases where an event causes a reversion to some previous step, go back to the step indicated and continue resolving steps normally. This means that multiple Special Events may take place in one month.", MsgBoxStyle.OkOnly, "Change in House/House Relations")
    End Sub


#Region "Campaign Events"
    Public Function CampaignEvent(Mission As eMission) As eEvent
        Dim TempEvent As eEvent = eEvent.None
        Select Case Mission
            Case eMission.Cadre_Duty, eMission.Garrison_Duty, eMission.Security_Duty
                Select Case CInt(D(1, 6) & D(1, 6))
                    Case 11 To 16 : TempEvent = eEvent.None
                    Case 21 To 36 : TempEvent = eEvent.Civil_Disturbance
                    Case 41 To 46 : TempEvent = eEvent.Sporadic_Uprisings
                    Case 51 To 54 : TempEvent = eEvent.Rebellion
                    Case 55 : TempEvent = eEvent.Betrayal
                    Case 56 : TempEvent = eEvent.Treachery
                    Case 61 To 62 : TempEvent = eEvent.Logistics_Failure
                    Case 63 : TempEvent = eEvent.Reinforcements
                    Case 64 To 65 : TempEvent = eEvent.Attrition
                    Case 66 : TempEvent = MajorEvent()
                End Select
            Case eMission.Offensive_Campaign, eMission.Defensive_Campaign
                Select Case D(1, 6) & D(1, 6)
                    Case 11 To 42 : TempEvent = eEvent.None
                    Case 43 To 44 : TempEvent = eEvent.Civil_Disturbance
                    Case 45 : TempEvent = eEvent.Sporadic_Uprisings
                    Case 46 : TempEvent = eEvent.Rebellion
                    Case 51 : TempEvent = eEvent.Betrayal
                    Case 52 To 53 : TempEvent = eEvent.Treachery
                    Case 54 To 56 : TempEvent = eEvent.Logistics_Failure
                    Case 61 To 63 : TempEvent = eEvent.Reinforcements
                    Case 64 To 65 : TempEvent = eEvent.Attrition
                    Case 66 : TempEvent = MajorEvent()
                End Select
            Case eMission.Relief_Duty, eMission.Planetary_Assault
                Select Case D(1, 6) & D(1, 6)
                    Case 11 To 46 : TempEvent = eEvent.None
                    Case 51 To 53 : TempEvent = eEvent.Betrayal
                    Case 54 To 55 : TempEvent = eEvent.Treachery
                    Case 56 To 61 : TempEvent = eEvent.Logistics_Failure
                    Case 62 To 63 : TempEvent = eEvent.Reinforcements
                    Case 64 To 65 : TempEvent = eEvent.Attrition
                    Case 66 : TempEvent = MajorEvent()
                End Select
            Case eMission.Riot_Duty, eMission.Siege_Duty
                Select Case D(1, 6) & D(1, 6)
                    Case 11 To 42 : TempEvent = eEvent.None
                    Case 43 To 44 : TempEvent = eEvent.Civil_Disturbance
                    Case 45 : TempEvent = eEvent.Sporadic_Uprisings
                    Case 46 : TempEvent = eEvent.Rebellion
                    Case 51 : TempEvent = eEvent.Betrayal
                    Case 52 To 53 : TempEvent = eEvent.Treachery
                    Case 54 To 56 : TempEvent = eEvent.Logistics_Failure
                    Case 61 To 653 : TempEvent = eEvent.Reinforcements
                    Case 64 To 65 : TempEvent = eEvent.Attrition
                    Case 66 : TempEvent = MajorEvent()
                End Select
            Case eMission.Recon_Raid, eMission.Objective_Raid, eMission.Divisionary_Raid, eMission.Guerrilla_Warfare
                Select Case D(1, 6) & D(1, 6)
                    Case 11 To 54 : TempEvent = eEvent.None
                    Case 55 To 56 : TempEvent = eEvent.Betrayal
                    Case 61 To 62 : TempEvent = eEvent.Treachery
                    Case 63 : TempEvent = eEvent.Logistics_Failure
                    Case 64 : TempEvent = eEvent.Reinforcements
                    Case 65 : TempEvent = eEvent.Attrition
                    Case 66 : TempEvent = MajorEvent()
                End Select
            Case eMission.Retainer
                TempEvent = eEvent.None
        End Select
        Return TempEvent
    End Function
    Public Function MajorEvent() As eEvent
        Dim TempEvent As eEvent = eEvent.None
        Select Case D(1, 6) & D(1, 6)
            Case 11 To 14 : TempEvent = eEvent.Internal_Dissension
            Case 15 To 16 : TempEvent = eEvent.Armistice
            Case 21 To 22 : TempEvent = eEvent.Change_of_Allegiance
            Case 23 To 25 : TempEvent = eEvent.ComStar_Activity
            Case 26 To 32 : TempEvent = eEvent.Periphery_Contact
            Case 33 To 41 : TempEvent = eEvent.Major_Campaign
            Case 42 To 43 : TempEvent = eEvent.Technological_Advance
            Case 44 To 46 : TempEvent = eEvent.Star_League_Facility
            Case 51 To 56 : TempEvent = eEvent.Fall_of_Major_World
            Case 61 To 63 : TempEvent = eEvent.Death_of_Major_Personage
            Case 64 To 66 : TempEvent = eEvent.Change_in_House_House_Relationship
        End Select
        Return TempEvent
    End Function
#End Region
#End Region


    Dim Campaign As Boolean = False
    Dim MercForceSize As Single = 0

    Private Sub cmdAdvance_Click(sender As Object, e As EventArgs) Handles cmdAdvance.Click

        Campaign = False
        MercForceSize = Math.Round(Mercenary.SquadCount / 108, 0, MidpointRounding.AwayFromZero)

        If cmdPaySalaries.Text <> "Salaries Paid" Then
            MsgBox("Please pay your troop salaries before advancing!", MsgBoxStyle.OkOnly, "Accountant")
            Exit Sub
        End If
        If cmdProvideMaintenance.Text <> "Maintenance Paid" Then
            MsgBox("Please perform maintenance before advancing!", MsgBoxStyle.OkOnly, "Technician")
            Exit Sub
        End If
        If cmdPayOverhead.Text <> "Overhead Paid" Then
            MsgBox("Please pay your unit overhead before advancing!", MsgBoxStyle.OkOnly, "Administrator")
            Exit Sub
        End If
        If cmdProvideSupplies.Text <> "Unit Supplied" Then
            MsgBox("Please allocate supplies before advancing!", MsgBoxStyle.OkOnly, "Supply Officer")
            Exit Sub
        End If

        CurrentTurn = CurrentTurn.AddMonths(1)

        RecoverWounded() 'Recover the wounded
        Pillage()
        UpdateSupplyOrders()
        CheckForMutiny()

        'Check end of contract
        If CurrentContract.Length_Remaining <= 0 And CurrentContract.Active Then
            If CurrentContract.Paid = False Then
                CurrentContract.Paid = True
                Dim Cash As Int64 = CurrentContract.Value_Total_Completion - CurrentContract.Value_Total_Advance
                If Cash = 0 Then
                    DebugEvent("Contract over but all payment was already provided.", True)
                Else
                    DebugEvent("Contract over, " & Cash & " now paid for contract completion", True)
                End If
                Mercenary.Cash_On_Hand += Cash
                MsgBox("Your contract has been completed!" & vbCrLf & "€" & Format(Cash, "###,##0") & " Paid!", MsgBoxStyle.OkOnly, "Contract Complete")

                cmdProvideSupplies.Enabled = True
            Else
                cmdProvideSupplies.Enabled = False
            End If
            CurrentContract.Active = False
            Forces_Friendly.Clear()
            Forces_Enemy.Clear()
            cmdPaySalaries.Enabled = False
            cmdPayOverhead.Enabled = False
            cmdProvideMaintenance.Enabled = False

        Else
            'Reset the buttons
            cmdProvideSupplies.Enabled = True
            cmdProvideSupplies.Text = "Provide Supplies"
            cmdProvideSupplies.Enabled = True

            cmdPaySalaries.Enabled = True
            cmdPaySalaries.Text = "Pay Salaries"
            cmdPaySalaries.Enabled = False

            cmdPayOverhead.Enabled = True
            cmdPayOverhead.Text = "Pay Overhead"
            cmdPayOverhead.Enabled = False

            cmdProvideMaintenance.Enabled = True
            cmdProvideMaintenance.Text = "Provide Maintenance"
            cmdProvideMaintenance.Enabled = False
        End If
        If CurrentContract.Active Then
            cmdFindContract.Enabled = False
            cmdBreak.Enabled = True
        Else
            cmdFindContract.Enabled = True
            cmdBreak.Enabled = False
        End If
        cmdAdvance.Enabled = False


        Try
            If CurrentContract.Active Then
                CurrentContract.Length_Remaining -= 1

                EnemyRepairs()
                EnemyCommittment()

                If Enemy.LaunchCampaign() Then
                    Campaign = True
                    DebugEvent("Enemy launches campaign!", True)
                Else
                    If LastMajorEvent <> eEvent.ComStar_Activity Then LastMajorEvent = eEvent.None
                    Dim TempEvent As eEvent = CampaignEvent(CurrentContract.Mission)
                    DebugEvent("Event: " & TempEvent.ToString, True)
                    Select Case TempEvent
                        Case eEvent.None : Process_Event_None()
                        Case eEvent.Civil_Disturbance : Process_Event_Civil_Disturbance()
                        Case eEvent.Sporadic_Uprisings : Process_Event_Sporadic_Uprisings()
                        Case eEvent.Rebellion : Process_Event_Rebellion()
                        Case eEvent.Betrayal : Process_Event_Betrayal()
                        Case eEvent.Treachery : Process_Event_Treachery()
                        Case eEvent.Logistics_Failure : Process_Event_Logistics_Failure()
                        Case eEvent.Reinforcements : Process_Event_Reinforcements()
                        Case eEvent.Attrition : Process_Event_Attrition()

                    'Major Events.  None of these have any effect in the campaign
                        Case eEvent.Internal_Dissension : Process_Event_Internal_Dissension()
                        Case eEvent.Armistice : Process_Event_Armistice()
                        Case eEvent.Change_of_Allegiance : Process_Event_Change_of_Allegiance()
                        Case eEvent.ComStar_Activity : Process_Event_ComStar_Activity()
                        Case eEvent.Periphery_Contact : Process_Event_Periphery_Contact()
                        Case eEvent.Major_Campaign : Process_Event_Major_Campaign()
                        Case eEvent.Technological_Advance : Process_Event_Technological_Advance()
                        Case eEvent.Star_League_Facility : Process_Event_Star_League_Facility()
                        Case eEvent.Fall_of_Major_World : Process_Event_Fall_of_Major_World()
                        Case eEvent.Death_of_Major_Personage : Process_Event_Death_of_Major_Personage()
                        Case eEvent.Change_in_House_House_Relationship : Process_Event_Change_in_House_House_Relationship()
                    End Select
                End If


                If Campaign Then
                    'Aerospace Operations
                    Dim EnemyAS As Int64 = Enemy.Aerospace
                    Dim FriendlyAS As Int64 = Friendly.Aerospace

                    'Need to ensure has fighters
                    If FriendlyAS > 0 Or EnemyAS > 0 Then
                        Select Case AerospaceOperations(FriendlyAS, EnemyAS)
                            Case eAerospaceOperationsResult.ES
                                CombatResults(Forces_Enemy, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                EnemyAS = Enemy.Aerospace
                                CombatResults(Forces_Friendly, D(1, 6) * D(1, 6), D(1, 6) * D(1, 6), True)
                                FriendlyAS = Friendly.Aerospace
                                If CurrentContract.Mission = eMission.Planetary_Assault Or CurrentContract.Mission = eMission.Relief_Duty Then
                                    MsgBox("Can not do mission this month!")
                                End If

                                If FriendlyAS > 0 Then
                                    If MsgBox("You lost air superiority, would you like to launch a sortie to prevent that, even though it will increase aerospace losses?", MsgBoxStyle.YesNo, "Air Superiority Lost!") = MsgBoxResult.Yes Then
                                        CombatResults(Forces_Friendly, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                        EnemyAS = 0
                                    End If
                                End If
                            Case eAerospaceOperationsResult.A
                                CombatResults(Forces_Enemy, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                EnemyAS = Enemy.Aerospace
                                CombatResults(Forces_Friendly, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                FriendlyAS = Friendly.Aerospace

                                If CurrentContract.Mission = eMission.Planetary_Assault Or CurrentContract.Mission = eMission.Relief_Duty Then
                                    MsgBox("Can not do mission this month!")
                                    'TODO:  Wtf do I do with this information?
                                End If
                            Case eAerospaceOperationsResult.FS
                                CombatResults(Forces_Enemy, D(1, 6) * D(1, 6), D(1, 6) * D(1, 6), True)
                                EnemyAS = Enemy.Aerospace
                                CombatResults(Forces_Friendly, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                FriendlyAS = Friendly.Aerospace + Mercenary.Aerospace

                                If D(1, 2) = 1 Then
                                    MsgBox("Enemy forces prevent you from using your air superiority to assist maneuver and combat operations!", vbOKOnly, "Air Superiority Denied!")
                                    CombatResults(Forces_Enemy, D(1, 3) * D(1, 6), D(1, 3) * D(1, 6), True)
                                    FriendlyAS = 0
                                End If
                        End Select
                    Else
                        DebugEvent("No aerospace forces present", True)
                    End If

                    'Maneuver (and Combat) Operations
                    Dim DecE As Boolean = False
                    Dim DecF As Boolean = False
                    Dim CampaignFinished As Boolean = False
                    Do
                        Dim CanRollAgain As Boolean = False
                        Dim ManeuverResults As eManeuverOperationsResult

                        If MercForceSize > Forces_Friendly.Count Then 'If mercs most of the unit, their skill is used
                            ManeuverResults = ManeuverOperations(Friendly.Maneuver + Mercenary.Maneuver + FriendlyAS, Enemy.Maneuver + EnemyAS, Leader.Skills(eSkill.Tactics), Enemy.OverallSkill)
                        Else
                            ManeuverResults = ManeuverOperations(Friendly.Maneuver + Mercenary.Maneuver + FriendlyAS, Enemy.Maneuver + EnemyAS, Friendly.OverallSkill, Enemy.OverallSkill)
                        End If

                        Select Case ManeuverResults
                            Case eManeuverOperationsResult.NP
                                'No progress
                                CanRollAgain = True
                                MsgBox("There was no progress in maneuver operations this month.", vbOKOnly, "Boring Month")

                            Case eManeuverOperationsResult.Att
                                'Attrition
                                MsgBox("Maneuver operations this month caused attrition on both forces.", vbOKOnly, "Attrition")
                                CombatResults(Forces_Enemy, 0, (D(2, 6) - 2) / 100 * Enemy.SquadCount)
                                CombatResults(Forces_Friendly, 0, (D(2, 6) - 2) / 100 * Enemy.SquadCount)
                                CanRollAgain = True

                            Case eManeuverOperationsResult.Sk
                                'Skirmish
                                MsgBox("Maneuver operations this month created a skirmish!", vbOKOnly, "Skirmish!")
                                Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6), Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6), Enemy.SquadCount, Mercenary.Dishonorable)
                                CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                CanRollAgain = True

                            Case eManeuverOperationsResult.Bat
                                'Battle
                                Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Enemy.SquadCount, Mercenary.Dishonorable)
                                MsgBox("Maneuver operations this month created a battle!", vbOKOnly, "Battle!")
                                CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                CanRollAgain = True

                            Case eManeuverOperationsResult.Con
                                'Continued campaign. go to Campaign Outcome
                                CanRollAgain = True
                                CampaignOutcome(Mercenary, CurrentContract.Mission, False, False, False, False)

                            Case eManeuverOperationsResult.DecE
                                'Decisive Enemy Outcome
                                DecE = True
                                CanRollAgain = False
                                If MsgBox("Our forces were completely outmaneuvered.  Do you wish to request the enemies accept the honor of war and allow us to abandon the planet?", vbYesNo, "Request Retreat with Honor") = MsgBoxResult.Yes Then
                                    If D(1, 10) = 10 Then 'I made it 10% of the time the enemy would not let friendlies leave when beaten
                                        MsgBox("The enemy does not accept our request, and continues to attack!", vbOKOnly, "Enemy Dishonor!")
                                        Enemy.Dishonorable = True
                                        Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                        Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Enemy.SquadCount, Mercenary.Dishonorable)
                                        CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                        CanRollAgain = True
                                    Else
                                        MsgBox("The enemy has accepted our request to leave the planet.", vbOKOnly, "Left with Honor")
                                        CampaignOutcome(Mercenary, CurrentContract.Mission, False, True, False, False)
                                        CurrentContract.Active = False
                                        'TODO:  Contract abandoned!  What to do when player runs?  Merc3055 states they lose all remaining money.  Troops usually ransomed back for 100K, but may take all equipment, surrendering to other mercs is -25% of battlefield losses to each other
                                    End If
                                Else
                                    'Player does not accept defeat!
                                    'Battle Fought immediately with no dishonor!
                                    Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                    Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Enemy.SquadCount, Mercenary.Dishonorable)
                                    MsgBox("Enemy pressed their attack!", vbOKOnly, "Enemy Overrun!")
                                    CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                End If




                            Case eManeuverOperationsResult.DecF
                                'Decisive Friendly Outcome
                                DecF = True
                                CanRollAgain = False
                                If D(1, 10) = 10 Then 'I made enemies 10% of the time not leave if they were decisively beaten
                                    MsgBox("Our forces completely outmaneuvered the enemy, who have refused to abandon the planet, so we continue to press the attack.", vbOKOnly, "Enemies fight till death!")
                                    Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                    Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Enemy.SquadCount, Mercenary.Dishonorable)
                                    CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                    CanRollAgain = True
                                Else
                                    'Enemy requests honor
                                    'TODO:  If mercs not in charge, they do not get this question!
                                    If MsgBox("Our forces completely outmaneuvered the enemy, who have requested we abide by the honors of war and allow them to leave the planet.  Do you wish to accept their terms?", vbYesNo, "Enemy surrender!") = MsgBoxResult.No Then
                                        Mercenary.Dishonorable = True
                                        'Battle Fought immediately dishonor!
                                        Dim FriendlySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Friendly.SquadCount + Mercenary.SquadCount, Enemy.Dishonorable)
                                        Dim EnemySize As Int32 = GetSquadsCommitted(D(1, 6) + 5, Enemy.SquadCount, Mercenary.Dishonorable)
                                        MsgBox("With dishonor, our forces pressed the attack!", vbOKOnly, "Dishonorable Attack!")
                                        CombatOperations(FriendlySize, EnemySize, False, FriendlyAS, EnemyAS)
                                    Else
                                        'CampaignOutcome(Mercenary, Contract.Mission, True, False, False, False)
                                        CampaignFinished = True
                                        MsgBox("Your contract has been completed.", vbOKOnly, "Contract Completion")
                                        CurrentContract.Length_Remaining = 0
                                    End If
                                End If


                        End Select

                        If CanRollAgain Then
                            If MsgBox("Maneuver operations inconclusive, would you like to try again?", MsgBoxStyle.YesNo, "Inconclusive") <> MsgBoxResult.Yes Then Exit Do
                        Else
                            Exit Do
                        End If

                    Loop

                    MsgBox("Campaign this month ends.", vbOKOnly, "Campaign Over")
                    If CampaignFinished = False Then
                        CampaignOutcome(Friendly, CurrentContract.Mission, DecF, DecE, Mercenary.Dishonorable, Enemy.Dishonorable)
                        'TODO:  Salvage
                        TitleIncome(Leader.Title)
                    End If
                    Mercenary.Dishonorable = False
                    Enemy.Dishonorable = False
                End If
            End If

        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

        'Refresh contracts
        frmNegotiate.UpdateContractList()

        RefreshLabels()


        'MsgBox("Saving game!", MsgBoxStyle.OkOnly, "Save!")
        'TODO:  Save this unit (ensure I am also saving date and all the current enemies and shit too!)
    End Sub


    Private Sub TitleIncome(Title As String)
        Dim YearlyIncome As Int32 = 0
        Select Case Title.ToLower
            Case "knight" : YearlyIncome = 25000
            Case "baronet" : YearlyIncome = 50000
            Case "baron" : YearlyIncome = 100000
            Case "viscount" : YearlyIncome = 500000
            Case "count" : YearlyIncome = 1500000
            Case "marquess" : YearlyIncome = 10000000
            Case "duke" : YearlyIncome = 100000000
            Case Else : YearlyIncome = 0
        End Select
        If YearlyIncome > 0 Then
            YearlyIncome = Math.Round(YearlyIncome / 12, 0, MidpointRounding.AwayFromZero)
            DebugEvent("Additional income from title of " & Title & " adds " & YearlyIncome & " cash to the unit.", True)
            Mercenary.Cash_On_Hand += YearlyIncome
        End If
    End Sub
    Private Sub EnemyRepairs()
        Try
            'Next, lets get a list of all salvaged equipment so the enemy can repair it
            Dim SalvagedEnemies As New List(Of cSquad)
            For Each R As cRegiment In Forces_Enemy
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If S.Quality = eQuality.Salvaged Then
                                    SalvagedEnemies.Add(S)
                                End If
                            Next
                        Next
                    Next
                Next
            Next

            If SalvagedEnemies.Count = 0 Then
                Exit Sub
            Else
                DebugEvent("Enemy conducting repairs...", True)
            End If
            Dim RegainedSquads As Int32 = Math.Floor((D(2, 6) - 2) / 10 * SalvagedEnemies.Count)
            Dim ActualRegained As Int32 = 0
            Do
                Dim I As Int32 = D(1, RegainedSquads) - 1
                SalvagedEnemies(I).Quality = eQuality.LikeNew
                DebugEvent("Enemy repaired " & SalvagedEnemies(I).UnitType.ToString, True)
                ActualRegained += 1
                RegainedSquads -= 1
                If RegainedSquads < 0 Or SalvagedEnemies.Count = 0 Then Exit Do
            Loop
            DebugEvent(ActualRegained & " enemy squads were repaired from salvaged quality", True)
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub EnemyCommittment()
        DebugEvent("Checking enemy committment levels...", True)
        Try
            'Now lets update enemy committment level
            Dim ER As Int32 = EnemyCommitment(CurrentContract.Mission)
            If ER > Enemy.Regiments.Count Then
                DebugEvent("Enemy was reinforced by " & (ER - Enemy.Regiments.Count) & " regiments!", True)
                For I As Int32 = 0 To (ER - Enemy.Regiments.Count)
                    Enemy.AddRegiment()
                Next
                Forces_Enemy.Clear()
                For Each R As cRegiment In Enemy.Regiments.Values
                    Forces_Enemy.Add(R)
                Next
            ElseIf ER < Enemy.Regiments.Count Then
                Dim EnemyRegimentList As New List(Of cRegiment)
                For Each R As cRegiment In Enemy.Regiments.Values
                    EnemyRegimentList.Add(R)
                Next
                Dim RegimentsToRemove As Int32 = Enemy.Regiments.Count - ER
                DebugEvent("Enemy withdrew " & RegimentsToRemove & " regiments of troops", True)
                Do 'Units removed
                    Dim R As Int32 = D(1, EnemyRegimentList.Count) - 1
                    Enemy.Regiments.Remove(EnemyRegimentList(R).ID)
                    RegimentsToRemove -= 1
                    If RegimentsToRemove = 0 Or EnemyRegimentList.Count = 0 Or Enemy.Regiments.Count = 0 Then Exit Do
                Loop
            Else
                DebugEvent("There was no change in enemy force sizes.", True)
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Function GetSquadsCommitted(DieRoll As Int32, ForceSize As Int32, DishonorableOpponent As Boolean) As Int32
        Try
            Dim TempReturn As Int32 = 0
            Dim TempCol As Int32
            Select Case ForceSize
                Case 1 : TempCol = 1
                Case 2 To 4 : TempCol = 2
                Case 5 To 12 : TempCol = 3
                Case 13 To 36 : TempCol = 4
                Case 37 To 108 : TempCol = 5
                Case > 108 : TempCol = 6
            End Select
            If DishonorableOpponent Then TempCol += 1

            Select Case TempCol
                Case 1
                    TempReturn = 1
                Case 2
                    Select Case DieRoll
                        Case 1 To 3 : TempReturn = 1
                        Case > 3 : TempReturn = 4
                    End Select
                Case 3
                    Select Case DieRoll
                        Case 1 To 2 : TempReturn = 1
                        Case 3 To 5 : TempReturn = 4
                        Case > 5 : TempReturn = 12
                    End Select
                Case 4
                    Select Case DieRoll
                        Case 1 : TempReturn = 1
                        Case 2 To 3 : TempReturn = 4
                        Case 4 To 5 : TempReturn = 12
                        Case > 5 : TempReturn = 36
                    End Select
                Case 5
                    Select Case DieRoll
                        Case 1 : TempReturn = 1
                        Case 2 : TempReturn = 4
                        Case 3 To 4 : TempReturn = 12
                        Case > 4 : TempReturn = 36
                    End Select
                Case >= 6
                    Select Case DieRoll
                        Case 1 : TempReturn = 1
                        Case 2 : TempReturn = 4
                        Case 3 : TempReturn = 12
                        Case 4 : TempReturn = 36
                        Case 5 : TempReturn = 108
                        Case 6 : TempReturn = ForceSize * 0.2
                        Case 7 : TempReturn = ForceSize * 0.4
                        Case 8 : TempReturn = ForceSize * 0.6
                        Case 9 : TempReturn = ForceSize * 0.8
                        Case > 9 : TempReturn = ForceSize
                    End Select
            End Select
            Return TempReturn
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Function

    Private Sub CombatResults(Force As List(Of cRegiment), Destroyed As Int32, Salvaged_Wounded As Int32, Optional Air As Boolean = False)
        Try
            DebugEvent("CombatResults: " & Destroyed & " Destroyed, " & Salvaged_Wounded & " Wounded" & IIf(Air, ", Aerospace", "Ground"), True)

            Dim Squad_List As New List(Of cSquad)

            For Each R As cRegiment In Force
                For Each B As cBattalion In R.Battalions.Values
                    For Each C As cCompany In B.Companies.Values
                        For Each L As cLance In C.Lances.Values
                            For Each S As cSquad In L.Squads.Values
                                If Air Then
                                    If S.UnitType = eUnitType.Aircraft Or S.UnitType = eUnitType.Fighter_Heavy Or S.UnitType = eUnitType.Fighter_Medium Or S.UnitType = eUnitType.Fighter_Light Or S.UnitType = eUnitType.LAM_Light Or S.UnitType = eUnitType.LAM_Medium Then Squad_List.Add(S)
                                Else
                                    If S.UnitType <> eUnitType.Aircraft And S.UnitType <> eUnitType.Fighter_Heavy And S.UnitType <> eUnitType.Fighter_Medium And S.UnitType <> eUnitType.Fighter_Light Then Squad_List.Add(S)
                                End If
                            Next
                        Next
                    Next
                Next
            Next

            DebugEvent("Combat Results: ", True)

            'Destroyed
            Do
                Dim Pos As Int32 = D(1, Squad_List.Count) - 1
                If Squad_List(Pos).Quality <> eQuality.Destroyed And Squad_List(Pos).State = eSquadState.Active Then
                    Squad_List(Pos).Quality = eQuality.Destroyed
                    Destroyed -= 1
                    DebugEvent("  " & Squad_List(Pos).UnitType.ToString & " Destroyed", True)
                End If
                Dim Left As Int32 = 0
                For I As Int32 = 0 To Squad_List.Count - 1
                    If Squad_List(I).Quality <> eQuality.Destroyed And Squad_List(Pos).State = eSquadState.Active Then Left += 1
                Next
                If Left = 0 Or Destroyed = 0 Then Exit Do
            Loop

            'Salvaged and Wounded
            Do
                Dim Pos As Int32 = D(1, Squad_List.Count) - 1
                If Squad_List(Pos).Quality <> eQuality.Destroyed And Squad_List(Pos).Quality <> eQuality.Salvaged And Squad_List(Pos).State = eSquadState.Active Then
                    Squad_List(Pos).Quality = eQuality.Salvaged
                    Salvaged_Wounded -= 1
                    DebugEvent("  " & Squad_List(Pos).UnitType.ToString & " Wounded", True)
                End If
                Dim Left As Int32 = 0
                For I As Int32 = 0 To Squad_List.Count - 1
                    If Squad_List(Pos).Quality <> eQuality.Destroyed And Squad_List(Pos).Quality <> eQuality.Salvaged And Squad_List(Pos).State = eSquadState.Active Then Left += 1
                Next
                If Left = 0 Or Salvaged_Wounded = 0 Then Exit Do
            Loop
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub


    Private Sub cmdProvideSupplies_Click(sender As Object, e As EventArgs) Handles cmdProvideSupplies.Click
        Try
            If Mercenary.SquadCount > Mercenary.Supplies_On_Hand Then
                MsgBox("You do not have enough supplies for your unit!!!", MsgBoxStyle.Critical)
                For Each R As cRegiment In Mercenary.Regiments.Values
                    R.SupplyMonthsMissed += 1
                Next
            Else
                For Each R As cRegiment In Mercenary.Regiments.Values
                    R.SupplyMonthsMissed -= 1
                    If R.SupplyMonthsMissed < 0 Then R.SupplyMonthsMissed = 0
                Next
                MsgBox(Format(Mercenary.SquadCount, "###,###") & " supply removed for troops", vbOKOnly, "Supplies")
                Mercenary.Supplies_On_Hand -= Mercenary.SquadCount
            End If
            cmdProvideSupplies.Enabled = False
            cmdProvideSupplies.Text = "Unit Supplied"
            cmdPaySalaries.Enabled = True
            RefreshLabels()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub cmdPaySalaries_Click(sender As Object, e As EventArgs) Handles cmdPaySalaries.Click
        Try
            If Mercenary.Salaries > Mercenary.Cash_On_Hand Then
                MsgBox("You do not have enough money to pay for your troops salaries!!!  This could result in mutinies!", MsgBoxStyle.OkOnly, "Insufficient Funds")
                For Each R As cRegiment In Mercenary.Regiments.Values
                    R.BackPayMonths += 1
                    R.BackPayDue += R.Salary
                Next
            Else
                Dim TempSal As Int64 = 0
                For Each R As cRegiment In Mercenary.Regiments.Values
                    R.BackPayMonths -= 1
                    If R.BackPayMonths < 0 Then R.BackPayMonths = 0
                    TempSal += R.Salary
                Next
                MsgBox(Format(TempSal, "###,###") & " C-Bills removed to pay salaries", vbOKOnly, "Payroll")
                Mercenary.Cash_On_Hand -= TempSal
            End If
            cmdPaySalaries.Enabled = False
            cmdPaySalaries.Text = "Salaries Paid"
            cmdPayOverhead.Enabled = True
            RefreshLabels()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub cmdPayOverhead_Click(sender As Object, e As EventArgs) Handles cmdPayOverhead.Click
        Try
            Dim TempOverhead As Int64 = (Mercenary.Salaries + Mercenary.SquadCount * 500) * (Mercenary.Overhead / 100)
            If TempOverhead > Mercenary.Cash_On_Hand Then
                MsgBox("You do not have enough money to pay for unit overhead!!!", MsgBoxStyle.OkOnly, "Insufficient Funds")
                Dim TempVal As Int64 = "0" & InputBox("You can change the overhead, but lowering it under 15 may result in reduced morale!", "Overhead", 15)
                If ((Mercenary.Maintenance_Cost + Mercenary.Salaries) * (TempVal / 100)) > Mercenary.Cash_On_Hand Then
                    MsgBox("That is still more money than you have available, please try again", vbOKOnly, "Insufficient Funds")
                    Exit Sub
                Else
                    Mercenary.Overhead = TempVal
                End If
                MsgBox(Format(TempVal, "###,###") & " C-Bills removed to pay overhead", vbOKOnly, "Overhead")
                Mercenary.Cash_On_Hand -= TempVal
            Else
                Dim TempVal As Int64 = (Mercenary.Salaries + Mercenary.SquadCount * 500) * (Mercenary.Overhead / 100)
                MsgBox(Format(TempVal, "###,###") & " C-Bills removed to pay overhead", vbOKOnly, "Overhead")
                Mercenary.Cash_On_Hand -= TempVal
            End If
            If Mercenary.Overhead = 0 Then
                For Each R As cRegiment In Mercenary.Regiments.Values
                    R.MonthsWithoutOverhead += 1
                Next
            End If
            cmdPayOverhead.Enabled = False
            cmdPayOverhead.Text = "Overhead Paid"
            cmdProvideMaintenance.Enabled = True
            RefreshLabels()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub cmdProvideMaintenance_Click(sender As Object, e As EventArgs) Handles cmdProvideMaintenance.Click
        Try
            Dim TempMaintCost As Int32 = 0
            'TODO: can be percentage of (p40)
            Dim AnyMissed As Boolean = False
            For Each R As cRegiment In Mercenary.Regiments.Values
                If R.SupplyMonthsMissed > 0 Then
                    TempMaintCost += R.Maintenance_Cost * 2
                    AnyMissed = True
                Else
                    TempMaintCost += R.Maintenance_Cost
                End If
            Next
            If AnyMissed Then
                MsgBox("Due to your troops not having adequate supplies, some maintenance costs have doubled!", MsgBoxStyle.OkOnly, "Insufficient Supplies")
            End If

            If TempMaintCost > Mercenary.Cash_On_Hand Then
                Dim TempVal As Int32 = InputBox("You do not have enough money to pay for all of your unit maintenance!!!" & vbCrLf & "Please put in a percentage of maintenance you would like to perform." & vbCrLf & "Note that units that are not maintained, will not be able to fight, and may be lost in a defeat!", "Enter Maintenance Amount", 100)
                If (TempMaintCost * TempVal / 100) > Mercenary.Cash_On_Hand Then
                    MsgBox("That is still more money than you have available, please try again", vbOKOnly, "Insufficient Funds")
                    Exit Sub
                End If
                Mercenary.Cash_On_Hand -= TempMaintCost * TempVal / 100
                MsgBox(Format(TempMaintCost * TempVal / 100, "###,###") & " C-Bills removed from cash", vbOKOnly, "Paid Maintenance")
            Else
                MsgBox("You paid " & Format(TempMaintCost, "###,###") & " for maintenance.", MsgBoxStyle.OkOnly, "Maintenance")
                Mercenary.Cash_On_Hand -= TempMaintCost
            End If
            cmdProvideMaintenance.Enabled = False
            cmdProvideMaintenance.Text = "Maintenance Paid"
            cmdAdvance.Enabled = True
            RefreshLabels()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub


    Private Sub cmdOverhead_Click(sender As Object, e As EventArgs) Handles cmdOverhead.Click
        Dim Ov As String = InputBox("Enter new overhead.  Overhead can be 5, 10, 15, or 20 (percent) with penalties if not 15% or higher", "Overhead", "15")
        Try
            If CInt(Ov) < 0 Then Ov = "0"
            If CInt(Ov) > 20 Then Ov = "20"
            Mercenary.Overhead = CInt(Ov)
            RefreshLabels()
        Catch ex As Exception
            MsgBox("Overhead needs to be a numerical value between 0 and 20", MsgBoxStyle.Critical)
            Exit Sub
        End Try
    End Sub

    Private Sub cmdFindContract_Click(sender As Object, e As EventArgs) Handles cmdFindContract.Click
        Me.Enabled = False
        frmNegotiate.Show()
    End Sub

    Private Sub cmdBreak_Click(sender As Object, e As EventArgs) Handles cmdBreak.Click
        If MsgBox("Are you sure you wish to break this contract?  There can be penalties!", vbYesNo, "Break Contract") = MsgBoxResult.No Then Exit Sub
        CurrentContract.Active = False
        CurrentContract.Paid = True 'Lost any chance at money
        cmdBreak.Enabled = False
        cmdFindContract.Enabled = True

        'TODO:  Break contract penalties.  p64 comstar mediation, taking bribes, double-crossing, and surrendering
        RefreshLabels()
    End Sub




    Private Sub cmdSalvage_Click(sender As Object, e As EventArgs) Handles cmdSalvage.Click
        Me.Enabled = False
        frmModify.Show()
        'New squads are 5x its monthly maintenance cost
        'Salvaged squads are 3x its monthly maintenance cost
        'Destroyed squads are 2x its monthly maintenance cost

        'Any destroyed squads can be raised to salvage by spending 5x maintenance cost
        'Salvaged squads can be raised to new by spending 3x maintenance cost

        'purchase new squads
        RefreshLabels()
    End Sub


End Class


