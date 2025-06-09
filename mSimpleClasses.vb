Public Module mSimpleClasses
    Public Class cCapacity
        Public Mechs As Int32
        Public Aerospace As Int32
        Public Aircraft As Int32
        Public SmallCraft As Int32
        Public LAM As Int32
        Public Armor_Light As Int32
        Public Armor_Heavy As Int32
        Public Infantry_Platoons As Int32
        Public Dropships As Int32
        Public Cargo As Single
        Public Fuel As Single
        Public Burn As Single

        Public Crew_Officers As Int32
        Public Crew_Enlisted As Int32
        Public Crew_Gunners As Int32
        Public Crew_Bay As Int32
        Public Pass1 As Int32
        Public Pass2 As Int32
        Public Pass3 As Int32


    End Class
    Public Class cStrength
        Dim Aerospace As Single
        Dim Maneuver As Single
        Dim Combat As Single
    End Class
    Public Class cSupplyOrder
        Public OrderID As Int32
        Public Supplies As Int64
        Public Delay As Int32
        Public Lost As Boolean
    End Class
    Public Class cContract
        Public Active As Boolean = False
        Public Paid As Boolean
        Public Mission As eMission
        Public Mission_General As String = ""
        Public Length As Int32 'In Months
        Public Remuneration As Int32 'C-Bills per Squad
        Public Guarantees As eGuarantee
        Public Command As eCommandRights
        Public Salvage As eSalvageRights
        Public TransportFee As Single       'Multiplier.  Employer provided transport is always free
        Public SupplyFee As Single          'Multiplier
        Public Length_Remaining As Int32

        Public AdvanceAmount As Single 'Percentage advance
        Public FeeAmount As Single     '5% comstar fees sometimes
        Public Value_Total_Advance As Int64
        Public Value_Total_Completion As Int64

        Public ID As String


        Public EmployerDiplomacy As Int32 = D(1, 6) - 1

        Public Position_Mission As Int32 = 0
        Public Position_Length As Int32 = 0
        Public Position_Remuneration As Int32 = 0
        Public Position_Guarantees As Int32 = 0
        Public Position_Command As Int32 = 0
        Public Position_Transport As Int32 = 0
        Public Position_Supply As Int32 = 0
        Public Position_Salvage As Int32 = 0

        Public Position_Length_Final As Int32 = 0
        Public Position_Remuneration_Final As Int32 = 0
        Public Position_Guarantees_Final As Int32 = 0
        Public Position_Command_Final As Int32 = 0
        Public Position_Transport_Final As Int32 = 0
        Public Position_Supply_Final As Int32 = 0
        Public Position_Salvage_Final As Int32 = 0

        Public Negotiated_Length As Boolean = False
        Public Negotiated_Remuneration As Boolean = False
        Public Negotiated_Guarantees As Boolean = False
        Public Negotiated_Command As Boolean = False
        Public Negotiated_Transport As Boolean = False
        Public Negotiated_Supply As Boolean = False
        Public Negotiated_Salvage As Boolean = False

        Public Contract_Distance As Int32 = 0 '0 = Same System, -1 = Same Planet, 1-10 is # of jumps away

        Public Sub New()
            Try
                'Officially these are named 'FC 56605-002-6'.  It has year in there, and # per day, and faction, not sure what is what so I made below
                Select Case Leader.Faction
                    Case eFaction.Davion : ID = "FS "
                    Case eFaction.Steiner : ID = "LC "
                    Case eFaction.Marik : ID = "FWL "
                    Case eFaction.Kurita : ID = "DC "
                    Case eFaction.Liao : ID = "CC "
                    Case eFaction.ComStar : ID = "CS "
                    Case eFaction.Other : ID = "O "
                End Select
                If CurrentTurn.Year = 1 Then CurrentTurn = New DateTime(3025, 1, 1)
                ID &= Mid(CurrentTurn.Year, 3, 2) & Format(CurrentTurn.DayOfYear, "000") & "-" & Format(Now.Millisecond, "000") & "-" & Format(D(1, 100) - 1, "00")

                Position_Mission = EmployerPosition()
                GetActualMission()
                Position_Length = EmployerPosition()
                Position_Remuneration = EmployerPosition()
                Position_Guarantees = EmployerPosition()
                Position_Command = EmployerPosition()
                Position_Transport = EmployerPosition()
                Position_Supply = EmployerPosition()
                Position_Salvage = EmployerPosition()

                Contract_Distance = D(2, 6) - 2
                If Contract_Distance = 0 Then
                    If D(1, 6) > 3 Then Contract_Distance = -1 Else Contract_Distance = 0
                End If

                DebugEvent("Employer Positions (Diplomacy " & EmployerDiplomacy & ")")
                DebugEvent("  Mission: " & Mission_General & " (" & Mission.ToString & " specifically)")
                DebugEvent("  Length: " & GetContractLength(Position_Length) & " (" & Position_Length & ")")
                DebugEvent("  Remuneration: " & GetContractRemuneration(Position_Remuneration) & " (" & Position_Remuneration & ")")
                DebugEvent("  Guarantees: " & GetContractGuarantees(Position_Guarantees) & " (" & Position_Guarantees & ")")
                DebugEvent("  Command: " & GetContractCommandRights(Position_Command) & " (" & Position_Command & ")")
                DebugEvent("  Transport: " & GetContractTransportFees(Position_Transport) & " (" & Position_Transport & ")")

                DebugEvent("  Supply: " & GetContractSupplyFees(Position_Supply) & " (" & Position_Supply & ")")
                DebugEvent("  Salvage: " & GetContractSalvageRights(Position_Salvage) & " (" & Position_Salvage & ")")
                DebugEvent("  Distance: " & Contract_Distance)

                Active = True
            Catch ex As Exception
                Debug.Print(ex.Message.ToString)
            End Try
        End Sub
        Private Sub GetActualMission()
            Select Case Position_Mission
                Case 0 To 5, 96 To 100 'Retainer
                    Mission_General = "Retainer"
                    Select Case D(1, 10)
                        Case 1 : Mission = eMission.Riot_Duty
                        Case 2 : Mission = eMission.Siege_Duty
                        Case 3 : Mission = eMission.Offensive_Campaign
                        Case 4 : Mission = eMission.Defensive_Campaign
                        Case 5 : Mission = eMission.Planetary_Assault
                        Case 6 : Mission = eMission.Relief_Duty
                        Case 7 : Mission = eMission.Guerrilla_Warfare
                        Case 8 : Mission = eMission.Garrison_Duty
                        Case 9 : Mission = eMission.Cadre_Duty
                        Case 10 : Mission = eMission.Security_Duty
                    End Select
                Case 6 To 15
                    Mission_General = "Minor Raid" 'Recon Raid, Objective Raid
                    If D(1, 2) = 1 Then
                        Mission = eMission.Recon_Raid
                    Else
                        Mission = eMission.Objective_Raid
                    End If
                Case 16 To 35
                    Mission_General = "Territorial Campaign" 'Riot Duty, Siege
                    If D(1, 2) = 1 Then
                        Mission = eMission.Riot_Duty
                    Else
                        Mission = eMission.Siege_Duty
                    End If
                Case 36 To 55
                    Mission_General = "Combat Campaign" 'Offensive Campaign, Defensive Campaign
                    If D(1, 2) = 1 Then
                        Mission = eMission.Offensive_Campaign
                    Else
                        Mission = eMission.Defensive_Campaign
                    End If
                Case 56 To 70
                    Mission_General = "Invasion" 'Planetary Assault, Relief Duty
                    If D(1, 2) = 1 Then
                        Mission = eMission.Planetary_Assault
                    Else
                        Mission = eMission.Relief_Duty
                    End If
                Case 71 To 80
                    Mission_General = "Major Raid" 'Diversionary, Guerrilla Warfare
                    If D(1, 2) = 1 Then
                        Mission = eMission.Divisionary_Raid
                    Else
                        Mission = eMission.Guerrilla_Warfare
                    End If
                Case 81 To 95
                    Mission_General = "Static Defense" 'Garrison, Cadre, Security Duty
                    Select Case D(1, 3)
                        Case 1 : Mission = eMission.Garrison_Duty
                        Case 2 : Mission = eMission.Cadre_Duty
                        Case 3 : Mission = eMission.Security_Duty
                    End Select
            End Select
            If Mission_General = "" Then
                DebugEvent("Oops.  This should never be empty.")
            End If
        End Sub

    End Class
End Module
