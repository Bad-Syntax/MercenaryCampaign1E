Public Class cForce
    Public UnitName As String = ""
    Public InitialUP As Int64 = 0
    Public CurrentUP As Int64 = 0
    Public Cash_On_Hand As Int64 = 0

    Public Supplies_On_Hand As Int64 = 0
    Public Supply_Orders As New List(Of cSupplyOrder)
    Public SupplyID As Int32 = 0

    Public Consumables As Int64 = 0

    Public OverallSkill As Int32 = 0

    Public Regiments As New Dictionary(Of String, cRegiment)

    Public Overhead As Int32 = 15 'p40 - This is per force, it is not done at lower levels

    Public MoraleMod As Int32 = 0 'From losses, supplies, overhead, etc
    Public ReputationMod As Int32 = 0
    Public SquadsStart As Int64 = 0

    'These are purely here to get unit names like 1st Air Regiment and 5th Mech Regiment and stuff
    Public SubUnitList_Aerospace As Int32 = 0
    Public SubUnitList_Air As Int32 = 0
    Public SubUnitList_Infantry As Int32 = 0
    Public SubUnitList_Militia As Int32 = 0
    Public SubUnitList_Armor As Int32 = 0
    Public SubUnitList_Mech As Int32 = 0

    Public Capacity_Available As New cCapacity
    Public Capacity_Used As New cCapacity

    Public Dishonorable As Boolean = False

    'TODO:  For dragoon rating, not tracked yet tho
    Public Missions_Completed As Int32 = 0
    Public Missions_Lost As Int32 = 0
    Public Missions_Uncomplete As Int32 = 0
    Public Contract_Breaches As Int32 = 0

    Public Sub New()
        OverallSkill = GetTacticalSkill(eForceComposition.Mech)
    End Sub

    Public Sub New(ForceType As eForceType, Optional MercUnitRegiments As Int64 = 0)
        Try
            Select Case ForceType
                Case eForceType.Player
                    InitialUP = GetInitialUP()
                    CurrentUP = InitialUP
                Case eForceType.Enemy
                    For I As Int32 = 1 To EnemyCommitment(CurrentContract.Mission)
                        AddRegiment()
                    Next
                Case eForceType.Friendly
                    Dim TempReg As Int32 = FriendlyCommitment(CurrentContract.Mission)
                    TempReg -= MercUnitRegiments
                    If TempReg <= 0 Then Exit Sub 'Entire force is the mercs
                    For I As Int32 = 1 To TempReg
                        AddRegiment()
                    Next
            End Select
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub
    Public Sub AddRegiment()
        Dim TempUnit As New cRegiment("")
        TempUnit.ForceComposition = GetRegiment_Enemy(CurrentContract.Mission)
        TempUnit.Experience = GetGenericQuality()
        TempUnit.LeaderSkill = GetTacticalSkill(TempUnit.ForceComposition)
        FillSquads(TempUnit)
        Dim TempName As String = ""
        Select Case TempUnit.ForceComposition
            Case eForceComposition.Aerospace
                SubUnitList_Aerospace += 1
                TempName = GetNumberString(SubUnitList_Aerospace) & " Aerospace Wing"
            Case eForceComposition.Air
                SubUnitList_Air += 1
                TempName = GetNumberString(SubUnitList_Air) & " Air Regiment"
            Case eForceComposition.Infantry
                SubUnitList_Infantry += 1
                TempName = GetNumberString(SubUnitList_Infantry) & " Infantry Regiment"
            Case eForceComposition.Militia
                SubUnitList_Militia += 1
                TempName = GetNumberString(SubUnitList_Militia) & " Militia Regiment"
            Case eForceComposition.Armor
                SubUnitList_Armor += 1
                TempName = GetNumberString(SubUnitList_Armor) & " Armor Regiment"
            Case eForceComposition.Mech
                SubUnitList_Mech += 1
                TempName = GetNumberString(SubUnitList_Mech) & " Mech Regiment"
        End Select
        TempUnit.OriginalSize = TempUnit.SquadCount
        TempUnit.Name = TempName
        DebugEvent("  " & TempName & " (" & TempUnit.SquadCount & " Squads)", True)
        Regiments.Add(TempName, TempUnit)
    End Sub
    Public Sub DoMutiny()
        Try
            Select Case GetMutiny()
                Case eMutiny.A_Morale_Loss
                    MoraleMod -= D(1, 6)
                Case eMutiny.B_Desertions
                    Dim Desertions As Int32 = D(2, 6) - 2 * 10
                Case eMutiny.C_Disobedient_Troops
                    If (D(2, 6) + Leader.Skills(eSkill.Leadership)) <= 7 Then
                        MsgBox("TODO:  All troops refuse to campaign!", vbOKOnly, "Mutiny!")
                    Else
                        MsgBox("TODO:  Some troops refuse to campaign!", vbOKOnly, "Mutiny!")
                    End If
                Case eMutiny.D_Unsanctioned_Attack
                    If (D(2, 6) + Leader.Skills(eSkill.Leadership)) <= 7 Then
                        If Forces_Enemy.Count = 0 Then
                            MsgBox("TODO:  All troops plunder locals!", vbOKOnly, "Mutiny!")
                        Else
                            MsgBox("TODO:  All troops attack enemy!", vbOKOnly, "Mutiny!")
                        End If
                    Else
                        If Forces_Enemy.Count = 0 Then
                            MsgBox("TODO:  Some Mutinous troops plunder locals!", vbOKOnly, "Mutiny!")
                        Else
                            MsgBox("TODO:  Some Mutinous troops attack!", vbOKOnly, "Mutiny!")
                        End If
                    End If
                Case eMutiny.E_Plunder_Locals
                    'TODO:  This continues 1d6 months after mutiny over
                    MsgBox("TODO:  Mutinous troops pillage locals, and unit is now considered to be in hostile territory!", vbOKOnly, "Mutiny!")
                Case eMutiny.F_Riot
                    If (D(2, 6) + Leader.Skills(eSkill.Leadership)) <= 7 Then
                        MsgBox("TODO:  All troops riot!", vbOKOnly, "Mutiny!")
                        Cash_On_Hand = 0
                        Supplies_On_Hand = 0
                    Else
                        MsgBox("TODO:  Some troops riot!", vbOKOnly, "Mutiny!")
                    End If
                Case eMutiny.G_Treachery
                    MsgBox("TODO:  All mutinous troops surrender to first enemy encountered!", vbOKOnly, "Mutiny!")
                Case eMutiny.H_Assassination
                    MsgBox("TODO:  Commander assassinated!", vbOKOnly, "Mutiny!")
            End Select
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Public Function LaunchCampaign() As Boolean 'Just for enemies
        Dim DieRoll As Int32 = D(2, 6)
        'TODO:  This doesn't feel right (p55 MW1e), and should happen far more often IMO
        Dim ModifiedRoll As Int32 = (Forces_Enemy.Count - Forces_Friendly.Count) + DieRoll
        If ModifiedRoll <= Forces_Enemy.Count Then
            DebugEvent("Enemy launches campaign as modified die roll of " & ModifiedRoll & " is less than or equal to 2d6 roll of " & DieRoll & ".", True)
            Return True
        Else
            DebugEvent("Enemy avoids campaign as modified die roll of " & ModifiedRoll & " is greater than than 2d6 roll of " & DieRoll & ".", True)
            Return False
        End If
    End Function
    Public Function Supplies_Pending() As Int64
        Dim TempVal As Int64 = 0
        For Each Order As cSupplyOrder In Supply_Orders
            TempVal += Order.Supplies
        Next
        Return TempVal
    End Function
    Public Function GetInitialUP() As Int64
        Dim UP As Int64 = 0
        Dim DieRoll As Int32 = D(2, 6)
        Select Case DieRoll + Leader.Skills(eSkill.Leadership)
            Case <= 2 : UP = 10000
            Case 3 : UP = 20000
            Case 4 : UP = 35000
            Case 5 : UP = 50000
            Case 6 : UP = 65000
            Case 7 : UP = 80000
            Case 8 : UP = 100000
            Case 9 : UP = 130000
            Case 10 : UP = 160000
            Case 11 : UP = 200000
            Case 12 : UP = 250000
            Case 13 : UP = 300000
            Case 14 : UP = 350000
            Case 15 : UP = 400000
            Case 16 : UP = 450000
            Case 17 : UP = 500000
            Case > 17 : UP = 600000
        End Select
        CurrentUP = UP
        DebugEvent("Initial UP calculation was a 2d6 roll of " & DieRoll & " plus leadership value of " & Leader.Skills(eSkill.Leadership) & " and resulted in " & UP & " Unit Points.", True)
        Return UP
    End Function
    Public Function GetMutiny() As eMutiny
        Dim DieRoll As Int32 = D(1, 36)
        DebugEvent("Mutiny Die Roll is " & DieRoll, True)
        Select Case DieRoll
            Case 1 To 8 : Return eMutiny.A_Morale_Loss
            Case 9 To 15 : Return eMutiny.B_Desertions
            Case 16 To 22 : Return eMutiny.C_Disobedient_Troops
            Case 23 To 26 : Return eMutiny.D_Unsanctioned_Attack
            Case 27 To 29 : Return eMutiny.E_Plunder_Locals
            Case 30 To 32 : Return eMutiny.F_Riot
            Case 33 To 35 : Return eMutiny.G_Treachery
            Case 36 : Return eMutiny.H_Assassination
        End Select
        Return eMutiny.A_Morale_Loss
    End Function
    Public Function DragoonRating() As eDragoonRating
        Dim TempVal As Int32 = 0
        TempVal += Missions_Completed * 5
        TempVal -= Missions_Lost * 10
        TempVal -= Contract_Breaches * 25
        TempVal -= Missions_Uncomplete * 25
        TempVal += Leader.Skills(eSkill.Leadership) * 4

        Select Case OverallSkill
            Case 0, 1, 2 : TempVal += 40
            Case 3 : TempVal += 20
            Case 4 : TempVal += 10
            Case 5, 6, 7 : TempVal += 5
        End Select

        Select Case Support_Generated() / Support_Required()'+5 per 10% over 60% support generated (so 100% is +20)
            Case >= 1 : TempVal += 20
            Case >= 0.9 : TempVal += 15
            Case >= 0.8 : TempVal += 10
            Case >= 0.7 : TempVal += 5
            Case Else : TempVal += 0
        End Select

        '+5 per 10% over 50% transport (25% for 100% DS capacity)
        Dim TotalCap As Int32 = Capacity_Available.Aircraft + Capacity_Available.Aerospace + Capacity_Available.Armor_Light + Capacity_Available.Armor_Heavy + Capacity_Available.Infantry_Platoons + Capacity_Available.Mechs + Capacity_Available.LAM
        Dim UsedCap As Int32 = Capacity_Used.Aircraft + Capacity_Used.Aerospace + Capacity_Used.Armor_Light + Capacity_Used.Armor_Heavy + Capacity_Used.Infantry_Platoons + Capacity_Used.Mechs + Capacity_Used.LAM
        Select Case UsedCap / TotalCap
            Case >= 1 : TempVal += 25
            Case >= 0.9 : TempVal += 20
            Case >= 0.8 : TempVal += 15
            Case >= 0.7 : TempVal += 10
            Case >= 0.6 : TempVal += 5
            Case Else : TempVal += 0
        End Select


        If Jumpship_Count() > 0 Then TempVal += 10

        'Not applicable
        '           +5 per 10% over 30% of upgraded tech

        Select Case TempVal
            Case < 0 : Return eDragoonRating.F
            Case 0 To 11 : Return eDragoonRating.D_Minus
            Case 12 To 33 : Return eDragoonRating.D
            Case 34 To 45 : Return eDragoonRating.D_Plus
            Case 46 To 56 : Return eDragoonRating.C_Minus
            Case 57 To 75 : Return eDragoonRating.C
            Case 76 To 85 : Return eDragoonRating.C_Plus
            Case 86 To 94 : Return eDragoonRating.B_Minus
            Case 95 To 111 : Return eDragoonRating.B
            Case 112 To 120 : Return eDragoonRating.B_Plus
            Case 121 To 132 : Return eDragoonRating.A_Minus
            Case 133 To 144 : Return eDragoonRating.A
            Case >= 145 : Return eDragoonRating.A_Plus
        End Select
    End Function



    'The following are common for all classes
    Public Function Dropship_Count() As Int32
        Dim Dropships As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            Dropships += R.Dropship_Count
        Next
        Return Dropships
    End Function
    Public Function Jumpship_Count() As Int32
        Dim Jumpships As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            Jumpships += R.Jumpship_Count
        Next
        Return Jumpships
    End Function
    Public Function Get_Capacity_Available() As cCapacity
        Capacity_Available.Aerospace = 0
        Capacity_Available.Aircraft = 0
        Capacity_Available.Armor_Heavy = 0
        Capacity_Available.Mechs = 0
        Capacity_Available.Armor_Light = 0
        Capacity_Available.Dropships = 0
        Capacity_Available.LAM = 0
        Capacity_Available.Infantry_Platoons = 0
        For Each R As cRegiment In Regiments.Values
            R.Get_Capacity_Available(Capacity_Available)
        Next
        Return Capacity_Available
    End Function
    Public Function Get_Capacity_Used() As cCapacity
        Dim Capacity_Used As New cCapacity
        Capacity_Used.Aerospace = 0
        Capacity_Used.Aircraft = 0
        Capacity_Used.Armor_Heavy = 0
        Capacity_Used.Mechs = 0
        Capacity_Used.Armor_Light = 0
        Capacity_Used.Dropships = 0
        Capacity_Used.LAM = 0
        Capacity_Used.Infantry_Platoons = 0
        For Each R As cRegiment In Regiments.Values
            R.Get_Capacity_Used(Capacity_Used)
        Next
        Return Capacity_Used
    End Function
    Public Function Reputation() As Int64
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Reputation
        Next
        Return TempVal + ReputationMod
    End Function
    Public Function Cost() As Int64
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Cost
        Next
        Return TempVal
    End Function
    Public Function Maintenance_Cost() As Int64
        Dim SPR As Int64 = 0
        Dim SPG As Int64 = 0
        For Each R As cRegiment In Regiments.Values
            SPR += R.Support_Required
            SPG += R.Support_Generated
        Next
        Return (SPR - SPG) * 5000
    End Function
    Public Function Support_Required() As Int64
        Dim SPR As Int64 = 0
        For Each R As cRegiment In Regiments.Values
            SPR += R.Support_Required
        Next
        Return SPR
    End Function
    Public Function Support_Generated() As Int64
        Dim SPG As Int64 = 0
        For Each R As cRegiment In Regiments.Values
            SPG += R.Support_Generated
        Next
        Return SPG
    End Function
    Public Function Salaries() As Int64
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Salary
        Next
        Return TempVal
    End Function
    Public Function Aerospace() As Single
        Dim TempVal As Single = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Aerospace
        Next
        Return TempVal
    End Function
    Public Function Combat() As Single
        Dim TempVal As Single = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Combat
        Next
        Return TempVal
    End Function
    Public Function Maneuver() As Single
        Dim TempVal As Single = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Maneuver
        Next
        Return TempVal
    End Function
    Public Function SupportSquads() As Int64
        Dim TempVal As Int64 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.SupportSquads
        Next
        Return TempVal
    End Function
    Public Function SquadCount() As Int32
        Dim TempVal As Int64 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.SquadCount
        Next
        Return TempVal
    End Function
    Public Function ForceMorale() As Single
        Dim TempVal As Single = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.TotalMorale
        Next
        If TempVal > 0 Then TempVal = TempVal / SquadCount()
        Return TempVal + MoraleMod
    End Function
    Public Function CountUnitType(UnitType As eUnitType) As Int32
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.CountUnitType(UnitType)
        Next
        Return TempVal
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.Personnel(InfantryOnly)
        Next
        Return TempVal
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        Dim TempVal As Int32 = 0
        For Each R As cRegiment In Regiments.Values
            TempVal += R.CountByState(SquadState)
        Next
        Return TempVal
    End Function
End Class

'Regiments/Battalions/Companies/Lances all have this data
Public MustInherit Class cSmallerFormation
    Public ID As String = Guid.NewGuid.ToString
    Public Name As String = ""
    Public Experience As eExperience
    Public LeaderSkill As Int32 = 0

    Public BackPayDue As Int64 = 0
    Public BackPayMonths As Int32 = 0
    Public SupplyMonthsMissed As Int32 = 0
    Public MonthsWithoutOverhead As Int32 = 0
    Public OriginalSize As Int32 = 0

    Public PlayingItSafe As Int32 = 0 'Page 64, this number reduces rolls for danger/damage/elimination/plunder.  Double this value is a saving roll to be discovered, and if discovered, comstar mediation
    Public PillagingCivilians As Int32 = 0
End Class
Public Class cRegiment
    Inherits cSmallerFormation
    Public Battalions As New Dictionary(Of String, cBattalion)

    Public MoraleMod As Int32 = 0 'From losses, supplies, overhead, etc
    Public ReputationMod As Int32 = 0
    Public SquadsStart As Int64 = 0
    Public ForceComposition As eForceComposition

    Public Sub New(UnitName As String)
        Name = UnitName
    End Sub


    'The following are common for all classes
    Public Function Dropship_Count() As Int32
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Dropship_Count
        Next
        Return TempVal
    End Function
    Public Function Jumpship_Count() As Int32
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Jumpship_Count
        Next
        Return TempVal
    End Function
    Public Sub Get_Capacity_Available(ByRef Capacity As cCapacity)
        For Each B As cBattalion In Battalions.Values
            B.Get_Capacity_Available(Capacity)
        Next
    End Sub
    Public Sub Get_Capacity_Used(ByRef Capacity_Used As cCapacity)
        For Each B As cBattalion In Battalions.Values
            B.Get_Capacity_Used(Capacity_Used)
        Next
    End Sub
    Public Function Reputation() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Reputation
        Next
        Return TempVal
    End Function
    Public Function Cost() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Cost
        Next
        Return TempVal
    End Function
    Public Function Maintenance_Cost() As Int64
        Dim SPR As Int64 = 0
        Dim SPG As Int64 = 0
        For Each B As cBattalion In Battalions.Values
            SPR += B.Support_Required
            SPG += B.Support_Generated
        Next
        Return (SPR - SPG) * 5000
    End Function
    Public Function Support_Required() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Support_Required
        Next
        Return TempVal
    End Function
    Public Function Support_Generated() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Support_Generated
        Next
        Return TempVal
    End Function
    Public Function Salary() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Salary
        Next
        Return TempVal
    End Function
    Public Function Aerospace() As Single
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Aerospace
        Next
        Return TempVal
    End Function
    Public Function Combat() As Single
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Combat
        Next
        Return TempVal
    End Function
    Public Function Maneuver() As Single
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Maneuver
        Next
        Return TempVal
    End Function
    Public Function SupportSquads() As Int64
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.SupportSquads
        Next
        Return TempVal
    End Function
    Public Function SquadCount() As Int32
        Dim TempVal As Int64 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.SquadCount
        Next
        Return TempVal
    End Function
    Public Function TotalMorale() As Single
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.TotalMorale
        Next
        Return TempVal
    End Function
    Public Function CountUnitType(UnitType As eUnitType) As Int32
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.CountUnitType(UnitType)
        Next
        Return TempVal
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempVal As Int32 = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.Personnel(InfantryOnly)
        Next
        Return TempVal
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        Dim TempVal As Single = 0
        For Each B As cBattalion In Battalions.Values
            TempVal += B.CountByState(SquadState)
        Next
        Return TempVal
    End Function
End Class
Public Class cBattalion
    Inherits cSmallerFormation
    Public Companies As New Dictionary(Of String, cCompany)
    Public Sub New(UnitName As String)
        Name = UnitName
    End Sub

    'The following are common for all classes
    Public Function Dropship_Count() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Dropship_Count
        Next
        Return TempVal
    End Function
    Public Function Jumpship_Count() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Jumpship_Count
        Next
        Return TempVal
    End Function
    Public Sub Get_Capacity_Available(ByRef Capacity As cCapacity)
        For Each C As cCompany In Companies.Values
            C.Get_Capacity_Available(Capacity)
        Next
    End Sub
    Public Sub Get_Capacity_Used(ByRef Capacity_Used As cCapacity)
        For Each C As cCompany In Companies.Values
            C.Get_Capacity_Used(Capacity_Used)
        Next
    End Sub
    Public Function Reputation() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Reputation
        Next
        Return TempVal
    End Function
    Public Function Cost() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Cost
        Next
        Return TempVal
    End Function
    Public Function Maintenance_Cost() As Int64
        Dim SPR As Int64 = 0
        Dim SPG As Int64 = 0
        For Each C As cCompany In Companies.Values
            SPR += C.Support_Required
            SPG += C.Support_Generated
        Next
        Return (SPR - SPG) * 5000
    End Function
    Public Function Support_Required() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Support_Required
        Next
        Return TempVal
    End Function
    Public Function Support_Generated() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Support_Generated
        Next
        Return TempVal
    End Function
    Public Function Salary() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Salary
        Next
        Return TempVal
    End Function
    Public Function Aerospace() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Aerospace
        Next
        Return TempVal
    End Function
    Public Function Combat() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Combat
        Next
        Return TempVal
    End Function
    Public Function Maneuver() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Maneuver
        Next
        Return TempVal
    End Function
    Public Function SupportSquads() As Int64
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.SupportSquads
        Next
        Return TempVal
    End Function
    Public Function SquadCount() As Int32
        Dim TempVal As Int64 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.SquadCount
        Next
        Return TempVal
    End Function
    Public Function TotalMorale() As Single
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.TotalMorale
        Next
        Return TempVal
    End Function
    Public Function CountUnitType(UnitType As eUnitType) As Int32
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.CountUnitType(UnitType)
        Next
        Return TempVal
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempVal As Int32 = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.Personnel(InfantryOnly)
        Next
        Return TempVal
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        Dim TempVal As Single = 0
        For Each C As cCompany In Companies.Values
            TempVal += C.CountByState(SquadState)
        Next
        Return TempVal
    End Function

End Class
Public Class cCompany
    Inherits cSmallerFormation
    Public Lances As New Dictionary(Of String, cLance)
    Public Sub New(UnitName As String)
        Name = UnitName
    End Sub

    'The following are common for all classes
    Public Function Dropship_Count() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Dropship_Count
        Next
        Return TempVal
    End Function
    Public Function Jumpship_Count() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Jumpship_Count
        Next
        Return TempVal
    End Function
    Public Sub Get_Capacity_Available(ByRef Capacity As cCapacity)
        For Each L As cLance In Lances.Values
            L.Get_Capacity_Available(Capacity)
        Next
    End Sub
    Public Sub Get_Capacity_Used(ByRef Capacity_Used As cCapacity)
        For Each L As cLance In Lances.Values
            L.Get_Capacity_Used(Capacity_Used)
        Next
    End Sub
    Public Function Reputation() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Reputation
        Next
        Return TempVal
    End Function
    Public Function Cost() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Cost
        Next
        Return TempVal
    End Function
    Public Function Maintenance_Cost() As Int64
        Dim SPR As Int64 = 0
        Dim SPG As Int64 = 0
        For Each L As cLance In Lances.Values
            SPR += L.Support_Required
            SPG += L.Support_Generated
        Next
        Return (SPR - SPG) * 5000
    End Function
    Public Function Support_Required() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Support_Required
        Next
        Return TempVal
    End Function
    Public Function Support_Generated() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Support_Generated
        Next
        Return TempVal
    End Function
    Public Function Salary() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Salary
        Next
        Return TempVal
    End Function
    Public Function Aerospace() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Aerospace
        Next
        Return TempVal
    End Function
    Public Function Combat() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Combat
        Next
        Return TempVal
    End Function
    Public Function Maneuver() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Maneuver
        Next
        Return TempVal
    End Function
    Public Function SupportSquads() As Int64
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.SupportSquads
        Next
        Return TempVal
    End Function
    Public Function SquadCount() As Int32
        Dim TempVal As Int64 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.SquadCount
        Next
        Return TempVal
    End Function
    Public Function TotalMorale() As Single
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.TotalMorale
        Next
        Return TempVal
    End Function
    Public Function CountUnitType(UnitType As eUnitType) As Int32
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.CountUnitType(UnitType)
        Next
        Return TempVal
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempVal As Int32 = 0
        For Each L As cLance In Lances.Values
            TempVal += L.Personnel(InfantryOnly)
        Next
        Return TempVal
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        Dim TempVal As Single = 0
        For Each L As cLance In Lances.Values
            TempVal += L.CountByState(SquadState)
        Next
        Return TempVal
    End Function

End Class
Public Class cLance
    Inherits cSmallerFormation
    Public Squads As New Dictionary(Of String, cSquad)
    Public Sub New(UnitName As String)
        Name = UnitName
    End Sub

    'The following are common for all classes
    Public Function Dropship_Count() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Dropship_Count
        Next
        Return TempVal
    End Function
    Public Function Jumpship_Count() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Jumpship_Count
        Next
        Return TempVal
    End Function
    Public Sub Get_Capacity_Available(ByRef Capacity As cCapacity)
        For Each S As cSquad In Squads.Values
            S.Get_Capacity_Available(Capacity)
        Next
    End Sub
    Public Sub Get_Capacity_Used(ByRef Capacity_Used As cCapacity)
        For Each S As cSquad In Squads.Values
            S.Get_Capacity_Used(Capacity_Used)
        Next
    End Sub
    Public Function Reputation() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Reputation
        Next
        Return TempVal
    End Function
    Public Function Cost() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Cost
        Next
        Return TempVal
    End Function
    Public Function Maintenance_Cost() As Int64
        Dim SPR As Int64 = 0
        Dim SPG As Int64 = 0
        For Each S As cSquad In Squads.Values
            SPR += S.Support_Required
            SPG += S.Support_Generated
        Next
        Return (SPR - SPG) * 5000
    End Function
    Public Function Support_Required() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            If S.State = eSquadState.Active Then TempVal += S.Support_Required
        Next
        Return TempVal
    End Function
    Public Function Support_Generated() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            If S.State = eSquadState.Active Then TempVal += S.Support_Generated
        Next
        Return TempVal
    End Function
    Public Function Salary() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Salary
        Next
        Return TempVal
    End Function
    Public Function Aerospace() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Aerospace
        Next
        Return TempVal
    End Function
    Public Function Combat() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Combat
        Next
        Return TempVal
    End Function
    Public Function Maneuver() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Maneuver
        Next
        Return TempVal
    End Function
    Public Function SupportSquads() As Int64
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.SupportSquads
        Next
        Return TempVal
    End Function
    Public Function SquadCount() As Int32
        Dim TempVal As Int64 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.SquadCount
        Next
        Return TempVal
    End Function
    Public Function TotalMorale() As Single
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Morale
        Next
        Return TempVal
    End Function
    Public Function CountUnitType(UnitType As eUnitType) As Int32
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.CountUnitType(UnitType)
        Next
        Return TempVal
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempVal As Int32 = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.Personnel(InfantryOnly)
        Next
        Return TempVal
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        Dim TempVal As Single = 0
        For Each S As cSquad In Squads.Values
            TempVal += S.CountByState(SquadState)
        Next
        Return TempVal
    End Function
End Class




Public Class cSquad
    Public ID As String = Guid.NewGuid.ToString
    Public UnitType As eUnitType = eUnitType.Mech_Medium
    Public Element As String = ""
    Public Rank As String = ""
    Public Pilot As String = ""
    Public Nickname As String = ""
    Public Quality As eQuality = eQuality.LikeNew
    Public Experience As eExperience = eExperience.Regular
    Public Hireling As Boolean = False
    Public HirelingRate As Single = 0
    Public State As eSquadState = eSquadState.Active
    Public ReputationMod As Int32 = 0
    Public MoraleMod As Int32 = 0


    'The following are called from parent classes
    Public Function Dropship_Count() As Single
        If State <> eSquadState.Killed And State <> eSquadState.Wounded Then
            If UnitType.ToString.ToLower.Contains("dropship") Then Return 1 Else Return 0
        End If
    End Function
    Public Function Jumpship_Count() As Single
        If State <> eSquadState.Killed And State <> eSquadState.Wounded Then
            If UnitType.ToString.ToLower.Contains("jumpship") Then Return 1 Else Return 0
        End If
    End Function
    Public Sub Get_Capacity_Available(ByRef Capacity As cCapacity)
        If State <> eSquadState.Killed Then
            Select Case UnitType
                Case eUnitType.Leopard_Dropship
                    Capacity.Mechs += 4
                    Capacity.Aerospace += 2
                    Capacity.Cargo += 34
                    Capacity.Fuel += 137
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 2
                    Capacity.Crew_Enlisted += 4
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 12
                Case eUnitType.Union_Dropship
                    Capacity.Mechs += 12
                    Capacity.Aerospace += 2
                    Capacity.Cargo += 74
                    Capacity.Fuel += 215
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 3
                    Capacity.Crew_Enlisted += 5
                    Capacity.Crew_Gunners += 6
                    Capacity.Crew_Bay += 28
                Case eUnitType.Overlord_Dropship
                    Capacity.Mechs += 36
                    Capacity.Aerospace += 6
                    Capacity.Cargo += 132
                    Capacity.Fuel += 306
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 8
                    Capacity.Crew_Enlisted += 29
                    Capacity.Crew_Gunners += 6
                    Capacity.Crew_Bay += 84
                Case eUnitType.Fury_Dropship
                    Capacity.Infantry_Platoons += 4
                    Capacity.Armor_Light += 8
                    Capacity.Cargo += 497.5
                    Capacity.Fuel += 140
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 2
                    Capacity.Crew_Enlisted += 3
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 152
                Case eUnitType.Gazelle_Dropship
                    'MW1E had this as also having 2 infantry platoons
                    Capacity.Armor_Heavy += 15
                    Capacity.Cargo += 68
                    Capacity.Fuel += 209
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 2
                    Capacity.Crew_Enlisted += 5
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 120
                Case eUnitType.Seeker_Dropship
                    Capacity.Infantry_Platoons += 4
                    Capacity.Armor_Light += 40 'MW1E had it at 64
                    Capacity.Cargo += 1632
                    Capacity.Fuel += 389
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 4
                    Capacity.Crew_Enlisted += 13
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 312
                Case eUnitType.Triumph_Dropship
                    Capacity.Infantry_Platoons += 4 'MW1E had 5
                    Capacity.Armor_Heavy += 45
                    Capacity.Armor_Light += 8 'MW1E had 0
                    Capacity.Cargo += 749.5
                    Capacity.Fuel += 383
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 3
                    Capacity.Crew_Enlisted += 9
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 512
                Case eUnitType.Condor_Dropship
                    Capacity.Infantry_Platoons += 12
                    Capacity.Armor_Light += 20
                    Capacity.Cargo += 1651
                    Capacity.Fuel += 208
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 4
                    Capacity.Crew_Enlisted += 17
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 436
                Case eUnitType.Excaliber_Dropship
                    Capacity.Infantry_Platoons += 12
                    Capacity.Armor_Light += 0 'MW1E had 24
                    Capacity.Armor_Heavy += 90 'MW1E had 100
                    Capacity.Mechs += 12
                    Capacity.Cargo += 440
                    Capacity.Fuel += 300
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 9
                    Capacity.Crew_Enlisted += 37
                    Capacity.Crew_Gunners += 4
                    Capacity.Crew_Bay += 1080
                Case eUnitType.Scout_Jumpship
                    Capacity.Dropships += 1
                    Capacity.SmallCraft += 1
                    Capacity.Cargo += 518 + 518.5
                    Capacity.Fuel += 46
                    Capacity.Burn += 9.77
                    Capacity.Crew_Officers += 3
                    Capacity.Crew_Enlisted += 15
                    Capacity.Crew_Bay += 5
                Case eUnitType.Merchant_Jumpship
                    Capacity.Dropships += 2
                    Capacity.SmallCraft += 2
                    Capacity.Cargo += 657
                    Capacity.Fuel += 85
                    Capacity.Burn += 19.75
                    Capacity.Crew_Officers += 4
                    Capacity.Crew_Enlisted += 17
                    Capacity.Crew_Bay += 10
                Case eUnitType.Invader_Jumpship 'PPC Variant
                    Capacity.Dropships += 3
                    Capacity.SmallCraft += 2
                    Capacity.Cargo += 285.5 + 286
                    Capacity.Fuel += 50
                    Capacity.Burn += 19.75
                    Capacity.Crew_Officers += 4
                    Capacity.Crew_Enlisted += 19
                    Capacity.Crew_Gunners += 1
                    Capacity.Crew_Bay += 10
                    Capacity.Pass3 += 30
                Case eUnitType.Star_Lord_Jumpship
                    Capacity.Dropships += 6
                    Capacity.SmallCraft += 4
                    Capacity.Cargo += 661
                    Capacity.Fuel += 100
                    Capacity.Burn += 39.52
                    Capacity.Crew_Officers += 5
                    Capacity.Crew_Enlisted += 25
                    Capacity.Crew_Bay += 20
                    Capacity.Pass3 += 50
                Case eUnitType.Monolith_Jumpship
                    Capacity.Dropships += 9
                    Capacity.SmallCraft += 6
                    Capacity.Cargo += 1427.5
                    Capacity.Fuel += 68
                    Capacity.Burn += 39.52
                    Capacity.Crew_Officers += 5
                    Capacity.Crew_Enlisted += 25
                    Capacity.Crew_Bay += 30


                    'Not supported yet
                Case eUnitType.Avenger_Dropship
                    Capacity.Cargo += 126
                    Capacity.Fuel += 160
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 3
                    Capacity.Crew_Enlisted += 8
                    Capacity.Crew_Gunners += 4
                Case eUnitType.Leopard_CV_Dropship
                    Capacity.Aerospace += 6
                    Capacity.Cargo += 87
                    Capacity.Fuel += 137
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 2
                    Capacity.Crew_Enlisted += 4
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 12
                Case eUnitType.Intruder_Dropship
                    Capacity.Infantry_Platoons += 4
                    Capacity.Aerospace += 2
                    Capacity.Cargo += 1108
                    Capacity.Fuel += 300
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 5
                    Capacity.Crew_Enlisted += 18
                    Capacity.Crew_Gunners += 7
                    Capacity.Crew_Bay += 116
                Case eUnitType.Buccaneer_Dropship
                    Capacity.Cargo += 2308.5
                    Capacity.Fuel += 160
                    Capacity.Burn += 2.82
                    Capacity.Crew_Officers += 2
                    Capacity.Crew_Enlisted += 8
                    Capacity.Crew_Gunners += 2
                Case eUnitType.Achilles_Dropship
                    Capacity.Cargo += 160.5
                    Capacity.Aerospace = 2
                    Capacity.SmallCraft = 2
                    Capacity.Infantry_Platoons = 1
                    Capacity.Fuel += 300
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 5
                    Capacity.Crew_Enlisted += 18
                    Capacity.Crew_Gunners += 7
                    Capacity.Crew_Bay += 42
                Case eUnitType.Monarch_Dropship
                    Capacity.Cargo += 1187.5
                    Capacity.Fuel += 112
                    Capacity.Burn += 3.37
                    Capacity.Crew_Officers += 6
                    Capacity.Crew_Enlisted += 28
                    Capacity.Pass1 += 200
                Case eUnitType.Fortress_Dropship
                    Capacity.Mechs += 12
                    Capacity.Armor_Heavy += 12
                    Capacity.Infantry_Platoons += 3
                    Capacity.Cargo += 415.5
                    Capacity.Fuel += 400
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 7
                    Capacity.Crew_Enlisted += 26
                    Capacity.Crew_Gunners += 9
                    Capacity.Crew_Bay += 204
                Case eUnitType.Vengeance_Dropship
                    Capacity.Aerospace += 40
                    Capacity.SmallCraft += 3
                    Capacity.Cargo += 225
                    Capacity.Fuel += 570
                    Capacity.Burn += 1.84
                    Capacity.Crew_Officers += 3
                    Capacity.Crew_Enlisted += 3
                    Capacity.Crew_Gunners += 6
                    Capacity.Crew_Bay += 95
                Case eUnitType.Mule_Dropship
                    Capacity.Cargo += 2715 + 2715 + 2714
                    Capacity.Fuel += 319
                    Capacity.Burn += 4.22
                    Capacity.Crew_Officers += 4
                    Capacity.Crew_Enlisted += 13
                    Capacity.Crew_Gunners += 3
                Case eUnitType.Mammoth_Dropship
                    Capacity.Cargo += 18624 * 2 + 0.5
                    Capacity.Fuel += 420
                    Capacity.Burn += 8.37
                    Capacity.Crew_Officers += 6
                    Capacity.Crew_Enlisted += 27
                    Capacity.Crew_Gunners += 2
                    Capacity.Crew_Bay += 20
                Case eUnitType.Behemoth_Dropship
                    Capacity.Cargo += 37486 * 2
                    Capacity.Fuel += 600
                    Capacity.Burn += 8.83
                    Capacity.Crew_Officers += 9
                    Capacity.Crew_Enlisted += 41
                    Capacity.Crew_Gunners += 3
                    Capacity.Crew_Bay += 100
            End Select
        End If
    End Sub
    Public Sub Get_Capacity_Used(ByRef Capacity_Used As cCapacity)
        If State <> eSquadState.Killed Then
            Select Case UnitType
                Case eUnitType.Mech_Light, eUnitType.Mech_Medium, eUnitType.Mech_Assault, eUnitType.Mech_Heavy
                    Capacity_Used.Mechs += 1
                Case eUnitType.Fighter_Light, eUnitType.Fighter_Medium, eUnitType.Fighter_Heavy
                    Capacity_Used.Aerospace += 1
                Case eUnitType.LAM_Light, eUnitType.LAM_Medium
                    Capacity_Used.LAM += 1
                Case eUnitType.Infantry_Regular, eUnitType.Infantry_Motorized, eUnitType.Infantry_Jump, eUnitType.Infantry_Scout, eUnitType.Support
                    Capacity_Used.Infantry_Platoons += 1
                Case eUnitType.Armor_Light
                    Capacity_Used.Armor_Light += 1
                Case eUnitType.Armor_Heavy
                    Capacity_Used.Armor_Heavy += 1
                Case eUnitType.Artillery
                    Capacity_Used.Infantry_Platoons += 1
                    Capacity_Used.Armor_Heavy += 1
                Case eUnitType.Aircraft
                    Capacity_Used.Aircraft += 1
                Case eUnitType.Infantry_Regular_Airmobile, eUnitType.Infantry_Motorized_Airmobile, eUnitType.Infantry_Jump_Airmobile, eUnitType.Infantry_Scout_Airmobile
                    Capacity_Used.Infantry_Platoons += 1
                    Capacity_Used.Armor_Light += 1
                Case eUnitType.Condor_Dropship, eUnitType.Excaliber_Dropship, eUnitType.Fury_Dropship, eUnitType.Gazelle_Dropship, eUnitType.Leopard_Dropship, eUnitType.Overlord_Dropship, eUnitType.Seeker_Dropship, eUnitType.Triumph_Dropship, eUnitType.Union_Dropship, eUnitType.Avenger_Dropship, eUnitType.Leopard_CV_Dropship, eUnitType.Intruder_Dropship, eUnitType.Buccaneer_Dropship, eUnitType.Achilles_Dropship, eUnitType.Monarch_Dropship, eUnitType.Fortress_Dropship, eUnitType.Vengeance_Dropship, eUnitType.Mule_Dropship, eUnitType.Mammoth_Dropship, eUnitType.Behemoth_Dropship
                    Capacity_Used.Dropships += 1
            End Select
        End If
    End Sub
    Public Function Reputation() As Int32
        If State <> eSquadState.Killed Then
            Dim TempRep As Int32 = 0
            If Experience = eExperience.Elite Then TempRep += 2
            If Experience = eExperience.Veteran Then TempRep += 1
            If UnitType = eUnitType.Mech_Light Then TempRep += 1
            If UnitType = eUnitType.Mech_Medium Then TempRep += 2
            If UnitType = eUnitType.Mech_Heavy Then TempRep += 3
            If UnitType = eUnitType.Mech_Assault Then TempRep += 4
            If UnitType = eUnitType.LAM_Light Or UnitType = eUnitType.LAM_Medium Then TempRep += 3
            If UnitType = eUnitType.Fighter_Heavy Or UnitType = eUnitType.Fighter_Light Or UnitType = eUnitType.Fighter_Medium Then TempRep += 3
            'Debug.Print("Reputation for " & UnitType.ToString & " with pilot " & Experience.ToString & " is " & TempRep)
            Return TempRep + ReputationMod
        Else
            Return 0
        End If
    End Function
    Public Function Cost() As Int32
        If State <> eSquadState.Killed Then
            If Hireling Then Return 0
            Dim TempCost As Int32 = GetUnitCost(UnitType)
            Select Case Experience
                Case eExperience.Green : TempCost *= 0.5
                Case eExperience.Regular : TempCost *= 1
                Case eExperience.Veteran : TempCost *= 2.5
                Case eExperience.Elite : TempCost *= 5
            End Select
            If Quality = eQuality.Salvaged Then
                TempCost *= 0.4
            ElseIf Quality = eQuality.Destroyed Then
                TempCost *= 0.2
            End If
            Return TempCost
        Else
            Return 0
        End If
    End Function
    Public Function Maintenance_Cost() As Int64
        If State <> eSquadState.Killed Then
            Dim SPR As Int64 = 0
            Dim SPG As Int64 = 0
            SPR += Support_Required()
            SPG += Support_Generated()
            Return (SPR - SPG) * 5000
        Else
            Return 0
        End If
    End Function

    Public Function Support_Required() As Int32
        If State <> eSquadState.Killed And Quality <> eQuality.Destroyed Then
            Dim Maintenance As Int32 = 0
            Select Case UnitType
                Case eUnitType.Mech_Light
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 15
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 30
                    End If
                Case eUnitType.Mech_Medium
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 25
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 63
                    End If
                Case eUnitType.Mech_Heavy
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 35
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 123
                    End If
                Case eUnitType.Mech_Assault
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 45
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 225
                    End If
                Case eUnitType.Fighter_Light
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 12
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 40 'Should all AF be the same?
                    End If
                Case eUnitType.Fighter_Medium
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 12
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 40 'Should all AF be the same?
                    End If
                Case eUnitType.Fighter_Heavy
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 12
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 40 'Should all AF be the same?
                    End If
                Case eUnitType.LAM_Light
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 25
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 50
                    End If
                Case eUnitType.LAM_Medium
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 35
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 123
                    End If
                Case eUnitType.Infantry_Regular
                    Maintenance += 1
                Case eUnitType.Infantry_Jump
                    Maintenance += 2
                Case eUnitType.Infantry_Scout
                    Maintenance += 1
                Case eUnitType.Infantry_Motorized
                    Maintenance += 2
                Case eUnitType.Infantry_Regular_Airmobile
                    Maintenance += 5
                Case eUnitType.Infantry_Jump_Airmobile
                    Maintenance += 6
                Case eUnitType.Infantry_Scout_Airmobile
                    Maintenance += 5
                Case eUnitType.Infantry_Motorized_Airmobile
                    Maintenance += 6
                Case eUnitType.Armor_Light
                    Maintenance += 7
                Case eUnitType.Armor_Heavy
                    Maintenance += 13
                Case eUnitType.Artillery
                    Maintenance += 1
                Case eUnitType.Aircraft
                    Maintenance += 5
                Case eUnitType.Support
                    Maintenance += 0
                Case eUnitType.Condor_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Excaliber_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 100
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 200
                    End If
                Case eUnitType.Fury_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 40
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 80
                    End If
                Case eUnitType.Gazelle_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 50
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 100
                    End If
                Case eUnitType.Leopard_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 60
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 120
                    End If
                Case eUnitType.Overlord_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 80
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 160
                    End If
                Case eUnitType.Seeker_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 65
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 130
                    End If
                Case eUnitType.Triumph_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 80
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 160
                    End If
                Case eUnitType.Union_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 70
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 140
                    End If
                Case eUnitType.Invader_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 85
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 170
                    End If
                Case eUnitType.Merchant_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 80
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 160
                    End If
                Case eUnitType.Monolith_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 100
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 200
                    End If
                Case eUnitType.Scout_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Star_Lord_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 90
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 180
                    End If

                    'These are from DS&JS
                Case eUnitType.Avenger_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 65
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 130
                    End If
                Case eUnitType.Achilles_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 80
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 160
                    End If
                Case eUnitType.Intruder_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Fortress_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 110
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 220
                    End If
                Case eUnitType.Leopard_CV_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 65
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 65
                    End If
                Case eUnitType.Vengeance_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Buccaneer_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 50
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 100
                    End If
                Case eUnitType.Mule_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Monarch_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 75
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 150
                    End If
                Case eUnitType.Mammoth_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 150
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 300
                    End If
                Case eUnitType.Behemoth_Dropship
                    If Quality = eQuality.LikeNew Then
                        Maintenance += 250
                    ElseIf Quality = eQuality.Salvaged Then
                        Maintenance += 500
                    End If

            End Select

            Return Maintenance
        Else
            Return 0
        End If

    End Function
    Public Function Support_Generated() As Int32
        If State <> eSquadState.Killed Then
            Dim BaseSupport As Int32 = 0

            Select Case UnitType
                Case eUnitType.Mech_Light, eUnitType.Mech_Medium, eUnitType.Mech_Heavy, eUnitType.Mech_Assault
                    If Experience = eExperience.Elite Then BaseSupport += 1
                Case eUnitType.Fighter_Light, eUnitType.Fighter_Medium, eUnitType.Fighter_Heavy
                    If Experience = eExperience.Elite Then BaseSupport += 1
                Case eUnitType.LAM_Light, eUnitType.LAM_Medium
                    If Experience = eExperience.Elite Then BaseSupport += 1
                Case eUnitType.Support
                    If Experience = eExperience.Green Then
                        BaseSupport += 5
                    ElseIf Experience = eExperience.Regular Then
                        BaseSupport += 10
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSupport += 20
                    ElseIf Experience = eExperience.Elite Then
                        BaseSupport += 30
                    End If
            End Select
            'TODO:  More detailed data from DS&JS.
            '       Every 5 Crew = 1 Support Point
            '       Engineers:  (crew skill / crew skill -1 quality)

            Return BaseSupport
        Else
            Return 0
        End If

    End Function
    Public Function Salary() As Int32
        'TODO:  Campaign Operations states JS salary is 750/ea, DS salary is 1000/ea.  60% if green, 160% if vet, 320% if elite.  Probably half for enlisted.
        If State <> eSquadState.Killed Then
            Dim BaseSalary As Int32 = 0
            Select Case UnitType
                Case eUnitType.Mech_Light, eUnitType.Mech_Medium, eUnitType.Mech_Heavy, eUnitType.Mech_Assault, eUnitType.Fighter_Light, eUnitType.Fighter_Medium, eUnitType.Fighter_Heavy, eUnitType.LAM_Light, eUnitType.LAM_Medium
                    If Experience <= eExperience.Green Then
                        BaseSalary = 400
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 600
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 1000
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 2000
                    End If
                Case eUnitType.Infantry_Scout, eUnitType.Infantry_Scout_Airmobile
                    If Experience <= eExperience.Green Then
                        BaseSalary = 150
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 300
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 600
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 1200
                    End If

                Case eUnitType.Infantry_Regular, eUnitType.Infantry_Jump, eUnitType.Infantry_Motorized, eUnitType.Infantry_Regular_Airmobile, eUnitType.Infantry_Jump_Airmobile, eUnitType.Infantry_Motorized_Airmobile, eUnitType.Artillery
                    '7 man squads
                    If Experience <= eExperience.Green Then
                        BaseSalary = 1050
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 1750
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 3500
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 7000
                    End If

                Case eUnitType.Armor_Light, eUnitType.Armor_Heavy
                    If Experience <= eExperience.Green Then
                        BaseSalary = 1000
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 1750
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 2500
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 5000
                    End If
                Case eUnitType.Aircraft
                    If Experience <= eExperience.Green Then
                        BaseSalary = 250
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 400
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 900
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 1500
                    End If
                Case eUnitType.Support
                    'Tech Salary
                    If Experience <= eExperience.Green Then
                        BaseSalary = 200
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary = 400
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary = 750
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary = 1500
                    End If
                    'Their squad of 7 astechs and assistants
                    If Experience <= eExperience.Green Then
                        BaseSalary += 1050
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary += 1750
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary += 3500
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary += 7000
                    End If



                Case eUnitType.Condor_Dropship
                    BaseSalary = 7600 '4 Officers, 17 Enlisted, 3 Gunners, 436 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Excaliber_Dropship
                    BaseSalary = 7650 '9 Officers, 37 Enlisted, 4 Gunners, 1080 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If
                Case eUnitType.Fury_Dropship
                    BaseSalary = 3100 '2 Officers, 3 Enlisted, 3 Gunners, 152 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Gazelle_Dropship
                    BaseSalary = 3600 '2 Officers, 5 Enlisted, 3 Gunners, 120 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Leopard_Dropship
                    BaseSalary = 3350 '2 Officers, 4 Enlisted, 3 Gunners, 12 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Overlord_Dropship
                    BaseSalary = 13150 '8 Officers, 29 Enlisted, 6 Gunners, 84 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Triumph_Dropship
                    BaseSalary = 5050 '3 Officers, 9 Enlisted, 3 Gunners, 512 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Union_Dropship
                    BaseSalary = 4800 '3 Officers, 5 Enlisted, 6 Gunners, 28 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Seeker_Dropship
                    BaseSalary = 7500 'DS&JS did not list this salary, so I just made it up.   '4 Officers, 13 Enlisted, 3 Gunners, 312 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If


'Not supported yet:
                Case eUnitType.Avenger_Dropship
                    BaseSalary = 5050 '3 Officers, 8 Enlisted, 4 Gunners
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Leopard_CV_Dropship
                    BaseSalary = 3350 '2 Officers, 4 Enlisted, 3 Gunners, 12 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Intruder_Dropship
                    BaseSalary = 17300 '5 Officers, 18 Enlisted, 7 Gunners, 116 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Buccaneer_Dropship
                    BaseSalary = 4300 '2 Officers, 8 Enlisted, 2 Gunners
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Achilles_Dropship
                    BaseSalary = 10800 '5 Officers, 18 Enlisted, 7 Gunners, 42 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Monarch_Dropship
                    BaseSalary = 10150 '6 Officers, 28 Enlisted, 200 1st Class Passengers
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Fortress_Dropship
                    BaseSalary = 12900 '7 Officers, 26 Enlisted, 9 Gunners, 204 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Vengeance_Dropship
                    BaseSalary = 10450 '3 Officers, 3 Enlisted, 6 Gunners, 95 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Mule_Dropship
                    BaseSalary = 5800 '4 Officers, 13 Enlisted, 3 Gunners
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Mammoth_Dropship
                    BaseSalary = 10700 '6 Officers, 27 Enlisted, 2 Gunners, 20 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Behemoth_Dropship
                    BaseSalary = 15800 '9 Officers, 41 Enlisted, 3 Gunners, 100 Bay
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.6
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.5
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.4
                    End If

                Case eUnitType.Invader_Jumpship
                    BaseSalary = 4150
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.583333333333
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.6666666666667
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.5
                    End If

                Case eUnitType.Merchant_Jumpship
                    BaseSalary = 3000
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.583333333333
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.6666666666667
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.5
                    End If
                Case eUnitType.Monolith_Jumpship
                    BaseSalary = 5600
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.583333333333
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.6666666666667
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.5
                    End If
                Case eUnitType.Scout_Jumpship
                    BaseSalary = 2000
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.583333333333
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.6666666666667
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.5
                    End If
                Case eUnitType.Star_Lord_Jumpship
                    BaseSalary = 3500
                    If Experience <= eExperience.Green Then
                        BaseSalary *= 0.583333333333
                    ElseIf Experience = eExperience.Regular Then
                        BaseSalary *= 1
                    ElseIf Experience = eExperience.Veteran Then
                        BaseSalary *= 1.6666666666667
                    ElseIf Experience >= eExperience.Elite Then
                        BaseSalary *= 2.5
                    End If
            End Select
            If Hireling Then
                HirelingRate = GetHirelingRate(UnitType, Experience, State)
                BaseSalary *= HirelingRate
            End If
            Return BaseSalary
        Else
            Return 0
        End If
    End Function
    Public Function Aerospace() As Single
        If State <> eSquadState.Killed And State <> eSquadState.Wounded Then
            Dim TempAerospace As Single = 0
            Select Case UnitType
                Case eUnitType.Fighter_Light : TempAerospace = 3
                Case eUnitType.Fighter_Medium : TempAerospace = 4
                Case eUnitType.Fighter_Heavy : TempAerospace = 5
                Case eUnitType.LAM_Light : TempAerospace = 1
                Case eUnitType.LAM_Medium : TempAerospace = 2
                Case eUnitType.Aircraft : TempAerospace = 2
            End Select
            Return TempAerospace * StrengthMult(Experience)
        Else
            Return 0
        End If
    End Function
    Public Function Combat() As Single
        If State <> eSquadState.Killed And State <> eSquadState.Wounded Then
            Dim TempCombat As Single = 0
            Select Case UnitType
                Case eUnitType.Mech_Light : TempCombat = 3
                Case eUnitType.Mech_Medium : TempCombat = 4
                Case eUnitType.Mech_Heavy : TempCombat = 5
                Case eUnitType.Mech_Assault : TempCombat = 6
                Case eUnitType.LAM_Light : TempCombat = 2
                Case eUnitType.LAM_Medium : TempCombat = 3
                Case eUnitType.Infantry_Regular : TempCombat = 1
                Case eUnitType.Infantry_Jump : TempCombat = 1
                Case eUnitType.Infantry_Motorized : TempCombat = 2
                Case eUnitType.Infantry_Regular_Airmobile : TempCombat = 2
                Case eUnitType.Infantry_Jump_Airmobile : TempCombat = 2
                Case eUnitType.Infantry_Scout_Airmobile : TempCombat = 1
                Case eUnitType.Infantry_Motorized_Airmobile : TempCombat = 3
                Case eUnitType.Armor_Light : TempCombat = 3
                Case eUnitType.Armor_Heavy : TempCombat = 4
                Case eUnitType.Artillery : TempCombat = 4
            End Select
            Return TempCombat * StrengthMult(Experience)
        Else
            Return 0
        End If
    End Function
    Public Function Maneuver() As Single
        If State <> eSquadState.Killed And State <> eSquadState.Wounded Then
            Dim TempManeuver As Single
            Select Case UnitType
                Case eUnitType.Mech_Light : TempManeuver = 6
                Case eUnitType.Mech_Medium : TempManeuver = 5
                Case eUnitType.Mech_Heavy : TempManeuver = 4
                Case eUnitType.Mech_Assault : TempManeuver = 3
                Case eUnitType.LAM_Light : TempManeuver = 6
                Case eUnitType.LAM_Medium : TempManeuver = 5
                Case eUnitType.Infantry_Regular : TempManeuver = 1
                Case eUnitType.Infantry_Jump : TempManeuver = 3
                Case eUnitType.Infantry_Scout : TempManeuver = 3
                Case eUnitType.Infantry_Motorized : TempManeuver = 2
                Case eUnitType.Infantry_Regular_Airmobile : TempManeuver = 3
                Case eUnitType.Infantry_Jump_Airmobile : TempManeuver = 5
                Case eUnitType.Infantry_Scout_Airmobile : TempManeuver = 5
                Case eUnitType.Infantry_Motorized_Airmobile : TempManeuver = 4
                Case eUnitType.Armor_Light : TempManeuver = 3
                Case eUnitType.Armor_Heavy : TempManeuver = 2
                Case eUnitType.Artillery : TempManeuver = 1
            End Select
            Return TempManeuver * StrengthMult(Experience)
        Else
            Return 0
        End If
    End Function
    Public Function SupportSquads() As Int32
        If UnitType <> eUnitType.Support Then
            Return 0
        Else
            If State = eSquadState.Active Or State = eSquadState.Wounded Then Return 1
        End If
    End Function
    Public Function SquadCount() As Int32
        If UnitType = eUnitType.Support Then
            Return 0
        Else
            If State = eSquadState.Active Or State = eSquadState.Wounded Then Return 1
        End If
    End Function
    Public Function Morale() As Int32
        If UnitType <> eUnitType.Support And State <> eSquadState.Killed Then
            Dim TempVal As Int32 = 0
            If Experience = eExperience.Elite Then
                TempVal = 4
            ElseIf Experience = eExperience.Veteran Then
                TempVal = 3
            ElseIf Experience = eExperience.Regular Then
                TempVal = 2
            ElseIf Experience = eExperience.Green Then
                TempVal = 1
            End If
            Return TempVal + MoraleMod
        Else
            Return 0
        End If
    End Function
    Public Function CountUnitType(SquadUnitType As eUnitType) As Int32
        If SquadUnitType = UnitType Then Return 1 Else Return 0
    End Function
    Public Function Personnel(Optional InfantryOnly As Boolean = False) As Int32
        Dim TempPersonnel As Int32 = 0
        If InfantryOnly Then
            Select Case UnitType
                Case eUnitType.Infantry_Regular : TempPersonnel += 7
                Case eUnitType.Infantry_Motorized : TempPersonnel += 10
                Case eUnitType.Infantry_Jump : TempPersonnel += 7'TODO:  These should be smaller
                Case eUnitType.Infantry_Scout : TempPersonnel += 3
                Case eUnitType.Infantry_Regular_Airmobile : TempPersonnel += 10
                Case eUnitType.Infantry_Motorized_Airmobile : TempPersonnel += 10
                Case eUnitType.Infantry_Jump_Airmobile : TempPersonnel += 10'TODO:  These should be smaller
                Case eUnitType.Infantry_Scout_Airmobile : TempPersonnel += 10
            End Select
        Else
            Select Case UnitType'Some of these values taken from bay requirements in Tech Manual
                Case eUnitType.Mech_Light : TempPersonnel += 1'Assumed
                Case eUnitType.Mech_Medium : TempPersonnel += 1'Assumed
                Case eUnitType.Mech_Heavy : TempPersonnel += 1'Assumed
                Case eUnitType.Mech_Assault : TempPersonnel += 1'Assumed
                Case eUnitType.Fighter_Light : TempPersonnel += 1'Assumed
                Case eUnitType.Fighter_Medium : TempPersonnel += 1'Assumed
                Case eUnitType.Fighter_Heavy : TempPersonnel += 1'Assumed
                Case eUnitType.LAM_Light : TempPersonnel += 1'Assumed
                Case eUnitType.LAM_Medium : TempPersonnel += 1'Assumed

                Case eUnitType.Infantry_Regular : TempPersonnel += 7
                Case eUnitType.Infantry_Motorized : TempPersonnel += 10
                Case eUnitType.Infantry_Jump : TempPersonnel += 7'TODO:  These should be smaller
                Case eUnitType.Infantry_Scout : TempPersonnel += 3
                Case eUnitType.Infantry_Regular_Airmobile : TempPersonnel += 10
                Case eUnitType.Infantry_Motorized_Airmobile : TempPersonnel += 10
                Case eUnitType.Infantry_Jump_Airmobile : TempPersonnel += 10'TODO:  These should be smaller
                Case eUnitType.Infantry_Scout_Airmobile : TempPersonnel += 10

                Case eUnitType.Armor_Light : TempPersonnel += 5'From TM
                Case eUnitType.Armor_Heavy : TempPersonnel += 8'From TM
                Case eUnitType.Artillery : TempPersonnel += 30'Assumed
                Case eUnitType.Aircraft : TempPersonnel += 1'Assumed
                Case eUnitType.Support : TempPersonnel += 5 'Assumed

                    'I do not add personnel from dropships, as this is only used for transport costs, and they would clearly self transport, but I left them in here for the fuck of it
                    'Dropships do not include bay personnel, that should be covered above
                    'Case eUnitType.Leopard_Dropship : TempPersonnel += 9'From TRO3057R
                    'Case eUnitType.Union_Dropship : TempPersonnel += 14'From TRO3057R
                    'Case eUnitType.Overlord_Dropship : TempPersonnel += 43'From TRO3057R
                    'Case eUnitType.Fury_Dropship : TempPersonnel += 8'From TRO3057R
                    'Case eUnitType.Gazelle_Dropship : TempPersonnel += 10'From TRO3057R
                    'Case eUnitType.Seeker_Dropship : TempPersonnel += 20'From TRO3057R
                    'Case eUnitType.Triumph_Dropship : TempPersonnel += 15'From TRO3057R
                    'Case eUnitType.Condor_Dropship : TempPersonnel += 24'From TRO3057R
                    'Case eUnitType.Excaliber_Dropship : TempPersonnel += 50'From TRO3057R

                    'Case eUnitType.Scout_Jumpship : TempPersonnel += 18'From TRO3057R
                    'Case eUnitType.Invader_Jumpship : TempPersonnel += 24'From TRO3057R
                    'Case eUnitType.Monolith_Jumpship : TempPersonnel += 30'From TRO3057R
                    'Case eUnitType.Star_Lord_Jumpship : TempPersonnel += 30'From TRO3057R
                    'Case eUnitType.Merchant_Jumpship : TempPersonnel += 21'From TRO3057R

                    'Case eUnitType.Avenger_Dropship : TempPersonnel += 15'From TRO3057R
                    'Case eUnitType.Leopard_CV_Dropship : TempPersonnel += 9'From TRO3057R
                    'Case eUnitType.Intruder_Dropship : TempPersonnel += 31'From TRO3057R
                    'Case eUnitType.Buccaneer_Dropship : TempPersonnel += 12'From TRO3057R
                    'Case eUnitType.Achilles_Dropship : TempPersonnel += 30'From TRO3057R
                    'Case eUnitType.Monarch_Dropship : TempPersonnel += 34'From TRO3057R
                    'Case eUnitType.Fortress_Dropship : TempPersonnel += 42'From TRO3057R
                    'Case eUnitType.Vengeance_Dropship : TempPersonnel += 12'From TRO3057R
                    'Case eUnitType.Mule_Dropship : TempPersonnel += 20'From TRO3057R
                    'Case eUnitType.Mammoth_Dropship : TempPersonnel += 35'From TRO3057R
                    'Case eUnitType.Behemoth_Dropship : TempPersonnel += 53 'From TRO3057R
            End Select
        End If

        Return TempPersonnel
    End Function
    Public Function CountByState(SquadState As eSquadState) As Int32
        If SquadState = State Then Return 1 Else Return 0
    End Function
End Class