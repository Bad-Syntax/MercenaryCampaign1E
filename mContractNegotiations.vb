Public Module mContractNegotiations
#Region "Contract Negotiations"
    Public Function EmployerPosition() As Int32
        Return Choose(D(2, 6) - 1, 5, 15, 25, 35, 75, 95, 85, 55, 45, 45, 55)
    End Function
    Public Function GetContractLength(EmployerPosition As Int32) As Int32 'months
        Select Case EmployerPosition
            Case < 11 : Return 12
            Case 11 To 20 : Return 11
            Case 21 To 30 : Return 10
            Case 31 To 40 : Return 9
            Case 41 To 50 : Return 8
            Case 51 To 60 : Return 7
            Case 61 To 70 : Return 6
            Case 71 To 80 : Return 5
            Case 81 To 85 : Return 4
            Case 86 To 90 : Return 3
            Case 91 To 95 : Return 2
            Case > 95 : Return 1
        End Select
        Return 0
    End Function
    Public Function GetContractRemuneration(EmployerPosition As Int32) As Int32 'Note these are per week, not per month
        Select Case EmployerPosition
            Case < 11 : Return 25000
            Case 11 To 20 : Return 20000
            Case 21 To 30 : Return 17500
            Case 31 To 60 : Return 12000
            Case 61 To 80 : Return 10000
            Case 81 To 90 : Return 5000
            Case > 90 : Return 2500
        End Select
        Return 0
    End Function
    Public Function GetContractGuarantees(EmployerPosition As Int32) As eGuarantee
        Select Case EmployerPosition
            Case < 6 : Return eGuarantee.Prior_Payment
            Case 6 To 20 : Return eGuarantee.Advance_Completion
            Case 21 To 40 : Return eGuarantee.Minor_Intermediary
            Case 41 To 80 : Return eGuarantee.ComStar_Intermediary
            Case > 80 : Return eGuarantee.Payment_On_Completion
        End Select
        Return ""
    End Function
    Public Function GetContractCommandRights(EmployerPosition As Int32) As eCommandRights
        Select Case EmployerPosition
            Case < 6 : Return eCommandRights.Independent
            Case 6 To 20 : Return eCommandRights.Mercenary
            Case 21 To 60 : Return eCommandRights.Liaison
            Case 61 To 80 : Return eCommandRights.House
            Case > 80 : Return eCommandRights.Integrated
        End Select
        Return ""
    End Function
    Public Function GetContractTransportFees(EmployerPosition As Int32) As Single
        Select Case EmployerPosition
            Case < 6 : Return 3
            Case 6 To 15 : Return 2.5
            Case 16 To 30 : Return 2
            Case 31 To 50 : Return 1.5
            Case 51 To 80 : Return 1
            Case 81 To 90 : Return 0.5
            Case > 90 : Return 0
        End Select
        Return 0
    End Function
    Public Function GetContractSupplyFees(EmployerPosition As Int32) As Single
        Select Case EmployerPosition
            Case < 6 : Return 1.5
            Case 6 To 15 : Return 1.25
            Case 16 To 30 : Return 1
            Case 31 To 50 : Return 0.75
            Case 51 To 80 : Return 0.5
            Case 81 To 90 : Return 0.25
            Case > 90 : Return 0
        End Select
        Return 0
    End Function
    Public Function GetContractSalvageRights(EmployerPosition As Int32) As eSalvageRights
        Select Case EmployerPosition
            Case < 11 : Return eSalvageRights.Merc_Claims
            Case 11 To 30 : Return eSalvageRights.Payment_In_Kind
            Case 31 To 50 : Return eSalvageRights.Prize_Court_Outright_Grant
            Case 51 To 70 : Return eSalvageRights.Prize_Court_Payment_In_Kind
            Case 71 To 85 : Return eSalvageRights.Employer_Compensation
            Case > 85 : Return eSalvageRights.Employer_Claims
        End Select
        Return 0
    End Function

    'TODO:  These SHOULD be getting called!!!
    Public Function PrizeCourt() As Int32
        Return Choose(D(2, 6) - 1, 5, 10, 15, 20, 25, 30, 40, 50, 60, 70, 80)
    End Function
    Public Function EmployerSalvage() As Single
        Return Choose(D(2, 6) - 1, 1.5, 1.225, 1, 1, 0.9, 0.8, 0.7, 0.6, 0.5, 0.5, 0.4)
    End Function
    Public Function MercSalvage() As Single
        Return Choose(D(2, 6) - 1, 3, 2.75, 2.5, 2.25, 2, 1.75, 1.5, 1.25, 1, 1, 0.75)
    End Function


    Public Function MissionModifier_Length(Mission As eMission) As Int32
        Select Case Mission
            Case eMission.Cadre_Duty : Return 8
            Case eMission.Garrison_Duty : Return 6
            Case eMission.Security_Duty : Return 4
            Case eMission.Riot_Duty : Return 3
            Case eMission.Siege_Duty : Return 3
            Case eMission.Relief_Duty : Return -3
            Case eMission.Planetary_Assault : Return -6
            Case eMission.Offensive_Campaign : Return 0
            Case eMission.Defensive_Campaign : Return 0
            Case eMission.Recon_Raid : Return -6
            Case eMission.Objective_Raid : Return -6
            Case eMission.Divisionary_Raid : Return -3
            Case eMission.Guerrilla_Warfare : Return 6
            Case eMission.Retainer : Return 12
        End Select
        Return 0
    End Function
    Public Function MissionModifier_Command(Mission As eMission) As Int32
        Select Case Mission
            Case eMission.Cadre_Duty : Return -20
            Case eMission.Garrison_Duty : Return -20
            Case eMission.Security_Duty : Return -20
            Case eMission.Riot_Duty
            Case eMission.Siege_Duty
            Case eMission.Relief_Duty
            Case eMission.Planetary_Assault
            Case eMission.Offensive_Campaign
            Case eMission.Defensive_Campaign
            Case eMission.Recon_Raid : Return -50
            Case eMission.Objective_Raid : Return -50
            Case eMission.Divisionary_Raid : Return -30
            Case eMission.Guerrilla_Warfare : Return -60
            Case eMission.Retainer
        End Select
        Return 0
    End Function
    Public Function MissionModifier_Remuneration(Mission As eMission) As Single
        Select Case Mission
            Case eMission.Cadre_Duty : Return 0.5
            Case eMission.Garrison_Duty : Return 0.5
            Case eMission.Security_Duty : Return 0.75
            Case eMission.Riot_Duty : Return 0.75
            Case eMission.Siege_Duty : Return 1
            Case eMission.Relief_Duty : Return 1.25
            Case eMission.Planetary_Assault : Return 1.5
            Case eMission.Offensive_Campaign : Return 1
            Case eMission.Defensive_Campaign : Return 1
            Case eMission.Recon_Raid : Return 1.5
            Case eMission.Objective_Raid : Return 1.75
            Case eMission.Divisionary_Raid : Return 2
            Case eMission.Guerrilla_Warfare : Return 1.75
            Case eMission.Retainer : Return 1
        End Select
        Return 1
    End Function
#End Region
End Module
