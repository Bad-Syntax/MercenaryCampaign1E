Imports System.ComponentModel

Public Class frmNegotiate
    Dim Negotiating_Points As Int32 = 0

    'Each dropship on each jump is 50K C-Bills
    'Can rent space aboard commercial dropships, that will not land in combat zones, +10% per jump, 10K per BM/AF/Support and 500 per Infantry and 2500 per Vehicle
    'Combat dropships are 10x the cost to use

    Private Sub frmNegotiate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO:  Remove this, just here for testing:
        nudPosLength.Value = 0
        nudPosRemuneration.Value = 10
        nudPosGuarantees.Value = 100
        nudPosCommandRights.Value = 20
        nudPosTransport.Value = 80
        nudPosSupply.Value = 30
        nudPosSalvage.Value = 10

        UpdateContractList()

        Label1.Visible = False
        cbArea.Visible = False
        Label10.Visible = False
        nudNegPts.Visible = False
        labPointsLeft.Visible = False
        labArea.Visible = False
        grpNegotiate.Visible = False
        cmdAccept.Visible = False
        cmdNegotiate.Visible = True

        'Show some initial values:
        nudPosLength_ValueChanged(Nothing, Nothing)
        nudPosRemuneration_ValueChanged(Nothing, Nothing)
        nudPosGuarantees_ValueChanged(Nothing, Nothing)
        nudPosCommandRights_ValueChanged(Nothing, Nothing)
        nudPosTransport_ValueChanged(Nothing, Nothing)
        nudPosSupply_ValueChanged(Nothing, Nothing)
        nudPosSalvage_ValueChanged(Nothing, Nothing)

    End Sub
    Public Sub UpdateContractList()
        Dim Leads As Int32 = ContractLeads(Mercenary.DragoonRating)
        lstContracts.Items.Clear()
        If Leads = 0 Then
            lstContracts.Items.Add("No leads at this time")
            cmdNegotiate.Visible = False
        Else
            ContractList.Clear()
            For I As Int32 = 1 To Leads
                Dim TempContract As New cContract
                ContractList.Add(TempContract.ID, TempContract)
                lstContracts.Items.Add(TempContract.ID & "  " & TempContract.Mission_General)
            Next
            cmdNegotiate.Visible = True
        End If
        lstContracts.SelectedIndex = 0
    End Sub
    Private Function ContractLeads(DragoonRating As eDragoonRating) As Int32
        Dim LeadRoll As Int32 = D(2, 6)
        Select Case Mercenary.DragoonRating
            Case eDragoonRating.A, eDragoonRating.A_Minus, eDragoonRating.A_Plus : LeadRoll += 5
            Case eDragoonRating.B, eDragoonRating.B_Minus, eDragoonRating.B_Plus : LeadRoll += 3
            Case eDragoonRating.C, eDragoonRating.C_Minus, eDragoonRating.C_Plus : LeadRoll += 2
            Case eDragoonRating.D, eDragoonRating.D_Minus, eDragoonRating.D_Plus : LeadRoll += 1
            Case eDragoonRating.F
        End Select
        Dim TempLeads As Int32 = 0
        Select Case LeadRoll
            Case 2, 3 : TempLeads = 0
            Case 4, 5 : TempLeads = 1
            Case 6, 7, 8, 9 : TempLeads = 2
            Case 10, 11, 12 : TempLeads = 3
            Case > 12 : TempLeads = 4
        End Select
        '       p102, the # of contract offers.  I can use this per month, after which none are available and player must advance
        '       2-3=0, 4-5=1, 6-9=2, 10-12=3, 13+=4, adding +1/+2/+3/+5 for D/C/B/A dragoon rating
        Return TempLeads
    End Function


    Private Function PlayerNegotiatingPoints(Reputation As Int32, Diplomacy As Int32) As Int32
        Return Reputation + Diplomacy * 10 - CurrentContract.EmployerDiplomacy * 10
    End Function
    Private Sub cmdNegotiate_Click(sender As Object, e As EventArgs) Handles cmdNegotiate.Click
        If lstContracts.SelectedIndex < 0 Or lstContracts.SelectedItem = "No leads at this time" Then
            MsgBox("You do not have any more contract opportunities available this month." & vbCrLf & "You must advance to the next month to (hopefully) find new opportunities.", MsgBoxStyle.OkOnly, "No leads on new contracts")
            cmdNegotiate.Visible = False
            Exit Sub
        End If
        Try

            Label1.Visible = True
            cbArea.Visible = True
            Label10.Visible = True
            nudNegPts.Visible = True
            labPointsLeft.Visible = True
            labArea.Visible = True
            grpNegotiate.Visible = True
            cmdAccept.Visible = True
            cmdNegotiate.Visible = False


            cbArea.SelectedIndex = 0
            nudPosLength.Enabled = False
            nudPosRemuneration.Enabled = False
            nudPosGuarantees.Enabled = False
            nudPosCommandRights.Enabled = False
            nudPosTransport.Enabled = False
            nudPosSupply.Enabled = False
            nudPosSalvage.Enabled = False

            CurrentContract.Negotiated_Length = False
            CurrentContract.Negotiated_Remuneration = False
            CurrentContract.Negotiated_Guarantees = False
            CurrentContract.Negotiated_Command = False
            CurrentContract.Negotiated_Transport = False
            CurrentContract.Negotiated_Supply = False
            CurrentContract.Negotiated_Salvage = False

            Negotiating_Points = Mercenary.Reputation + Leader.Skills(eSkill.Diplomacy) * 10 - CurrentContract.EmployerDiplomacy * 10
            DebugEvent("Initial Negotiating Points from " & Mercenary.Reputation & " reputation unit with a leader diplmacy of " & Leader.Skills(eSkill.Diplomacy) & " and enemy diplomacy of " & CurrentContract.EmployerDiplomacy & " is " & Negotiating_Points, True)
            nudNegPts.Maximum = Math.Min(100, Negotiating_Points) 'Can't use more than we have
            nudNegPts.Value = 0

            Dim TotalDiff As Int32 = 0 'This is the # different between each 2 numbers, assuming my value is less than theirs
            If nudPosLength.Value >= CurrentContract.Position_Length Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosLength.Value - CurrentContract.Position_Length & " for length", True)
                Negotiating_Points += nudPosLength.Value - CurrentContract.Position_Length
                'TODO:  Maybe don't auto do length?
                'Negotiated_Length = True
            Else
                TotalDiff += CurrentContract.Position_Length - nudPosLength.Value
            End If
            If nudPosRemuneration.Value >= CurrentContract.Position_Remuneration Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosRemuneration.Value - CurrentContract.Position_Remuneration & " for remuneration", True)
                Negotiating_Points += nudPosRemuneration.Value - CurrentContract.Position_Remuneration
                CurrentContract.Negotiated_Remuneration = True
            Else
                TotalDiff += CurrentContract.Position_Remuneration - nudPosRemuneration.Value
            End If
            If nudPosGuarantees.Value >= CurrentContract.Position_Guarantees Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosGuarantees.Value - CurrentContract.Position_Guarantees & " for guarantees", True)
                Negotiating_Points += nudPosGuarantees.Value - CurrentContract.Position_Guarantees
                CurrentContract.Negotiated_Guarantees = True
            Else
                TotalDiff += CurrentContract.Position_Guarantees - nudPosGuarantees.Value
            End If
            If nudPosCommandRights.Value >= CurrentContract.Position_Command Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosCommandRights.Value - CurrentContract.Position_Command & " for command rights", True)
                Negotiating_Points += nudPosCommandRights.Value - CurrentContract.Position_Command
                CurrentContract.Negotiated_Command = True
            Else
                TotalDiff += CurrentContract.Position_Command - nudPosCommandRights.Value
            End If
            If nudPosTransport.Value >= CurrentContract.Position_Transport Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosTransport.Value - CurrentContract.Position_Transport & " for transport", True)
                Negotiating_Points += nudPosTransport.Value - CurrentContract.Position_Transport
                CurrentContract.Negotiated_Transport = True
            Else
                TotalDiff += CurrentContract.Position_Transport - nudPosTransport.Value
            End If
            If nudPosSupply.Value >= CurrentContract.Position_Supply Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosSupply.Value - CurrentContract.Position_Supply & " for supply", True)
                Negotiating_Points += nudPosSupply.Value - CurrentContract.Position_Supply
                CurrentContract.Negotiated_Supply = True
            Else
                TotalDiff += CurrentContract.Position_Supply - nudPosSupply.Value
            End If
            If nudPosSalvage.Value >= CurrentContract.Position_Salvage Then
                DebugEvent("Negotiating Points: " & Negotiating_Points & " adds " & nudPosSalvage.Value - CurrentContract.Position_Salvage & " for salvage", True)
                Negotiating_Points += nudPosSalvage.Value - CurrentContract.Position_Salvage
                CurrentContract.Negotiated_Salvage = True
            Else
                TotalDiff += CurrentContract.Position_Salvage - nudPosSalvage.Value
            End If

            labPointsLeft.Text = Negotiating_Points & " Negotiating Points Remaining"

            If TotalDiff <= Negotiating_Points Then
                MsgBox("You are such a reputable unit, your employer accepted all your demands!", vbOKOnly, "Success!")
                CurrentContract.Negotiated_Length = True
                CurrentContract.Negotiated_Remuneration = True
                CurrentContract.Negotiated_Guarantees = True
                CurrentContract.Negotiated_Command = True
                CurrentContract.Negotiated_Transport = True
                CurrentContract.Negotiated_Supply = True
                CurrentContract.Negotiated_Salvage = True

                Select Case cbArea.SelectedIndex
                    Case 0
                        CurrentContract.Position_Length_Final = nudPosLength.Value
                        CurrentContract.Negotiated_Length = True
                        labArea.Text = "Agreed:  " & GetContractLength(nudPosLength.Value)
                    Case 1
                        CurrentContract.Position_Remuneration_Final = nudPosRemuneration.Value
                        CurrentContract.Negotiated_Remuneration = True
                        labArea.Text = "Agreed:  " & GetContractRemuneration(nudPosRemuneration.Value)
                    Case 2
                        CurrentContract.Position_Guarantees_Final = nudPosGuarantees.Value
                        CurrentContract.Negotiated_Guarantees = True
                        labArea.Text = "Agreed:  " & GetContractGuarantees(nudPosGuarantees.Value).ToString
                    Case 3
                        CurrentContract.Position_Command_Final = nudPosCommandRights.Value
                        CurrentContract.Negotiated_Command = True
                        labArea.Text = "Agreed:  " & GetContractCommandRights(nudPosCommandRights.Value).ToString
                    Case 4
                        CurrentContract.Position_Transport_Final = nudPosTransport.Value
                        CurrentContract.Negotiated_Transport = True
                        labArea.Text = "Agreed:  " & GetContractTransportFees(nudPosTransport.Value)
                    Case 5
                        CurrentContract.Position_Supply_Final = nudPosSupply.Value
                        CurrentContract.Negotiated_Supply = True
                        labArea.Text = "Agreed:  " & GetContractSupplyFees(nudPosSupply.Value)
                    Case 6
                        CurrentContract.Position_Salvage_Final = nudPosSalvage.Value
                        CurrentContract.Negotiated_Salvage = True
                        labArea.Text = "Agreed:  " & GetContractSalvageRights(nudPosSalvage.Value).ToString
                End Select

                cmdAccept_Click(Nothing, Nothing)
            End If
            cbArea.Enabled = True
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Private Sub nudPosLength_ValueChanged(sender As Object, e As EventArgs) Handles nudPosLength.ValueChanged
        labDesLength.Text = GetContractLength(nudPosLength.Value) & " Months"
    End Sub
    Private Sub nudPosRemuneration_ValueChanged(sender As Object, e As EventArgs) Handles nudPosRemuneration.ValueChanged
        labDesRemuneration.Text = "€" & GetContractRemuneration(nudPosRemuneration.Value) & " Per Squad"
    End Sub
    Private Sub nudPosGuarantees_ValueChanged(sender As Object, e As EventArgs) Handles nudPosGuarantees.ValueChanged
        labDesGuarantees.Text = GetContractGuarantees(nudPosGuarantees.Value).ToString
    End Sub
    Private Sub nudPosCommandRights_ValueChanged(sender As Object, e As EventArgs) Handles nudPosCommandRights.ValueChanged
        labDesCommand.Text = GetContractCommandRights(nudPosCommandRights.Value).ToString
    End Sub
    Private Sub nudPosTransport_ValueChanged(sender As Object, e As EventArgs) Handles nudPosTransport.ValueChanged
        labDesTransport.Text = GetContractTransportFees(nudPosTransport.Value) * 100 & "%"
    End Sub
    Private Sub nudPosSupply_ValueChanged(sender As Object, e As EventArgs) Handles nudPosSupply.ValueChanged
        labDesSupply.Text = GetContractSupplyFees(nudPosSupply.Value) * 100 & "%"
    End Sub
    Private Sub nudPosSalvage_ValueChanged(sender As Object, e As EventArgs) Handles nudPosSalvage.ValueChanged
        labDesSalvage.Text = GetContractSalvageRights(nudPosSalvage.Value).ToString
    End Sub


    Private Sub cbArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbArea.SelectedIndexChanged
        Try
            Dim Negotiated As Boolean = False
            nudNegPts.Minimum = 0
            nudNegPts.Maximum = 100
            nudNegPts.Value = 0
            Dim TempVal As String = ""
            Select Case cbArea.SelectedIndex
                Case 0
                    Negotiated = CurrentContract.Negotiated_Length
                    If Negotiated Then TempVal = GetContractLength(CurrentContract.Position_Length_Final) Else TempVal = GetContractLength(CurrentContract.Position_Length)
                Case 1
                    Negotiated = CurrentContract.Negotiated_Remuneration
                    If Negotiated Then TempVal = GetContractRemuneration(CurrentContract.Position_Remuneration_Final) Else TempVal = GetContractRemuneration(CurrentContract.Position_Remuneration)
                Case 2
                    Negotiated = CurrentContract.Negotiated_Guarantees
                    If Negotiated Then TempVal = GetContractGuarantees(CurrentContract.Position_Guarantees_Final).ToString Else TempVal = GetContractGuarantees(CurrentContract.Position_Guarantees).ToString
                Case 3
                    Negotiated = CurrentContract.Negotiated_Command
                    If Negotiated Then TempVal = GetContractCommandRights(CurrentContract.Position_Command_Final).ToString Else TempVal = GetContractCommandRights(CurrentContract.Position_Command).ToString
                Case 4
                    Negotiated = CurrentContract.Negotiated_Transport
                    If Negotiated Then TempVal = GetContractTransportFees(CurrentContract.Position_Transport_Final) Else TempVal = GetContractTransportFees(CurrentContract.Position_Transport)
                Case 5
                    Negotiated = CurrentContract.Negotiated_Supply
                    If Negotiated Then TempVal = GetContractSupplyFees(CurrentContract.Position_Supply_Final) Else TempVal = GetContractSupplyFees(CurrentContract.Position_Supply)
                Case 6
                    Negotiated = CurrentContract.Negotiated_Salvage
                    If Negotiated Then TempVal = GetContractSalvageRights(CurrentContract.Position_Salvage_Final).ToString Else TempVal = GetContractSalvageRights(CurrentContract.Position_Salvage).ToString
            End Select
            If Negotiated Then
                labArea.Text = "Accepted Terms: " & TempVal
                cmdAccept.Enabled = False
                nudNegPts.Enabled = False
            Else
                labArea.Text = "Employer Proposed Terms: " & TempVal
                cmdAccept.Enabled = True
                nudNegPts.Enabled = True
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub nudNegPts_ValueChanged(sender As Object, e As EventArgs) Handles nudNegPts.ValueChanged
        Try
            If nudNegPts.Value > 0 Then cbArea.Enabled = False
            Select Case cbArea.SelectedIndex
                Case 0
                    labArea.Text = "Negotiated Terms:  " & GetContractLength(CurrentContract.Position_Length - nudNegPts.Value)
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Length - nudPosLength.Value), Negotiating_Points) 'Can't use more than we have
                Case 1
                    labArea.Text = "Negotiated Terms:  " & GetContractRemuneration(CurrentContract.Position_Remuneration - nudNegPts.Value)
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Remuneration - nudPosRemuneration.Value), Negotiating_Points) 'Can't use more than we have
                Case 2
                    labArea.Text = "Negotiated Terms:  " & GetContractGuarantees(CurrentContract.Position_Guarantees - nudNegPts.Value).ToString
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Guarantees - nudPosGuarantees.Value), Negotiating_Points) 'Can't use more than we have
                Case 3
                    labArea.Text = "Negotiated Terms:  " & GetContractCommandRights(CurrentContract.Position_Command - nudNegPts.Value).ToString
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Command - nudPosCommandRights.Value), Negotiating_Points) 'Can't use more than we have
                Case 4
                    labArea.Text = "Negotiated Terms:  " & GetContractTransportFees(CurrentContract.Position_Transport - nudNegPts.Value)
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Transport - nudPosTransport.Value), Negotiating_Points) 'Can't use more than we have
                Case 5
                    labArea.Text = "Negotiated Terms:  " & GetContractSupplyFees(CurrentContract.Position_Supply - nudNegPts.Value)
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Supply - nudPosSupply.Value), Negotiating_Points) 'Can't use more than we have
                Case 6
                    labArea.Text = "Negotiated Terms:  " & GetContractSalvageRights(CurrentContract.Position_Salvage - nudNegPts.Value).ToString
                    nudNegPts.Maximum = Math.Min(Math.Max(0, CurrentContract.Position_Salvage - nudPosSalvage.Value), Negotiating_Points) 'Can't use more than we have
            End Select
            nudNegPts.Minimum = nudNegPts.Value
            labPointsLeft.Text = Negotiating_Points - nudNegPts.Value & " Negotiating Points Remaining"

            'TODO:  If no points left, just accept everything else
            If Negotiating_Points - nudNegPts.Value = 0 Then
                'Negotiated_Length = True
                'Negotiated_Remuneration = True
                'Negotiated_Guarantees = True
                'Negotiated_Command = True
                'Negotiated_Transport = True
                'Negotiated_Supply = True
                'Negotiated_Salvage = True
                ' cmdAccept_Click(Nothing, Nothing)
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub cmdAccept_Click(sender As Object, e As EventArgs) Handles cmdAccept.Click
        Debug.Print("if negotating points = 0 accept everything else")
        Try
            'nudNegPts_ValueChanged(Nothing, Nothing)
            Select Case cbArea.SelectedIndex
                Case 0
                    CurrentContract.Position_Length_Final = CurrentContract.Position_Length - nudNegPts.Value
                    CurrentContract.Negotiated_Length = True
                    labArea.Text = "Agreed:  " & GetContractLength(CurrentContract.Position_Length_Final)
                Case 1
                    CurrentContract.Position_Remuneration_Final = CurrentContract.Position_Remuneration - nudNegPts.Value
                    CurrentContract.Negotiated_Remuneration = True
                    labArea.Text = "Agreed:  " & GetContractRemuneration(CurrentContract.Position_Remuneration_Final)
                Case 2
                    CurrentContract.Position_Guarantees_Final = CurrentContract.Position_Guarantees - nudNegPts.Value
                    CurrentContract.Negotiated_Guarantees = True
                    labArea.Text = "Agreed:  " & GetContractGuarantees(CurrentContract.Position_Guarantees_Final).ToString
                Case 3
                    CurrentContract.Position_Command_Final = CurrentContract.Position_Command - nudNegPts.Value
                    CurrentContract.Negotiated_Command = True
                    labArea.Text = "Agreed:  " & GetContractCommandRights(CurrentContract.Position_Command_Final).ToString
                Case 4
                    CurrentContract.Position_Transport_Final = CurrentContract.Position_Transport - nudNegPts.Value
                    CurrentContract.Negotiated_Transport = True
                    labArea.Text = "Agreed:  " & GetContractTransportFees(CurrentContract.Position_Transport_Final)
                Case 5
                    CurrentContract.Position_Supply_Final = CurrentContract.Position_Supply - nudNegPts.Value
                    CurrentContract.Negotiated_Supply = True
                    labArea.Text = "Agreed:  " & GetContractSupplyFees(CurrentContract.Position_Supply_Final)
                Case 6
                    CurrentContract.Position_Salvage_Final = CurrentContract.Position_Salvage - nudNegPts.Value
                    CurrentContract.Negotiated_Salvage = True
                    labArea.Text = "Agreed:  " & GetContractSalvageRights(CurrentContract.Position_Salvage_Final).ToString
            End Select
            nudNegPts.Enabled = False
            Negotiating_Points -= nudNegPts.Value

            cmdAccept.Enabled = False
            If CurrentContract.Negotiated_Length = True And CurrentContract.Negotiated_Remuneration = True And CurrentContract.Negotiated_Guarantees = True And CurrentContract.Negotiated_Command = True And CurrentContract.Negotiated_Transport = True And CurrentContract.Negotiated_Supply = True And CurrentContract.Negotiated_Salvage = True Then
                If MsgBox("Would you like to accept this contract?", vbYesNo, "Accept Contract?") = MsgBoxResult.Yes Then
                    frmOperations.cmdBreak.Enabled = True

                    Label1.Visible = False
                    cbArea.Visible = False
                    Label10.Visible = False
                    nudNegPts.Visible = False
                    labPointsLeft.Visible = False
                    labArea.Visible = False
                    grpNegotiate.Visible = False
                    cmdAccept.Visible = False
                    cmdNegotiate.Visible = False


                    Mercenary.Dishonorable = False

                    CurrentContract.Length = Math.Max(1, GetContractLength(CurrentContract.Position_Length_Final) + MissionModifier_Length(CurrentContract.Mission))
                    CurrentContract.Length_Remaining = CurrentContract.Length
                    CurrentContract.Remuneration = GetContractRemuneration(CurrentContract.Position_Remuneration_Final) * MissionModifier_Remuneration(CurrentContract.Mission) * 4
                    CurrentContract.Command = GetContractCommandRights(CurrentContract.Position_Command_Final + MissionModifier_Command(CurrentContract.Mission))
                    CurrentContract.Guarantees = GetContractGuarantees(CurrentContract.Position_Guarantees_Final)

                    If CurrentContract.Guarantees = eGuarantee.Prior_Payment Then
                        CurrentContract.AdvanceAmount = 1
                    ElseIf CurrentContract.Guarantees = eGuarantee.Advance_Completion Then
                        CurrentContract.AdvanceAmount = 0.5 'TODO:  This isn't really defined as a number, so I just did a half thing
                    ElseIf CurrentContract.Guarantees = eGuarantee.Minor_Intermediary Then
                        CurrentContract.AdvanceAmount = (D(2, 6) - 2) / 100
                    ElseIf CurrentContract.Guarantees = eGuarantee.ComStar_Intermediary Then
                        CurrentContract.AdvanceAmount = 0.25
                        CurrentContract.FeeAmount = 0.05
                    ElseIf CurrentContract.Guarantees = eGuarantee.Payment_On_Completion Then
                        CurrentContract.AdvanceAmount = 0
                    End If
                    CurrentContract.SupplyFee = GetContractSupplyFees(CurrentContract.Position_Supply_Final)
                    CurrentContract.TransportFee = GetContractTransportFees(CurrentContract.Position_Transport_Final)
                    CurrentContract.Salvage = GetContractSalvageRights(CurrentContract.Position_Salvage_Final)
                    frmOperations.Enabled = True
                    frmOperations.RefreshLabels()

                    Mercenary.SquadsStart = Mercenary.SquadCount 'So we can track money excluding any losses


                    'TODO:  Transport costs:
                    '       Could accept free transport by employer
                    '       Or reimburse for some of all of own units being transported
                    If Mercenary.Jumpship_Count = 0 And Mercenary.Dropship_Count = 0 Then
                        'No transports
                        'Add 10K per Mech/Fighter
                        'Add 2.5k per Vehicle
                        'Add 500 per individual infantry
                    Else
                        Dim TransportationCosts As Int64 = 0

                        Dim MechCount As Int32 = 0
                        Dim AerospaceCount As Int32 = 0
                        Dim ArmorCount As Int32 = 0
                        Dim InfantryCount As Int32 = 0
                        Dim LAMCount As Int32 = 0

                        For Each R As cRegiment In Mercenary.Regiments.Values
                            MechCount += R.CountUnitType(eUnitType.Mech_Assault)
                            MechCount += R.CountUnitType(eUnitType.Mech_Heavy)
                            MechCount += R.CountUnitType(eUnitType.Mech_Medium)
                            MechCount += R.CountUnitType(eUnitType.Mech_Light)
                            LAMCount += R.CountUnitType(eUnitType.LAM_Light)
                            LAMCount += R.CountUnitType(eUnitType.LAM_Medium)
                            ArmorCount += R.CountUnitType(eUnitType.Armor_Light)
                            ArmorCount += R.CountUnitType(eUnitType.Armor_Heavy)
                            AerospaceCount += R.CountUnitType(eUnitType.Fighter_Heavy)
                            AerospaceCount += R.CountUnitType(eUnitType.Fighter_Medium)
                            AerospaceCount += R.CountUnitType(eUnitType.Fighter_Light)
                        Next

                        InfantryCount = Math.Max(0, Mercenary.Capacity_Used.Infantry_Platoons - Mercenary.Capacity_Available.Infantry_Platoons)

                        MechCount = Math.Max(0, Mercenary.Capacity_Used.Mechs - Mercenary.Capacity_Available.Mechs)
                        AerospaceCount = Math.Max(0, Mercenary.Capacity_Used.Aerospace - Mercenary.Capacity_Available.Aerospace) + Math.Max(0, Mercenary.Capacity_Used.Aircraft - Mercenary.Capacity_Available.Aircraft)
                        LAMCount = Math.Max(0, Mercenary.Capacity_Used.LAM - Mercenary.Capacity_Available.LAM)

                        ArmorCount = Math.Max(0, Mercenary.Capacity_Used.Armor_Heavy - Mercenary.Capacity_Available.Armor_Heavy) + Math.Max(0, Mercenary.Capacity_Used.Armor_Light - Mercenary.Capacity_Available.Armor_Light)

                        TransportationCosts += 10000 * (MechCount + AerospaceCount)
                        TransportationCosts += 2500 * ArmorCount
                        TransportationCosts += 500 * Mercenary.Personnel(True)
                        TransportationCosts += 50000 * Math.Max(0, Mercenary.Capacity_Used.Dropships - Mercenary.Capacity_Available.Dropships)
                    End If





                    CurrentContract.Value_Total_Completion = (CurrentContract.Remuneration * Mercenary.SquadsStart + CurrentContract.SupplyFee * Mercenary.SquadsStart * 500) * CurrentContract.Length * (1 - CurrentContract.FeeAmount)
                    CurrentContract.Value_Total_Advance = CurrentContract.Value_Total_Completion * CurrentContract.AdvanceAmount

                    CurrentContract.Active = True
                    CurrentContract.Paid = False

                    If CurrentContract.AdvanceAmount > 0 Then 'Go ahead and add some money!
                        Mercenary.Cash_On_Hand += CurrentContract.Value_Total_Advance
                        'currentcontract.Value_Total_Advance = 0
                    End If

                    Forces_Friendly = New List(Of cRegiment)
                    Friendly = New cForce(eForceType.Friendly, Math.Round(Mercenary.SquadCount / 108, 0, MidpointRounding.AwayFromZero)) 'This is done once for the whole campaign
                    For Each R As cRegiment In Friendly.Regiments.Values
                        Forces_Friendly.Add(R)
                    Next
                    For Each R As cRegiment In Mercenary.Regiments.Values
                        Forces_Friendly.Add(R)
                    Next

                    'Now add enemies facing for this contract
                    Enemy = New cForce(eForceType.Enemy)
                    For Each R As cRegiment In Enemy.Regiments.Values
                        Forces_Enemy.Add(R)
                    Next

                    Dim NVPContract() As String = Split(lstContracts.SelectedItem, "  ")
                    CurrentContract = ContractList(NVPContract(0))

                    Me.Hide()
                    frmOperations.cmdProvideSupplies.Enabled = True
                    frmOperations.cmdProvideMaintenance.Enabled = False
                    frmOperations.cmdPayOverhead.Enabled = False
                    frmOperations.cmdPaySalaries.Enabled = False
                    frmOperations.cmdProvideSupplies.Enabled = False
                    frmOperations.cmdFindContract.Enabled = False
                    frmOperations.cmdBreak.Enabled = True
                    frmOperations.Show()
                    frmOperations.Activate()
                    frmOperations.RefreshLabels()
                    frmOperations.cmdBreak.Visible = True
                Else
                    'Since negotiations were abandoned, contract is gone with it
                    Label1.Visible = False
                    cbArea.Visible = False
                    Label10.Visible = False
                    nudNegPts.Visible = False
                    labPointsLeft.Visible = False
                    labArea.Visible = False
                    grpNegotiate.Visible = False
                    cmdAccept.Visible = False
                    cmdNegotiate.Visible = True
                    CurrentContract = Nothing
                End If

                'remove this contract from the list
                Dim NVP() As String = Split(lstContracts.SelectedItem, "  ")
                lstContracts.SelectedIndex = Math.Max(0, lstContracts.SelectedIndex - 1)
                lstContracts.Items.Remove(ContractList(NVP(0)).ID & "  " & ContractList(NVP(0)).Mission_General)
                ContractList.Remove(NVP(0))
                If ContractList.Count = 0 Then
                    lstContracts.Items.Add("No leads at this time")
                    cmdNegotiate.Visible = False
                End If
            Else
                cbArea.Enabled = True
            End If
            cmdAccept.Enabled = True
            frmOperations.cmdProvideSupplies.Enabled = True
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

    Private Sub frmNegotiate_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing 'User clicked 'x'
        frmOperations.Enabled = True
        frmOperations.Show()
        frmOperations.Activate()
    End Sub

    Private Sub lstContracts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstContracts.SelectedIndexChanged
        If lstContracts.SelectedIndex < 0 Then Exit Sub
        nudPosLength.Enabled = True
        nudPosRemuneration.Enabled = True
        nudPosGuarantees.Enabled = True
        nudPosCommandRights.Enabled = True
        nudPosTransport.Enabled = True
        nudPosSupply.Enabled = True
        nudPosSalvage.Enabled = True
        frmOperations.cmdBreak.Visible = False

        Label1.Visible = False
        cbArea.Visible = False
        Label10.Visible = False
        nudNegPts.Visible = False
        labPointsLeft.Visible = False
        labArea.Visible = False
        grpNegotiate.Visible = False
        cmdAccept.Visible = False
        cmdNegotiate.Visible = True

        If lstContracts.SelectedItem = "No leads at this time" Or lstContracts.Items.Count = 0 Then
            cmdNegotiate.Visible = False
            CurrentContract = Nothing
            Exit Sub
        Else
            Dim NVP() As String = Split(lstContracts.SelectedItem, "   ")
            NVP = Split(lstContracts.SelectedItem, "  ")
            CurrentContract = ContractList(NVP(0))
            labContract.Text = "Contract Negotiations For " & CurrentContract.Mission_General & " Contract" & vbCrLf & CurrentContract.ID
            cmdNegotiate.Visible = True
        End If
    End Sub

    Private Sub cmdWalk_Click(sender As Object, e As EventArgs) Handles cmdWalk.Click
        CurrentContract = Nothing

        'remove this contract from the list
        Dim NVP() As String = Split(lstContracts.SelectedItem, "  ")
        lstContracts.SelectedIndex = Math.Max(0, lstContracts.SelectedIndex - 1)
        lstContracts.Items.Remove(ContractList(NVP(0)).ID & "  " & ContractList(NVP(0)).Mission_General)
        ContractList.Remove(NVP(0))
        If ContractList.Count = 0 Then
            lstContracts.Items.Add("No leads at this time")
            cmdNegotiate.Visible = False
        End If

        Label1.Visible = False
        cbArea.Visible = False
        Label10.Visible = False
        nudNegPts.Visible = False
        labPointsLeft.Visible = False
        labArea.Visible = False
        grpNegotiate.Visible = False
        cmdAccept.Visible = False
        cmdNegotiate.Visible = True
    End Sub

    Private Sub cmdReturn_Click(sender As Object, e As EventArgs) Handles cmdReturn.Click
        frmOperations.Enabled = True
        Me.Hide()
        frmOperations.Show()
        frmOperations.Activate()
        frmOperations.RefreshLabels()
    End Sub
End Class

