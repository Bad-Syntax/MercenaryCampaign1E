Public Class frmCharacter

    Public CanonCharacter As Boolean = True
    Public CharacterValid As Boolean = True
    Public UniversityFailed As Boolean = False

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        If cbName.Text = "" Then cbName.Text = Environment.UserName
        If CharacterValid = False And CanonCharacter <> True Then
            MsgBox("This is not a valid character, please fix and try again")
            Exit Sub
        End If
        Leader.Name = cbName.SelectedItem
        Leader.Rank = txtRank.Text
        Leader.Faction = cbFaction.SelectedIndex

        Leader.BOD = nudBOD.Value
        Leader.CHA = nudCHA.Value
        Leader.LRN = nudLRN.Value
        Leader.DEX = nudDEX.Value
        Leader.Ability_Thick_Skin = clbInborn.CheckedItems.Contains(0)
        Leader.Ability_Glass_Jaw = clbInborn.CheckedItems.Contains(1)
        Leader.Ability_Peripheral_Vision = clbInborn.CheckedItems.Contains(2)
        Leader.Ability_Sixth_Sense = clbInborn.CheckedItems.Contains(3)
        Leader.Ability_Family_Friend = clbInborn.CheckedItems.Contains(4)
        Leader.Ability_Family_Feud = clbInborn.CheckedItems.Contains(5)
        Select Case cbHandedness.SelectedIndex
            Case 0 : Leader.Ability_Handedness = eHandedness.Right
            Case 1 : Leader.Ability_Handedness = eHandedness.Left
            Case 2 : Leader.Ability_Handedness = eHandedness.Natural_Right
            Case 3 : Leader.Ability_Handedness = eHandedness.Natural_Left
            Case 4 : Leader.Ability_Handedness = eHandedness.Either
            Case 5 : Leader.Ability_Handedness = eHandedness.Both
        End Select
        Leader.Ability_Natural_Aptitude = lstNaturalAptitude.SelectedIndex
        Leader.Skills.Clear()
        Leader.Skills.Add(eSkill.Athletics, nudAthletics.Value)
        Leader.Skills.Add(eSkill.Bow_Blade, nudBowBlade.Value)
        Leader.Skills.Add(eSkill.Brawling, nudBrawling.Value)
        Leader.Skills.Add(eSkill.Computer, nudComputer.Value)
        Leader.Skills.Add(eSkill.Diplomacy, nudDiplomacy.Value)
        Leader.Skills.Add(eSkill.Driver, nudDriver.Value)
        Leader.Skills.Add(eSkill.Engineering, nudEngineering.Value)
        Leader.Skills.Add(eSkill.Gunnery_Aerospace, nudGunneryAerospace.Value)
        Leader.Skills.Add(eSkill.Gunnery_Artillery, nudGunneryArtillery.Value)
        Leader.Skills.Add(eSkill.Gunnery_Mech, nudGunneryMech.Value)
        Leader.Skills.Add(eSkill.Interrogation, nudInterrogation.Value)
        Leader.Skills.Add(eSkill.JumpShip_Piloting_Navigation, nudJumpshipPilotingNavigation.Value)
        Leader.Skills.Add(eSkill.Land_Management, nudLandManagement.Value)
        Leader.Skills.Add(eSkill.Leadership, nudLeadership.Value)
        Leader.Skills.Add(eSkill.Mechanical, nudMechanical.Value)
        Leader.Skills.Add(eSkill.Medical_First_Aid, nudMedicalFirstAid.Value)
        Leader.Skills.Add(eSkill.Piloting_Aerospace, nudPilotingAerospace.Value)
        Leader.Skills.Add(eSkill.Piloting_Mech, nudPilotingMech.Value)
        Leader.Skills.Add(eSkill.Pistol, nudPistol.Value)
        Leader.Skills.Add(eSkill.Rifle, nudRifle.Value)
        Leader.Skills.Add(eSkill.Rogue, nudRogue.Value)
        Leader.Skills.Add(eSkill.Streetwise, nudStreetwise.Value)
        Leader.Skills.Add(eSkill.Survival, nudSurvival.Value)
        Leader.Skills.Add(eSkill.Tactics, nudTactics.Value)
        Leader.Skills.Add(eSkill.Technician, nudTechnician.Value)

        If cbCareer.SelectedIndex = 0 Then 'MechWarrior
            Leader.PrimaryMech = GetMechAssignment(nudPrimaryMech.Value)
            Leader.BackupMech = GetMechAssignment(nudBackupMech.Value)
        ElseIf cbCareer.SelectedIndex = 1 Then 'Aerospace Pilot
            Leader.Fighter = GetFighterAssignment(nudFighter.Value)
        ElseIf cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3 Then 'Tech
        ElseIf cbCareer.SelectedIndex = 4 Or cbCareer.SelectedIndex = 5 Then 'Scout
            Leader.ScoutVehicle = cbScoutVehicle.SelectedItem.ToString
        End If

        Leader.XP_Total = nudXP.Value
        Leader.Title = cbTitle.SelectedValue

        Me.Visible = False
        frmCreate.Show()
    End Sub
    Private Sub Update_Character_Cost()
        CanonCharacter = False
        Dim TempCost As Int32 = 150
        Select Case nudBOD.Value
            Case 2 : TempCost += 60
            Case 3 : TempCost += 30
            Case 4 : TempCost += 15
            Case 5 : TempCost += 5
            Case 6 : TempCost += 0
            Case 7 : TempCost -= 10
            Case 8 : TempCost -= 20
            Case 9 : TempCost -= 40
            Case 10 : TempCost -= 80
            Case 11 : TempCost -= 150
            Case 12 : TempCost -= 300
        End Select
        Select Case nudDEX.Value
            Case 2 : TempCost += 70
            Case 3 : TempCost += 35
            Case 4 : TempCost += 20
            Case 5 : TempCost += 10
            Case 6 : TempCost += 0
            Case 7 : TempCost -= 15
            Case 8 : TempCost -= 30
            Case 9 : TempCost -= 50
            Case 10 : TempCost -= 100
            Case 11 : TempCost -= 200
            Case 12 : TempCost -= 400
        End Select
        Select Case nudLRN.Value
            Case 2 : TempCost += 80
            Case 3 : TempCost += 40
            Case 4 : TempCost += 25
            Case 5 : TempCost += 15
            Case 6 : TempCost += 0
            Case 7 : TempCost -= 20
            Case 8 : TempCost -= 40
            Case 9 : TempCost -= 60
            Case 10 : TempCost -= 125
            Case 11 : TempCost -= 250
            Case 12 : TempCost -= 500
        End Select
        Select Case nudCHA.Value
            Case 2 : TempCost += 50
            Case 3 : TempCost += 25
            Case 4 : TempCost += 15
            Case 5 : TempCost += 5
            Case 6 : TempCost += 0
            Case 7 : TempCost -= 10
            Case 8 : TempCost -= 20
            Case 9 : TempCost -= 40
            Case 10 : TempCost -= 80
            Case 11 : TempCost -= 150
            Case 12 : TempCost -= 300
        End Select
        If clbInborn.CheckedIndices.Contains(0) Then 'Thick Skin
            TempCost -= 20
        End If
        If clbInborn.CheckedIndices.Contains(1) Then 'Glass Jaw
            TempCost += 15
        End If
        If clbInborn.CheckedIndices.Contains(2) Then 'Peripheral Vision
            TempCost -= 15
        End If
        If clbInborn.CheckedIndices.Contains(3) Then 'Sixth Sense
            TempCost -= 20
        End If
        If clbInborn.CheckedIndices.Contains(4) Then 'Family Friend
            TempCost -= 25
        End If
        If clbInborn.CheckedIndices.Contains(5) Then 'Family Feud
            TempCost += 15
        End If
        If lstNaturalAptitude.SelectedIndex > 0 Then TempCost -= 10
        Select Case cbHandedness.SelectedIndex
            Case 0, 1 'Right or Left, fine
            Case 2, 3 : TempCost -= 5 'Natural Right/Left
            Case 4 : TempCost -= 20 'Either
            Case 5 : TempCost -= 40 'Both
        End Select


        If (cbCareer.SelectedIndex = 0 Or cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3) And nudLRN.Value >= 6 Then
            chkAcademy.Visible = True
        Else
            chkAcademy.Visible = False
            chkAcademy.Checked = False
        End If

        If (cbCareer.SelectedIndex = 0 Or cbCareer.SelectedIndex = 1 Or cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3) And nudLRN.Value >= 8 And (cbFaction.SelectedIndex = 0 Or cbFaction.SelectedIndex = 1) Then
            chkUniversity.Visible = True
        Else
            chkUniversity.Checked = False
            chkUniversity.Visible = False
        End If

        Dim SkillCount As Int32 = 0
        If nudAthletics.Value > 0 Then SkillCount += 1
        If nudBowBlade.Value > 0 Then SkillCount += 1
        If nudBrawling.Value > 0 Then SkillCount += 1
        If nudComputer.Value > 0 Then SkillCount += 1
        If nudDiplomacy.Value > 0 Then SkillCount += 1
        If nudDriver.Value > 0 Then SkillCount += 1
        If nudEngineering.Value > 0 Then SkillCount += 1
        If nudGunneryAerospace.Value > 0 Then SkillCount += 1
        If nudGunneryArtillery.Value > 0 Then SkillCount += 1
        If nudGunneryMech.Value > 0 Then SkillCount += 1
        If nudInterrogation.Value > 0 Then SkillCount += 1
        If nudJumpshipPilotingNavigation.Value > 0 Then SkillCount += 1
        If nudLandManagement.Value > 0 Then SkillCount += 1
        If nudLeadership.Value > 0 Then SkillCount += 1
        If nudMechanical.Value > 0 Then SkillCount += 1
        If nudMedicalFirstAid.Value > 0 Then SkillCount += 1
        If nudPilotingAerospace.Value > 0 Then SkillCount += 1
        If nudPilotingMech.Value > 0 Then SkillCount += 1
        If nudPistol.Value > 0 Then SkillCount += 1
        If nudRifle.Value > 0 Then SkillCount += 1
        If nudRogue.Value > 0 Then SkillCount += 1
        If nudStreetwise.Value > 0 Then SkillCount += 1
        If nudSurvival.Value > 0 Then SkillCount += 1
        If nudTactics.Value > 0 Then SkillCount += 1
        If nudTechnician.Value > 0 Then SkillCount += 1

        Dim SkillTotal As Int32 = 0
        SkillTotal += nudAthletics.Value
        SkillTotal += nudBowBlade.Value
        SkillTotal += nudBrawling.Value
        SkillTotal += nudComputer.Value
        SkillTotal += nudDiplomacy.Value
        SkillTotal += nudDriver.Value
        SkillTotal += nudEngineering.Value
        SkillTotal += nudGunneryAerospace.Value
        SkillTotal += nudGunneryArtillery.Value
        SkillTotal += nudGunneryMech.Value
        SkillTotal += nudInterrogation.Value
        SkillTotal += nudJumpshipPilotingNavigation.Value
        SkillTotal += nudLandManagement.Value
        SkillTotal += nudLeadership.Value
        SkillTotal += nudMechanical.Value
        SkillTotal += nudMedicalFirstAid.Value
        SkillTotal += nudPilotingAerospace.Value
        SkillTotal += nudPilotingMech.Value
        SkillTotal += nudPistol.Value
        SkillTotal += nudRifle.Value
        SkillTotal += nudRogue.Value
        SkillTotal += nudStreetwise.Value
        SkillTotal += nudSurvival.Value
        SkillTotal += nudTactics.Value
        SkillTotal += nudTechnician.Value

        TempCost -= nudPrimaryMech.Value * 20
        TempCost -= nudBackupMech.Value * 20

        Dim SkillCost As Int32 = 0
        SkillCost += GetSkillCost(nudAthletics)
        SkillCost += GetSkillCost(nudBowBlade)
        SkillCost += GetSkillCost(nudBrawling)
        SkillCost += GetSkillCost(nudComputer)
        SkillCost += GetSkillCost(nudDiplomacy)
        SkillCost += GetSkillCost(nudDriver)
        SkillCost += GetSkillCost(nudEngineering)
        SkillCost += GetSkillCost(nudGunneryAerospace)
        SkillCost += GetSkillCost(nudGunneryArtillery)
        SkillCost += GetSkillCost(nudGunneryMech)
        SkillCost += GetSkillCost(nudInterrogation)
        SkillCost += GetSkillCost(nudJumpshipPilotingNavigation)
        SkillCost += GetSkillCost(nudLandManagement)
        SkillCost += GetSkillCost(nudLeadership)
        SkillCost += GetSkillCost(nudMechanical)
        SkillCost += GetSkillCost(nudMedicalFirstAid)
        SkillCost += GetSkillCost(nudPilotingAerospace)
        SkillCost += GetSkillCost(nudPilotingMech)
        SkillCost += GetSkillCost(nudPistol)
        SkillCost += GetSkillCost(nudRifle)
        SkillCost += GetSkillCost(nudRogue)
        SkillCost += GetSkillCost(nudStreetwise)
        SkillCost += GetSkillCost(nudSurvival)
        SkillCost += GetSkillCost(nudTactics)
        SkillCost += GetSkillCost(nudTechnician)
        TempCost -= SkillCost


        labAthletics.Text = "Atletics" & IIf(nudAthletics.Value > 0, " (" & Skill_Roll(nudBOD.Value, nudDEX.Value) - nudAthletics.Value & "+)", "")
        labBowBlade.Text = "Bow/Blade" & IIf(nudBowBlade.Value > 0, " (" & Skill_Roll(nudDEX.Value) - nudBowBlade.Value & "+)", "")
        labBrawling.Text = "Brawling" & IIf(nudBrawling.Value > 0, " (" & Skill_Roll(nudBOD.Value) - nudBrawling.Value & "+)", "")
        labComputer.Text = "Computer" & IIf(nudComputer.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudComputer.Value & "+)", "")
        labDiplomacy.Text = "Diplomacy" & IIf(nudDiplomacy.Value > 0, " (" & Skill_Roll(nudCHA.Value) - nudDiplomacy.Value & "+)", "")
        labDriver.Text = "Driver" & IIf(nudDriver.Value > 0, " (" & Skill_Roll(nudDEX.Value) - nudDriver.Value & "+)", "")
        labEngineering.Text = "Engineering" & IIf(nudEngineering.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudEngineering.Value & "+)", "")
        labGunneryAerospace.Text = "Gunnery/Aerospace" & IIf(nudGunneryAerospace.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudGunneryAerospace.Value & "+)", "")
        labGunneryArtillery.Text = "Gunnery/Artillery" & IIf(nudGunneryArtillery.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudGunneryArtillery.Value & "+)", "")
        labGunneryMech.Text = "Gunnery/Mech" & IIf(nudGunneryMech.Value > 0, " (" & Skill_Roll(nudDEX.Value) - nudGunneryMech.Value & "+)", "")
        labInterrogation.Text = "Interrogation" & IIf(nudInterrogation.Value > 0, " (" & Skill_Roll(nudLRN.Value, nudCHA.Value) - nudInterrogation.Value & "+)", "")
        labJumpshipPilotingNavigation.Text = "Jumpship Piloting/Navigation" & IIf(nudJumpshipPilotingNavigation.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudJumpshipPilotingNavigation.Value & "+)", "")
        labLandManagement.Text = "Land Management" & IIf(nudLandManagement.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudLandManagement.Value & "+)", "")
        labLeadership.Text = "Leadership" & IIf(nudLeadership.Value > 0, " (" & Skill_Roll(nudCHA.Value) - nudLeadership.Value & "+)", "")
        labMechanical.Text = "Mechanical" & IIf(nudMechanical.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudMechanical.Value & "+)", "")
        labMedicalFirstAid.Text = "Medical/First Aid" & IIf(nudMedicalFirstAid.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudMedicalFirstAid.Value & "+)", "")
        labPilotingAerospace.Text = "Piloting/Aerospace" & IIf(nudPilotingAerospace.Value > 0, " (" & Skill_Roll(nudLRN.Value, nudDEX.Value) - nudPilotingAerospace.Value & "+)", "")
        labPilotingMech.Text = "Piloting/Mech" & IIf(nudPilotingMech.Value > 0, " (" & Skill_Roll(nudLRN.Value, nudDEX.Value) - nudPilotingMech.Value & "+)", "")
        labPistol.Text = "Pistol" & IIf(nudPistol.Value > 0, " (" & Skill_Roll(nudDEX.Value) - nudPistol.Value & "+)", "")
        labRifle.Text = "Rifle" & IIf(nudRifle.Value > 0, " (" & Skill_Roll(nudDEX.Value) - nudRifle.Value & "+)", "")
        labRogue.Text = "Rogue" & IIf(nudRogue.Value > 0, " (" & Skill_Roll(nudLRN.Value, nudDEX.Value) - nudRogue.Value & "+)", "")
        labStreetwise.Text = "Streetwise" & IIf(nudStreetwise.Value > 0, " (" & Skill_Roll(nudCHA.Value) - nudStreetwise.Value & "+)", "")
        labSurvival.Text = "Survival" & IIf(nudSurvival.Value > 0, " (" & Skill_Roll(nudBOD.Value) - nudSurvival.Value & "+)", "")
        labTactics.Text = "Tactics" & IIf(nudTactics.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudTactics.Value & "+)", "")
        labTechnician.Text = "Technician" & IIf(nudTechnician.Value > 0, " (" & Skill_Roll(nudLRN.Value) - nudTechnician.Value & "+)", "")

        If chkAcademy.Checked Then TempCost -= 75
        If chkUniversity.Checked Then TempCost -= 100

        If cbCareer.SelectedIndex = 4 Or cbCareer.SelectedIndex = 5 Then 'Scout
            TempCost -= nudInformants.Value * 5
            TempCost -= nudUsefulContacts.Value * 15
            TempCost -= nudProminentContacts.Value * 30
            Select Case cbScoutVehicle.SelectedIndex
                Case 0 : TempCost -= 0 'None
                Case 1 : TempCost -= 15 'Jeep
                Case 2 : TempCost -= 20 'Light Truck
                Case 3 : TempCost -= 20 'Light Van
                Case 4 : TempCost -= 25 'Skimmer
            End Select
        ElseIf cbCareer.SelectedIndex = 0 Then 'MechWarrior
            If nudPrimaryMech.Value < 0 Then
                TempCost -= nudPrimaryMech.Value * 15
            ElseIf nudPrimaryMech.Value > 0 Then
                TempCost -= nudPrimaryMech.Value * 20
            End If
            If chkBackupMech.Checked Then
                TempCost -= 50
                TempCost -= nudBackupMech.Value * 20
            End If
        ElseIf cbCareer.SelectedIndex = 1 Then 'Aerospace Pilot
            If nudFighter.Value < 0 Then
                TempCost -= nudFighter.Value * 15
            ElseIf nudFighter.Value > 0 Then
                TempCost -= nudFighter.Value * 20
            End If
        End If



        If TempCost < 0 Then
            labCP.Text = "Maximum 150 character points, currently " & TempCost & " CP remaining"
            CharacterValid = False
        ElseIf SkillTotal > nudLRN.Value * 4 Then
            labCP.Text = "Maximum skill points is equal to 4 times the LRN value"
            CharacterValid = False
        ElseIf SkillCount > nudLRN.Value Then
            labCP.Text = "Maximum number of skills is equal to the LRN value"
            CharacterValid = False
        Else
            labCP.Text = TempCost & " CP Remaining"
            CharacterValid = True
        End If

        Update_Character_XP()
    End Sub
    Private Function GetMechAssignment(DRM As Int32) As Int32
        Return Choose(D(2, 6) + DRM + 5, 15, 20, 15, 25, 30, 20, 20, 25, 35, 40, 45, 50, 55, 60, 65, 70, 55, 75, 85, 60, 80, 65, 75, 90)
    End Function
    Private Function GetFighterAssignment(DRM As Int32) As Int32
        Return Choose(D(2, 6) + DRM + 5, 20, 20, 20, 30, 20, 40, 30, 40, 20, 50, 30, 50, 60, 40, 50, 70, 60, 80, 70, 80, 90, 100, 80, 90, 100)
    End Function
    Private Function GetSkillCost(nud As NumericUpDown) As Int32
        If nud.Enabled = False Then
            Return 0 'Must be part of a package
        ElseIf nud.Minimum > 0 Then
            If nud.Minimum = 1 Then
                Return Choose(nud.Value + 1, 0, 0, 10, 30, 60, 110, 190, 310, 470)
            ElseIf nud.Minimum = 2 Then
                Return Choose(nud.Value + 1, 0, 0, 0, 20, 50, 100, 180, 300, 460)
            Else
                Debug.Print("This should never happen")
                Return Choose(nud.Value + 1, 0, 20, 30, 50, 80, 130, 210, 330, 490)
            End If
        Else
            Return Choose(nud.Value + 1, 0, 20, 30, 50, 80, 130, 210, 330, 490)
        End If
    End Function
    Private Sub cbCareer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCareer.SelectedIndexChanged
        nudUsefulContacts.Visible = False
        nudProminentContacts.Visible = False
        nudInformants.Visible = False
        labUsefulContacts.Visible = False
        labProminentContacts.Visible = False
        labInformants.Visible = False

        labScoutVehicle.Visible = False
        cbScoutVehicle.Visible = False

        nudFighter.Visible = False
        labFighter.Visible = False

        nudPrimaryMech.Visible = False
        nudBackupMech.Visible = False
        chkBackupMech.Visible = False
        labPrimaryMech.Visible = False
        labBackupMech.Visible = False

        If cbCareer.SelectedIndex = 0 Then 'MechWarrior
            labCareer.Text = "Min Dexterity 5.  Min Learn 5."
            nudPrimaryMech.Visible = True
            chkBackupMech.Visible = True
            labPrimaryMech.Visible = True
        ElseIf cbCareer.SelectedIndex = 1 Then 'Aerospace Pilot
            labCareer.Text = "Min Dexterity 6.  Min Learn 6."
            nudFighter.Visible = True
            labFighter.Visible = True
        ElseIf cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3 Then 'Technician
            labCareer.Text = "Min Dexterity 4.  Min Learn 7."
        ElseIf cbCareer.SelectedIndex = 4 Or cbCareer.SelectedIndex = 5 Then 'Scout
            labCareer.Text = "Min Learn 5.  Min Charisma 6."
            nudUsefulContacts.Visible = True
            nudProminentContacts.Visible = True
            nudInformants.Visible = True
            labUsefulContacts.Visible = True
            labProminentContacts.Visible = True
            labInformants.Visible = True
            labScoutVehicle.Visible = True
            cbScoutVehicle.Visible = True
        End If
        Update_Character_Cost()
    End Sub
    Private Function Saving_Roll(Attribute1 As Int32, Optional Attribute2 As Int32 = 0) As Int32
        Dim TempA As Int32 = Attribute1
        If Attribute2 >= 2 Then TempA = Math.Floor((Attribute1 + Attribute2) / 2)
        Return Choose(TempA - 1, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2)
    End Function
    Private Function Skill_Roll(Attribute1 As Int32, Optional Attribute2 As Int32 = 0) As Int32
        Dim TempA As Int32 = Attribute1
        If Attribute2 >= 2 Then TempA = Math.Floor((Attribute1 + Attribute2) / 2)
        Return Choose(TempA - 1, 12, 11, 10, 10, 9, 8, 8, 8, 7, 7, 6)
    End Function
    Private Sub frmCharacter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbCareer.SelectedIndex = 0
        cbFaction.SelectedIndex = 0
        cbHandedness.SelectedIndex = 0
        lstNaturalAptitude.SelectedIndex = 0
        cbScoutVehicle.SelectedIndex = 0
        cbTitle.SelectedIndex = 0
        cbScoutVehicle.SelectedIndex = 0
        cbName.Text = Environment.UserName
    End Sub

    Private Sub ValidatePackages()
        nudComputer.Enabled = True
        nudDiplomacy.Minimum = 0
        nudDriver.Enabled = True
        nudEngineering.Enabled = True
        nudEngineering.Minimum = 0
        nudGunneryAerospace.Enabled = True
        nudGunneryAerospace.Minimum = 0
        nudGunneryArtillery.Enabled = True
        nudGunneryMech.Enabled = True
        nudGunneryMech.Minimum = 0
        nudJumpshipPilotingNavigation.Enabled = True
        nudLeadership.Enabled = True
        nudMechanical.Enabled = True
        nudMechanical.Minimum = 0
        nudPilotingAerospace.Enabled = True
        nudPilotingAerospace.Minimum = 0
        nudPilotingMech.Enabled = True
        nudPilotingMech.Minimum = 0
        nudPistol.Enabled = True
        nudRifle.Enabled = True
        nudRogue.Minimum = 0
        nudStreetwise.Value = 0
        nudSurvival.Enabled = True
        nudTactics.Enabled = True
        nudTechnician.Enabled = True
        nudTechnician.Minimum = 0




        Select Case cbCareer.SelectedIndex
            Case 0 'MechWarrior
                If chkAcademy.Checked Then '75 CP, LRN 6+
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudTechnician.Value = 2
                    nudTechnician.Enabled = False
                    nudPistol.Value = 1
                    nudPistol.Enabled = False
                    nudLeadership.Value = 1
                    nudLeadership.Enabled = False
                    nudSurvival.Value = 1
                    nudSurvival.Enabled = False
                ElseIf chkUniversity.Checked Then '100 CP, LRN 8+, AND is Davion/Steiner, needing an 11+ with -1 per attribute score (any type) over 9
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudPilotingAerospace.Value = 2
                    nudPilotingAerospace.Enabled = False
                    nudDriver.Value = 1
                    nudDriver.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudGunneryArtillery.Value = 2
                    nudGunneryArtillery.Enabled = False
                    nudRifle.Value = 1
                    nudRifle.Enabled = False
                    nudLeadership.Value = 2
                    nudLeadership.Enabled = False
                    nudTactics.Value = 2
                    nudTactics.Enabled = False
                Else
                    nudPilotingMech.Minimum = 1
                    nudPilotingMech.Value = 1
                    nudGunneryMech.Minimum = 1
                    nudGunneryMech.Value = 1
                End If
            Case 1 'Aerospace Pilot
                If chkUniversity.Checked Then '100 CP, LRN 8+, AND is Davion/Steiner, needing an 11+ with -1 per attribute score (any type) over 9
                    nudPilotingAerospace.Value = 2
                    nudPilotingAerospace.Enabled = False
                    nudGunneryAerospace.Value = 2
                    nudGunneryAerospace.Enabled = False
                    nudEngineering.Value = 3
                    nudEngineering.Enabled = False
                    nudTechnician.Value = 2
                    nudTechnician.Enabled = False
                    nudComputer.Value = 2
                    nudComputer.Enabled = False
                    nudMechanical.Value = 1
                    nudMechanical.Enabled = False
                    nudTactics.Value = 1
                    nudTactics.Enabled = False
                    nudJumpshipPilotingNavigation.Value = 1
                    nudJumpshipPilotingNavigation.Enabled = False
                Else
                    nudPilotingAerospace.Value = 1
                    nudPilotingAerospace.Minimum = 1
                    nudGunneryAerospace.Value = 1
                    nudGunneryAerospace.Minimum = 1
                End If
            Case 2 'Technician/Mechanical
                If chkAcademy.Checked Then '75 CP, LRN 6+
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudTechnician.Value = 2
                    nudTechnician.Enabled = False
                    nudPistol.Value = 1
                    nudPistol.Enabled = False
                    nudLeadership.Value = 1
                    nudLeadership.Enabled = False
                    nudSurvival.Value = 1
                    nudSurvival.Enabled = False
                ElseIf chkUniversity.Checked Then '100 CP, LRN 8+, AND is Davion/Steiner, needing an 11+ with -1 per attribute score (any type) over 9
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudPilotingAerospace.Value = 2
                    nudPilotingAerospace.Enabled = False
                    nudDriver.Value = 1
                    nudDriver.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudGunneryArtillery.Value = 2
                    nudGunneryArtillery.Enabled = False
                    nudRifle.Value = 1
                    nudRifle.Enabled = False
                    nudLeadership.Value = 2
                    nudLeadership.Enabled = False
                    nudTactics.Value = 2
                    nudTactics.Enabled = False
                Else
                    nudTechnician.Value = 1
                    nudMechanical.Value = 1
                    nudTechnician.Minimum = 1
                    nudMechanical.Minimum = 1
                End If
            Case 3 'Technician/Engineering
                If chkAcademy.Checked Then '75 CP, LRN 6+
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudTechnician.Value = 2
                    nudTechnician.Enabled = False
                    nudPistol.Value = 1
                    nudPistol.Enabled = False
                    nudLeadership.Value = 1
                    nudLeadership.Enabled = False
                    nudSurvival.Value = 1
                    nudSurvival.Enabled = False
                ElseIf chkUniversity.Checked Then '100 CP, LRN 8+, AND is Davion/Steiner, needing an 11+ with -1 per attribute score (any type) over 9
                    nudPilotingMech.Value = 2
                    nudPilotingMech.Enabled = False
                    nudPilotingAerospace.Value = 2
                    nudPilotingAerospace.Enabled = False
                    nudDriver.Value = 1
                    nudDriver.Enabled = False
                    nudGunneryMech.Value = 2
                    nudGunneryMech.Enabled = False
                    nudGunneryArtillery.Value = 2
                    nudGunneryArtillery.Enabled = False
                    nudRifle.Value = 1
                    nudRifle.Enabled = False
                    nudLeadership.Value = 2
                    nudLeadership.Enabled = False
                    nudTactics.Value = 2
                    nudTactics.Enabled = False
                Else
                    nudTechnician.Value = 1
                    nudEngineering.Value = 1
                    nudTechnician.Minimum = 1
                    nudEngineering.Minimum = 1
                End If
            Case 4 'Scout/Streetwise
                nudRogue.Value = 1
                nudRogue.Minimum = 1
                nudStreetwise.Value = 1
                nudStreetwise.Minimum = 1
            Case 5 'Scout/Diplomacy
                nudRogue.Value = 1
                nudRogue.Minimum = 1
                nudDiplomacy.Value = 1
                nudDiplomacy.Minimum = 1
        End Select
    End Sub
    Private Sub chkAcademy_CheckedChanged(sender As Object, e As EventArgs) Handles chkAcademy.CheckedChanged
        ValidatePackages()
        Update_Character_Cost()
    End Sub
    Private Sub chkUniversity_CheckedChanged(sender As Object, e As EventArgs) Handles chkUniversity.CheckedChanged
        If chkUniversity.Checked Then
            Dim Target As Int32 = 11
            If nudBOD.Value > 9 Then Target -= nudBOD.Value - 9
            If nudDEX.Value > 9 Then Target -= nudDEX.Value - 9
            If nudCHA.Value > 9 Then Target -= nudCHA.Value - 9
            If nudLRN.Value > 9 Then Target -= nudLRN.Value - 9
            If D(2, 6) < Target Then
                MsgBox("You failed to qualify for a university!")
                chkUniversity.Checked = False
            End If
        Else

        End If
        ValidatePackages()
        Update_Character_Cost()
    End Sub

    Private Sub ResetForm()
        For I As Int32 = 0 To clbInborn.Items.Count - 1
            clbInborn.SetItemChecked(I, False)
        Next

        nudAthletics.Minimum = 0
        nudBowBlade.Minimum = 0
        nudBrawling.Minimum = 0
        nudComputer.Minimum = 0
        nudDiplomacy.Minimum = 0
        nudDriver.Minimum = 0
        nudEngineering.Minimum = 0
        nudGunneryAerospace.Minimum = 0
        nudGunneryArtillery.Minimum = 0
        nudGunneryMech.Minimum = 0
        nudInterrogation.Minimum = 0
        nudJumpshipPilotingNavigation.Minimum = 0
        nudLandManagement.Minimum = 0
        nudLeadership.Minimum = 0
        nudMechanical.Minimum = 0
        nudMedicalFirstAid.Minimum = 0
        nudPilotingAerospace.Minimum = 0
        nudPilotingMech.Minimum = 0
        nudPistol.Minimum = 0
        nudRifle.Minimum = 0
        nudRogue.Minimum = 0
        nudStreetwise.Minimum = 0
        nudSurvival.Minimum = 0
        nudTactics.Minimum = 0
        nudTechnician.Minimum = 0

        nudAthletics.Enabled = True
        nudBowBlade.Enabled = True
        nudBrawling.Enabled = True
        nudComputer.Enabled = True
        nudDiplomacy.Enabled = True
        nudDriver.Enabled = True
        nudEngineering.Enabled = True
        nudGunneryAerospace.Enabled = True
        nudGunneryArtillery.Enabled = True
        nudGunneryMech.Enabled = True
        nudInterrogation.Enabled = True
        nudJumpshipPilotingNavigation.Enabled = True
        nudLandManagement.Enabled = True
        nudLeadership.Enabled = True
        nudMechanical.Enabled = True
        nudMedicalFirstAid.Enabled = True
        nudPilotingAerospace.Enabled = True
        nudPilotingMech.Enabled = True
        nudPistol.Enabled = True
        nudRifle.Enabled = True
        nudRogue.Enabled = True
        nudStreetwise.Enabled = True
        nudSurvival.Enabled = True
        nudTactics.Enabled = True
        nudTechnician.Enabled = True
        nudUsefulContacts.Enabled = True
        nudProminentContacts.Enabled = True
        nudInformants.Enabled = True
        nudUsefulContacts2.Enabled = True
        nudProminentContacts2.Enabled = True
        nudInformants2.Enabled = True


        nudAthletics.Value = 0
        nudBowBlade.Value = 0
        nudBrawling.Value = 0
        nudComputer.Value = 0
        nudDiplomacy.Value = 0
        nudDriver.Value = 0
        nudEngineering.Value = 0
        nudGunneryAerospace.Value = 0
        nudGunneryArtillery.Value = 0
        nudGunneryMech.Value = 0
        nudInterrogation.Value = 0
        nudJumpshipPilotingNavigation.Value = 0
        nudLandManagement.Value = 0
        nudLeadership.Value = 0
        nudMechanical.Value = 0
        nudMedicalFirstAid.Value = 0
        nudPilotingAerospace.Value = 0
        nudPilotingMech.Value = 0
        nudPistol.Value = 0
        nudRifle.Value = 0
        nudRogue.Value = 0
        nudStreetwise.Value = 0
        nudSurvival.Value = 0
        nudTactics.Value = 0
        nudTechnician.Value = 0

        cbName.Text = ""
        cbFaction.SelectedValue = ""
        nudBOD.Minimum = 2
        nudCHA.Minimum = 2
        nudDEX.Minimum = 2
        nudLRN.Minimum = 2

        nudBOD.Enabled = True
        nudCHA.Enabled = True
        nudDEX.Enabled = True
        nudLRN.Enabled = True

        nudBOD.Value = 6
        nudCHA.Value = 6
        nudDEX.Value = 6
        nudLRN.Value = 6
        txtRank.Text = ""
        txtRank.Enabled = True

        nudUsefulContacts.Value = 0
        nudProminentContacts.Value = 0
        nudInformants.Value = 0
        nudUsefulContacts2.Value = 0
        nudProminentContacts2.Value = 0
        nudInformants2.Value = 0
        cbHandedness.SelectedIndex = 0
        nudPrimaryMech.Value = 0
        nudBackupMech.Value = 0
        nudFighter.Value = 0
        CanonCharacter = False
        lstNaturalAptitude.SelectedIndex = 0
        chkAcademy.Checked = False
        chkUniversity.Checked = False
        chkBackupMech.Checked = False
        labUsefulContacts2.Visible = False
        labProminentContacts2.Visible = False
        labInformants2.Visible = False

        cbTitle.Enabled = True
        nudXP.Enabled = True

        'Leader.Fighter = 0
        'Leader.ScoutVehicle = ""
        'Leader.Titles = ""
        'Leader.PrimaryMech = 0
        'Leader.BackupMech = 0
        'Leader.XP_Total = 0
        'Leader.XP_Availabile = 0

        cbTitle.SelectedIndex = 0
        cmdFinishedGeneration.Visible = True

        cmdFinishedGeneration.Visible = True
        labXP.Visible = False
        nudXP.Visible = False
        labCP.Visible = True
        labCPXP.Visible = False
        cbTitle.Visible = False
        labTitle.Visible = False
        nudInformants2.Visible = False
        nudProminentContacts2.Visible = False
        nudUsefulContacts2.Visible = False
        chkAcademy.Enabled = True
        chkUniversity.Enabled = True
        cbFaction.Enabled = True

        cbCareer.Enabled = True
        cbHandedness.Enabled = True
        nudPrimaryMech.Enabled = True
        chkBackupMech.Enabled = True

        clbInborn.Enabled = True
        lstNaturalAptitude.Enabled = True
    End Sub

    Private Sub cmdFinishedGeneration_Click(sender As Object, e As EventArgs) Handles cmdFinishedGeneration.Click
        cmdFinishedGeneration.Visible = False
        labXP.Visible = True
        nudXP.Visible = True
        labCP.Visible = False
        labCPXP.Visible = True

        nudBackupMech.Enabled = False
        nudFighter.Enabled = False
        nudPrimaryMech.Enabled = False
        cbScoutVehicle.Enabled = False
        chkAcademy.Enabled = False
        chkUniversity.Enabled = False
        clbInborn.Enabled = False
        lstNaturalAptitude.Enabled = False
        cbCareer.Enabled = False
        cbHandedness.Enabled = False

        cbTitle.Visible = True
        labTitle.Visible = True
        nudInformants2.Visible = True
        labInformants2.Visible = True
        nudProminentContacts2.Visible = True
        labProminentContacts2.Visible = True
        nudUsefulContacts2.Visible = True
        labUsefulContacts2.Visible = True


        nudBOD.Minimum = nudBOD.Value
        nudCHA.Minimum = nudCHA.Value
        nudDEX.Minimum = nudDEX.Value
        nudLRN.Minimum = nudLRN.Value

        nudBOD.Maximum = Math.Min(12, nudBOD.Value + 6)
        nudCHA.Maximum = Math.Min(12, nudCHA.Value + 6)
        nudDEX.Maximum = Math.Min(12, nudDEX.Value + 6)
        nudLRN.Maximum = Math.Min(12, nudLRN.Value + 6)

        nudAthletics.Minimum = nudAthletics.Value
        nudBowBlade.Minimum = nudBowBlade.Value
        nudBrawling.Minimum = nudBrawling.Value
        nudComputer.Minimum = nudComputer.Value
        nudDiplomacy.Minimum = nudDiplomacy.Value
        nudDriver.Minimum = nudDriver.Value
        nudEngineering.Minimum = nudEngineering.Value
        nudGunneryAerospace.Minimum = nudGunneryAerospace.Value
        nudGunneryArtillery.Minimum = nudGunneryArtillery.Value
        nudGunneryMech.Minimum = nudGunneryMech.Value
        nudInterrogation.Minimum = nudInterrogation.Value
        nudJumpshipPilotingNavigation.Minimum = nudJumpshipPilotingNavigation.Value
        nudLandManagement.Minimum = nudLandManagement.Value
        nudLeadership.Minimum = nudLeadership.Value
        nudMechanical.Minimum = nudMechanical.Value
        nudMedicalFirstAid.Minimum = nudMedicalFirstAid.Value
        nudPilotingAerospace.Minimum = nudPilotingAerospace.Value
        nudPilotingMech.Minimum = nudPilotingMech.Value
        nudPistol.Minimum = nudPistol.Value
        nudRifle.Minimum = nudRifle.Value
        nudRogue.Minimum = nudRogue.Value
        nudStreetwise.Minimum = nudStreetwise.Value
        nudSurvival.Minimum = nudSurvival.Value
        nudTactics.Minimum = nudTactics.Value
        nudTechnician.Minimum = nudTechnician.Value


        nudAthletics.Enabled = True
        nudBowBlade.Enabled = True
        nudBrawling.Enabled = True
        nudComputer.Enabled = True
        nudDiplomacy.Enabled = True
        nudDriver.Enabled = True
        nudEngineering.Enabled = True
        nudGunneryAerospace.Enabled = True
        nudGunneryArtillery.Enabled = True
        nudGunneryMech.Enabled = True
        nudInterrogation.Enabled = True
        nudJumpshipPilotingNavigation.Enabled = True
        nudLandManagement.Enabled = True
        nudLeadership.Enabled = True
        nudMechanical.Enabled = True
        nudMedicalFirstAid.Enabled = True
        nudPilotingAerospace.Enabled = True
        nudPilotingMech.Enabled = True
        nudPistol.Enabled = True
        nudRifle.Enabled = True
        nudRogue.Enabled = True
        nudStreetwise.Enabled = True
        nudSurvival.Enabled = True
        nudTactics.Enabled = True
        nudTechnician.Enabled = True

        nudXP.Value = Leader.XP_Total

        labCP.Visible = True
        If cbCareer.SelectedIndex = 0 Then 'MechWarriors
            labCP.Text = "Mech Weight:  " & Leader.PrimaryMech & " tons"
            If Leader.BackupMech > 0 Then
                labCP.Text &= ", Backup Mech Weight: " & Leader.BackupMech
            End If
        ElseIf cbCareer.SelectedIndex = 1 Then 'Pilot
            labCP.Text = "Fighter Weight:  " & Leader.Fighter
        ElseIf cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3 Then 'Tech
            labCP.Visible = False
            labCP.Text = ""
        ElseIf cbCareer.SelectedIndex = 4 Or cbCareer.SelectedIndex = 5 Then 'Scout
            labCP.Text = "Vehicle:  " & Leader.ScoutVehicle
        End If

    End Sub



#Region "Updates Only"
    Private Sub nudBOD_ValueChanged(sender As Object, e As EventArgs) Handles nudBOD.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudDEX_ValueChanged(sender As Object, e As EventArgs) Handles nudDEX.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudLRN_ValueChanged(sender As Object, e As EventArgs) Handles nudLRN.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudCHA_ValueChanged(sender As Object, e As EventArgs) Handles nudCHA.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudAthletics_ValueChanged(sender As Object, e As EventArgs) Handles nudAthletics.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudBowBlade_ValueChanged(sender As Object, e As EventArgs) Handles nudBowBlade.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudBrawling_ValueChanged(sender As Object, e As EventArgs) Handles nudBrawling.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudComputer_ValueChanged(sender As Object, e As EventArgs) Handles nudComputer.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudDiplomacy_ValueChanged(sender As Object, e As EventArgs) Handles nudDiplomacy.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudDriver_ValueChanged(sender As Object, e As EventArgs) Handles nudDriver.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudEngineering_ValueChanged(sender As Object, e As EventArgs) Handles nudEngineering.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudGunneryAerospace_ValueChanged(sender As Object, e As EventArgs) Handles nudGunneryAerospace.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudGunneryArtillery_ValueChanged(sender As Object, e As EventArgs) Handles nudGunneryArtillery.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudGunneryMech_ValueChanged(sender As Object, e As EventArgs) Handles nudGunneryMech.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudInterrogation_ValueChanged(sender As Object, e As EventArgs) Handles nudInterrogation.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudJumpshipPilotingNavigation_ValueChanged(sender As Object, e As EventArgs) Handles nudJumpshipPilotingNavigation.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudLandManagement_ValueChanged(sender As Object, e As EventArgs) Handles nudLandManagement.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudLeadership_ValueChanged(sender As Object, e As EventArgs) Handles nudLeadership.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudMechanical_ValueChanged(sender As Object, e As EventArgs) Handles nudMechanical.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudMedicalFirstAid_ValueChanged(sender As Object, e As EventArgs) Handles nudMedicalFirstAid.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudPilotingAerospace_ValueChanged(sender As Object, e As EventArgs) Handles nudPilotingAerospace.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudPilotingMech_ValueChanged(sender As Object, e As EventArgs) Handles nudPilotingMech.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudPistol_ValueChanged(sender As Object, e As EventArgs) Handles nudPistol.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudRifle_ValueChanged(sender As Object, e As EventArgs) Handles nudRifle.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudRogue_ValueChanged(sender As Object, e As EventArgs) Handles nudRogue.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudStreetwise_ValueChanged(sender As Object, e As EventArgs) Handles nudStreetwise.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudSurvival_ValueChanged(sender As Object, e As EventArgs) Handles nudSurvival.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudTactics_ValueChanged(sender As Object, e As EventArgs) Handles nudTactics.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudTechnician_ValueChanged(sender As Object, e As EventArgs) Handles nudTechnician.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub cbFaction_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbFaction.SelectedIndexChanged
        Update_Character_Cost()
    End Sub

    Private Sub lstHandedness_SelectedIndexChanged(sender As Object, e As EventArgs)
        Update_Character_Cost()
    End Sub

    Private Sub clbInborn_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbInborn.ItemCheck
        Update_Character_Cost()
    End Sub

    Private Sub lstNaturalAptitude_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstNaturalAptitude.SelectedIndexChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudPrimaryMech_ValueChanged(sender As Object, e As EventArgs) Handles nudPrimaryMech.ValueChanged
        Update_Character_Cost()
    End Sub

    Private Sub nudBackupMech_ValueChanged(sender As Object, e As EventArgs) Handles nudBackupMech.ValueChanged
        Update_Character_Cost()
    End Sub



    Private Sub cbScoutVehicle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbScoutVehicle.SelectedIndexChanged
        Update_Character_Cost()
    End Sub


    Private Sub nudFighter_ValueChanged(sender As Object, e As EventArgs) Handles nudFighter.ValueChanged
        Update_Character_Cost()
    End Sub
    Private Sub nudInformants2_ValueChanged(sender As Object, e As EventArgs) Handles nudInformants2.ValueChanged
        Update_Character_XP()
    End Sub

    Private Sub nudUsefulContacts2_ValueChanged(sender As Object, e As EventArgs) Handles nudUsefulContacts2.ValueChanged
        Update_Character_XP()
    End Sub

    Private Sub nudProminentContacts2_ValueChanged(sender As Object, e As EventArgs) Handles nudProminentContacts2.ValueChanged
        Update_Character_XP()
    End Sub

    Private Sub cbTitle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTitle.SelectedIndexChanged
        Update_Character_XP()
    End Sub

#End Region

    Private Sub cbName_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbName.SelectedValueChanged
        ResetForm()
        Select Case cbName.SelectedItem
            Case "Daniel Sorenson" 'Sorenson's Sabres Scenario Pack, Page 17
                txtRank.Text = "Captain"
                cbFaction.SelectedValue = "Kurita"
                nudBOD.Value = 7
                nudCHA.Value = 7
                nudDEX.Value = 9
                nudLRN.Value = 10
                cbHandedness.SelectedIndex = 0
                nudBowBlade.Value = 2
                nudGunneryMech.Value = 5
                nudLeadership.Value = 6
                nudPilotingMech.Value = 4
                nudPistol.Value = 3
                nudRogue.Value = 3
                nudStreetwise.Value = 2
                nudSurvival.Value = 4
                nudTactics.Value = 6
                nudTechnician.Value = 5
                Leader.PrimaryMech = 75 'Marauder
                Leader.XP_Total = 158000
                Leader.XP_Availabile = 10160
                nudProminentContacts2.Value = 1
                chkAcademy.Checked = True
                CanonCharacter = True
                chkBackupMech.Enabled = False
                cbTitle.SelectedIndex = 3
            Case "Cranston Snord" 'Cranston' Snord's Irregulars Scenario Pack, Page 9
                txtRank.Text = "Captain"
                cbFaction.SelectedValue = "Steiner"
                nudBOD.Value = 8
                nudCHA.Value = 8
                nudDEX.Value = 8
                nudLRN.Value = 11
                cbHandedness.SelectedIndex = 2
                nudDiplomacy.Value = 4
                nudGunneryMech.Value = 6
                nudLeadership.Value = 5
                nudPilotingMech.Value = 7
                nudPistol.Value = 5
                nudStreetwise.Value = 5
                nudSurvival.Value = 4
                nudTactics.Value = 4
                nudTechnician.Value = 3
                clbInborn.SetItemChecked(3, True) 'Sixth Sense
                Leader.PrimaryMech = 70 'ARC-2R Archer
                Leader.XP_Total = 155000
                Leader.XP_Availabile = 20000
                nudProminentContacts2.Value = 2
                nudUsefulContacts2.Value = 2
                nudInformants2.Value = 15
                CanonCharacter = True
                chkBackupMech.Enabled = False
            Case "Grayson Death Carlyle"
                txtRank.Text = "Colonel"
                cbFaction.SelectedValue = "Steiner"
                nudBOD.Value = 9
                nudCHA.Value = 10
                nudDEX.Value = 10
                nudLRN.Value = 12
                nudAthletics.Value = 2
                nudComputer.Value = 5
                nudDiplomacy.Value = 3
                nudDriver.Value = 3
                nudGunneryMech.Value = 4
                nudLeadership.Value = 5
                nudPilotingMech.Value = 5
                nudPistol.Value = 4
                nudRifle.Value = 5
                nudSurvival.Value = 2
                nudTactics.Value = 6
                nudTechnician.Value = 4
                Leader.PrimaryMech = 75 'MAD-3R Marauder
                Leader.XP_Total = 135000
                Leader.XP_Availabile = 20000
                CanonCharacter = True
                chkBackupMech.Enabled = False
            Case "Gerald Cameron-Jones" 'Rolling Thunder Scenario Pack, page 14
                txtRank.Text = "Force Captain"
                cbFaction.SelectedValue = "Marik"
                nudBOD.Value = 8
                nudCHA.Value = 9
                nudDEX.Value = 10
                nudLRN.Value = 10
                cbHandedness.SelectedIndex = 0
                nudDiplomacy.Value = 4
                nudGunneryMech.Value = 5
                nudLeadership.Value = 4
                nudPilotingMech.Value = 5
                nudPistol.Value = 4
                nudRogue.Value = 4
                nudStreetwise.Value = 2
                nudSurvival.Value = 4
                nudTactics.Value = 4
                nudTechnician.Value = 4
                Leader.PrimaryMech = 55 'SHD-2H Shadow Hawk
                cbTitle.SelectedIndex = 3
                Leader.XP_Total = 126000
                Leader.XP_Availabile = 9010
                chkAcademy.Checked = True
                nudProminentContacts2.Value = 2
                nudUsefulContacts2.Value = 2
                CanonCharacter = True
            Case "Patrick M. Kell" 'Kell Hounds Scenario Pack, page 28
                txtRank.Text = "Lieutenant Colonel"
                cbFaction.SelectedValue = "Steiner"
                nudBOD.Value = 9
                nudCHA.Value = 8
                nudDEX.Value = 10
                nudLRN.Value = 11
                cbHandedness.SelectedIndex = 0
                nudBrawling.Value = 5
                nudDiplomacy.Value = 2
                nudGunneryMech.Value = 7
                nudLeadership.Value = 5
                nudPilotingMech.Value = 8
                nudPistol.Value = 4
                nudRogue.Value = 4
                nudSurvival.Value = 2
                nudTactics.Value = 6
                nudTechnician.Value = 3
                Leader.PrimaryMech = 65 'TDR-5S Thunderbolt
                'cbTitle.SelectedIndex = 3
                Leader.XP_Total = 185000
                Leader.XP_Availabile = 35000
                'chkAcademy.Checked = True
                nudProminentContacts2.Value = 1
                nudUsefulContacts2.Value = 2
                CanonCharacter = True

            Case "Marcus Barton" 'McCarron's AC p14 - NOT CURRENTLY SUPPORTED
                txtRank.Text = "Colonel"
                cbFaction.SelectedValue = "Liao"
                nudBOD.Value = 10
                nudCHA.Value = 10
                nudDEX.Value = 11
                nudLRN.Value = 8
                cbHandedness.SelectedIndex = 1
                nudBowBlade.Value = 3
                nudDiplomacy.Value = 3
                'Gambling 2
                nudPilotingMech.Value = 6
                nudGunneryMech.Value = 7
                nudLeadership.Value = 4
                nudMedicalFirstAid.Value = 2
                nudPistol.Value = 3
                nudRifle.Value = 3
                'Strategy 3
                nudStreetwise.Value = 4
                'Support Weapons 2
                nudSurvival.Value = 2
                nudTactics.Value = 3
                nudTechnician.Value = 4
                nudBrawling.Value = 3

                Leader.PrimaryMech = 80 'AWS-8Q Awesome, PIloting 1, Gunnery 0
                'cbTitle.SelectedIndex = 3
                'Leader.XP_Total = 185000
                'Leader.XP_Availabile = 35000
                'chkAcademy.Checked = True
                'nudProminentContacts2.Value = 1
                'nudUsefulContacts2.Value = 2
                CanonCharacter = True

        End Select
        If CanonCharacter Then
            cmdFinishedGeneration_Click(Nothing, Nothing)
            nudUsefulContacts.Enabled = False
            nudProminentContacts.Enabled = False
            nudInformants.Enabled = False
            nudUsefulContacts2.Enabled = False
            nudProminentContacts2.Enabled = False
            nudInformants2.Enabled = False
            nudAthletics.Enabled = False
            nudBowBlade.Enabled = False
            nudBrawling.Enabled = False
            nudComputer.Enabled = False
            nudDiplomacy.Enabled = False
            nudDriver.Enabled = False
            nudEngineering.Enabled = False
            nudGunneryAerospace.Enabled = False
            nudGunneryArtillery.Enabled = False
            nudGunneryMech.Enabled = False
            nudInterrogation.Enabled = False
            nudJumpshipPilotingNavigation.Enabled = False
            nudLandManagement.Enabled = False
            nudLeadership.Enabled = False
            nudMechanical.Enabled = False
            nudMedicalFirstAid.Enabled = False
            nudPilotingAerospace.Enabled = False
            nudPilotingMech.Enabled = False
            nudPistol.Enabled = False
            nudRifle.Enabled = False
            nudRogue.Enabled = False
            nudStreetwise.Enabled = False
            nudSurvival.Enabled = False
            nudTactics.Enabled = False
            nudTechnician.Enabled = False
            cbTitle.Enabled = False
            nudBOD.Enabled = False
            nudCHA.Enabled = False
            nudDEX.Enabled = False
            nudLRN.Enabled = False
            nudXP.Enabled = False
            txtRank.Enabled = False
            cbFaction.Enabled = False
            labCP.Visible = True
            If cbCareer.SelectedIndex = 0 Then 'MechWarriors
                labCP.Text = "Mech Weight:  " & Leader.PrimaryMech & " tons"
                If Leader.BackupMech > 0 Then
                    labCP.Text &= ", Backup Mech Weight: " & Leader.BackupMech
                End If
            ElseIf cbCareer.SelectedIndex = 1 Then 'Pilot
                labCP.Text = "Fighter Weight:  " & Leader.Fighter
            ElseIf cbCareer.SelectedIndex = 2 Or cbCareer.SelectedIndex = 3 Then 'Tech
                labCP.Visible = False
                labCP.Text = ""
            ElseIf cbCareer.SelectedIndex = 4 Or cbCareer.SelectedIndex = 5 Then 'Scout
                labCP.Text = "Vehicle:  " & Leader.ScoutVehicle
            End If
        Else
            Update_Character_Cost()
        End If
    End Sub

    Private Sub chkBackupMech_CheckedChanged(sender As Object, e As EventArgs) Handles chkBackupMech.CheckedChanged
        If chkBackupMech.Checked Then
            nudBackupMech.Visible = True
            labBackupMech.Visible = True
        Else
            nudBackupMech.Visible = False
            labBackupMech.Visible = False
        End If
        Update_Character_Cost()
    End Sub
    Private Sub nudXP_ValueChanged(sender As Object, e As EventArgs) Handles nudXP.ValueChanged
        Select Case nudXP.Value
            Case 0
                labXP.Text = "XP Gained in Career"
            Case < 3001
                labXP.Text = "XP Gained in Career (Green)"
            Case < 15001
                labXP.Text = "XP Gained in Career (Regular)"'1 point to an attribute, +1 skill level in 2 skills
            Case < 100001
                labXP.Text = "XP Gained in Career (Veteran)" '1 point to an attribute, +1 skill level in 2 skills
            Case < 200001
                labXP.Text = "XP Gained in Career (Elite)" '1 point to an attribute, +1 skill level in 2 skills
            Case < 300001
                labXP.Text = "XP Gained in Career (Heroic)" '1 point to an attribute, +1 skill level in 2 skills
            Case < 400001
                labXP.Text = "XP Gained in Career (Legendary)" '1 point to an attribute, +1 skill level in 2 skills
            Case < 500001
                labXP.Text = "XP Gained in Career (Godlike)" '1 point to an attribute, +1 skill level in 2 skills
        End Select
        labCPXP.Text = nudXP.Value / 10 & " CP Available"
        Update_Character_XP()
    End Sub

    Private Sub Update_Character_XP()
        Dim TempCost As Int32 = 0

        Dim FreeAttributes As Int32 = 0
        If nudXP.Value > 3000 Then FreeAttributes += 1
        If nudXP.Value > 15000 Then FreeAttributes += 1
        If nudXP.Value > 100000 Then FreeAttributes += 1
        If nudXP.Value > 200000 Then FreeAttributes += 1
        If nudXP.Value > 300000 Then FreeAttributes += 1
        If nudXP.Value > 400000 Then FreeAttributes += 1
        Dim FreeSkill1 As Int32 = FreeAttributes
        Dim FreeSkill2 As Int32 = FreeAttributes

        Dim Diff As Int32 = 0
        Dim NewDiff As Int32 = 0

        Diff = nudLRN.Value - nudLRN.Minimum
        If Diff > 0 And FreeAttributes > 0 Then
            NewDiff = Math.Max(0, Diff - FreeAttributes)
            FreeAttributes -= Diff
        End If
        If NewDiff > 0 Then TempCost += Choose(NewDiff, 100, 300, 700, 1300, 2100, 3100)

        Diff = nudDEX.Value - nudDEX.Minimum
        If Diff > 0 And FreeAttributes > 0 Then
            NewDiff = Math.Max(0, Diff - FreeAttributes)
            FreeAttributes -= Diff
        End If
        If NewDiff > 0 Then TempCost += Choose(NewDiff, 80, 240, 480, 880, 1440, 2240)

        'I assume CHA is higher on priority over BOD for the leader
        Diff = nudCHA.Value - nudCHA.Minimum
        If Diff > 0 And FreeAttributes > 0 Then
            NewDiff = Math.Max(0, Diff - FreeAttributes)
            FreeAttributes -= Diff
        End If
        If NewDiff > 0 Then TempCost += Choose(NewDiff, 60, 180, 360, 660, 1110, 1710)

        Diff = nudBOD.Value - nudBOD.Minimum
        If Diff > 0 And FreeAttributes > 0 Then
            NewDiff = Math.Max(0, Diff - FreeAttributes)
            FreeAttributes -= Diff
        End If
        If NewDiff > 0 Then TempCost += Choose(NewDiff, 60, 180, 360, 660, 1110, 1710)

        Dim SkillDiff As New List(Of Point)
        SkillDiff.Add(New Point(nudAthletics.Value, nudAthletics.Minimum))
        SkillDiff.Add(New Point(nudBowBlade.Value, nudBowBlade.Minimum))
        SkillDiff.Add(New Point(nudBrawling.Value, nudBrawling.Minimum))
        SkillDiff.Add(New Point(nudComputer.Value, nudComputer.Minimum))
        SkillDiff.Add(New Point(nudDiplomacy.Value, nudDiplomacy.Minimum))
        SkillDiff.Add(New Point(nudDriver.Value, nudDriver.Minimum))
        SkillDiff.Add(New Point(nudEngineering.Value, nudEngineering.Minimum))
        SkillDiff.Add(New Point(nudGunneryAerospace.Value, nudGunneryAerospace.Minimum))
        SkillDiff.Add(New Point(nudGunneryArtillery.Value, nudGunneryArtillery.Minimum))
        SkillDiff.Add(New Point(nudGunneryMech.Value, nudGunneryMech.Minimum))
        SkillDiff.Add(New Point(nudInterrogation.Value, nudInterrogation.Minimum))
        SkillDiff.Add(New Point(nudJumpshipPilotingNavigation.Value, nudJumpshipPilotingNavigation.Minimum))
        SkillDiff.Add(New Point(nudLandManagement.Value, nudLandManagement.Minimum))
        SkillDiff.Add(New Point(nudLeadership.Value, nudLeadership.Minimum))
        SkillDiff.Add(New Point(nudMechanical.Value, nudMechanical.Minimum))
        SkillDiff.Add(New Point(nudMedicalFirstAid.Value, nudMedicalFirstAid.Minimum))
        SkillDiff.Add(New Point(nudPilotingAerospace.Value, nudPilotingAerospace.Minimum))
        SkillDiff.Add(New Point(nudPilotingMech.Value, nudPilotingMech.Minimum))
        SkillDiff.Add(New Point(nudPistol.Value, nudPistol.Minimum))
        SkillDiff.Add(New Point(nudRifle.Value, nudRifle.Minimum))
        SkillDiff.Add(New Point(nudRogue.Value, nudRogue.Minimum))
        SkillDiff.Add(New Point(nudStreetwise.Value, nudStreetwise.Minimum))
        SkillDiff.Add(New Point(nudSurvival.Value, nudSurvival.Minimum))
        SkillDiff.Add(New Point(nudTactics.Value, nudTactics.Minimum))
        SkillDiff.Add(New Point(nudTechnician.Value, nudTechnician.Minimum))
        TempCost += GetNewSkillXPCost(SkillDiff, FreeSkill1, FreeSkill2)


        TempCost += nudInformants2.Value * 5
        TempCost += nudUsefulContacts2.Value * 15
        TempCost += nudProminentContacts2.Value * 30
        If cbTitle.SelectedIndex = 1 Then
            TempCost += 1000
        ElseIf cbTitle.SelectedIndex = 2 Then
            TempCost += 1500
        ElseIf cbTitle.SelectedIndex = 3 Then
            TempCost += 3000
        End If

        If Leader.XP_Availabile > 0 Then TempCost = (Leader.XP_Total - Leader.XP_Availabile) / 10

        labCPXP.Text = Format(TempCost * 10, "###,##0") & " XP Used, " & Format(nudXP.Value - TempCost * 10, "###,##0") & " XP Remaining"
    End Sub
    Private Function GetNewSkillXPCost(SkilLDiff As List(Of Point), ByRef FreeSkill1 As Int32, ByRef FreeSkill2 As Int32) As Int32
        Dim TempCost As Int32 = 0
        For Each P As Point In SkilLDiff
            If P.X > P.Y Then 'X=Value, Y=Min, so X-Y = Change
                Dim Diff As Int32 = P.X - P.Y
                If FreeSkill1 > FreeSkill2 Then
                    Dim NewDiff As Int32 = Math.Max(0, Diff - FreeSkill1)
                    FreeSkill1 -= NewDiff
                    If NewDiff > 0 Then TempCost += CostForSkillIncrease(P.X - NewDiff, P.X)
                Else
                    Dim NewDiff As Int32 = Math.Max(0, Diff - FreeSkill2)
                    FreeSkill2 -= NewDiff
                    If NewDiff > 0 Then TempCost += CostForSkillIncrease(P.X - NewDiff, P.X)
                End If
            End If
        Next
        Return TempCost
    End Function
    Private Function CostForSkillIncrease(OldLevel As Int32, NewLevel As Int32) As Int32
        If OldLevel < 0 Then OldLevel = 0
        If NewLevel > 8 Then NewLevel = 8
        Dim LevelCost() As Int32 = {50, 75, 125, 175, 250, 325, 425, 550}
        Dim TempCost As Int32 = 0
        For X As Int32 = OldLevel To NewLevel - 1
            TempCost += LevelCost(X)
        Next
        Return TempCost
    End Function

    Private Sub cmdStartOver_Click(sender As Object, e As EventArgs) Handles cmdStartOver.Click
        ResetForm()
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Hide()
        frmMain.Show()
    End Sub


End Class