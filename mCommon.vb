Public Module mCommon
    Public Leader As New cCharacter
    Public CurrentContract As cContract
    Public CurrentTurn As New DateTime(3025, 1, 1)
    Public Mercenary As New cForce
    Public Friendly As New cForce
    Public Enemy As New cForce

    Public Forces_Friendly As New List(Of cRegiment)
    Public Forces_Enemy As New List(Of cRegiment)

    Public Availability As Dictionary(Of eUnitType, Boolean)
    Public Availability_S As Dictionary(Of eUnitType, Boolean)
    Public Availability_D As Dictionary(Of eUnitType, Boolean)

    Public ContractList As New Dictionary(Of String, cContract)

#Region "Dice"
    Dim Rand As Random '(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Int32)) 'Initialize random number generator
    Public Function D(Optional DieNumber As Int32 = 1, Optional DieSize As Int32 = 6) As Int32
        If Rand Is Nothing Then Rand = New Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Int32))
        'Dim Dies() As String = {"⚀", "⚁", "⚂", "⚃", "⚄", "⚅"}
        'Dim TempRollGraphic As String = ""
        Dim TempRoll As Int32 = 0
        For I As Int32 = 1 To DieNumber
            Dim Roll As Int32 = Rand.Next(1, DieSize + 1)
            'TempRollGraphic &= Dies(Roll - 1)
            TempRoll += Roll
        Next
        'Debug.Print(DieNumber & "d" & DieSize & " = " & TempRoll & " " & TempRollGraphic)
        Return TempRoll
    End Function
#End Region



#Region "Not really used right now, and probably will not be"
    Public Function Availability_Mech(MechName As String) As Boolean
        Dim Availability As Int32 = 2
        Select Case MechName.ToLower
            Case "locust lct-1v" : Availability = 4
            Case "wasp wsp-1a" : Availability = 4
            Case "stinger stg-3r" : Availability = 4
            Case "commando com-2d" : Availability = 9
            Case "javelin jvn-10n" : Availability = 7
            Case "spider sdr-5v" : Availability = 10
            Case "urbanmech um-r60" : Availability = 6
            Case "valkyrie vlk-qa" : Availability = 5
            Case "firestarter fs9-h" : Availability = 5
            Case "jenner jr7-d" : Availability = 9
            Case "ostscout ott-7j" : Availability = 7
            Case "panther pnt-9r" : Availability = 8

            Case "assassin asn-21" : Availability = 9
            Case "cicada cda-2a" : Availability = 8
            Case "clint clnt-2-3t" : Availability = 8
            Case "hermes ii her-2s" : Availability = 9
            Case "vulcan vl-2t" : Availability = 6
            Case "whitworth wth-1" : Availability = 8
            Case "blackjack bj-1" : Availability = 10
            Case "hatchetman hct-3f" : Availability = 9
            Case "phoenix hawk pxh-1h" : Availability = 5
            Case "vindicator vnd-1r" : Availability = 7
            Case "centurion cn9-a" : Availability = 7
            Case "enforcer enf-4r" : Availability = 8
            Case "hunchback hbk-4g" : Availability = 7
            Case "trebuchet tbt-5n" : Availability = 5
            Case "dervish dv-6m" : Availability = 4
            Case "griffin grf-1n" : Availability = 5
            Case "shadow hawk shd-2h" : Availability = 6
            Case "scorpion scp-1n" : Availability = 9
            Case "wolverine wvr-6r" : Availability = 7

            Case "dragon drg-1n" : Availability = 9
            Case "ostroc osr-2c" : Availability = 8
            Case "ostsol otl-4d" : Availability = 6
            Case "quickdraw qkd-4g" : Availability = 7
            Case "rifleman rfl-3n" : Availability = 6
            Case "catapult cplt-c1" : Availability = 10
            Case "crusader crd-3r" : Availability = 4
            Case "jagermech jm6-s" : Availability = 9
            Case "thunderbolt tdr-5s" : Availability = 6
            Case "archer arc-2r" : Availability = 6
            Case "grasshopper ghr-5h" : Availability = 8
            Case "warhammer whm-6r" : Availability = 7
            Case "marauder mad-3r" : Availability = 7
            Case "orion on1-k" : Availability = 6

            Case "awesome aws-8q" : Availability = 7
            Case "charger cgr-1a1" : Availability = 9
            Case "goliath gol-1h" : Availability = 10
            Case "victor vtr-9b" : Availability = 10
            Case "zeus zeu-6s" : Availability = 8
            Case "battlemaster blr-1g" : Availability = 6
            Case "stalker stk-3f" : Availability = 6
            Case "cyclops cp 10-z" : Availability = 8
            Case "banshee bnc-3e" : Availability = 7
            Case "atlas as7-d" : Availability = 8

                'Really not sure of these, so I made them up
            Case "stinger lam stg-a5" : Availability = 12
            Case "wasp lam wsp-105" : Availability = 12
            Case "phoenix hawk lam pxh-hk2" : Availability = 12

        End Select
        If D(2, 6) >= Availability Then Return True Else Return False
    End Function
    Public Function Availability_Fighter(FighterName As String) As Boolean
        Dim Availability As Int32 = 2
        Select Case FighterName.ToLower
            Case "sparrowhawk spr-h5" : Availability = 8
            Case "corsair csr-v12" : Availability = 7
            Case "stuka stu-k5" : Availability = 10
            Case "sl-21 sholagar" : Availability = 5
            Case "sl-17 shilone" : Availability = 6
            Case "sl-15 slayer" : Availability = 8
            Case "tr-7 thrush" : Availability = 6
            Case "tr-10 transit" : Availability = 9
            Case "tr-13 transgressor" : Availability = 6
            Case "f-10 cheetah" : Availability = 5
            Case "f-90 stingray" : Availability = 5
            Case "f-100 riever" : Availability = 8
            Case "syd-21 seydlitz" : Availability = 8
            Case "lcf-r15 lucifer" : Availability = 5
            Case "chp-w5 chippewa" : Availability = 9
            Case "sabre sb-27" : Availability = 7
            Case "centurion cnt-1d" : Availability = 10
            Case "hellcat hct-213" : Availability = 4
            Case "lightning ltn-g15" : Availability = 9
            Case "eagle egl-r6" : Availability = 5
            Case "thunderbird trb-d36" : Availability = 5

        End Select
        If D(2, 6) >= Availability Then Return True Else Return False
    End Function
    Public Function Availability_Infantry(DesiredWeapon As String) As String
        Dim Availability As Int32 = 2
        Select Case DesiredWeapon.ToLower
            Case "rifle" : Availability = 2
            Case "machine gun" : Availability = 8
            Case "flamer" : Availability = 9
            Case "laser" : Availability = 11
            Case "srm" : Availability = 11
        End Select
        If D(2, 6) >= Availability Then Return DesiredWeapon Else Return "Rifle"
    End Function
    Public Function Weight_Mech(MechName As String) As Int32
        Dim Weight As Int32 = 0
        Select Case MechName.ToLower
            Case "locust lct-1v" : Weight = 20
            Case "wasp wsp-1a" : Weight = 20
            Case "stinger stg-3r" : Weight = 20
            Case "commando com-2d" : Weight = 25
            Case "javelin jvn-10n" : Weight = 30
            Case "spider sdr-5v" : Weight = 30
            Case "urbanmech um-r60" : Weight = 30
            Case "valkyrie vlk-qa" : Weight = 30
            Case "firestarter fs9-h" : Weight = 35
            Case "jenner jr7-d" : Weight = 35
            Case "ostscout ott-7j" : Weight = 35
            Case "panther pnt-9r" : Weight = 35

            Case "assassin asn-21" : Weight = 40
            Case "cicada cda-2a" : Weight = 40
            Case "clint clnt-2-3t" : Weight = 40
            Case "hermes ii her-2s" : Weight = 40
            Case "vulcan vl-2t" : Weight = 40
            Case "whitworth wth-1" : Weight = 40
            Case "blackjack bj-1" : Weight = 45
            Case "hatchetman hct-3f" : Weight = 45
            Case "phoenix hawk pxh-1h" : Weight = 45
            Case "vindicator vnd-1r" : Weight = 45
            Case "centurion cn9-a" : Weight = 50
            Case "enforcer enf-4r" : Weight = 50
            Case "hunchback hbk-4g" : Weight = 50
            Case "trebuchet tbt-5n" : Weight = 50
            Case "dervish dv-6m" : Weight = 55
            Case "griffin grf-1n" : Weight = 55
            Case "shadow hawk shd-2h" : Weight = 55
            Case "scorpion scp-1n" : Weight = 55
            Case "wolverine wvr-6r" : Weight = 55

            Case "dragon drg-1n" : Weight = 60
            Case "ostroc osr-2c" : Weight = 60
            Case "ostsol otl-4d" : Weight = 60
            Case "quickdraw qkd-4g" : Weight = 60
            Case "rifleman rfl-3n" : Weight = 60
            Case "catapult cplt-c1" : Weight = 65
            Case "crusader crd-3r" : Weight = 65
            Case "jagermech jm6-s" : Weight = 65
            Case "thunderbolt tdr-5s" : Weight = 65
            Case "archer arc-2r" : Weight = 70
            Case "grasshopper ghr-5h" : Weight = 70
            Case "warhammer whm-6r" : Weight = 70
            Case "marauder mad-3r" : Weight = 75
            Case "orion on1-k" : Weight = 75

            Case "awesome aws-8q" : Weight = 80
            Case "charger cgr-1a1" : Weight = 80
            Case "goliath gol-1h" : Weight = 80
            Case "victor vtr-9b" : Weight = 80
            Case "zeus zeu-6s" : Weight = 80
            Case "battlemaster blr-1g" : Weight = 85
            Case "stalker stk-3f" : Weight = 85
            Case "cyclops cp 10-z" : Weight = 90
            Case "banshee bnc-3e" : Weight = 95
            Case "atlas as7-d" : Weight = 100

                'Really not sure of these, so I made them up
            Case "stinger lam stg-a5" : Weight = 30
            Case "wasp lam wsp-105" : Weight = 30
            Case "phoenix hawk lam pxh-hk2" : Weight = 55

        End Select
        If D(2, 6) >= Weight Then Return True Else Return False
    End Function
    Public Function Weight_Fighter(FighterName As String) As Boolean
        Dim Weight As Int32 = 0
        Select Case FighterName.ToLower
            Case "sparrowhawk spr-h5" : Weight = 30
            Case "corsair csr-v12" : Weight = 50
            Case "stuka stu-k5" : Weight = 100

            Case "sl-21 sholagar" : Weight = 35
            Case "sl-17 shilone" : Weight = 65
            Case "sl-15 slayer" : Weight = 80

            Case "tr-7 thrush" : Weight = 25
            Case "tr-10 transit" : Weight = 50
            Case "tr-13 transgressor" : Weight = 75

            Case "f-10 cheetah" : Weight = 25
            Case "f-90 stingray" : Weight = 60
            Case "f-100 riever" : Weight = 100

            Case "syd-21 seydlitz" : Weight = 20
            Case "lcf-r15 lucifer" : Weight = 65
            Case "chp-w5 chippewa" : Weight = 90

            Case "sabre SB-27" : Weight = 25
            Case "centurion cnt-1d" : Weight = 30
            Case "hellcat hct-213" : Weight = 60
            Case "lightning ltn-g15" : Weight = 50
            Case "eagle egl-r6" : Weight = 75
            Case "thunderbird trb-d36" : Weight = 100
        End Select
        Return Weight
    End Function
    Public Function ShipLocation() As eShipLocation
        Select Case D(1, 6)
            Case 1 To 2 : Return eShipLocation.Nearest_Spaceport
            Case 3 To 4 : Return eShipLocation.In_Orbit
            Case 5 : Return eShipLocation.Elsewhere_In_System
            Case 6 : Return eShipLocation.Neighboring_System
        End Select
    End Function
#End Region

    Public Function Availability_Jumpship(JumpshipType As eUnitType, Quality As eQuality) As Boolean
        Dim Availability As Int32 = 2
        'If frmCreate.chkActualCosts.Checked Then
        Select Case JumpshipType
                Case eUnitType.Scout_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 8
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Merchant_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 8
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Invader_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 8
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Star_Lord_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Availability = 10
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 11
                    End If
                Case eUnitType.Monolith_Jumpship
                    If Quality = eQuality.LikeNew Then
                        Availability = 11
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 10
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 12
                    End If
            End Select
        'Else
        'Select Case JumpshipType
        'Case eUnitType.Scout_Jumpship : If Quality = eQuality.LikeNew Then Availability = 6 Else Availability = 4
        'Case eUnitType.Merchant_Jumpship : If Quality = eQuality.LikeNew Then Availability = 5 Else Availability = 4
        'Case eUnitType.Invader_Jumpship : If Quality = eQuality.LikeNew Then Availability = 8 Else Availability = 6
        'Case eUnitType.Star_Lord_Jumpship : If Quality = eQuality.LikeNew Then Availability = 9 Else Availability = 7
        'Case eUnitType.Monolith_Jumpship : If Quality = eQuality.LikeNew Then Availability = 10 Else Availability = 8
        'End Select
        'End If

        'If frmCreate.chkActualCosts.Checked Then
        DebugEvent("Sales Agent paid 300 C-Bills to look for leads.", True)
        Mercenary.Cash_On_Hand -= 300 'agent gets €300 per Month
        'End If

        If D(2, 6) >= Availability Then
            'If frmCreate.chkActualCosts.Checked Then
            DebugEvent("Sales Agent paid " & GetUnitCost(JumpshipType) * 0.001 & " C-Bills for commission on sale.", True)
            Mercenary.Cash_On_Hand -= GetUnitCost(JumpshipType) * 0.001 'Agent gets .1% for commission
            'End If
            Return True
        Else
            Return False
        End If
    End Function
    Public Function Availability_Dropship(DropshipType As eUnitType, Quality As eQuality) As Boolean
        Dim Availability As Int32 = 2
        'If frmCreate.chkActualCosts.Checked Then
        Select Case DropshipType
                Case eUnitType.Leopard_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 8
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 7
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Union_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 8
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 7
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Overlord_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 10
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Fury_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 8
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 7
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Gazelle_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 7
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 6
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 7
                    End If
                Case eUnitType.Seeker_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 10
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 11
                    End If
                Case eUnitType.Triumph_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 10
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Condor_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 8
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Excaliber_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 11
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 10
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Avenger_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 8
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 7
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Achilles_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 11
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Intruder_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 8
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Fortress_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 12
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 11
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 12
                    End If
                Case eUnitType.Leopard_CV_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 10
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 11
                    End If
                Case eUnitType.Vengeance_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 11
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 10
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 10
                    End If
                Case eUnitType.Buccaneer_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 7
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 6
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 8
                    End If
                Case eUnitType.Mule_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 7
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 6
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 7
                    End If
                Case eUnitType.Monarch_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 9
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 10
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 9
                    End If
                Case eUnitType.Mammoth_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 11
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 9
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 11
                    End If
                Case eUnitType.Behemoth_Dropship
                    If Quality = eQuality.LikeNew Then
                        Availability = 12
                    ElseIf Quality = eQuality.Salvaged Then
                        Availability = 11
                    ElseIf Quality = eQuality.Destroyed Then
                        Availability = 12
                    End If
            End Select
        'Else
        'Availability = 2 'Always available under MW1E rules
        'End If

        'If frmCreate.chkActualCosts.Checked Then
        DebugEvent("Sales Agent paid 300 C-Bills to look for leads.", True)
        Mercenary.Cash_On_Hand -= 300 'agent gets €300 per Month
        'End If

        If D(2, 6) >= Availability Then
            'If frmCreate.chkActualCosts.Checked Then
            DebugEvent("Sales Agent paid " & GetUnitCost(DropshipType) * 0.001 & " C-Bills for commission on sale.", True)
            Mercenary.Cash_On_Hand -= GetUnitCost(DropshipType) * 0.001 'Agent gets .1% for commission
            'End If
            Return True
        Else
            DebugEvent("Ship was not available.", True)
            Return False
        End If
    End Function
    Public Function Available_Leads() As Int32() 'This returns a number of modified die rolls
        'TODO:  Rename to Available_Ship_Leads
        Try
            Dim DieRoll As Int32 = D(2, 6)
            If Leader.Skills(eSkill.Diplomacy) = 4 And Leader.Skills(eSkill.Diplomacy) = 5 Then
                DieRoll += 1
            ElseIf Leader.Skills(eSkill.Diplomacy) >= 6 And Leader.Skills(eSkill.Diplomacy) <= 8 Then
                DieRoll += 2
            End If
            Select Case Math.Round((Leader.CHA + Leader.LRN) / 2, 0, MidpointRounding.AwayFromZero)
                Case 1 To 4 : DieRoll -= 2
                Case 5 To 6 : DieRoll -= 1
            End Select
            DieRoll = Math.Min(14, Math.Max(0, DieRoll))

            Dim Leads As Int32 = 0
            Dim AvMod As Int32 = 0
            Select Case DieRoll
                Case 0 To 3
                    Return {0}
                Case 4
                    Leads = 1
                    AvMod = 2
                Case 5
                    Leads = 1
                    AvMod = 1
                Case 6
                    Leads = 1
                Case 7
                    Leads = 2
                    AvMod = 1
                Case 8
                    Leads = 2
                Case 9
                    Leads = 2
                    AvMod = -1
                Case 10
                    Leads = 3
                    AvMod = 1
                Case 11
                    Leads = 3
                Case 12
                    Leads = 4
                Case 13
                    Leads = 4
                    AvMod = -1
                Case 14
                    Leads = 5
                    AvMod = -1
            End Select
            DebugEvent("Available Leads: " & Leads & ", Modifier: " & AvMod, True)
            Dim TempReturn(Leads) As Int32
            For R As Int32 = 1 To Leads
                Dim LeadRoll As Int32 = D(2, 6)
                TempReturn(R - 1) = LeadRoll + AvMod
                DebugEvent("  Lead:  " & LeadRoll + AvMod & "+", True)
            Next
            Return TempReturn
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Function
    Public Function GetUnitCost(UnitType As eUnitType) As Single
        Dim TempCost As Single = 0
        Select Case UnitType
            Case eUnitType.Mech_Light : TempCost = 460
            Case eUnitType.Mech_Medium : TempCost = 780
            Case eUnitType.Mech_Heavy : TempCost = 1140
            Case eUnitType.Mech_Assault : TempCost = 1640
            Case eUnitType.Fighter_Light : TempCost = 400
            Case eUnitType.Fighter_Medium : TempCost = 700
            Case eUnitType.Fighter_Heavy : TempCost = 1000
            Case eUnitType.LAM_Light : TempCost = 800
            Case eUnitType.LAM_Medium : TempCost = 1460
            Case eUnitType.Infantry_Regular : TempCost = 50
            Case eUnitType.Infantry_Motorized : TempCost = 80
            Case eUnitType.Infantry_Jump : TempCost = 100
            Case eUnitType.Infantry_Scout : TempCost = 80
            Case eUnitType.Infantry_Regular_Airmobile : TempCost = 70
            Case eUnitType.Infantry_Motorized_Airmobile : TempCost = 100
            Case eUnitType.Infantry_Jump_Airmobile : TempCost = 120
            Case eUnitType.Infantry_Scout_Airmobile : TempCost = 100
            Case eUnitType.Armor_Light : TempCost = 200
            Case eUnitType.Armor_Heavy : TempCost = 600
            Case eUnitType.Artillery : TempCost = 800
            Case eUnitType.Aircraft : TempCost = 70
            Case eUnitType.Support : TempCost = 100
            Case eUnitType.Leopard_Dropship : TempCost = 16808'6000  
            Case eUnitType.Union_Dropship : TempCost = 21471'16000  
            Case eUnitType.Overlord_Dropship : TempCost = 32114'43000  
            Case eUnitType.Fury_Dropship : TempCost = 16447'3000  
            Case eUnitType.Gazelle_Dropship : TempCost = 17705'4000  
            Case eUnitType.Seeker_Dropship : TempCost = 22226'10000  
            Case eUnitType.Triumph_Dropship : TempCost = 34593'45000  
            Case eUnitType.Condor_Dropship : TempCost = 26600'30000  
            Case eUnitType.Excaliber_Dropship : TempCost = 42934'75000  

            Case eUnitType.Scout_Jumpship : TempCost = 38954'30000  
            Case eUnitType.Invader_Jumpship : TempCost = 67713'50000  
            Case eUnitType.Monolith_Jumpship : TempCost = 165683'100000  
            Case eUnitType.Star_Lord_Jumpship : TempCost = 114176'75000  
            Case eUnitType.Merchant_Jumpship : TempCost = 53463'40000  

            Case eUnitType.Avenger_Dropship : TempCost = 23390'7000  
            Case eUnitType.Achilles_Dropship : TempCost = 42801'35000  
            Case eUnitType.Intruder_Dropship : TempCost = 28957'20000  
            Case eUnitType.Fortress_Dropship : TempCost = 33770'80000  
            Case eUnitType.Leopard_CV_Dropship : TempCost = 16808'6000  
            Case eUnitType.Vengeance_Dropship : TempCost = 38300'35000  
            Case eUnitType.Buccaneer_Dropship : TempCost = 10842'10000  
            Case eUnitType.Mule_Dropship : TempCost = 16324'30000  
            Case eUnitType.Monarch_Dropship : TempCost = 15392'20000  
            Case eUnitType.Mammoth_Dropship : TempCost = 46525'120000  
            Case eUnitType.Behemoth_Dropship : TempCost = 63152 '200000  
        End Select
        Return TempCost
    End Function


    Dim Name_Given As New Dictionary(Of String, Int32)
    Dim Name_Surname As New Dictionary(Of String, Int32)
    Dim Name_Given_Sum As Int64 = 0
    Dim Name_Surname_Sum As Int64 = 0
    Public Sub LoadNames()
        DebugEvent("Loading name files...", True)
        Dim NVP() As String

        For Each G As String In System.IO.File.ReadAllLines("Names\Given.txt")
            If G.Trim.StartsWith("#") = False Then
                NVP = Split(G, vbTab)
                Name_Given_Sum += CLng(NVP(1))
                Name_Given.Add(NVP(0) & vbTab & NVP(2), Name_Given_Sum)
            End If
        Next

        For Each S As String In System.IO.File.ReadAllLines("Names\Surname.txt")
            If S.Trim.StartsWith("#") = False Then
                NVP = Split(S, vbTab)
                Name_Surname_Sum += CLng(NVP(1))
                Name_Surname.Add(NVP(0), Name_Surname_Sum)
            End If
        Next
    End Sub
    Public Function GetName() As String
        Dim FN As Int64 = D(1, Name_Given_Sum) - 1
        Dim LN As Int64 = D(1, Name_Surname_Sum) - 1
        Dim FirstName As String = ""
        Dim LastName As String = ""
        Dim NVP() As String
        For Each First In Name_Given
            If First.Value <= FN Then
                NVP = Split(First.Key, vbTab)
                FirstName = NVP(0)
            End If
        Next
        For Each Last In Name_Surname
            If Last.Value <= LN Then
                LastName = Last.Key
            End If
        Next
        Return FirstName & " " & LastName
    End Function
    Public Function GetNumberString(Number As Int32) As String
        Dim Prefix As String = "th"
        Select Case Number
            Case 1 : Prefix = "st"
            Case 2 : Prefix = "nd"
            Case 3 : Prefix = "rd"
            Case 23, 33, 43, 53, 63, 73, 83, 93, 103 : Prefix = "rd"
            Case Else
                Prefix = "th"
        End Select
        'If Right("00" & Number.ToString, 2) = "11" Then Prefix = "th"
        Return Number & Prefix
    End Function
    Public Function FriendlyCommitment(Mission As eMission) As Int32 'This is done at the start of the contract
        Dim DieRoll As Int32 = D(2, 6)
        Dim TempVal As Int32 = 0
        Select Case Mission
            Case eMission.Cadre_Duty : TempVal = Choose(DieRoll - 1, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 5)
            Case eMission.Garrison_Duty : TempVal = Choose(DieRoll - 1, 4, 3, 2, 1, 1, 1, 1, 2, 3, 4, 5)
            Case eMission.Security_Duty : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 1, 1, 1, 1, 2, 3, 4)
            Case eMission.Offensive_Campaign : TempVal = Choose(DieRoll - 1, 2, 2, 2, 3, 4, 4, 5, 6, 7, 8, 9)
            Case eMission.Defensive_Campaign : TempVal = Choose(DieRoll - 1, 5, 4, 3, 2, 2, 3, 4, 5, 6, 7, 8)
            Case eMission.Planetary_Assault : TempVal = Choose(DieRoll - 1, 15, 13, 11, 9, 7, 5, 6, 8, 10, 12, 14)
            Case eMission.Relief_Duty : TempVal = Choose(DieRoll - 1, 3, 3, 2, 2, 1, 2, 3, 4, 5, 6, 7)
            Case eMission.Riot_Duty : TempVal = Choose(DieRoll - 1, 4, 3, 2, 1, 1, 1, 1, 1, 2, 3, 4)
            Case eMission.Recon_Raid : TempVal = Choose(DieRoll - 1, 4, 3, 2, 1, 1, 1, 1, 2, 3, 4, 5)
            Case eMission.Objective_Raid : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 1, 2, 1, 2, 3, 4, 5)
            Case eMission.Divisionary_Raid : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 2, 3, 2, 2, 3, 4, 6)
            Case eMission.Siege_Duty : TempVal = Choose(DieRoll - 1, 1, 2, 2, 3, 4, 4, 4, 5, 6, 7, 8)
            Case eMission.Guerrilla_Warfare : TempVal = Choose(DieRoll - 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 4)
        End Select
        DebugEvent("Friendly forces commit " & TempVal & " regiments to the " & Mission.ToString & " mission!", True)
        Return TempVal
    End Function
    Public Function EnemyCommitment(Mission As eMission) As Int32  'This is done monthly
        Dim DieRoll As Int32 = D(2, 6)
        Dim TempVal As Int32 = 0
        Select Case Mission
            Case eMission.Cadre_Duty : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 1, 2, 1, 2, 3, 4, 5)
            Case eMission.Garrison_Duty : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 2, 3, 2, 2, 3, 4, 6)
            Case eMission.Security_Duty : TempVal = Choose(DieRoll - 1, 4, 3, 2, 1, 1, 1, 1, 1, 2, 3, 4)
            Case eMission.Offensive_Campaign : TempVal = Choose(DieRoll - 1, 5, 4, 3, 2, 2, 3, 4, 5, 6, 7, 8)
            Case eMission.Defensive_Campaign : TempVal = Choose(DieRoll - 1, 2, 2, 2, 3, 4, 4, 5, 6, 7, 8, 9)
            Case eMission.Planetary_Assault : TempVal = Choose(DieRoll - 1, 13, 11, 9, 7, 5, 3, 4, 6, 8, 10, 12)
            Case eMission.Relief_Duty : TempVal = Choose(DieRoll - 1, 1, 2, 2, 3, 4, 4, 4, 5, 6, 7, 8)
            Case eMission.Riot_Duty : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 1, 1, 1, 1, 2, 3, 4)
            Case eMission.Recon_Raid : TempVal = Choose(DieRoll - 1, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 5)
            Case eMission.Objective_Raid : TempVal = Choose(DieRoll - 1, 4, 3, 2, 1, 1, 1, 1, 2, 3, 4, 5)
            Case eMission.Divisionary_Raid : TempVal = Choose(DieRoll - 1, 5, 4, 3, 2, 2, 3, 4, 5, 6, 7, 8)
            Case eMission.Siege_Duty : TempVal = Choose(DieRoll - 1, 3, 3, 2, 2, 1, 2, 3, 4, 5, 6, 7)
            Case eMission.Guerrilla_Warfare : TempVal = Choose(DieRoll - 1, 3, 2, 1, 1, 1, 1, 1, 1, 2, 3, 4)
        End Select
        DebugEvent("Enemy forces commit " & TempVal & " regiments to the " & Mission.ToString & " mission!", True)
        Return TempVal
    End Function
    Public Function GetTacticalSkill(ForceComposition As eForceComposition) As Int32
        Dim DieRoll As Int32 = D(2, 6)
        If ForceComposition = eForceComposition.Militia Then DieRoll -= 3
        Select Case DieRoll
            Case < 5 : Return 1
            Case 5 To 6 : Return 2
            Case 7 To 8 : Return 3
            Case 9 : Return 4
            Case 10 : Return 5
            Case 11 : Return 6
            Case 12 : Return 7
        End Select
        Return 0
    End Function
    Public Function GetRegiment_Friendly(Mission As eMission) As eForceComposition
        Select Case Mission
            Case eMission.Cadre_Duty, eMission.Garrison_Duty, eMission.Security_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Infantry, eForceComposition.Militia, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Air, eForceComposition.Mech, eForceComposition.Mech)
            Case eMission.Offensive_Campaign, eMission.Defensive_Campaign
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Air, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Militia, eForceComposition.Aerospace)
            Case eMission.Planetary_Assault, eMission.Relief_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Aerospace, eForceComposition.Infantry, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Aerospace, eForceComposition.Aerospace)
            Case eMission.Riot_Duty, eMission.Siege_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Air)
            Case eMission.Recon_Raid, eMission.Objective_Raid, eMission.Divisionary_Raid, eMission.Guerrilla_Warfare
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Aerospace, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Aerospace, eForceComposition.Aerospace)
        End Select

    End Function
    Public Function GetRegiment_Enemy(Mission As eMission) As eForceComposition
        Select Case Mission
            Case eMission.Cadre_Duty, eMission.Garrison_Duty, eMission.Security_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Aerospace, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Aerospace, eForceComposition.Aerospace)
            Case eMission.Offensive_Campaign, eMission.Defensive_Campaign
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Air, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Militia, eForceComposition.Aerospace)
            Case eMission.Planetary_Assault, eMission.Relief_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Aerospace, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Air, eForceComposition.Aerospace)
            Case eMission.Riot_Duty, eMission.Siege_Duty
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Militia, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Mech, eForceComposition.Militia, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Air)
            Case eMission.Recon_Raid, eMission.Objective_Raid, eMission.Divisionary_Raid, eMission.Guerrilla_Warfare
                Return Choose(D(2, 6) - 1, eForceComposition.Aerospace, eForceComposition.Mech, eForceComposition.Armor, eForceComposition.Infantry, eForceComposition.Infantry, eForceComposition.Militia, eForceComposition.Infantry, eForceComposition.Armor, eForceComposition.Air, eForceComposition.Mech, eForceComposition.Mech)
        End Select
    End Function
    Public Function GetGenericQuality() As eExperience
        Select Case D(2, 6)
            Case 2 To 5 : Return eExperience.Green
            Case 6 To 8 : Return eExperience.Regular
            Case 9 To 11 : Return eExperience.Veteran
            Case 12 : Return eExperience.Elite
        End Select
        Return eExperience.Regular
    End Function
    Public Function GetUnitSize(Aerospace As Boolean) As Int32
        If Aerospace Then 'TODO:  This is *NOT* in the rules, but these regiments are 18 vs 108, always.  I completely made up the table
            Select Case D(2, 6)
                Case 2 : Return 6
                Case 3 : Return 12
                Case 4 : Return 14
                Case 5 : Return 18
                Case 6 : Return 18
                Case 7 : Return 18
                Case 8 : Return 18
                Case 9 : Return 18
                Case 10 : Return 16
                Case 11 : Return 10
                Case 12 : Return 8
            End Select
        Else
            Select Case D(2, 6)
                Case 2 : Return 40 + D(5, 6)
                Case 3 : Return 50 + D(5, 6)
                Case 4 : Return 60 + D(5, 6)
                Case 5 : Return 70 + D(5, 6)
                Case 6 : Return 80 + D(5, 6)
                Case 7 : Return 90 + D(5, 6)
                Case 8 : Return 100 + D(5, 6)
                Case 9 : Return 110 + D(5, 6)
                Case 10 : Return 120 + D(5, 6)
                Case 11 : Return 135 + D(5, 6)
                Case 12 : Return 150 + D(5, 6)
            End Select
        End If
    End Function
    Public Function GetSquadType(RegimentType As eForceComposition) As eUnitType
        Select Case RegimentType
            Case eForceComposition.Aerospace
                Select Case D(2, 6)
                    Case 2 : Return eUnitType.LAM_Light
                    Case 3 To 11 : Return eUnitType.Fighter_Medium'TODO:  Not sure what weight to use here
                    Case 12 : Return eUnitType.LAM_Medium
                End Select
            Case eForceComposition.Air
                Return eUnitType.Aircraft
            Case eForceComposition.Infantry
                Select Case D(2, 6)
                    Case 2 : Return eUnitType.Infantry_Regular_Airmobile'TODO:  Not sure which one to use here
                    Case 3, 4, 10 : Return eUnitType.Infantry_Jump
                    Case 5, 6, 8, 9 : Return eUnitType.Infantry_Regular
                    Case 7 : Return eUnitType.Infantry_Motorized
                    Case 11, 12 : Return eUnitType.Artillery
                End Select
            Case eForceComposition.Militia
                Select Case D(2, 6)
                    Case 2, 12 : Return eUnitType.Infantry_Motorized
                    Case 3 : Return eUnitType.Armor_Light
                    Case 10 : Return eUnitType.Artillery
                    Case 11 : Return eUnitType.Armor_Heavy
                    Case 4 To 9 : Return eUnitType.Infantry_Regular
                End Select
            Case eForceComposition.Armor
                Select Case D(2, 6)
                    Case 2, 12 : Return eUnitType.Artillery
                    Case 3 To 7 : Return eUnitType.Armor_Light
                    Case 8 To 10 : Return eUnitType.Armor_Heavy
                    Case 11 : Return eUnitType.Infantry_Motorized
                End Select

            Case eForceComposition.Mech
                Select Case D(2, 6)
                    Case 2, 12 : Return eUnitType.Mech_Assault
                    Case 3, 4, 9, 10, 11 : Return eUnitType.Mech_Heavy
                    Case 5, 6 : Return eUnitType.Mech_Medium
                    Case 7, 8 : Return eUnitType.Mech_Light
                End Select
        End Select
    End Function
    Public Function StrengthMult(Experience As eExperience) As Single
        Dim Mult As Single = 1
        If Experience = eExperience.Elite Then
            Mult = 2
        ElseIf Experience = eExperience.Veteran Then
            Mult = 1.5
        ElseIf Experience = eExperience.Regular Then
            Mult = 1
        ElseIf Experience = eExperience.Green Then
            Mult = 0.75
        End If
        Return Mult
    End Function
    Public Function GetHirelingRate(UnitType As eUnitType, Experience As eExperience, State As eSquadState) As Single
        If State <> eSquadState.Killed Then
            Dim TempRate As Int32 = Choose(D(1, 6), -5, -10, -15, -20, -25, -30)
            If Experience = eExperience.Green Then
                TempRate -= 10
            ElseIf Experience = eExperience.Veteran Then
                TempRate += 10
            ElseIf Experience = eExperience.Elite Then
                TempRate += 20
            End If
            Select Case UnitType
                Case eUnitType.Mech_Light : TempRate -= 10
                Case eUnitType.Mech_Medium : TempRate += 0
                Case eUnitType.Mech_Heavy : TempRate += 10
                Case eUnitType.Mech_Assault : TempRate += 20
                Case eUnitType.Fighter_Light : TempRate += 0
                Case eUnitType.Fighter_Medium : TempRate += 5
                Case eUnitType.Fighter_Heavy : TempRate += 10
                Case eUnitType.LAM_Light : TempRate += 5
                Case eUnitType.LAM_Medium : TempRate += 10
                Case eUnitType.Infantry_Regular : TempRate -= 30
                Case eUnitType.Infantry_Jump : TempRate -= 20
                Case eUnitType.Infantry_Scout : TempRate -= 5
                Case eUnitType.Infantry_Motorized : TempRate += 0
                Case eUnitType.Infantry_Regular_Airmobile : TempRate -= 25
                Case eUnitType.Infantry_Jump_Airmobile : TempRate -= 15
                Case eUnitType.Infantry_Scout_Airmobile : TempRate += 0
                Case eUnitType.Infantry_Motorized_Airmobile : TempRate += 5
                Case eUnitType.Armor_Light : TempRate -= 15
                Case eUnitType.Armor_Heavy : TempRate -= 10
                Case eUnitType.Artillery : TempRate -= 10
                Case eUnitType.Aircraft : TempRate -= 10
                Case eUnitType.Support : TempRate += 10
            End Select
            Return 1 + (TempRate / 100)
        Else
            Return 0
        End If
    End Function

    Public Sub FillSquads(TempUnit As cRegiment)
        Dim TempSize As Int32 = GetUnitSize(IIf((TempUnit.ForceComposition = eForceComposition.Air Or TempUnit.ForceComposition = eForceComposition.Aerospace), True, False))
        Dim TempSquadList As New List(Of cSquad)
        Do
            Dim Chunk As Int32 = D(2, 6)
            If Chunk > TempSize Then Chunk = TempSize
            Dim TempType As eUnitType = GetSquadType(TempUnit.ForceComposition)
            For S As Int32 = 1 To Chunk
                Dim TempSquad As New cSquad
                TempSquad.State = eSquadState.Active
                TempSquad.Pilot = GetName()
                TempSquad.UnitType = TempType
                TempSquadList.Add(TempSquad)
            Next
            TempSize -= Chunk
            If TempSize = 0 Then Exit Do
        Loop

        'Now I just have a big ass list of pretty similar units, lets sort them
        TempSquadList.Sort(Function(X, Y) X.UnitType.CompareTo(Y.UnitType))
        TempSquadList = TempSquadList.OrderBy(Function(x) x.UnitType).ToList()

        'And lets build the unit
        Dim CurrentRegiment As Int32 = 1
        Dim CurrentBattalion As Int32 = 1
        Dim CurrentCompany As Int32 = 1
        Dim CurrentLance As Int32 = 1

        Dim Counter As Int32 = 0
        For Each S As cSquad In TempSquadList
            Counter += 1
            If Counter > IIf((TempUnit.ForceComposition = eForceComposition.Air Or TempUnit.ForceComposition = eForceComposition.Aerospace), 2, 4) Then
                CurrentLance += 1
                Counter = 1
            End If
            If CurrentLance = 4 Then
                CurrentCompany += 1
                CurrentLance = 1
            End If
            If CurrentCompany = 4 Then
                CurrentBattalion += 1
                CurrentCompany = 1
                CurrentLance = 1
            End If

            If TempUnit.Battalions.ContainsKey(CurrentBattalion) = False Then
                TempUnit.Battalions.Add(CurrentBattalion, New cBattalion(CurrentBattalion))
            End If
            If TempUnit.Battalions(CurrentBattalion).Companies.ContainsKey(CurrentCompany) = False Then
                TempUnit.Battalions(CurrentBattalion).Companies.Add(CurrentCompany, New cCompany(CurrentCompany))
            End If
            If TempUnit.Battalions(CurrentBattalion).Companies(CurrentCompany).Lances.ContainsKey(CurrentLance) = False Then
                TempUnit.Battalions(CurrentBattalion).Companies(CurrentCompany).Lances.Add(CurrentLance, New cLance(CurrentLance))
            End If
            TempUnit.Battalions(CurrentBattalion).Companies(CurrentCompany).Lances(CurrentLance).Squads.Add(S.ID, S)
        Next

        'Now go back through and name them
        For Each B As cBattalion In TempUnit.Battalions.Values
            B.Name = GetNumberString(B.Name)
            For Each C As cCompany In B.Companies.Values
                C.Name = GetNumberString(C.Name)
                For Each L As cLance In C.Lances.Values
                    L.Name = GetNumberString(L.Name)
                Next
            Next
        Next
    End Sub
    Public Function DetermineSkill(Experience As eExperience) As eExperience
        Dim TempExperience As eExperience
        Select Case Experience
            Case eExperience.Green : TempExperience = Choose(D(1, 6), 0, 0, 1, 1, 1, 1)
            Case eExperience.Regular : TempExperience = Choose(D(1, 6), 1, 1, 2, 2, 2, 3)
            Case eExperience.Veteran : TempExperience = Choose(D(1, 6), 1, 2, 2, 3, 3, 4)
            Case eExperience.Elite : TempExperience = Choose(D(1, 6), 3, 3, 3, 4, 5, 6)
        End Select
        Return TempExperience
    End Function

    Public Sub GetOriginalSize(R As cRegiment)
        R.OriginalSize = R.SquadCount
        For Each B As cBattalion In R.Battalions.Values
            B.OriginalSize = B.SquadCount
            For Each C As cCompany In B.Companies.Values
                C.OriginalSize = C.SquadCount
                For Each L As cLance In C.Lances.Values
                    L.OriginalSize = L.SquadCount
                Next
            Next
        Next
    End Sub

    'These kinda suck, but may have some useful symbols in them:
    Public Function ShowRegiment(R As cRegiment) As String
        Dim TempState As String = ""
        Dim TempVal As String = R.Name & " Regiment:" & vbCrLf
        For Each B As cBattalion In R.Battalions.Values
            TempVal &= " " & B.Name & " Battalion:" & vbCrLf
            For Each C As cCompany In B.Companies.Values
                TempVal &= "  " & C.Name & " Company:  "
                For Each L As cLance In C.Lances.Values
                    TempVal &= " "
                    For Each S As cSquad In L.Squads.Values
                        Select Case S.UnitType
                            Case eUnitType.Mech_Light : TempVal &= "L"'⍂
                            Case eUnitType.Mech_Medium : TempVal &= "M"'⍌
                            Case eUnitType.Mech_Heavy : TempVal &= "H"'⍁
                            Case eUnitType.Mech_Assault : TempVal &= "A"'⍓
                            Case eUnitType.LAM_Light : TempVal &= "L"'L*
                            Case eUnitType.LAM_Medium : TempVal &= "M"'M*
                            Case eUnitType.Fighter_Light : TempVal &= "F"'L*"'★
                            Case eUnitType.Fighter_Medium : TempVal &= "F"'M*"'★
                            Case eUnitType.Fighter_Heavy : TempVal &= "F"'H*"'★
                            Case eUnitType.Aircraft : TempVal &= "R"'☆
                            Case eUnitType.Infantry_Jump : TempVal &= "I"'♢
                            Case eUnitType.Infantry_Motorized : TempVal &= "I"'♢
                            Case eUnitType.Infantry_Regular : TempVal &= "I"'♢
                            Case eUnitType.Infantry_Scout : TempVal &= "I"'♢
                            Case eUnitType.Infantry_Jump_Airmobile : TempVal &= "I"'⌺
                            Case eUnitType.Infantry_Motorized_Airmobile : TempVal &= "I"'⌺
                            Case eUnitType.Infantry_Regular_Airmobile : TempVal &= "I"'⌺
                            Case eUnitType.Infantry_Scout_Airmobile : TempVal &= "I"'⌺
                            Case eUnitType.Armor_Light : TempVal &= "V"'⌻
                            Case eUnitType.Armor_Heavy : TempVal &= "V"'⌼
                            Case eUnitType.Artillery : TempVal &= "Y"'◉
                            Case eUnitType.Support : TempVal &= "S"'☩
                            Case eUnitType.Monolith_Jumpship, eUnitType.Star_Lord_Jumpship, eUnitType.Invader_Jumpship, eUnitType.Merchant_Jumpship, eUnitType.Scout_Jumpship : TempVal &= "J" 'ǂ
                            Case Else : TempVal &= "D" 'Ọ  Must be a dropship
                        End Select
                    Next
                    TempState = " ("
                    For Each S As cSquad In L.Squads.Values
                        Select Case S.State
                            Case eSquadState.Wounded : TempState &= "⛨"
                            Case eSquadState.Killed : TempState &= "☠"
                            Case eSquadState.Active : TempState &= "☐" '☑
                        End Select
                    Next
                    TempState &= ") "
                    If TempState.Replace("☐", "") <> " () " Then TempVal &= TempState

                Next
                TempVal &= vbCrLf '1 line per company
            Next
        Next
        Return TempVal.Trim.Replace("Company Company", "Company").Replace("Battalion Battalion", "Battalion").Replace("Regiment Regiment", "Regiment")
        '
        'Ȼ C-Bills


    End Function


    '█ <=90%
    '░ <=75%
    '▒ <=50%
    '▓ <=25%

    '◕◒◔ Supply
    'Ȼ Salaries
    '☹ Morale/Mutiny
    '⛨  Wounded

    '
    '†ΔΦθπϮ፨ỻ
    'ᗐᗑᗒᗓᗔᗕᗖᗗᗘᗙᗚᗛ
    '⌹⌽⌾⍃⍄⍣◈
    '▶▷▼▽◀◁▲△
    '◆◇
    '◍◎●
    '◐◑◓
    '◧◨◩◪
    '☠☢☣🔧
    '⛟♥
    '⚀⚁⚂⚃⚄⚅

    Public Function ShowRegimentStrength(R As cRegiment) As String
        Dim TempVal As String = "" 'R.Name & " "
        For Each B As cBattalion In R.Battalions.Values
            For Each C As cCompany In B.Companies.Values
                For Each L As cLance In C.Lances.Values
                    Dim Count_Active As Int32 = 0
                    Dim Count_Wounded As Int32 = 0
                    Dim Count_KIA As Int32 = 0
                    For Each S As cSquad In L.Squads.Values
                        Select Case S.State
                            Case eSquadState.Wounded : Count_Wounded += 1
                            Case eSquadState.Killed : Count_KIA += 1
                            Case eSquadState.Active : Count_Active += 1
                        End Select
                    Next
                    Select Case (Count_Wounded + Count_KIA) / Count_Active
                        Case <= 0.25 : TempVal &= "▓"
                        Case <= 0.5 : TempVal &= "▒"
                        Case <= 0.75 : TempVal &= "░"
                        Case <= 0.9 : TempVal &= "█"
                    End Select
                Next
            Next
        Next
        If R.BackPayMonths > 0 Then TempVal &= "Ȼ"
        If R.SupplyMonthsMissed > 0 Then TempVal &= "🔧"
        If R.MonthsWithoutOverhead > 0 Then TempVal &= "⚔"
        Return TempVal

        '◕◒◔ Supply
        'Ȼ Salaries
        '⚔ Morale/Mutiny

        '
        '†ΔΦθπϮ፨ỻ
        'ᗐᗑᗒᗓᗔᗕᗖᗗᗘᗙᗚᗛ
        '⌹⌽⌾⍃⍄⍣◈
        '▶▷▼▽◀◁▲△
        '◆◇
        '◍◎●
        '◐◑◓
        '◧◨◩◪
        '☠☢☣
        '⛟


    End Function

    Public EventList As New List(Of String)
    Public Sub DebugEvent(Data As String, Optional Logging As Boolean = True)
        Data = CurrentTurn.Year & Format(CurrentTurn.Month, "00") & Format(CurrentTurn.Day, "00") & ": " & Data
        Debug.Print(Data)
        If Logging Then EventList.Add(Data)
    End Sub

End Module