Public Class cCharacter
    Public Name As String
    Public Rank As String
    Public Faction As eFaction
    Public BOD As Int32 = 0
    Public DEX As Int32 = 0
    Public LRN As Int32 = 0
    Public CHA As Int32 = 0
    Public Ability_Thick_Skin As Boolean = False
    Public Ability_Glass_Jaw As Boolean = False
    Public Ability_Peripheral_Vision As Boolean = False
    Public Ability_Sixth_Sense As Boolean = False
    Public Ability_Family_Friend As Boolean = False
    Public Ability_Family_Feud As Boolean = False
    Public Ability_Handedness As eHandedness = eHandedness.Right
    Public Ability_Natural_Aptitude As eSkill = eSkill.None
    Public Skills As New Dictionary(Of eSkill, Int32)
    Public XP_Total As Int32 = 0
    Public XP_Availabile As Int32 = 0
    Public Title As String
    Public PrimaryMech As Int32 = 0
    Public PrimaryNud As Int32 = 0
    Public BackupMech As Int32 = 0
    Public BackupNud As Int32 = 0
    Public Fighter As Int32 = 0
    Public FighterNud As Int32 = 0
    Public ScoutVehicle As String = ""

    Public Academy As Boolean = False
    Public University As Boolean = False
    Public Informants As Int32 = 0
    Public Informants2 As Int32 = 0
    Public UsefulContacts As Int32 = 0
    Public UsefulContacts2 As Int32 = 0
    Public ProminentContacts As Int32 = 0
    Public ProminentContacts2 As Int32 = 0
    Public Career As String

    Public Sub New()
        Skills(eSkill.Athletics) = 0
        Skills(eSkill.Bow_Blade) = 0
        Skills(eSkill.Brawling) = 0
        Skills(eSkill.Computer) = 0
        Skills(eSkill.Diplomacy) = 0
        Skills(eSkill.Driver) = 0
        Skills(eSkill.Engineering) = 0
        Skills(eSkill.Gunnery_Aerospace) = 0
        Skills(eSkill.Gunnery_Artillery) = 0
        Skills(eSkill.Gunnery_Mech) = 0
        Skills(eSkill.Interrogation) = 0
        Skills(eSkill.JumpShip_Piloting_Navigation) = 0
        Skills(eSkill.Land_Management) = 0
        Skills(eSkill.Leadership) = 0
        Skills(eSkill.Mechanical) = 0
        Skills(eSkill.Medical_First_Aid) = 0
        Skills(eSkill.Piloting_Aerospace) = 0
        Skills(eSkill.Piloting_Mech) = 0
        Skills(eSkill.Pistol) = 0
        Skills(eSkill.Rifle) = 0
        Skills(eSkill.Rogue) = 0
        Skills(eSkill.Streetwise) = 0
        Skills(eSkill.Survival) = 0
        Skills(eSkill.Tactics) = 0
        Skills(eSkill.Technician) = 0
    End Sub

    Public Sub Create(NVP() As String)
        Try
            Dim I As Int32 = 0 'Start with 1, as 0 is the 'leaderdata' value that shows this is the leader
            I += 1 : Name = NVP(I)
        I += 1 : Rank = NVP(I)

        I += 1
        Select Case NVP(I).ToLower
            Case "davion" : Faction = eFaction.Davion
            Case "steiner" : Faction = eFaction.Steiner
            Case "marik" : Faction = eFaction.Marik
            Case "liao" : Faction = eFaction.Liao
            Case "comstar" : Faction = eFaction.ComStar
            Case "kurita" : Faction = eFaction.Kurita
            Case Else : Faction = eFaction.Other
        End Select

        I += 1 : BOD = CInt(NVP(I))
        I += 1 : DEX = CInt(NVP(I))
        I += 1 : LRN = CInt(NVP(I))
        I += 1 : CHA = CInt(NVP(I))
        I += 1 : If NVP(I) = "Ability_Thick_Skin" Then Ability_Thick_Skin = True
        I += 1 : If NVP(I) = "Ability_Glass_Jaw" Then Ability_Glass_Jaw = True
        I += 1 : If NVP(I) = "Ability_Peripheral_Vision" Then Ability_Peripheral_Vision = True
        I += 1 : If NVP(I) = "Ability_Sixth_Sense" Then Ability_Sixth_Sense = True
        I += 1 : If NVP(I) = "Ability_Family_Friend" Then Ability_Family_Friend = True
        I += 1 : If NVP(I) = "Ability_Family_Feud" Then Ability_Family_Feud = True

        I += 1
        Select Case NVP(I)
            Case "Right" : Ability_Handedness = eHandedness.Right
            Case "Left" : Ability_Handedness = eHandedness.Left
            Case "Natural_Right" : Ability_Handedness = eHandedness.Natural_Right
            Case "Natural_Left" : Ability_Handedness = eHandedness.Natural_Left
            Case "Either" : Ability_Handedness = eHandedness.Either
            Case "Both" : Ability_Handedness = eHandedness.Both
        End Select

        I += 1
        Select Case NVP(I)
            Case "Athletics" : Ability_Natural_Aptitude = eSkill.Athletics
            Case "Bow_Blade" : Ability_Natural_Aptitude = eSkill.Bow_Blade
            Case "Brawling" : Ability_Natural_Aptitude = eSkill.Brawling
            Case "Computer" : Ability_Natural_Aptitude = eSkill.Computer
            Case "Diplomacy" : Ability_Natural_Aptitude = eSkill.Diplomacy
            Case "Driver" : Ability_Natural_Aptitude = eSkill.Driver
            Case "Engineering" : Ability_Natural_Aptitude = eSkill.Engineering
            Case "Gunnery_Aerospace" : Ability_Natural_Aptitude = eSkill.Gunnery_Aerospace
            Case "Gunnery_Artillery" : Ability_Natural_Aptitude = eSkill.Gunnery_Artillery
            Case "Gunnery_Mech" : Ability_Natural_Aptitude = eSkill.Gunnery_Mech
            Case "Interrogation" : Ability_Natural_Aptitude = eSkill.Interrogation
            Case "JumpShip_Piloting_Navigation" : Ability_Natural_Aptitude = eSkill.JumpShip_Piloting_Navigation
            Case "Land_Management" : Ability_Natural_Aptitude = eSkill.Land_Management
            Case "Leadership" : Ability_Natural_Aptitude = eSkill.Leadership
            Case "Mechanical" : Ability_Natural_Aptitude = eSkill.Mechanical
            Case "Medical_First_Aid" : Ability_Natural_Aptitude = eSkill.Medical_First_Aid
            Case "Piloting_Aerospace" : Ability_Natural_Aptitude = eSkill.Piloting_Aerospace
            Case "Piloting_Mech" : Ability_Natural_Aptitude = eSkill.Piloting_Mech
            Case "Pistol" : Ability_Natural_Aptitude = eSkill.Pistol
            Case "Rifle" : Ability_Natural_Aptitude = eSkill.Rifle
            Case "Rogue" : Ability_Natural_Aptitude = eSkill.Rogue
            Case "Streetwise" : Ability_Natural_Aptitude = eSkill.Streetwise
            Case "Survival" : Ability_Natural_Aptitude = eSkill.Survival
            Case "Tactics" : Ability_Natural_Aptitude = eSkill.Tactics
            Case "Technician" : Ability_Natural_Aptitude = eSkill.Technician
                Case Else 'Unknown or more likely blank or none, skip
            End Select



        I += 1 : XP_Total = CInt(NVP(I))
        I += 1 : XP_Availabile = CInt(NVP(I))
        I += 1 : PrimaryMech = CInt(NVP(I))
        I += 1 : PrimaryNud = CInt(NVP(I))
        I += 1 : BackupMech = CInt(NVP(I))
        I += 1 : BackupNud = CInt(NVP(I))
        I += 1 : Fighter = CInt(NVP(I))
        I += 1 : FighterNud = CInt(NVP(I))
        I += 1 : ScoutVehicle = NVP(I)
        I += 1 : If NVP(I) = "Academy" Then Academy = True
        I += 1 : If NVP(I) = "University" Then University = True
        I += 1 : Informants = CInt(NVP(I))
        I += 1 : Informants2 = CInt(NVP(I))
        I += 1 : UsefulContacts = CInt(NVP(I))
        I += 1 : UsefulContacts2 = CInt(NVP(I))
        I += 1 : ProminentContacts = CInt(NVP(I))
        I += 1 : ProminentContacts2 = CInt(NVP(I))
        I += 1 : Career = NVP(I)
        I += 1 : Title = NVP(I)

        I += 1 : Skills(eSkill.Athletics) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Bow_Blade) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Brawling) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Computer) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Diplomacy) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Driver) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Engineering) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Gunnery_Aerospace) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Gunnery_Artillery) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Gunnery_Mech) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Interrogation) = CInt(NVP(I))
        I += 1 : Skills(eSkill.JumpShip_Piloting_Navigation) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Land_Management) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Leadership) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Mechanical) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Medical_First_Aid) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Piloting_Aerospace) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Piloting_Mech) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Pistol) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Rifle) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Rogue) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Streetwise) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Survival) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Tactics) = CInt(NVP(I))
        I += 1 : Skills(eSkill.Technician) = CInt(NVP(I))


        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub
    Public Function Save() As String
        Dim Outline As String = ""
        Outline &= Name & vbTab
        Outline &= Rank & vbTab
        Outline &= Faction.ToString & vbTab
        Outline &= BOD & vbTab
        Outline &= DEX & vbTab
        Outline &= LRN & vbTab
        Outline &= CHA & vbTab
        Outline &= IIf(Ability_Thick_Skin, "Ability_Thick_Skin", "") & vbTab
        Outline &= IIf(Ability_Glass_Jaw, "Ability_Glass_Jaw", "") & vbTab
        Outline &= IIf(Ability_Peripheral_Vision, "Ability_Peripheral_Vision", "") & vbTab
        Outline &= IIf(Ability_Sixth_Sense, "Ability_Sixth_Sense", "") & vbTab
        Outline &= IIf(Ability_Family_Friend, "Ability_Family_Friend", "") & vbTab
        Outline &= IIf(Ability_Family_Feud, "Ability_Family_Feud", "") & vbTab
        Outline &= Ability_Handedness.ToString & vbTab
        Outline &= Ability_Natural_Aptitude.ToString & vbTab
        Outline &= XP_Total & vbTab
        Outline &= XP_Availabile & vbTab
        Outline &= PrimaryMech & vbTab
        Outline &= PrimaryNud & vbTab
        Outline &= BackupMech & vbTab
        Outline &= BackupNud & vbTab
        Outline &= Fighter & vbTab
        Outline &= FighterNud & vbTab
        Outline &= ScoutVehicle & vbTab
        Outline &= IIf(Academy, "Academy", "") & vbTab
        Outline &= IIf(University, "University", "") & vbTab
        Outline &= Informants & vbTab
        Outline &= Informants2 & vbTab
        Outline &= UsefulContacts & vbTab
        Outline &= UsefulContacts2 & vbTab
        Outline &= ProminentContacts & vbTab
        Outline &= ProminentContacts2 & vbTab
        Outline &= Career & vbTab
        Outline &= Title & vbTab

        Outline &= Skills(eSkill.Athletics) & vbTab
        Outline &= Skills(eSkill.Bow_Blade) & vbTab
        Outline &= Skills(eSkill.Brawling) & vbTab
        Outline &= Skills(eSkill.Computer) & vbTab
        Outline &= Skills(eSkill.Diplomacy) & vbTab
        Outline &= Skills(eSkill.Driver) & vbTab
        Outline &= Skills(eSkill.Engineering) & vbTab
        Outline &= Skills(eSkill.Gunnery_Aerospace) & vbTab
        Outline &= Skills(eSkill.Gunnery_Artillery) & vbTab
        Outline &= Skills(eSkill.Gunnery_Mech) & vbTab
        Outline &= Skills(eSkill.Interrogation) & vbTab
        Outline &= Skills(eSkill.JumpShip_Piloting_Navigation) & vbTab
        Outline &= Skills(eSkill.Land_Management) & vbTab
        Outline &= Skills(eSkill.Leadership) & vbTab
        Outline &= Skills(eSkill.Mechanical) & vbTab
        Outline &= Skills(eSkill.Medical_First_Aid) & vbTab
        Outline &= Skills(eSkill.Piloting_Aerospace) & vbTab
        Outline &= Skills(eSkill.Piloting_Mech) & vbTab
        Outline &= Skills(eSkill.Pistol) & vbTab
        Outline &= Skills(eSkill.Rifle) & vbTab
        Outline &= Skills(eSkill.Rogue) & vbTab
        Outline &= Skills(eSkill.Streetwise) & vbTab
        Outline &= Skills(eSkill.Survival) & vbTab
        Outline &= Skills(eSkill.Tactics) & vbTab
        Outline &= Skills(eSkill.Technician) & vbTab

        Return Outline
    End Function

End Class

Public Enum eHandedness
    Right
    Left
    Natural_Right
    Natural_Left
    Either
    Both
End Enum
Public Enum eSkill
    None
    Athletics
    Bow_Blade
    Brawling
    Computer
    Diplomacy
    Driver
    Engineering
    Gunnery_Aerospace
    Gunnery_Artillery
    Gunnery_Mech
    Interrogation
    JumpShip_Piloting_Navigation
    Land_Management
    Leadership
    Mechanical
    Medical_First_Aid
    Piloting_Aerospace
    Piloting_Mech
    Pistol
    Rifle
    Rogue
    Streetwise
    Survival
    Tactics
    Technician
End Enum