Public Module mEnums
    Public Enum eExperience
        Really_Green = 0
        Green = 1
        Regular = 2
        Veteran = 3
        Elite = 4
        Heroic = 5
        Legendary = 6
    End Enum
    Public Enum eQuality
        LikeNew = 0
        Salvaged = 1 '40% Cost
        Destroyed = 2
    End Enum
    Public Enum eUnitType
        Mech_Light
        Mech_Medium
        Mech_Heavy
        Mech_Assault
        Fighter_Light
        Fighter_Medium
        Fighter_Heavy
        LAM_Light
        LAM_Medium
        Infantry_Regular
        Infantry_Motorized
        Infantry_Jump
        Infantry_Scout
        Infantry_Regular_Airmobile
        Infantry_Motorized_Airmobile
        Infantry_Jump_Airmobile
        Infantry_Scout_Airmobile
        Armor_Light
        Armor_Heavy
        Artillery
        Aircraft
        Support
        Leopard_Dropship
        Union_Dropship
        Overlord_Dropship
        Fury_Dropship
        Gazelle_Dropship
        Seeker_Dropship
        Triumph_Dropship
        Condor_Dropship
        Excaliber_Dropship
        Scout_Jumpship
        Invader_Jumpship
        Monolith_Jumpship
        Star_Lord_Jumpship
        Merchant_Jumpship

        'Not supported yet:
        Avenger_Dropship
        Leopard_CV_Dropship
        Intruder_Dropship
        Buccaneer_Dropship
        Achilles_Dropship
        Monarch_Dropship
        Fortress_Dropship
        Vengeance_Dropship
        Mule_Dropship
        Mammoth_Dropship
        Behemoth_Dropship
    End Enum
    Public Enum eMission
        Cadre_Duty
        Garrison_Duty
        Security_Duty
        Riot_Duty
        Siege_Duty
        Relief_Duty
        Planetary_Assault
        Offensive_Campaign
        Defensive_Campaign
        Recon_Raid
        Objective_Raid
        Divisionary_Raid
        Guerrilla_Warfare
        Retainer
    End Enum
    Public Enum eEvent
        None
        Civil_Disturbance
        Sporadic_Uprisings
        Rebellion
        Betrayal 'Use table on p64
        Treachery
        Logistics_Failure
        Reinforcements
        Attrition
        'Major Events:
        Internal_Dissension
        Armistice
        Change_of_Allegiance
        ComStar_Activity
        Periphery_Contact
        Major_Campaign
        Technological_Advance
        Star_League_Facility
        Fall_of_Major_World
        Death_of_Major_Personage
        Change_in_House_House_Relationship
    End Enum
    Public Enum eFaction
        Davion
        Kurita
        Liao
        Steiner
        Marik
        ComStar
        Other
    End Enum
    Public Enum eMutiny
        A_Morale_Loss = 0
        B_Desertions = 1
        C_Disobedient_Troops = 2
        D_Unsanctioned_Attack = 3
        E_Plunder_Locals = 4
        F_Riot = 5
        G_Treachery = 6
        H_Assassination = 7
    End Enum
    Public Enum eGuarantee
        Prior_Payment         'Full payment before mission
        Advance_Completion    'Some percent before mission, rest after
        Minor_Intermediary    '2d6-2% advance, rest after
        ComStar_Intermediary  'Can pull up to 25% in advance, but looses 5% total fee to comstar (and employer pays +5%)
        Payment_On_Completion 'Full payment after mission completion
    End Enum
    Public Enum eCommandRights
        Independent     'Merc unit does its own thing
        Mercenary       'Incorporated with larger force, but under command of independent mercenary officer
        Liaison         'Attached to a higher command, with employer-appointed liaisono
        House           'Merc commander directly answerable to house officer
        Integrated      'Integrated with house unit, and house troops/officers may replace/augment mercs
    End Enum
    Public Enum eSalvageRights
        Merc_Claims                     'All salvage belongs to merc
        Payment_In_Kind                 'All salvage belongs to merc, but employer deducts that salvage from their pay (or if paid in full, they can purchase this).
        Prize_Court_Outright_Grant      'After 2d6 Months using the prize court table they get % of salvage back
        Prize_Court_Payment_In_Kind     'After 2d6 Months using the prize court table they get % of payment based on salvage value
        Employer_Compensation           'All salvage belongs to employer, but they pay mercs for the slavage using Employer Award column
        Employer_Claims                 'All salvage belongs to employer
    End Enum
    Public Enum eForceComposition
        Air
        Aerospace
        Mech
        Armor
        Infantry
        Militia
    End Enum
    Public Enum eSquadState
        Active = 0
        Wounded = 1
        Killed = 2
    End Enum
    Public Enum eForceType
        Player
        Friendly
        Enemy
    End Enum
    Public Enum eShipLocation
        Nearest_Spaceport
        In_Orbit
        Elsewhere_In_System
        Neighboring_System
    End Enum

    Public Enum eAerospaceOperationsResult
        ES = -1 'Enemy Superiority
        A = 0 'Attrition
        FS = 1 'Friendly Superiority
    End Enum
    Public Enum eManeuverOperationsResult
        NP 'No Progress
        Att 'Attrition
        Sk 'Skirmish
        Bat 'Battle
        DecF 'Decisive Friendly
        DecE 'Decisive Enemy
        Con 'Continued Campaign
    End Enum

    Public Enum eDragoonRating
        A_Plus
        A
        A_Minus
        B_Plus
        B
        B_Minus
        C_Plus
        C
        C_Minus
        D_Plus
        D
        D_Minus
        F
    End Enum
End Module
