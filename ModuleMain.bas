Attribute VB_Name = "ModuleMain"
'Type pour les caractéristiques d'une station météo
'de référence pour la vérif au gel
Type StationMétéo
    monNom As String * 20
    monNumDpt As String * 3  'Numéro de département
    monAltitude As Integer
    monHRE As Integer       'Hiver Rigoureux Exceptionnel
    monHRNE As Integer      'Hiver Rigoureux Non Exceptionnel
    monHC As Integer        'Hiver Courant
End Type

'Type pour les options sur les matériaux de base
'ou de fondation (cf Fenêtre VB frmOptionsMat)
Type OptionsMat
    monFichPersoSTR As String 'Fichier de structures personnelles
    mesMatPersoNonAutorisés As String 'String contenant les indices des
        'matériaux non autorisés avec un blanc à la fin comme séparateur
        'dans la collection de matériaux de base/fondation personnels
    mesMatCERTUNonAutorisés As String 'String contenant les indices des
        'matériaux non autorisés avec un blanc à la fin comme séparateur
        'dans la collection de matériaux de base/fondation du CERTU
End Type

'Type pour les options générales (cf Fenêtre VB frmOptionsGen)
Type OptionsGen
    maDuréeService As Byte
    maCroisAnnuel As Byte
    maCoulSurf As Long
    maCoulBase1 As Long
    maCoulBase2 As Long
    maCoulFond1 As Long
    maCoulFond2 As Long
    maVerifGel As Byte
    monIndStationRef As Byte 'Indice de la station météo de référence
    'Champs suivants pour les caractéristiques d'une agglomération
    'où l'on veut faire une vérif au gel
    monAltiAgglo As Integer
    monTailleAgglo As Integer '0 ==> Inf à 100 000 Hab.
                              '1 ==> Entre 100 000 et 1 000 000 Hab.
                              '2 ==> Sup à 1 000 000 Hab.
    monCoefAgglo As Single    '1 ==> Inf à 100 000 Hab.
                              '0.9 ==> Entre 100 000 et 1 000 000 Hab.
                              '0.8 ==> Sup à 1 000 000 Hab.
End Type

'Collection contenant toutes les données lues dans un fichier .urb
Public maColLectFich As New Collection

'Variable indiquant si l'ouverture de la fenetre
'a été OK (Form_Initialize event sans erreur)
Public monOuverture As Boolean

'Variable contenant les stations météo de références
Public monTabStation(1 To 84) As StationMétéo
Public monNbStation As Integer 'Valeur initiale = 0

'Variable contenant les options générales
Public mesOptionsGen As OptionsGen
'Variable contenant les options matériaux
Public mesOptionsMat As OptionsMat

'Variables contenant les collections contenant
'les matériaux de surface
Public maColMatSurf As New Collection
Public maColMatComposant As New Collection

'Variables contenant les collections contenant les structures
'et les matériaux de base et de fondation du CERTU
Public maColStructCERTU As New Collection
Public maColMatBFCERTU As New Collection

'Variables contenant les collections contenant les structures
'et les matériaux de base et de fondation personnels
Public maColStructPerso As New Collection
Public maColMatBFPerso As New Collection

'Variable globale contenant la fenêtre mère MDI
Public fMainForm As frmMain
'Variable globale contenant la fenêtre fille de l'étude en cours
Public monEtude As Form

'Variable indiquant si on ouvre une nouvelle étude
Public maNewEtude As Boolean

'Constante donnant l'indice de la station de LYON
'dans le tableau des stations
Public Const IndiceStationLYON As Integer = 65

'Constante donnant l'épaisseur totale maximale des carottes Q1 et Q2
Public Const EpTotMaxEcran As Integer = 3575
'c'est 65 *55, 65 twips pour 1 cm, épaisseur max totale 55 cm

'Constantes d'entêtes de fichier MTS et STR
Public Const ENTETE_MTS As String = "Fichier de matériaux de surfaces"
Const ENTETE_STR_v100 As String = "Fichier de structures de chaussées"
'Compatible avec les versions Struct-urb <= 1.00.0002
Const ENTETE_STR_v103 As String = "Fichier de structures de chaussées pour Struct-Urb version >= 1.00.0003"
Const ENTETE_STR_v200 As String = "Fichier de structures de chaussées pour Struct-urb version >= 2.0.0"
'Constantes indiquant la fin du commentaire RTF
Public Const FIN_COMMENT As String = "###Fin commentaire###"

'Constantes pour le type de voie
Public Const TypeVoieInconnu As Integer = 0
Public Const TypeVoieDesserte As Integer = 1
Public Const TypeVoieDistribution As Integer = 2
Public Const TypeVoieTraficLourd As Integer = 3
Public Const TypeVoieBus As Integer = 4
Public Const TypeVoieParking As Integer = 5

Public Const TypeGiratoireDistribution As Integer = 6
Public Const TypeGiratoireTraficLourd As Integer = 7

'Constantes pour le type d'onglet
Public Const OngletVoie As Integer = 0
Public Const OngletTrafic As Integer = 1
Public Const OngletStruct As Integer = 3
Public Const OngletCAM As Integer = 4
Public Const OngletPF As Integer = 2
Public Const OngletSurf As Integer = 5
Public Const OngletGel As Integer = 6

'Constantes pour l'épaisseur par défaut en cm dans une carotte
Public Const EpParDefaut As Integer = 10

'Constantes pour l'épaisseur du lit de pose en cm dans une carotte
Public Const EpLitPose As Integer = 4

'Constantes les hivers de référence
Public Const HE As Integer = 1
Public Const HRNE As Integer = 2
Public Const HC As Integer = 3

'Constantes pour le type de condition de chantier
Public Const TypeChantierQ1 As Integer = 1
Public Const TypeChantierQ2 As Integer = 2

'Constantes pour le type d'étude
Public Const TypeEtudeStandard As Integer = 1
Public Const TypeEtudeGiratoire As Integer = 2

'Constantes pour le sol support
Public Const TresGelif As Integer = 1
Public Const PeuGelif As Integer = 2
Public Const NonGelif As Integer = 3

'Constantes pour la couche de forme non gélive
Public Const NonTraité As Integer = 1
Public Const Traité As Integer = 2

'Constantes pour les formats de fichiers *.urb de Struct-Urb
'Pour la version finale, rajout d'une ligne contenant l'indice de gel perso
'et l'état de la case à cocher correspondante dans l'onglet gel
'Pour la version béta = sites pilotes cette ligne n'existe pas
Public Const FormatFichierVersionBeta As Byte = 1
Public Const FormatFichierVersionFinale As Byte = 2

'Constante donnant une valeur entière à rajouter au type de voie stocké dans
'le fichier *.urb lors de la sauvegarde pour savoir si on est en qualité de
'chantier Standard(Q1) (typevoie < 100) ou difficile (Q2) (typevoie > 100)
Public Const ChantierDifficile As Byte = 100

'Constante pour marquer la fin du titre de l'étude dans un fichier *.urb
Public Const FinTitre As String = "###FinTitre###"

'Constante pour indiquer chaussée hors gel
Public Const HorsGel As Long = 1000000000

'Constantes indiquant le type de structures
Public Const ToutType As Byte = 0
Public Const Souple As Byte = 1
Public Const Bitumineuse As Byte = 2
Public Const GTLH As Byte = 3
Public Const Beton As Byte = 4
Public Const Mixte As Byte = 5
Public Const PavesDalles As Byte = 6

'Constante indiquant que l'on a changé de type de structure et
'que la liste déroulante de l'onglet structure permettant d'en choisir
'une. On prend -2 car les index vont de 0 à NbElementsListe-1 et -1 sert à dire
'qu'il n'y a rien de sélectionner dans une combobox vb
Public Const ChangeTypeStruct As Integer = -2

'Constante pour les id de l'aide
Public Const IDhlp_VerifGel As Integer = 118 ' ch01s12.htm Partie Aide sur la vérif au gel
Public Const IDhlp_OngletVoie As Integer = 215 'ch02s07s01 Onglet Voie
Public Const IDhlp_OngletTrafic As Integer = 216 'ch02s07s02 Onglet Trafic
Public Const IDhlp_OngletStructure As Integer = 218 'ch02s07s04 Onglet Structure
Public Const IDhlp_OngletCAM As Integer = 219 'ch02s07s05 Onglet CAM
Public Const IDhlp_OngletPlateForme As Integer = 217 'ch02s07s03 Onglet PlateForme
Public Const IDhlp_OngletCoucheSurf As Integer = 220 'ch02s07s06 Onglet Couche de surface
Public Const IDhlp_OngletGel As Integer = 221 'ch02s07s07 Onglet Gel

Public Const IDhlp_WinAbout As Integer = 236 'ch02s11s04 Fenêtre a propos
Public Const IDhlp_WinPrint As Integer = 210 'ch02s04s06 Fenêtre Impression
Public Const IDhlp_WinOptionsGen As Integer = 226 'ch02s09s01 Fenêtre Options générales
Public Const IDhlp_WinOptionsMat As Integer = 72 ' Fenêtre Options Matériaux

Public Const IDhlp_NewSite As Integer = 205 'ch02s04s01 menu nouveau
Public Const IDhlp_OpenSite As Integer = 206 'ch02s04s02 menu ouvrir
Public Const IDhlp_SaveSite As Integer = 208 'ch02s04s04 menu sauver
Public Const IDhlp_SaveAsSite As Integer = 209 'ch02s04s05 menu sauver sous
Public Const IDhlp_CloseSite As Integer = 207 'ch02s04s03 menu fermer

Sub Main()

    'Récup du séparateur décimale . ou ,
    'fixé dans les paramètres régionaux de Windows
    TrouverCaractèreDécimalUtilisé
    
    'Chargement du fichier structure du CERTU
    If OuvrirFichierMatSurface(App.Path + "\CERTU.MTS", maColMatSurf) = False Then
        'Erreur on vide et on sort
        ViderCollection maColMatSurf
    Else
        '********************************
        'test Protection
        '********************************
        'Type de protection
        TYPPROTECTION = CPM
    
        ' Vérification de l'enregistrement
        If ProtectCheck("its00+-k") = "its00+-k" Then
        ' Affichage de la feuille principale
        'Lancement de la fenetre MDI mère de l'application
            Set fMainForm = New frmMain
            fMainForm.Show
        Else 'la licence n'a pas été validée on ferme
            End
        End If
        '********************************
    End If
End Sub


Public Sub FermerFenetre(uneForm As Form)
    'Permet de fermer la fenêtre passée en paramètres
    'et de remettre au premier plan la fenêtre fille courante
    'Correction d'un bug dans la gestin des fenêtre actives
    'Windows si on ouvre un sélectionneur de fichier, d'imprimante
    'ou de fontes.
    Unload uneForm
    fMainForm.Show
End Sub

Public Sub CentrerFenetreEcran(uneForm As Form)
    'Centrage d'une fenetre (= une Form VB) à l'écran
    uneForm.Top = (Screen.Height - uneForm.Height) / 2
    uneForm.Left = (Screen.Width - uneForm.Width) / 2
End Sub
 
Public Sub TrouverCoefCorrecteurAgglo(unTypeTaille As Integer)
    'Affectation du coefficient correcteur de l'Agglo
    'pour la vérif au gel suivant son type de taille
    If unTypeTaille = 0 Then
        '< à 100 000 Hab.
        mesOptionsGen.monCoefAgglo = 1
    ElseIf unTypeTaille = 1 Then
        'entre 100 000 et 1 000 000 Hab.
        mesOptionsGen.monCoefAgglo = 0.9
    ElseIf unTypeTaille = 2 Then
        '> à 1 000 000 Hab.
        mesOptionsGen.monCoefAgglo = 0.8
    Else
        MsgBox MsgErreurProg + MsgErreurTailleAgglo + MsgIn + "TrouverCoefCorrecteurAgglo", vbCritical
    End If
End Sub

Public Sub RécupérerOptionsGen()
    'Récupération des options générales par lecture des valeurs
    'de ces options stockées dans la base de registre
    
    mesOptionsGen.maDuréeService = GetSetting(App.Title, "OptionsGen", "DuréeService", 20)
    mesOptionsGen.maCroisAnnuel = GetSetting(App.Title, "OptionsGen", "CroissAnnuel", 1)
    mesOptionsGen.maCoulSurf = GetSetting(App.Title, "OptionsGen", "CouleurSurf", QBColor(11))
    mesOptionsGen.maCoulBase1 = GetSetting(App.Title, "OptionsGen", "CouleurBase1", QBColor(2))
    mesOptionsGen.maCoulBase2 = GetSetting(App.Title, "OptionsGen", "CouleurBase2", QBColor(10))
    mesOptionsGen.maCoulFond1 = GetSetting(App.Title, "OptionsGen", "CouleurFond1", QBColor(12))
    mesOptionsGen.maCoulFond2 = GetSetting(App.Title, "OptionsGen", "CouleurFond2", QBColor(13))
    
    mesOptionsGen.maVerifGel = GetSetting(App.Title, "OptionsGen", "VerifGel", 1)
    
    'Affectation de l'Agglomération des études
    mesOptionsGen.monTailleAgglo = GetSetting(App.Title, "OptionsGen", "TailleAgglo", 2)
    mesOptionsGen.monAltiAgglo = GetSetting(App.Title, "OptionsGen", "AltiAgglo", 200)
    TrouverCoefCorrecteurAgglo mesOptionsGen.monTailleAgglo
    
    'Affectation de la station de référence par défaut
    mesOptionsGen.monIndStationRef = GetSetting(App.Title, "OptionsGen", "StationRef", IndiceStationLYON)
End Sub

Public Sub RécupérerOptionsMat()
    'Récupération des options matériaux par lecture des valeurs
    'de ces options stockées dans la base de registre
    
    'Affectation du nom de fichier de structures personnelles
    mesOptionsMat.monFichPersoSTR = GetSetting(App.Title, "OptionsMat", "FichierPersoSTR", "")
    
    'Récupération de la chaine de caractères contenant les indices
    'des matériaux de base et de fondation non autorisés avec un blanc à
    'la fin comme séparateur (une dizaine de matériaux au maximum)
    'Au début chaine vide car tous autorisés
    mesOptionsMat.mesMatCERTUNonAutorisés = GetSetting(App.Title, "OptionsMat", "MatCERTUNonAutorisés", "")
    mesOptionsMat.mesMatPersoNonAutorisés = GetSetting(App.Title, "OptionsMat", "MatPersoNonAutorisés", "")
End Sub

Public Sub StockerOptionsMat()
    'Stockage des options matériaux dans la base de registre
    
    'Stockage du nom de fichier de structures personnelles
    SaveSetting App.Title, "OptionsMat", "FichierPersoSTR", mesOptionsMat.monFichPersoSTR
    
    'stockage de la chaine de caractères contenant les indices
    'des matériaux de base et de fondation non autorisés avec un blanc à
    'la fin comme séparateur (une dizaine de matériaux au maximum)
    SaveSetting App.Title, "OptionsMat", "MatCERTUNonAutorisés", mesOptionsMat.mesMatCERTUNonAutorisés
    SaveSetting App.Title, "OptionsMat", "MatPersoNonAutorisés", mesOptionsMat.mesMatPersoNonAutorisés
End Sub

Public Sub StockerOptionsGen()
    'Stockage des options générales dans la base de registre
    With mesOptionsGen
        SaveSetting App.Title, "OptionsGen", "DuréeService", .maDuréeService
        SaveSetting App.Title, "OptionsGen", "CroissAnnuel", .maCroisAnnuel
        
        SaveSetting App.Title, "OptionsGen", "CouleurSurf", .maCoulSurf
        SaveSetting App.Title, "OptionsGen", "CouleurBase1", .maCoulBase1
        SaveSetting App.Title, "OptionsGen", "CouleurBase2", .maCoulBase2
        SaveSetting App.Title, "OptionsGen", "CouleurFond1", .maCoulFond1
        SaveSetting App.Title, "OptionsGen", "CouleurFond2", .maCoulFond2
        
        SaveSetting App.Title, "OptionsGen", "VerifGel", .maVerifGel
        SaveSetting App.Title, "OptionsGen", "TailleAgglo", .monTailleAgglo
        SaveSetting App.Title, "OptionsGen", "AltiAgglo", .monAltiAgglo
        SaveSetting App.Title, "OptionsGen", "StationRef", .monIndStationRef
    End With
End Sub

Public Sub ChoisirCouleur(unePicCouleur As PictureBox)
    'Choix de la couleur parmi les couleurs systèmes disponibles
    'pour la PictureBox passée en paramètre
    With fMainForm
          ' Attribue à CancelError la valeur True
          .dlgCommonDialog.CancelError = True
          On Error GoTo ErrHandler
          ' Définit la propriété Flags
          .dlgCommonDialog.Flags = cdlCCRGBInit
          ' Affiche la boîte de dialogue Couleur
          .dlgCommonDialog.ShowColor
          ' Attribue à l'arrière-plan de la feuille la
          ' couleur sélectionnée
          unePicCouleur.BackColor = .dlgCommonDialog.Color
    End With
      
    Exit Sub

ErrHandler:
    ' L'utilisateur a cliqué sur Annuler
    'On ne fait rien
End Sub

Public Function OuvrirFichierStructures(unFileName As String, uneColStruct As Collection, uneColMat As Collection) As Boolean
    'Ouverture d'un fichier structure et remplissage
    'de la collection de structures passée en paramètre
    'et de la collection de matériaux de couche de base ou fondation
    'passée en paramètre
    
    'Retourne :
    '       - TRUE si pas d'erreur à la lecture
    '       - FALSE sinon
    
    Dim unFichSTR As Integer
    Dim unTypeMat As String, unMat As Object
    Dim unEntete As String, unNom As String, unAbrege As String
    Dim unComment As String, uneNorme As String
    Dim unCommentSuite As String, uneCoucheSurfSansEp As Integer
    Dim uneStruct As Structure, uneColPF As Collection
    Dim unYoung As Single, unPoisson As Single, i As Integer
    Dim unEpsilon As Single, unSigma As Single
    Dim uneCouSurf As String, uneCouBase As String, uneCouFond As String
    Dim unUtilVDes As Integer, unUtilVDis As Integer
    Dim unUtilVPL As Integer, unUtilVBus As Integer
    Dim unTauxRisque As Integer, unTypeCAM As Integer
    Dim unNbEssieuxMin As Long, unNbEssieuxMax As Long
    Dim uneEpSurf As Integer, unMaxPratique As String
    Dim uneEpBase1 As Integer, uneEpBase2 As Integer
    Dim uneEpFond1 As Integer, uneEpFond2 As Integer
    Dim unQm As Single, unMinTec As String
    Dim uneQual As String, unAGel As Single, unBGel As Single
    Dim uneSaisieComplète As Boolean, unNbTot As Long, unNumIndex As Integer
    Dim uneSizeFichSTR As Long, uneWinWait As frmWaitBar
    Dim unTypeChaussee As Byte, unTypeStructure As Byte, unUtilVParking As Integer
    Dim unUtilGDis As Integer, unUtilGPL As Integer
    
    'Correction de \\ par un seul \
    unFileName = CorrigerNomFichier(unFileName)
    'Calcul de la taille du fichier structure
    uneSizeFichSTR = FileLen(unFileName)
    
    'Ouverture de la fenêtre d'atente pour charger les structures
    'nouveau depuis la version 2, en réseau cela permet de patienter
    'lors de l'ouverture de Struct-Urb
    Set uneWinWait = New frmWaitBar
    uneWinWait.LabelWait = LabelWaitLoadStructures
    uneWinWait.Show
    DoEvents
    
    'Démarrage de la gestion des erreurs
    On Error GoTo erreurfichier_str
    unFichSTR = FreeFile(0)
    
    'Test de l'existence du fichier .str
    Open unFileName For Input As #unFichSTR
    'si Fichier inexistant ==> erreur 53
    'on va dans erreurfichier_str
    Close #unFichSTR 'On ferme pour faire le bon traitement
    
    Open unFileName For Binary As #unFichSTR
    'Test si fichier structures
    LireString unFichSTR, unEntete
    'If unEntete = ENTETE_STR_v100 Then
    If unEntete = ENTETE_STR_v100 Or unEntete = ENTETE_STR_v103 Then
        'Avertissement que cette version de Struct-Urb >= à 2.0.0
        'n'est pas compatible avec le fichier structure des versions <= 2.00.0000
        uneWinWait.Hide 'on cache la fenêtre d'attente
        unMsg = "Le fichier structure " + unFileName + ", de version 1.0, n'est plus compatible avec " + fMainForm.Caption
        unMsg = unMsg + Chr(13) + "Il faut réinstaller " + fMainForm.Caption + " ou contacter le Certu."
        MsgBox unMsg, vbCritical
        'Fonction retourne faux car erreur
        OuvrirFichierStructures = False
    ElseIf unEntete = ENTETE_STR_v200 Then
    'ElseIf unEntete = ENTETE_STR_v103 Then
        'Remplissage des structures et des matériaux de couches
        'de base ou de fondation pour un format de fichier structure
        'compatible avec Struct-Urb version >= 2.0.0000
        Do
            'Affichage de la progression de la lecture du fichier structure
            uneWinWait.ProgressBar1.Value = Int(Seek(unFichSTR) / (uneSizeFichSTR) * 100)
            'Lecture d 'un élément du fichier structure
            LireString unFichSTR, unTypeMat
            If unTypeMat = "Structure" Then
                'Lecture fichier pour une structure de chaussées
                Get #unFichSTR, , unNumIndex
                LireString unFichSTR, unAbrege
                LireString unFichSTR, uneCouSurf
                LireString unFichSTR, uneCouBase
                LireString unFichSTR, uneCouFond
                Get #unFichSTR, , uneCoucheSurfSansEp
                Get #unFichSTR, , uneSaisieComplète
                Get #unFichSTR, , unUtilVDes
                Get #unFichSTR, , unUtilVDis
                Get #unFichSTR, , unUtilVPL
                Get #unFichSTR, , unUtilVBus
                
                'Lecture de données supplémentaires pour le fichier structure v2
                'les nouveaux types de voies
                Get #unFichSTR, , unUtilVParking
                Get #unFichSTR, , unUtilGDis
                Get #unFichSTR, , unUtilGPL
                Get #unFichSTR, , unTypeChaussee
                
                Get #unFichSTR, , unTauxRisque
                Get #unFichSTR, , unTypeCAM
                
                'Lecture de données supplémentaires pour le fichier structure v2
                'le type de structure (souple, bitumineuse, GTLH,...)
                Get #unFichSTR, , unTypeStructure
                
                Get #unFichSTR, , unNbEssieuxMin
                Get #unFichSTR, , unNbEssieuxMax
                LireString unFichSTR, unComment
                LireString unFichSTR, unCommentSuite
                Do While unCommentSuite <> FIN_COMMENT
                    'Corrrection du problème lié à la présence
                    'de " dans le texte RTF
                    '==> début ou fin de string en lecture
                    unComment = unComment + unCommentSuite
                    LireString unFichSTR, unCommentSuite
                Loop
                
                'Création des structures
                Set uneStruct = New Structure
                uneStruct.monNumIndex = unNumIndex
                uneStruct.monAbrégé = unAbrege
                uneStruct.SetPropsInfo uneSaisieComplète, unUtilVDes, unUtilVDis, unUtilVPL, unUtilVBus, unComment, unTauxRisque, unNbEssieuxMin, unNbEssieuxMax, unTypeCAM
                uneStruct.SetComposition uneCouSurf, uneCouBase, uneCouFond, uneCoucheSurfSansEp
                
                'Stockage des infos rajoutés en version 2
                uneStruct.SetPropsInfoV2 unTypeChaussee, unTypeStructure, unUtilVParking, unUtilGDis, unUtilGPL
                
                'Lecture des données par plate-forme
                'et alimentation des col PF de la nouvelle structure
                ' on passe de 6 à 8 en version 2
                For i = 1 To 8
                    Set uneColPF = DonnerColPF(uneStruct, i)
                    Get #unFichSTR, , unNbTot
                    For j = 1 To (unNbTot \ 8) '8 colonnes de données par plate-forme
                        Get #unFichSTR, , uneEpSurf
                        Get #unFichSTR, , uneEpBase1
                        Get #unFichSTR, , uneEpBase2
                        Get #unFichSTR, , uneEpFond1
                        Get #unFichSTR, , uneEpFond2
                        Get #unFichSTR, , unQm
                        LireString unFichSTR, unMinTec
                        LireString unFichSTR, unMaxPratique
                        uneColPF.Add uneEpSurf
                        uneColPF.Add uneEpBase1
                        uneColPF.Add uneEpBase2
                        uneColPF.Add uneEpFond1
                        uneColPF.Add uneEpFond2
                        uneColPF.Add unQm
                        uneColPF.Add unMinTec
                        uneColPF.Add unMaxPratique
                    Next j
                Next i
                
                'Alimentation de la collection de structure
                uneColStruct.Add uneStruct
            
            ElseIf unTypeMat = "MatériauFondBase" Then
                'Lecture fichier pour un matériau de couche de
                'base ou de fondation
                LireString unFichSTR, unNom
                LireString unFichSTR, unAbrege
                LireString unFichSTR, uneNorme
                LireString unFichSTR, uneQual
                Get #unFichSTR, , unAGel
                Get #unFichSTR, , unBGel
                Get #unFichSTR, , unYoung
                Get #unFichSTR, , unPoisson
                Get #unFichSTR, , unEpsilon
                Get #unFichSTR, , unSigma
                LireString unFichSTR, unComment
                LireString unFichSTR, unCommentSuite
                Do While unCommentSuite <> FIN_COMMENT
                    'Corrrection du problème lié à la présence
                    'de " dans le texte RTF
                    '==> début ou fin de string en lecture
                    unComment = unComment + unCommentSuite
                    LireString unFichSTR, unCommentSuite
                Loop
                'Création des matériaux
                Set unMat = New Matériau
                unMat.SetProps unNom, unAbrege, uneNorme, unComment
                unMat.SetPropsPhysic unYoung, unPoisson, unEpsilon, unSigma
                RemplirQualitéGel unMat, uneQual, unAGel, unBGel
                'Alimentation de la collection matériau base/fondation
                uneColMat.Add unMat, unMat.monAbrégé
            ElseIf EOF(unFichSTR) = False Then
                'Cas où ce n'est pas la fin du fichier
                uneWinWait.Hide 'on cache la fenêtre d'attente
                MsgBox MsgErreurMatériauInconnu + ": " + unTypeMat + Chr(13) + MsgFichStruct + unFileName + MsgIncorrect, vbCritical
                'Sortie du programme
                OuvrirFichierStructures = False
                Close unFichSTR
                Exit Function
            End If
        Loop Until EOF(unFichSTR)
        'Valeur de retour mis à vrai car tout est OK
        OuvrirFichierStructures = True
    Else
        uneWinWait.Hide 'on cache la fenêtre d'attente
        MsgBox MsgFichStruct + unFileName + MsgIncorrect, vbCritical
        'Fonction retourne faux car erreur
        OuvrirFichierStructures = False
    End If
            
    'Fermeture du fichier structure
    Close unFichSTR
    
    'Fermeture de la fenêtre d'attente du chargement des structures
    Unload uneWinWait
    DoEvents
    
    'Sortie du programme et fin de la gestion d'erreur
    On Error GoTo 0
    Exit Function
    
erreurfichier_str:
    'Gestion d'erreur de lecture des fichiers structures *.STR
    unMsg = unMsg + Chr(13) + MsgFichStruct + unFileName + MsgRunError + Format(Err.Number) + " - " + Err.Description
    If Err.Number <> 53 Then
        'Cas d'une erreur différente de fichier introuvable
        unMsg = unMsg + Chr(13) + Chr(13) + MsgFichStruct + MsgIncorrect
    End If
    OuvrirFichierStructures = False
    'Fermeture de la fenêtre d'attente du chargement des structures
    Unload uneWinWait
    DoEvents
    'Affichage de l'erreur survenue
    MsgBox unMsg, vbCritical
    'Sortie du programme
    Close unFichSTR
    
    On Error GoTo 0
    Exit Function
End Function


Public Sub RemplirSpreadMat(unSpreadMat As vaSpread, uneColMat As Collection, uneChaineIndicesNonAutorisés As String)
    'Remplir un Spread avec les matériaux base/fondation
    'avec leur autorisation d'utilisation
    unSpreadMat.MaxRows = uneColMat.Count
    For i = 1 To uneColMat.Count
        unSpreadMat.Row = i
        'Remplissage de l'abrégé du matériau
        unSpreadMat.Col = 1
        unSpreadMat.Text = uneColMat(i).monAbrégé
        'Remplissage de son autorisation d'utilisation
        'par recherche dans une chaine contenant les indices
        'des matériaux non autorisés séparés par des blancs,
        'même le dernier est suivi d'un blanc
        unSpreadMat.Col = 2
        If uneColMat(i).monUtilisationAutorisée Then
            'Cas d'un indice de matériau autorisé
            unSpreadMat.Value = 1
        Else
            'Cas d'un indice de matériau non autorisé
            unSpreadMat.Value = 0
        End If
    Next i
End Sub

Public Sub AlimenterAutorisation(unFichierCERTU As Boolean)
    'Procédure alimentant les autorisations d'utilisation
    'des matériaux de base/fondation
    Dim uneColMat As Collection, unRes As Boolean
    Dim uneChaineIndicesNonAutorisés As String
    
    'en version 2, on ne fait plus rien ici
    Exit Sub
    
    If unFichierCERTU Then
        'Cas du fichier de structures CERTU
        Set uneColMat = maColMatBFCERTU
        uneChaineIndicesNonAutorisés = mesOptionsMat.mesMatCERTUNonAutorisés
    Else
        'Cas du fichier de structures personnelles
        Set uneColMat = maColMatBFPerso
        uneChaineIndicesNonAutorisés = mesOptionsMat.mesMatPersoNonAutorisés
    End If
    
    For i = 1 To uneColMat.Count
        unRes = InStr(1, uneChaineIndicesNonAutorisés, " " + Format(i) + " ")
        uneColMat(i).monUtilisationAutorisée = (unRes = 0)
        'En effet si unRes = 0 ==> i n'est pas dans la string uneChaineIndicesNonAutorisés
        'donc i n'est pas un indice de matériau non autorisé
    Next i
End Sub

Public Sub ActualiserFrameVerifGel(uneFrmDoc As Form)
    'Affichage ou non de la vérif au gel dans la frame
    'de la fenêtre fille passée en paramètre
    Dim uneVisu As Boolean
    
    'En version 2 bétatest on affiche que les indices de gel Ref et Admis
    'de la qualité choisie
    uneVisu = (uneFrmDoc.monTypeChantier = TypeChantierQ1)
    
    uneFrmDoc.TabData.TabEnabled(OngletGel) = (uneFrmDoc.monTypeVoie > 0) And mesOptionsGen.maVerifGel
    
    If mesOptionsGen.maVerifGel Then
        uneFrmDoc.LabelIndiceGel.Caption = LabelIGelCaption
        uneFrmDoc.LabelIndiceGel.Top = 0
    Else
        uneFrmDoc.LabelIndiceGel.Caption = MsgNoVerifGel
        uneFrmDoc.LabelIndiceGel.Top = uneFrmDoc.FrameGel.Height / 2 - 60
        If uneFrmDoc.TabData.Tab = OngletGel Then
            'Si l 'onglet gel est l'onglet actif, on passe dans l'onglet voie
            'l'onglet gel ayant été grisé avant
            uneFrmDoc.TabData.Tab = OngletVoie
        End If
    End If
    
    'Affichage éventuel des valeurs des valeurs
    'de l'indice de gel admissible pour Q1 et Q2
    uneFrmDoc.LabelIGelAdmin.Visible = mesOptionsGen.maVerifGel
    uneFrmDoc.LabelIGAdmQ1.Visible = mesOptionsGen.maVerifGel And uneVisu
    uneFrmDoc.LabelIGAdmQ2.Visible = mesOptionsGen.maVerifGel And Not uneVisu
    If uneFrmDoc.monIndiceGelAdmQ1 = 0 Then
        uneString1 = MsgInconnu
    ElseIf uneFrmDoc.monIndiceGelAdmQ1 = HorsGel Then
        'Cas d'indice de gel admin infini (pente <= 0.05)
        uneString1 = MsgChausseeHorsGel
    Else
        uneString1 = Trim(Format(uneFrmDoc.monIndiceGelAdmQ1, "## ### ###"))
    End If
    If uneFrmDoc.monIndiceGelAdmQ2 = 0 Then
        uneString2 = MsgInconnu
    ElseIf uneFrmDoc.monIndiceGelAdmQ2 = HorsGel Then
        'Cas d'indice de gel admin infini (pente <= 0.05)
        uneString2 = MsgChausseeHorsGel
    Else
        uneString2 = Trim(Format(uneFrmDoc.monIndiceGelAdmQ2, "## ### ###"))
    End If
    'Positionnement des différents labels d'affichage
    uneFrmDoc.LabelIGelAdmin.Caption = LabelIGelAdminCaption
    uneFrmDoc.LabelIGelAdmin.Left = uneFrmDoc.BtnGelQ1.Width + uneFrmDoc.BtnGelQ1.Left * 2
    'uneFrmDoc.LabelIGelAdmin.Left = (uneFrmDoc.FrameGel.Width - uneFrmDoc.LabelIGelAdmin.Width) / 2
    uneFrmDoc.LabelIGAdmQ1.Caption = uneString1
    uneFrmDoc.LabelIGAdmQ1.Left = uneFrmDoc.LabelIGelAdmin.Left + uneFrmDoc.LabelIGelAdmin.Width
    uneFrmDoc.LabelIGAdmQ2.Caption = uneString2
    uneFrmDoc.LabelIGAdmQ2.Left = uneFrmDoc.LabelIGelAdmin.Left + uneFrmDoc.LabelIGelAdmin.Width
    
    'Affichage éventuel des valeurs des valeurs
    'de l'indice de gel de référence corrigé pour Q1 et Q2
    uneFrmDoc.LabelIGelRef.Visible = mesOptionsGen.maVerifGel
    uneFrmDoc.LabelIGRefQ1.Visible = mesOptionsGen.maVerifGel And uneVisu
    uneFrmDoc.LabelIGRefQ2.Visible = mesOptionsGen.maVerifGel And Not uneVisu
    If uneFrmDoc.monIndiceGelRefQ1 = -1 Then
        uneString1 = MsgInconnu
    ElseIf uneFrmDoc.monIndiceGelRefQ1 = 0 Then
        uneString1 = "0" 'car l'affichage avec format ci-dessus rend vide pour 0
    Else
        uneString1 = Trim(Format(uneFrmDoc.monIndiceGelRefQ1, "## ###"))
    End If
    If uneFrmDoc.monIndiceGelRefQ2 = -1 Then
        uneString2 = MsgInconnu
    ElseIf uneFrmDoc.monIndiceGelRefQ2 = 0 Then
        uneString2 = "0" 'car l'affichage avec format ci-dessus rend vide pour 0
    Else
        uneString2 = Trim(Format(uneFrmDoc.monIndiceGelRefQ2, "## ###"))
    End If
    
    'Positionnement des différents labels d'affichage
    uneFrmDoc.LabelIGelRef.Caption = LabelIGelRefCaption
    uneFrmDoc.LabelIGelRef.Left = uneFrmDoc.BtnGelQ1.Width + uneFrmDoc.BtnGelQ1.Left * 2
    'uneFrmDoc.LabelIGelRef.Left = (uneFrmDoc.FrameGel.Width - uneFrmDoc.LabelIGelRef.Width) / 2
    uneFrmDoc.LabelIGRefQ1.Caption = uneString1
    uneFrmDoc.LabelIGRefQ1.Left = uneFrmDoc.LabelIGelRef.Left + uneFrmDoc.LabelIGelRef.Width
    uneFrmDoc.LabelIGRefQ2.Caption = uneString2
    uneFrmDoc.LabelIGRefQ2.Left = uneFrmDoc.LabelIGelRef.Left + uneFrmDoc.LabelIGelRef.Width
    
    'Indication dans le cas où la chaussée n'est pas protégée au gel pour Q1
    If uneFrmDoc.monIndiceGelRefQ1 > uneFrmDoc.monIndiceGelAdmQ1 And uneFrmDoc.monIndiceGelAdmQ1 > 0 Then
        uneFrmDoc.BtnGelQ1.Visible = True And mesOptionsGen.maVerifGel And uneVisu
        'Indication que la vérif au gel n'est pas ok
        uneVerifGel = False
    Else
        uneFrmDoc.BtnGelQ1.Visible = False
        'Indication que la vérif au gel est ok
        uneVerifGel = uneFrmDoc.monIndiceGelAdmQ1 > 0 And uneFrmDoc.monIndiceGelAdmQ1 <> HorsGel
    End If
    
    'Indication dans le cas où la chaussée n'est pas protégée au gel pour Q2
    If uneFrmDoc.monIndiceGelRefQ2 > uneFrmDoc.monIndiceGelAdmQ2 And uneFrmDoc.monIndiceGelAdmQ2 > 0 Then
        uneFrmDoc.BtnGelQ2.Visible = True And mesOptionsGen.maVerifGel And Not uneVisu
        'Indication que la vérif au gel n'est pas ok si on est en Q2
        If Not uneVisu Then uneVerifGel = False
    Else
        uneFrmDoc.BtnGelQ2.Visible = False
        'Indication que la vérif au gel est ok si on est en Q2
        If Not uneVisu Then uneVerifGel = uneFrmDoc.monIndiceGelAdmQ2 > 0 And uneFrmDoc.monIndiceGelAdmQ2 <> HorsGel
    End If
    'Pour la version 2 bétatest, le bouton gel Q2 est mis au même endroit
    'que le bouton gel Q1 (les top sont dèjà égaux)
    uneFrmDoc.BtnGelQ2.Left = uneFrmDoc.BtnGelQ1.Left
    'Alignement du bouton OK gel
    uneFrmDoc.BtnOKGel.Left = uneFrmDoc.BtnGelQ1.Left
    uneFrmDoc.BtnOKGel.Top = uneFrmDoc.BtnGelQ1.Top
    'Affichage du bouton indiquant que la chaussée est protégée au gel en Q1 ou Q2
    uneFrmDoc.BtnOKGel.Visible = uneVerifGel And mesOptionsGen.maVerifGel
End Sub

Public Sub SaisirEpaisseursCarottePourTestDessin()
    Dim uneListEp As String, uneString As String
    Dim unTabEp As Variant
    Dim unMinTecQ1 As String, unMinTecQ2 As String
    Dim unMaxPraQ1 As String, unMaxPraQ2 As String
    Dim unTrouvEpQ1 As Boolean, unTrouvEpQ2 As Boolean
    Dim unePos1 As Integer, unePos2 As Integer
    
    unTabEp = Array(unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1, unePos1)
    
    uneListEp = InputBox("Entrez les épaisseurs de la carotte de qualité Q1, puis celle de qualité Q2 séparées par des blancs, épaisseur nulle pour une couche absente :", , "2 2 11 12 13 14 oui oui oui 3 0 15 0 20 0 non non oui")
    If uneListEp <> "" Then
        For i = 1 To 17
            unePos2 = InStr(unePos1 + 1, uneListEp, " ")
            uneString = Mid(uneListEp, unePos1 + 1, unePos2 - unePos1 - 1)
            If i = 7 Then
                unDecInd = 3
                unMinTecQ1 = uneString
            ElseIf i = 8 Then
                unMaxPraQ1 = uneString
            ElseIf i = 9 Then
                unTrouvEpQ1 = (UCase(uneString) = "OUI")
            ElseIf i = 16 Then
                unMinTecQ2 = uneString
            ElseIf i = 17 Then
                unMaxPraQ2 = uneString
            Else
                unTabEp(i - unDecInd) = Val(uneString)
            End If
            unePos1 = unePos2
        Next i
        uneString = Mid(uneListEp, unePos1 + 1)
        unTrouvEpQ2 = (UCase(uneString) = "OUI")
        
        unDecInd = 0
        If unDecInd = 0 Then
            'Si unDecInd = 0 ==> Affichage pour vérif saisie
            'sinon on met undecInd = 1 ==> pas d'affichage de vérif
            For i = 1 To 18
                If i = 7 Then
                    unDecInd = 3
                    unMsg = unMsg + unMinTecQ1 + "-"
                ElseIf i = 8 Then
                    unMsg = unMsg + unMaxPraQ1 + "-"
                ElseIf i = 9 Then
                    unMsg = unMsg + Format(unTrouvEpQ1) + "-"
                ElseIf i = 16 Then
                    unMsg = unMsg + unMinTecQ2 + "-"
                ElseIf i = 17 Then
                    unMsg = unMsg + unMaxPraQ2 + "-"
                ElseIf i = 18 Then
                    unMsg = unMsg + Format(unTrouvEpQ2)
                Else
                    unMsg = unMsg + Format(unTabEp(i - unDecInd)) + "-"
                End If
            Next i
            MsgBox unMsg
        End If
        
        'Dessin des carottes Q1 et Q2
        DessinerCarottes fMainForm.ActiveForm, unTabEp, unMinTecQ1, unMaxPraQ1, unTrouvEpQ1, unMinTecQ2, unMaxPraQ2, unTrouvEpQ2
    End If
End Sub

Public Sub DessinerCarottes(uneFrmDoc As Form, unTabEp As Variant, unMinTecQ1 As String, unMaxPraQ1 As String, unTrouvEpQ1 As Boolean, unMinTecQ2 As String, unMaxPraQ2 As String, unTrouvEpQ2 As Boolean)
    'Dessin des carottes de qualité Q1 et Q2 d'une fenetre fille en
    'respectant les proportions des différentes épaisseurs en commençant
    'à partir de la couche symbolisant la plateforme et en remontant les couches
    'Echelle 1 cm représentée par EpTotMaxEcran / max Ep totale réelle (Q1,Q2)
    'et les Carottes totales doit être compris entre 0 et 55 cm
    Dim unCmEcran As Single, uneEchelle As Single
    Dim uneEpTotQ1 As Integer, uneEpTotQ2 As Integer
    
    'Calcul de l'échelle
    For i = 1 To 6
        'Calcul de l'épaisseur totale de la carotte Q1
        uneEpTotQ1 = uneEpTotQ1 + unTabEp(i)
    Next i
    For i = 7 To 12
        'Calcul de l'épaisseur totale de la carotte Q2
        uneEpTotQ2 = uneEpTotQ2 + unTabEp(i)
    Next i
        
    If uneEpTotQ1 = 0 And uneEpTotQ2 = 0 Then
        uneEchelle = EpTotMaxEcran / 5 / EpParDefaut
    ElseIf uneEpTotQ1 > uneEpTotQ2 Then
        uneEchelle = EpTotMaxEcran / uneEpTotQ1
    Else
        uneEchelle = EpTotMaxEcran / uneEpTotQ2
    End If
    
    'Dessin de la carotte Q1 de la fenetre fille
    'à partir de la plateforme
    DessinerCarotteQ1 uneFrmDoc, CInt(unTabEp(1)), CInt(unTabEp(2)), CInt(unTabEp(3)), CInt(unTabEp(4)), CInt(unTabEp(5)), CInt(unTabEp(6)), uneEchelle, unMinTecQ1, unMaxPraQ1, unTrouvEpQ1
    'Dessin de la carotte Q2 de la fenetre fille
    'à partir de la plateforme
    DessinerCarotteQ2 uneFrmDoc, CInt(unTabEp(7)), CInt(unTabEp(8)), CInt(unTabEp(9)), CInt(unTabEp(10)), CInt(unTabEp(11)), CInt(unTabEp(12)), uneEchelle, unMinTecQ2, unMaxPraQ2, unTrouvEpQ2
End Sub

Public Sub DessinerCarotteQ1(uneFrmDoc As Form, uneEpS1 As Integer, uneEpS2 As Integer, uneEpB1 As Integer, uneEpB2 As Integer, uneEpF1 As Integer, uneEpF2 As Integer, uneEchelle As Single, unMinTec As String, unMaxPra As String, unTrouvEp As Boolean)
    'Dessin de la carotte de qualité Q1 en respectant les proportions
    'des différentes épaisseurs en commençant à partir de la
    'couche symbolisant la plateforme et en remontant les couches
    Dim unTop As Long, uneSurfSansEpaisseur As Boolean
    Dim uneStruct As Structure
    Dim uneExistence As Boolean, unMatSurfComposé As Boolean
    Dim uneStringEpS1 As String, uneStringEpS2 As String
    Dim uneStringEpB1 As String, uneStringEpB2 As String
    Dim uneStringEpF1 As String, uneStringEpF2 As String
    
    unMatSurfComposé = False
    'En version 2 bétatest on n'affiche que la carotte de la qualité choisie Q1 ou Q2
    uneVisu = (uneFrmDoc.monTypeChantier = TypeChantierQ1)
    uneFrmDoc.LabelPFQ1.Visible = uneVisu
    uneFrmDoc.ShapePFQ1.Visible = uneVisu
    
    'Affectation des libellés des différentes couches
    Set uneStruct = DonnerStructChoisie(uneFrmDoc)
    If Not (uneStruct Is Nothing) Then
        'Cas d'une structure choisie, on affiche les matériaux
        'de ses couches
        With uneFrmDoc
            .LabelCFond2Q1.Caption = uneStruct.maCoucheFondation
            .LabelCFond1Q1.Caption = uneStruct.maCoucheFondation
            .LabelCBase2Q1.Caption = uneStruct.maCoucheBase
            .LabelCBase1Q1.Caption = uneStruct.maCoucheBase
            .LabelCSurfQ1.Caption = uneStruct.maCoucheSurface
        End With
        
        If uneStruct.maCoucheSurface <> "Aucune" Then
            unMatSurfComposé = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé)
        End If
    End If
    
    'Recherche si on a une structure avec une couche de surface sans épaisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Affectation des valeurs d'épaisseurs à afficher suivant
    'le résultat du dimensionnement
    If UCase(unMaxPra) = "OUI" Or unTrouvEp = False Then
        'Maximun pratique atteint ou pas d'épaisseur trouvée
        '(qui fait qussi affichage après premier choix de structure)
        uneStringEpS1 = "??"
        uneStringEpS2 = "??"
        uneStringEpB1 = "??"
        uneStringEpB2 = "??"
        uneStringEpF1 = "??"
        uneStringEpF2 = "??"
        uneFrmDoc.LabelEpTotQ1.Caption = "Tot = ?? cm"
    Else
        uneStringEpS1 = Format(uneEpS1)
        uneStringEpS2 = Format(uneEpS2)
        uneStringEpB1 = Format(uneEpB1)
        uneStringEpB2 = Format(uneEpB2)
        uneStringEpF1 = Format(uneEpF1)
        uneStringEpF2 = Format(uneEpF2)
        uneFrmDoc.LabelEpTotQ1.Caption = "Tot = " + Format(uneFrmDoc.DonnerEpaisseurTotale(1)) + " cm" '(uneEpS1 * Abs(Not uneSurfSansEpaisseur) + uneEpS2 + uneEpB1 + uneEpB2 + uneEpF1 + uneEpF2) + " cm"
    End If
    
    With uneFrmDoc
        'Affichages éventuels du min techno et du max pratique
        .LabelInfoMaxPQ1.Visible = (UCase(unMaxPra) = "OUI") And uneVisu
        .LabelInfoMaxPQ1.Top = .FrameCarotte.Height / 2
        .LabelInfoMinTechnoQ1.Visible = (UCase(unMinTec) = "OUI") And uneVisu
        .LabelEpTotQ1.Visible = UCase(unMaxPra) = "NON" And uneVisu
    
        'Initialisation du début du dessin
        unTop = .ShapePFQ1.Top
            
        'Couche de fondation 2
        uneExistence = (uneEpF2 > 0)
        If uneExistence Then
            .ShapeFond2Q1.Height = Int(uneEchelle * uneEpF2)
            .ShapeFond2Q1.Top = unTop - .ShapeFond2Q1.Height
            unTop = .ShapeFond2Q1.Top
            .LabelCFond2Q1.Top = unTop + (.ShapeFond2Q1.Height - .LabelCFond2Q1.Height) / 2
            .LabelCFond2Q1.Left = .ShapePFQ1.Left + (.ShapePFQ1.Width - .LabelCFond2Q1.Width) / 2
            .LabelFond2Q1.Top = .LabelCFond2Q1.Top
            .LabelFond2Q1.Caption = uneStringEpF2 + " cm"
            .LabelFond2Q1.Left = .ShapeFond2Q1.Left - .LabelFond2Q1.Width - 100
        End If
        .ShapeFond2Q1.Visible = uneExistence And uneVisu
        .LabelCFond2Q1.Visible = uneExistence And uneVisu
        .LabelFond2Q1.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de fondation 1
        uneExistence = (uneEpF1 > 0)
        If uneExistence Then
            .ShapeFond1Q1.Height = Int(uneEchelle * uneEpF1)
            .ShapeFond1Q1.Top = unTop - .ShapeFond1Q1.Height
            unTop = .ShapeFond1Q1.Top
            .LabelCFond1Q1.Top = unTop + (.ShapeFond1Q1.Height - .LabelCFond1Q1.Height) / 2
            .LabelCFond1Q1.Left = .ShapePFQ1.Left + (.ShapePFQ1.Width - .LabelCFond1Q1.Width) / 2
            .LabelFond1Q1.Top = .LabelCFond1Q1.Top
            .LabelFond1Q1.Caption = uneStringEpF1 + " cm"
            .LabelFond1Q1.Left = .ShapeFond1Q1.Left - .LabelFond1Q1.Width - 100
        End If
        .ShapeFond1Q1.Visible = uneExistence And uneVisu
        .LabelCFond1Q1.Visible = uneExistence And uneVisu
        .LabelFond1Q1.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de base 2
        uneExistence = (uneEpB2 > 0)
        If uneExistence Then
            .ShapeBase2Q1.Height = Int(uneEchelle * uneEpB2)
            .ShapeBase2Q1.Top = unTop - .ShapeBase2Q1.Height
            unTop = .ShapeBase2Q1.Top
            .LabelCBase2Q1.Top = unTop + (.ShapeBase2Q1.Height - .LabelCBase2Q1.Height) / 2
            .LabelCBase2Q1.Left = .ShapePFQ1.Left + (.ShapePFQ1.Width - .LabelCBase2Q1.Width) / 2
            .LabelBase2Q1.Top = .LabelCBase2Q1.Top
            .LabelBase2Q1.Caption = uneStringEpB2 + " cm"
            .LabelBase2Q1.Left = .ShapeBase2Q1.Left - .LabelBase2Q1.Width - 100
        End If
        .ShapeBase2Q1.Visible = uneExistence And uneVisu
        .LabelCBase2Q1.Visible = uneExistence And uneVisu
        .LabelBase2Q1.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de base 1
        uneExistence = (uneEpB1 > 0)
        If uneExistence Then
            .ShapeBase1Q1.Height = Int(uneEchelle * uneEpB1)
            .ShapeBase1Q1.Top = unTop - .ShapeBase1Q1.Height
            unTop = .ShapeBase1Q1.Top
            .LabelCBase1Q1.Top = unTop + (.ShapeBase1Q1.Height - .LabelCBase1Q1.Height) / 2
            .LabelCBase1Q1.Left = .ShapePFQ1.Left + (.ShapePFQ1.Width - .LabelCBase1Q1.Width) / 2
            .LabelBase1Q1.Top = .LabelCBase1Q1.Top
            .LabelBase1Q1.Caption = uneStringEpB1 + " cm"
            .LabelBase1Q1.Left = .ShapeBase1Q1.Left - .LabelBase1Q1.Width - 100
        End If
        .ShapeBase1Q1.Visible = uneExistence And uneVisu
        .LabelCBase1Q1.Visible = uneExistence And uneVisu
        .LabelBase1Q1.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de surface
        'uneEpS1 = Ep totale de la surface et uneEpS2 = 0 si matériau simple
        'sinon uneEpS1 et uneEpS2 non nulles pour un matériau de surface composé
        uneExistence = (uneEpS1 > 0)
        uneCSurfEnDallesOuPaves = (uneStruct.maCoucheSurface = "Dalles") Or (uneStruct.maCoucheSurface = "Pavés")
        .ShapeLitPoseQ1.Visible = uneCSurfEnDallesOuPaves And uneVisu
        .LabelLitPoseQ1.Visible = uneCSurfEnDallesOuPaves And uneVisu
        If uneExistence Then
            If uneCSurfEnDallesOuPaves Then
                'Cas d'une couche de surface en dalles ou pavés
                'Affichage du lit de pose
                .ShapeSurfQ1.Height = Int(uneEchelle * uneEpS1)
                .ShapeLitPoseQ1.Height = Int(uneEchelle * uneEpS2)
                .ShapeLitPoseQ1.Top = unTop - .ShapeLitPoseQ1.Height
                .LabelLitPoseQ1.Top = .ShapeLitPoseQ1.Top + (.ShapeLitPoseQ1.Height - .LabelLitPoseQ1.Height) / 2
                .LabelLitPoseQ1.Left = .ShapeLitPoseQ1.Left + (.ShapeLitPoseQ1.Width - .LabelLitPoseQ1.Width) / 2
                .LabelSurf2Q1.Caption = uneStringEpS2 + " cm"
                .LabelSurf2Q1.Top = .LabelLitPoseQ1.Top
                unTop = .ShapeLitPoseQ1.Top
            Else
                .ShapeSurfQ1.Height = Int(uneEchelle * (uneEpS1 + uneEpS2))
            End If
            .ShapeSurfQ1.Top = unTop - .ShapeSurfQ1.Height
            unTop = .ShapeSurfQ1.Top
            .LabelCSurfQ1.Top = unTop + (.ShapeSurfQ1.Height - .LabelCSurfQ1.Height) / 2
            .LabelCSurfQ1.Left = .ShapePFQ1.Left + (.ShapePFQ1.Width - .LabelCSurfQ1.Width) / 2
            If unMatSurfComposé Then
                'Cas d'un matériau de surface composé
                .LabelSurf1Q1.Top = .ShapeSurfQ1.Top + .ShapeSurfQ1.Height / 2 - .LabelSurf2Q1.Height
                .LabelSurf1Q1.Caption = .monMSComp1Q1 + " " + uneStringEpS1 + " cm"
                If uneEpS2 > 0 Then .LabelSurf1Q1.Caption = .LabelSurf1Q1.Caption + " +"
                .LabelSurf2Q1.Caption = .monMSComp2Q1 + " " + uneStringEpS2 + " cm"
                .LabelSurf2Q1.Top = .LabelSurf1Q1.Top + .LabelSurf1Q1.Height
            Else
                'Cas d'un matériau de surface simple
                .LabelSurf1Q1.Top = .LabelCSurfQ1.Top
                If uneSurfSansEpaisseur Then
                    'Cas d'une couche de surface sans épaisseur
                    .LabelSurf1Q1.Caption = ""
                Else
                    .LabelSurf1Q1.Caption = uneStringEpS1 + " cm"
                End If
            End If
            .LabelSurf1Q1.Left = .ShapeSurfQ1.Left - .LabelSurf1Q1.Width - 70
            .LabelSurf2Q1.Left = .LabelSurf1Q1.Left
        End If
        .ShapeSurfQ1.Visible = uneExistence And uneVisu
        .LabelCSurfQ1.Visible = uneExistence And uneVisu
        .LabelSurf1Q1.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        .LabelSurf2Q1.Visible = (uneEpS2 > 0) And UCase(unMaxPra) = "NON" And uneVisu
    End With
End Sub

Public Sub DessinerCarotteQ2(uneFrmDoc As Form, uneEpS1 As Integer, uneEpS2 As Integer, uneEpB1 As Integer, uneEpB2 As Integer, uneEpF1 As Integer, uneEpF2 As Integer, uneEchelle As Single, unMinTec As String, unMaxPra As String, unTrouvEp As Boolean)
    'Dessin de la carotte de qualité Q2 en respectant les proportions
    'des différentes épaisseurs en commençant à partir de la
    'couche symbolisant la plateforme et en remontant les couches
    Dim unTop As Long, uneVisu As Boolean
    Dim uneStruct As Structure
    Dim uneExistence As Boolean, unMatSurfComposé As Boolean
    Dim uneStringEpS1 As String, uneStringEpS2 As String
    Dim uneStringEpB1 As String, uneStringEpB2 As String
    Dim uneStringEpF1 As String, uneStringEpF2 As String
    
    unMatSurfComposé = False
    'En version 2 bétatest on n'affiche que la carotte de la qualité choisie Q1 ou Q2
    uneVisu = (uneFrmDoc.monTypeChantier = TypeChantierQ2)
    uneFrmDoc.LabelPFQ2.Visible = uneVisu
    uneFrmDoc.ShapePFQ2.Visible = uneVisu
    
    'Affectation des libellés des différentes couches
    Set uneStruct = DonnerStructChoisie(uneFrmDoc)
    If Not (uneStruct Is Nothing) Then
        'Cas d'une structure choisie, on affiche les matériaux
        'de ses couches
        With uneFrmDoc
            .LabelCFond2Q2.Caption = uneStruct.maCoucheFondation
            .LabelCFond1Q2.Caption = uneStruct.maCoucheFondation
            .LabelCBase2Q2.Caption = uneStruct.maCoucheBase
            .LabelCBase1Q2.Caption = uneStruct.maCoucheBase
            .LabelCSurfQ2.Caption = uneStruct.maCoucheSurface
        End With
        
        If uneStruct.maCoucheSurface <> "Aucune" Then
            unMatSurfComposé = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé)
        End If
    End If
    
    'Recherche si on a une structure avec une couche de surface sans épaisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Affectation des valeurs d'épaisseurs à afficher suivant
    'le résultat du dimensionnement
    If UCase(unMaxPra) = "OUI" Or unTrouvEp = False Then
        'Maximun pratique atteint ou pas d'épaisseur trouvée
        '(qui fait qussi affichage après premier choix de structure)
        uneStringEpS1 = "??"
        uneStringEpS2 = "??"
        uneStringEpB1 = "??"
        uneStringEpB2 = "??"
        uneStringEpF1 = "??"
        uneStringEpF2 = "??"
        uneFrmDoc.LabelEpTotQ2.Caption = "Tot = ?? cm"
    Else
        uneStringEpS1 = Format(uneEpS1)
        uneStringEpS2 = Format(uneEpS2)
        uneStringEpB1 = Format(uneEpB1)
        uneStringEpB2 = Format(uneEpB2)
        uneStringEpF1 = Format(uneEpF1)
        uneStringEpF2 = Format(uneEpF2)
        uneFrmDoc.LabelEpTotQ2.Caption = "Tot = " + Format(uneFrmDoc.DonnerEpaisseurTotale(2)) + " cm" 'Format(uneEpS1 * Abs(Not uneSurfSansEpaisseur) + uneEpS2 + uneEpB1 + uneEpB2 + uneEpF1 + uneEpF2) + " cm"
    End If
    
    With uneFrmDoc
        'Affichages éventuels du min techno et du max pratique
        .LabelInfoMaxPQ2.Visible = (UCase(unMaxPra) = "OUI") And uneVisu
        .LabelInfoMaxPQ2.Top = .FrameCarotte.Height / 2
        .LabelInfoMinTechnoQ2.Visible = (UCase(unMinTec) = "OUI") And uneVisu
        uneFrmDoc.LabelEpTotQ2.Visible = UCase(unMaxPra) = "NON" And uneVisu
        
        'Initialisation du début du dessin
        unTop = .ShapePFQ2.Top
            
        'Couche de fondation 2
        uneExistence = (uneEpF2 > 0)
        If uneExistence Then
            .ShapeFond2Q2.Height = CInt(uneEchelle * uneEpF2)
            .ShapeFond2Q2.Top = unTop - .ShapeFond2Q2.Height
            unTop = .ShapeFond2Q2.Top
            .LabelCFond2Q2.Top = unTop + (.ShapeFond2Q2.Height - .LabelCFond2Q2.Height) / 2
            .LabelCFond2Q2.Left = .ShapePFQ2.Left + (.ShapePFQ2.Width - .LabelCFond2Q2.Width) / 2
            .LabelFond2Q2.Top = .LabelCFond2Q2.Top
            .LabelFond2Q2.Caption = uneStringEpF2 + " cm"
        End If
        .ShapeFond2Q2.Visible = uneExistence And uneVisu
        .LabelCFond2Q2.Visible = uneExistence And uneVisu
        .LabelFond2Q2.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de fondation 1
        uneExistence = (uneEpF1 > 0)
        If uneExistence Then
            .ShapeFond1Q2.Height = CInt(uneEchelle * uneEpF1)
            .ShapeFond1Q2.Top = unTop - .ShapeFond1Q2.Height
            unTop = .ShapeFond1Q2.Top
            .LabelCFond1Q2.Top = unTop + (.ShapeFond1Q2.Height - .LabelCFond1Q2.Height) / 2
            .LabelCFond1Q2.Left = .ShapePFQ2.Left + (.ShapePFQ2.Width - .LabelCFond1Q2.Width) / 2
            .LabelFond1Q2.Top = .LabelCFond1Q2.Top
            .LabelFond1Q2.Caption = uneStringEpF1 + " cm"
        End If
        .ShapeFond1Q2.Visible = uneExistence And uneVisu
        .LabelCFond1Q2.Visible = uneExistence And uneVisu
        .LabelFond1Q2.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de base 2
        uneExistence = (uneEpB2 > 0)
        If uneExistence Then
            .ShapeBase2Q2.Height = CInt(uneEchelle * uneEpB2)
            .ShapeBase2Q2.Top = unTop - .ShapeBase2Q2.Height
            unTop = .ShapeBase2Q2.Top
            .LabelCBase2Q2.Top = unTop + (.ShapeBase2Q2.Height - .LabelCBase2Q2.Height) / 2
            .LabelCBase2Q2.Left = .ShapePFQ2.Left + (.ShapePFQ2.Width - .LabelCBase2Q2.Width) / 2
            .LabelBase2Q2.Top = .LabelCBase2Q2.Top
            .LabelBase2Q2.Caption = uneStringEpB2 + " cm"
        End If
        .ShapeBase2Q2.Visible = uneExistence And uneVisu
        .LabelCBase2Q2.Visible = uneExistence And uneVisu
        .LabelBase2Q2.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de base 1
        uneExistence = (uneEpB1 > 0)
        If uneExistence Then
            .ShapeBase1Q2.Height = CInt(uneEchelle * uneEpB1)
            .ShapeBase1Q2.Top = unTop - .ShapeBase1Q2.Height
            unTop = .ShapeBase1Q2.Top
            .LabelCBase1Q2.Top = unTop + (.ShapeBase1Q2.Height - .LabelCBase1Q2.Height) / 2
            .LabelCBase1Q2.Left = .ShapePFQ2.Left + (.ShapePFQ2.Width - .LabelCBase1Q2.Width) / 2
            .LabelBase1Q2.Top = .LabelCBase1Q2.Top
            .LabelBase1Q2.Caption = uneStringEpB1 + " cm"
        End If
        .ShapeBase1Q2.Visible = uneExistence And uneVisu
        .LabelCBase1Q2.Visible = uneExistence And uneVisu
        .LabelBase1Q2.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        
        'Couche de surface
        'uneEpS1 = Ep totale de la surface et uneEpS2 = 0 si matériau simple
        'sinon uneEpS1 et uneEpS2 non nulles pour un matériau de surface composé
        uneExistence = (uneEpS1 > 0)
        uneCSurfEnDallesOuPaves = (uneStruct.maCoucheSurface = "Dalles") Or (uneStruct.maCoucheSurface = "Pavés")
        .ShapeLitPoseQ2.Visible = uneCSurfEnDallesOuPaves And uneVisu
        .LabelLitPoseQ2.Visible = uneCSurfEnDallesOuPaves And uneVisu
        If uneExistence Then
            If uneCSurfEnDallesOuPaves Then
                'Cas d'une couche de surface en dalles ou pavés
                'Affichage du lit de pose
                .ShapeSurfQ2.Height = Int(uneEchelle * uneEpS1)
                .ShapeLitPoseQ2.Height = Int(uneEchelle * uneEpS2)
                .ShapeLitPoseQ2.Top = unTop - .ShapeLitPoseQ2.Height
                .LabelLitPoseQ2.Top = .ShapeLitPoseQ2.Top + (.ShapeLitPoseQ2.Height - .LabelLitPoseQ2.Height) / 2
                .LabelLitPoseQ2.Left = .ShapeLitPoseQ2.Left + (.ShapeLitPoseQ2.Width - .LabelLitPoseQ2.Width) / 2
                .LabelSurf2Q2.Caption = uneStringEpS2 + " cm"
                .LabelSurf2Q2.Top = .LabelLitPoseQ2.Top
                unTop = .ShapeLitPoseQ2.Top
            Else
            .ShapeSurfQ2.Height = CInt(uneEchelle * (uneEpS1 + uneEpS2))
            End If
            .ShapeSurfQ2.Top = unTop - .ShapeSurfQ2.Height
            unTop = .ShapeSurfQ2.Top
            .LabelCSurfQ2.Top = unTop + (.ShapeSurfQ2.Height - .LabelCSurfQ2.Height) / 2
            .LabelCSurfQ2.Left = .ShapePFQ2.Left + (.ShapePFQ2.Width - .LabelCSurfQ2.Width) / 2
            If unMatSurfComposé Then
                'Cas d'un matériau de surface composé (deux épaisseurs)
                .LabelSurf1Q2.Top = .ShapeSurfQ2.Top + .ShapeSurfQ2.Height / 2 - .LabelSurf2Q2.Height
                .LabelSurf1Q2.Caption = .monMSComp1Q2 + " " + uneStringEpS1 + " cm"
                If uneEpS2 > 0 Then .LabelSurf1Q2.Caption = .LabelSurf1Q2.Caption + " +"
                .LabelSurf2Q2.Caption = .monMSComp2Q2 + " " + uneStringEpS2 + " cm"
                .LabelSurf2Q2.Top = .LabelSurf1Q2.Top + .LabelSurf1Q2.Height + 10
            Else
                'Cas d'un matériau de surface simple
                .LabelSurf1Q2.Top = .LabelCSurfQ2.Top
                If uneSurfSansEpaisseur Then
                    'Cas d'une couche de surface sans épaisseur
                    .LabelSurf1Q2.Caption = ""
                Else
                    .LabelSurf1Q2.Caption = uneStringEpS1 + " cm"
                End If
            End If
        End If
        .ShapeSurfQ2.Visible = uneExistence And uneVisu
        .LabelCSurfQ2.Visible = uneExistence And uneVisu
        .LabelSurf1Q2.Visible = uneExistence And UCase(unMaxPra) = "NON" And uneVisu
        .LabelSurf2Q2.Visible = (uneEpS2 > 0) And UCase(unMaxPra) = "NON" And uneVisu
    End With
End Sub


Public Function VerifierSaisieOngletCourant() As Boolean
    'Afficher les messages d'erreur si une saisie est erronée dans
    'l'onglet courant de la fenetre fille active
    Dim uneForm As Form
    
    'Initialisation du code de retour
    VerifierSaisieOngletCourant = True
    
    If Forms.Count > 1 Then
        'Cas où il y a au moins une fenetre fille
        Set uneForm = fMainForm.ActiveForm
        If uneForm.TabData.Tab = OngletTrafic Then
            'Vérification dans l'onglet Trafic
            VerifierSaisieOngletCourant = VerifierMinMaxTraficIni(uneForm, uneForm.TextTrafIni.Text)
            If VerifierSaisieOngletCourant Then VerifierSaisieOngletCourant = VerifierMinMaxDuréeService(uneForm)
        End If
    End If
End Function

Public Function VerifierFinSaisie() As Boolean
    'Vérifier si une saisie est en cours dans un textbox,
    'il faut valider avant pour une impression, save, new, open,
    'quitter ou fermer
    Dim unMaskEdBox As MaskEdBox
    
    If monEtude Is Nothing Then
        VerifierFinSaisie = True
        Exit Function
    End If
    
    'Vérification que l'on ne soit pas encore de saisie dans un textbox
    'demandant une validation par sortie du champ ou par retour chariot
    If Not (monEtude.ActiveControl Is Nothing) Then
        Set unMaskEdBox = monEtude.MaskCAM
    
        'Cas où il y a un control actif
        If TypeOf monEtude.ActiveControl Is TextBox Then
            If monEtude.ActiveControl.ForeColor = QBColor(12) Then
                'Cas d'une texte box avec un texte rouge ==> saisie en cours
                MsgBox MsgFinirSaisie + DonnerNomTextBox, vbCritical
                VerifierFinSaisie = False
                monEtude.WindowState = vbNormal
            Else
                VerifierFinSaisie = True
            End If
        ElseIf monEtude.ActiveControl Is unMaskEdBox Then
            If unMaskEdBox.ForeColor = QBColor(12) Then
                'Cas du MaskCAM avec un texte rouge ==> saisie en cours
                MsgBox MsgFinirSaisie + MsgMaskCAM, vbCritical
                unMaskEdBox.SetFocus
                VerifierFinSaisie = False
                monEtude.WindowState = vbNormal
            Else
                VerifierFinSaisie = True
            End If
        Else
            VerifierFinSaisie = True
        End If
    Else
        VerifierFinSaisie = True
    End If
End Function


Public Function DonnerNomTextBox() As String
    If monEtude.ActiveControl.Name = "TextTrafIni" Then
        DonnerNomTextBox = MsgTextBoxTrafIni
    ElseIf monEtude.ActiveControl.Name = "TextTrafCUM" Then
        DonnerNomTextBox = MsgTextBoxTrafCum
    ElseIf monEtude.ActiveControl.Name = "TextDuréeS" Then
        DonnerNomTextBox = MsgTextBoxDuréeS
    ElseIf monEtude.ActiveControl.Name = "TextHAgglo" Then
        DonnerNomTextBox = MsgTextBoxHAgglo
    ElseIf monEtude.ActiveControl.Name = "TextIndGelPerso" Then
        DonnerNomTextBox = MsgTextIndGelPerso
    ElseIf monEtude.ActiveControl.Name = "TextEpaisseur" Then
        DonnerNomTextBox = MsgTextBoxEpais
    ElseIf monEtude.ActiveControl.Name = "TextPente" Then
        DonnerNomTextBox = MsgTextBoxPente
    End If
End Function

Public Sub MettreAJourFrameRésultat(uneForm As Form)
    'Mise à jour de l'affichage de la frame visualisant
    'les résultats et les carottes de qualité Q1 et Q2
    With uneForm
        If .monTraficCumulé = 0 Then
            uneString = MsgInconnu
        Else
            uneString = Format(.monTraficCumulé, "### ### ###")
        End If
        .LabelTraficCum.Caption = LabelTraficCumCaption + uneString
        
        If .monCAM = "" Then
            uneString = MsgInconnu
        Else
            'uneString = .monCAM
            uneString = Mid(.monCAM, 1, 1) + monCarDeci + Mid(.monCAM, 3)
            'On prend de part et d'autre du séparateur décimale
            'car le cam a un format #.##
        End If
        .LabelCAM.Caption = LabelCAMCaption + uneString
        
        If .monNEEquiv = 0 Then
            uneString = MsgInconnu
        Else
            uneString = Format(.monNEEquiv, "### ### ###")
        End If
        .LabelNEequiv.Caption = LabelNEequivCaption + uneString
        
        If .monIndicePF = 1 Then
            .OptionPF1.Value = True
        ElseIf .monIndicePF = 2 Then
            .OptionPF2.Value = True
        ElseIf .monIndicePF = 3 Then
             .OptionPF3.Value = True
       End If
        
        .FrameCarotte.Visible = (.monIndStructChoisie > 0)
    End With
End Sub

Public Sub MettreAJourOngletVoie(uneForm As Form)
    'Mise à jour de l'affichage de l'onglet voie
    If uneForm.maDate = "" Then
        'Affichage de la date du jour
        uneForm.LabelDate.Caption = Format(Date, "dd/mm/yyyy")
    Else
        uneForm.LabelDate.Caption = uneForm.maDate
    End If
    uneForm.TextTitre = uneForm.monTitreEtude
    uneForm.TextVar = uneForm.maVariante
End Sub

Public Sub MettreAJourOngletTrafic(uneForm As Form)
    'Mise à jour de l'affichage de l'onglet trafic
    With uneForm
        If .monTraficIni = 0 Then
            .TextTrafIni.Text = ""
        Else
            .TextTrafIni.Text = Format(.monTraficIni)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextTrafIni.Tag = .TextTrafIni.Text
        .TextTrafIni.ForeColor = QBColor(0) 'Car pas de modif
        
        If .maDuréeService = 0 Then
            .TextDuréeS.Text = Format(mesOptionsGen.maDuréeService)
        Else
            .TextDuréeS.Text = Format(.maDuréeService)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextDuréeS.Tag = .TextDuréeS.Text
        .TextDuréeS.ForeColor = QBColor(0) 'Car pas de modif
        
        If .maCroisAnnuel = 100 Then
            'Cas d'une nouvelle étude la croiss annuel est inconnu (= 100)
            'Les % vont de 0 à 5 % et
            'le tableau de controle optionCA de 0 à 5
            .OptionCA(mesOptionsGen.maCroisAnnuel).Value = True
        Else
            .OptionCA(.maCroisAnnuel).Value = True
        End If
        
        If .monTraficCumulé = 0 Then
            .TextTrafCUM.Text = ""
        Else
            .TextTrafCUM.Text = Format(.monTraficCumulé)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextTrafCUM.Tag = .TextTrafCUM.Text
        .TextTrafCUM.ForeColor = QBColor(0) 'Car pas de modif
    End With
End Sub

Public Sub MettreAJourOngletCoucheSurface(uneForm As Form, uneEpPrecTrouvQ1 As Integer, uneEpPrecTrouvQ2 As Integer)
    'Mise à jour du contenu Onglet couche de surface
    Dim unNomMat1 As String, unNomMat2 As String, unNomMat3 As String
    Dim uneEpMat1 As Integer, uneEpMat2 As Integer, uneEpMat3 As Integer
    Dim uneEpPrec As Integer, unN As Integer
    Dim lesComp As Collection
    Dim uneStruct As Structure
    Dim unMatComposé As Object
    
    With uneForm
        .LabelValEpPrecQ1.Caption = Format(uneEpPrecTrouvQ1)
        .LabelValEpPrecQ2.Caption = Format(uneEpPrecTrouvQ2)
        
        'Récup de la structure choisie
        Set uneStruct = DonnerStructChoisie(uneForm)
        If uneStruct Is Nothing Then Exit Sub
        
        'Récup du matériau composé en couche de surface
        Set unMatComposé = maColMatSurf(uneStruct.maCoucheSurface)
        If Not (TypeOf unMatComposé Is MatComposé) Then Exit Sub
        
        'On vide la liste des matériaux composants possibles
        'et la combobox des compositions possibles
        .ListViewMat.ListItems.Clear
        .ComboCompQ1.Clear
        .ComboCompQ2.Clear
        
        'Remplissage de la combobox combocomp des compositions possibles
        Set lesComp = unMatComposé.mesCompositions
        unNbComp = 3
        '3 = Nb de colonnes de composant
        unNbComposition = lesComp.Count \ (2 * unNbComp + 1)
        For j = 1 To unNbComposition
            unN = (j - 1) * (2 * unNbComp + 1) + 1
            'Récup de l'épaisseur préconisée de la composition j
            uneEpPrec = CInt(lesComp(unN))
            
            'Recup premier composant
            uneEpMat1 = CInt(lesComp(unN + 1))
            unNomMat1 = Format(lesComp(unN + 2))
            
            'Recup deuxième composant
            uneEpMat2 = CInt(lesComp(unN + 3))
            unNomMat2 = Format(lesComp(unN + 4))
                        
            'Recup troisième composant
            uneEpMat3 = CInt(lesComp(unN + 5))
            unNomMat3 = Format(lesComp(unN + 6))
                            
            If uneEpPrec = uneEpPrecTrouvQ1 Then
                'Cas où l'épaisseur préconisée trouvé Q1
                'correspond à celle de la composition j
                '===> Ajout aux compositions possibles
                AjouterComposition uneForm, .ComboCompQ1, uneEpMat1, unNomMat1, uneEpMat2, unNomMat2, uneEpMat3, unNomMat3, unN
            End If
             
             If uneEpPrec = uneEpPrecTrouvQ2 Then
                'Cas où l'épaisseur préconisée trouvé Q2
                'correspond à celle de la composition j
                '===> Ajout aux compositions possibles
                AjouterComposition uneForm, .ComboCompQ2, uneEpMat1, unNomMat1, uneEpMat2, unNomMat2, uneEpMat3, unNomMat3, unN
           End If
        Next j
        
        'Affichage des compositions Q1 et Q2
        If uneForm.monJustOpen Then
            'Cas où on vient d'ouvrir l'étude on met les indices stockés
            'et on n'y passe qu'une fois
            .ComboCompQ1.ListIndex = (.monIndCompQ1 - 1)
            .ComboCompQ2.ListIndex = (.monIndCompQ2 - 1)
            uneForm.monJustOpen = False
        Else
            'tous les autres cas ==> on remet à vide
            .ComboCompQ1.ListIndex = -1
            .ComboCompQ2.ListIndex = -1
        End If
    End With
End Sub

Public Sub MettreAJourOngletStructure(uneForm As Form, unUtilFichPerso As Byte, unIndStructChoisie As Integer)
    'Mise à jour du contenu Onglet Structure
    Dim uneColStruct As Collection
    Dim uneStruct As Structure
    
    'Changement de valeur de la case à cocher Utiliser un fichier perso
    'Comme elle vaut Grayed = 2 toutes les valeurs 0 ou 1 déclenchent
    'sont click event (cf FrmDocument code, checkfichperso_click)
    
    uneForm.CheckFichPerso.Value = unUtilFichPerso
    
    'Affichage de la case à cocher Utiliser Fichier personnel si
    'il y  un fichier personnel donné dans les options générales
    'ou si l'étude en utilise un.
    If uneForm.monFichPersoSTR = "" And mesOptionsMat.monFichPersoSTR = "" Then
        uneForm.CheckFichPerso.Visible = False
    End If
    
    'Affichage de la structure choisie dans la combobox combostruct
    If unIndStructChoisie > 0 Then
        'Cas du chargement d'une étude ayant déjà choisie une structure
        'Récup de la bonne collection de structures
        If uneForm.CheckFichPerso.Value = 0 Then
            Set uneColStruct = maColStructCERTU
        Else
            Set uneColStruct = maColStructPerso
        End If
        
        'En version 2, Déclenchement du radio bouton correspondant
        'au type de la strucutre chosie éventuelle
        Set uneStruct = uneColStruct(unIndStructChoisie)
        uneForm.OptionTypeStruct(uneStruct.monTypeStructure).Value = True
        
        uneStructTrouv = False
        i = 1
        Do
            'Parcours de la liste des structures
            Set uneStruct = uneColStruct(i)
            j = 0
            Do
                'Parcours des structures de la combobox combostruct
                'qui a été rempli avant par le CheckFichPerso_click
                'grâce au CheckFichPerso.value = du début
                
                'On cherche la structure de même abrégé et se trouvant
                'à la position i (tout ça car les abrégés de structures
                'ne sont pas uniques)
                unInd = uneForm.ComboStruct.ItemData(j)
                If uneForm.ComboStruct.List(j) = uneStruct.monAbrégé And unInd = unIndStructChoisie Then
                    uneForm.ComboStruct.ListIndex = j
                    uneStructTrouv = True
                End If
                j = j + 1
            Loop Until uneStructTrouv Or j = uneForm.ComboStruct.ListCount
            i = i + 1
        Loop Until uneStructTrouv Or i > uneColStruct.Count
    End If
End Sub

Public Sub ActiverTypeStructure(uneForm As Form)
    'Fonction rajoutée en version 2
    'Activation ou déactivation des radio boutons à cliquer
    'de l'étude active, donc de la form passée en paramètre,
    'permettant de choisir le type de structures dans l'onglet Structure
    'leur click remplira la liste des structures avec juste les structures du bon type
    
    Dim uneColStruct As Collection
    Dim uneStruct As Structure
    Dim unTypeVoieOK As Boolean
    
    If uneForm.CheckFichPerso.Value = 0 Or uneForm.LabelFichPerso.Caption = "" Then
        Set uneColStruct = maColStructCERTU
    Else
        Set uneColStruct = maColStructPerso
    End If
    
    'Vidage de la liste des structures
    uneForm.ComboStruct.Clear
    uneForm.ComboStruct.ListIndex = -1
    'Indication que l'on change de type de structure
    uneForm.ComboStruct.Tag = Format(ChangeTypeStruct)
        
    'Mise en désactivé de tous les radio boutons et décochage
    For i = Souple To PavesDalles
        uneForm.OptionTypeStruct(i).Enabled = False
        uneForm.OptionTypeStruct(i).Value = False
    Next i
    
    'Mise en activation éventuelle des radio boutons
    For i = 1 To uneColStruct.Count
        Set uneStruct = uneColStruct(i)
        'Calcul si la structure est utilisable dans le type de voie choisi
        unTypeVoieOK = (uneForm.monTypeVoie = TypeVoieDesserte) And (uneStruct.monUtilVDes = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieDistribution) And (uneStruct.monUtilVDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieTraficLourd) And (uneStruct.monUtilVPL = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieBus) And (uneStruct.monUtilVBus = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieParking) And (uneStruct.monUtilVParking = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireDistribution) And (uneStruct.monUtilGDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireTraficLourd) And (uneStruct.monUtilGPL = 1)
        
        If unTypeVoieOK Then
            uneForm.OptionTypeStruct(uneStruct.monTypeStructure).Enabled = True
        End If
    Next i
End Sub

Public Sub RemplirComboStructures(uneForm As Form, Optional unTypeStruct As Integer = ToutType)
    'Remplissage de la combobox listant les structures possibles
    'du fichier de structures perso ou CERTU,
    'donc ayant des matériaux de base et de fondation
    'autorisés et ayant un type de voie compatibles avec
    'celui de l'étude active, donc de la form passée en paramètre
    
    'EN PLUS EN VERSION 2, IL FAUT LES STRUCTURES DU BON TYPE
    Dim uneColStruct As Collection
    Dim uneColMatBF As Collection
    Dim unTypeVoieOK As Boolean
    Dim uneStruct As Structure, unMat As Matériau
    
    If uneForm.CheckFichPerso.Value = 0 Or uneForm.LabelFichPerso.Caption = "" Then
        Set uneColStruct = maColStructCERTU
        Set uneColMatBF = maColMatBFCERTU
    Else
        Set uneColStruct = maColStructPerso
        Set uneColMatBF = maColMatBFPerso
    End If
    
    unInd = 0
    uneForm.ComboStruct.Clear
    uneForm.ComboStruct.ListIndex = -1
    'Indication que l'on change de type de structure
    uneForm.ComboStruct.Tag = Format(ChangeTypeStruct)
    
    For i = 1 To uneColStruct.Count
        Set uneStruct = uneColStruct(i)
        
        'Vérification que cette structure a un type de voie
        'compatible avec celui de l'étude (= la form) active
        unTypeVoieOK = (uneForm.monTypeVoie = TypeVoieDesserte) And (uneStruct.monUtilVDes = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieDistribution) And (uneStruct.monUtilVDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieTraficLourd) And (uneStruct.monUtilVPL = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieBus) And (uneStruct.monUtilVBus = 1)
        
        'Pour la version 2 bétatest, on rajoute les VoieParking
        'et pour les voies giratoires
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieParking) And (uneStruct.monUtilVParking = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireDistribution) And (uneStruct.monUtilGDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireTraficLourd) And (uneStruct.monUtilGPL = 1)
        
        'EN PLUS EN VERSION 2, IL FAUT LES STRUCTURES DU BON TYPE
        unTypeVoieOK = unTypeVoieOK And (uneStruct.monTypeStructure = unTypeStruct)
        
        If unTypeVoieOK Then
            'Cas où la structure i est utilisable
            'pour le type de voie de l'etude active
            
            'Vérification que le matériau de base
            's'il existe est autorisé
            If uneStruct.maCoucheBase = "Aucune" Then
                unMatBaseOK = True
            Else
                Set unMat = uneColMatBF.Item(uneStruct.maCoucheBase)
                unMatBaseOK = unMat.monUtilisationAutorisée
            End If
            'Vérification que le matériau de fondation
            's'il existe est autorisé
            If uneStruct.maCoucheFondation = "Aucune" Then
                unMatFondOK = True
            Else
                Set unMat = uneColMatBF.Item(uneStruct.maCoucheFondation)
                unMatFondOK = unMat.monUtilisationAutorisée
            End If
            
            If unMatBaseOK And unMatFondOK Then
                'Ajout dans la liste des structures possibles
                uneForm.ComboStruct.AddItem uneColStruct(i).monAbrégé
                'Stockage dans l'item de la combobox de sa position dans
                'la collection des structures persos ou Certu
                uneForm.ComboStruct.ItemData(unInd) = i
                unInd = unInd + 1
            End If
        End If
    Next i
End Sub
        
Public Sub InhiberBoutonMat(uneForm As Form)
    If uneForm.ComboStruct.ListIndex = -1 Then
        'Cas où aucune structure n'a été choisie d'où
        'Inhibition des boutons de visu des matériaux surface simple,
        'base et fondation
        'Et affichage d'un libellé de boutons sans abrégé matériau
        uneForm.cmdInfoMS.Enabled = False
        uneForm.cmdInfoMS.Caption = LabelCmdInfoMS
        uneForm.CmdInfoMB.Enabled = False
        uneForm.CmdInfoMB.Caption = LabelCmdInfoMB
        uneForm.CmdInfoMF.Enabled = False
        uneForm.CmdInfoMF.Caption = LabelCmdInfoMF
        uneForm.LabelRisk.Visible = False
    End If
End Sub

Public Sub AfficherFichierPerso(uneForm As Form)
    'Affichage du fichier personnel dans l'onglet Structure
    With uneForm
        If (uneForm.monFichPersoSTR = mesOptionsMat.monFichPersoSTR Or uneForm.monFichPersoSTR = "") Then
            'Cas où le fichier perso utilisé est celui des options matériaux
            'ou n'existe pas ==> On fait le traitement
            .ComboStruct.ListIndex = -1 'Vide la structure choisie
            .FrameCarotte.Visible = False
            .LabelFichPerso.Visible = (.CheckFichPerso.Value = 1)
            If .CheckFichPerso.Value = 1 Then
                .LabelFichPerso.Caption = mesOptionsMat.monFichPersoSTR
            Else
                .LabelFichPerso.Caption = ""
            End If
            RemplirComboStructures uneForm
            InhiberBoutonMat uneForm
        ElseIf uneForm.monFichPersoSTR <> mesOptionsMat.monFichPersoSTR And uneForm.monFichPersoSTR <> "" And .CheckFichPerso.Value = 1 Then
            'Cas où le fichier perso existe sans être celui des options matériaux
            '==> On ne fait rien, retour comme avant le choix de travail
            'en fichier perso str ====> Utilisation structures CERTU
            .CheckFichPerso.Value = 0
            MsgBox MsgFichPersoKO1 + Chr(13) + Chr(13) + MsgFichPersoKO2, vbInformation
        End If
    End With
End Sub

Public Sub AfficherCarottes(uneForm As Form)
    Dim unTabEp As Variant
    Dim uneStruct As Structure
    Dim unMatSurf1Exist As Boolean
    Dim unMatSurf2Exist As Boolean
    Dim unMatBaseExist As Boolean
    Dim unMatFondExist As Boolean
    
    With uneForm
        'Récup de la strucutre choisie
        Set uneStruct = DonnerStructChoisie(uneForm)
        If uneStruct Is Nothing Then Exit Sub
        
        'Indication des couches existantes
        unMatSurf1Exist = Not (uneStruct.maCoucheSurface = "Aucune")
        If unMatSurf1Exist Then
            'Cas avec couche de surface
            unMatSurf2Exist = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé)
        Else
            'Cas sans couche de surface
            unMatSurf2Exist = unMatSurf1Exist
        End If
        unMatBaseExist = Not (uneStruct.maCoucheBase = "Aucune")
        unMatFondExist = Not (uneStruct.maCoucheFondation = "Aucune")
        
        'Recherche si on a une structure avec une couche de surface sans épaisseur
        uneSurfSansEpaisseur = (unMatSurf1Exist And uneStruct.maCoucheSurfSansEp = 1)
        
        'Récup du tableau de variant contenant les épaisseurs
        unTabEp = .monTabEp
        
        If .monEpQ1Trouv = False Then
            'Pas d'épaisseur trouvée pour Q1
            '==> on prend épaisseur par défaut
            'si aucune épaisseur trouvée pour Q2
            'sinon on prend les épaisseurs de Q2 d'où même carotte
            'sauf si on est en étude giratoire car pas d'épaisseur en Q2 dans ces structures
            If .monEpQ2Trouv And .OptionEtudeGiratoire.Value = False Then
                For i = 1 To 6
                    unTabEp(i) = unTabEp(i + 6)
                Next i
            Else
                'Les booleens False = 0 et True = -1
                'd'ou abs pour la valeur absolue pour avoir 0 ou 1
                If uneSurfSansEpaisseur Then
                    'Trois cm pour une bonne visu à l'écran d'une couche
                    'de surface sans épaisseur
                    unTabEp(1) = 3
                    unTabEp(2) = 0
                ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés" Then
                    'EpLitPose cm pour le lit de pose à l'écran d'une couche
                    'de surface en dalles ou pavés
                    unTabEp(1) = EpParDefaut
                    unTabEp(2) = EpLitPose
                Else
                    unTabEp(1) = EpParDefaut * Abs(unMatSurf1Exist) / (1 + Abs(unMatSurf2Exist))
                    unTabEp(2) = EpParDefaut * Abs(unMatSurf2Exist) / 2
                End If
                unTabEp(3) = EpParDefaut * Abs(unMatBaseExist)
                unTabEp(4) = 0
                unTabEp(5) = EpParDefaut * Abs(unMatFondExist)
                unTabEp(6) = 0
            End If
        End If
        
        If .monEpQ2Trouv = False Then
            'Pas d'épaisseur trouvée pour Q2
            '==> on prend épaisseur par défaut
            'si aucune épaisseur trouvée pour Q1
            'sinon on prend les épaisseurs de Q1 d'où même carotte
            If .monEpQ1Trouv Then
                For i = 7 To 12
                    unTabEp(i) = unTabEp(i - 6)
                Next i
            Else
                'Les booleens False = 0 et True = -1
                'd'ou abs pour la valeur absolue
                If uneSurfSansEpaisseur Then
                    'Trois cm pour une bonne visu à l'écran d'une couche
                    'de surface sans épaisseur
                    unTabEp(7) = 3
                    unTabEp(8) = 0
                ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés" Then
                    'EpLitPose cm pour le lit de pose à l'écran d'une couche
                    'de surface en dalles ou pavés
                    unTabEp(7) = EpParDefaut
                    unTabEp(8) = EpLitPose
                Else
                    unTabEp(7) = EpParDefaut * Abs(unMatSurf1Exist) / (1 + Abs(unMatSurf2Exist))
                    unTabEp(8) = EpParDefaut * Abs(unMatSurf2Exist) / 2
                End If
                unTabEp(9) = EpParDefaut * Abs(unMatBaseExist)
                unTabEp(10) = 0
                unTabEp(11) = EpParDefaut * Abs(unMatFondExist)
                unTabEp(12) = 0
            End If
        End If
        
        'Affectation du tableau de variant contenant les épaisseurs
        .monTabEp = unTabEp
        
        'Dessin des carottes Q1 et Q2
        DessinerCarottes uneForm, .monTabEp, .monMinTecQ1, .monMaxPraQ1, .monEpQ1Trouv, .monMinTecQ2, .monMaxPraQ2, .monEpQ2Trouv
    End With
End Sub

Public Function DonnerStructChoisie(uneForm As Form) As Structure
    'Retour de la structure choisie
    Dim uneColStruct As Collection
    
    If uneForm.ComboStruct.ListIndex = -1 Then
        Set DonnerStructChoisie = Nothing
    Else
        If uneForm.CheckFichPerso.Value = 0 Then
            Set uneColStruct = maColStructCERTU
        Else
            Set uneColStruct = maColStructPerso
        End If
        unIndStruct = uneForm.ComboStruct.ItemData(uneForm.ComboStruct.ListIndex)
        Set DonnerStructChoisie = uneColStruct(unIndStruct)
    End If
End Function

Public Function DonnerColStruct(uneForm As Form) As Collection
    'Retourne la collection de structures CERTU ou personnelles
    If uneForm.CheckFichPerso.Value = 0 Then
        Set DonnerColStruct = maColStructCERTU
    Else
        Set DonnerColStruct = maColStructPerso
    End If
End Function

Public Function OuvrirFichierMatSurface(unFileName As String, uneColMatSurf As Collection) As Boolean
    'Ouverture du fichier de matériaux de surface
    'Test de l'existence des fichiers CERTU.mts et CERTU.str
    'dans le répertoire de l'application GestionStructure
    Dim unFichMTS As Integer
    Dim unMsg As String
    Dim unTypeMat As String
    Dim uneErreur As Boolean, unMat As Object
    Dim unEntete As String, unNom As String, unAbrege As String
    Dim unComment As String, uneNorme As String
    Dim unCommentSuite As String
    Dim unYoung As Single, unPoisson As Single
    Dim unEpsilon As Single, unSigma As Single
    Dim unNbComp As Integer, unNbComposition As Integer
    Dim unTabEp(0 To 3) As Integer, unTabAbMat(1 To 3) As String
    Dim uneColComp As Collection
    Dim uneQual As String, unAGel As Single, unBGel As Single
    
    'Correction de \\ par un seul \
    unFileName = CorrigerNomFichier(unFileName)
    
    On Error GoTo ErreurFichierMTS
    unFichMTS = FreeFile(0)
    
    OuvrirFichierMatSurface = True
    
    'Test existence du fichier
    uneErreur = ("" = Dir(unFileName))
    If uneErreur = False Then
        'Traitement si aucune erreur
        'Test si fichier matériaux surface valide
        'Ouverture du fichier Matériaux de surface CERTU.MTS
        Open unFileName For Binary As #unFichMTS
        LireString unFichMTS, unEntete
        If unEntete = ENTETE_MTS Then
            'Remplissage des matériaux de surface
            Do
                LireString unFichMTS, unTypeMat
                If unTypeMat = "MatériauSimple" Or unTypeMat = "Matériau" Then
                    'Lecture fichier pour un matériau simple ou composant
                    LireString unFichMTS, unNom
                    LireString unFichMTS, unAbrege
                    LireString unFichMTS, uneNorme
                    LireString unFichMTS, uneQual
                    Get #unFichMTS, , unAGel
                    Get #unFichMTS, , unBGel
                    If unTypeMat = "Matériau" Then
                        Get #unFichMTS, , unYoung
                        Get #unFichMTS, , unPoisson
                        Get #unFichMTS, , unEpsilon
                        Get #unFichMTS, , unSigma
                    End If
                    LireString unFichMTS, unComment
                    LireString unFichMTS, unCommentSuite
                    Do While unCommentSuite <> FIN_COMMENT
                        'Corrrection du problème lié à la présence
                        'de " dans le texte RTF
                        '==> début ou fin de string en lecture
                        unComment = unComment + unCommentSuite
                        LireString unFichMTS, unCommentSuite
                    Loop
                    'Création des matériaux suivant leur type
                    If unTypeMat = "MatériauSimple" Then
                        Set unMat = New MatSimple
                        unMat.SetProps unNom, unAbrege, uneNorme, unComment
                    ElseIf unTypeMat = "Matériau" Then
                        Set unMat = New Matériau
                        unMat.SetProps unNom, unAbrege, uneNorme, unComment
                        unMat.SetPropsPhysic unYoung, unPoisson, unEpsilon, unSigma
                        'Alimentation de la collection des matériaux composants
                        maColMatComposant.Add unMat, unMat.monAbrégé
                    End If
                    RemplirQualitéGel unMat, uneQual, unAGel, unBGel
                    'Alimentation de la collection des matériaux de surface
                    maColMatSurf.Add unMat, unMat.monAbrégé
                ElseIf unTypeMat = "MatériauComposé" Then
                    'Lecture fichier pour un matériau composé
                    LireString unFichMTS, unNom
                    LireString unFichMTS, unAbrege
                    Get #unFichMTS, , unNbComposition
                    Get #unFichMTS, , unAGel
                    Get #unFichMTS, , unBGel
                    Set uneColComp = New Collection
                    unNbComp = 3
                    '3 = Nb de colonnes de composant et alimentation d'une
                    'collection contenant unNbcomp couples
                    '(épaisseur, composant) + épaisseur préconisée
                    For j = 1 To unNbComposition
                        unN = (j - 1) * (2 * unNbComp + 1) + 1
                        Get #unFichMTS, , unTabEp(0)
                        Get #unFichMTS, , unTabEp(1)
                        LireString unFichMTS, unTabAbMat(1)
                        Get #unFichMTS, , unTabEp(2)
                        LireString unFichMTS, unTabAbMat(2)
                        Get #unFichMTS, , unTabEp(3)
                        LireString unFichMTS, unTabAbMat(3)
                        uneColComp.Add unTabEp(0)
                        For i = 1 To 3
                            uneColComp.Add unTabEp(i)
                            uneColComp.Add unTabAbMat(i)
                        Next i
                    Next j
                    'Création des matériaux
                    Set unMat = New MatComposé
                    unMat.SetProps unNom, unAbrege
                    Set unMat.mesCompositions = uneColComp
                    RemplirQualitéGel unMat, "", unAGel, unBGel
                    'Alimentation de la collection des mat surface
                    maColMatSurf.Add unMat, unMat.monAbrégé
                ElseIf EOF(unFichMTS) = False Then
                    'Car unTypeMat ="" après la dernière lecture
                    'en fin de fichier
                    MsgBox MsgErreurMatériauInconnu + ": " + unTypeMat + Chr(13) + MsgFich + " " + unFileName + MsgIncorrect, vbCritical
                    OuvrirFichierMatSurface = False
                    Exit Function
                End If
            Loop Until EOF(unFichMTS)
        Else
            MsgBox MsgFich + " " + unFileName + MsgIncorrect, vbCritical
            OuvrirFichierMatSurface = False
            Exit Function
        End If
    Else
        OuvrirFichierMatSurface = False
        MsgBox MsgFich + " " + unFileName + MsgInexistant, vbCritical
    End If
    Close unFichMTS
    
    On Error GoTo 0
    Exit Function
    
ErreurFichierMTS:
    'Gestion d'erreur de lecture des fichiers
    '*.MTS
    'Cas d'erreur
    unMsg = MsgFich + " " + unFileName + MsgRunError + Format(Err.Number) + " - " + Err.Description + Chr(13)
    OuvrirFichierMatSurface = False
    MsgBox unMsg, vbCritical
    On Error GoTo 0
End Function



Public Sub AjouterInListMat(uneForm As Form, unNomMat As String)
    'Ajout dans le listview listviewmat d'un nom de matériau
    's'il n'y est pas déjà
    Dim unItemX As ListItem
    
    'Test si déjà présent dans les items du listview
    i = 1
    unDejaPresent = False
    Do While unDejaPresent = False And i <= uneForm.ListViewMat.ListItems.Count
        unDejaPresent = (uneForm.ListViewMat.ListItems(i).Text = unNomMat)
        i = i + 1
    Loop
    
    If unDejaPresent = False Then
        Set unItemX = uneForm.ListViewMat.ListItems.Add()
        unItemX.SmallIcon = 1
        unItemX.Text = unNomMat
    End If
End Sub

Public Sub AjouterComposition(uneForm As Form, uneComboBox As ComboBox, uneEpMat1 As Integer, unNomMat1 As String, uneEpMat2 As Integer, unNomMat2 As String, uneEpMat3 As Integer, unNomMat3 As String, unN As Integer)
    'Ajouter dans la combobox de composition passée en paramètre
    'une composition possible par rapport à l'épaisseur préconisée
    Dim unTitreComp As String
    
    unTitreComp = ""
    unNbMat = 0
    
    If uneEpMat1 > 0 And unNomMat1 <> "" Then
        AjouterInListMat uneForm, unNomMat1
        unTitreComp = unTitreComp + unNomMat1 + " " + Format(uneEpMat1) + " cm"
        unNbMat = 1
    End If
    
    If uneEpMat2 > 0 And unNomMat2 <> "" Then
        AjouterInListMat uneForm, unNomMat2
        unNbMat = unNbMat + 1
        If unNbMat > 1 Then unTitreComp = unTitreComp + " + "
        unTitreComp = unTitreComp + unNomMat2 + " " + Format(uneEpMat2) + " cm"
    End If
    
    If uneEpMat3 > 0 And unNomMat3 <> "" Then
        AjouterInListMat uneForm, unNomMat3
        unNbMat = unNbMat + 1
        If unNbMat > 1 Then unTitreComp = unTitreComp + " + "
        unTitreComp = unTitreComp + unNomMat3 + " " + Format(uneEpMat3) + " cm"
    End If
    
    uneComboBox.AddItem unTitreComp
    uneComboBox.ItemData(uneComboBox.ListCount - 1) = unN
End Sub

Public Sub ChangerCouleurCouches(uneForm As Form)
    'Changer les couleurs des couches des carottes Q1 et Q2 en prenant
    'les couleurs des options générales d'une Form = une étude
    With uneForm
        'Pour la qualité Q1
        .ShapeSurfQ1.FillColor = mesOptionsGen.maCoulSurf
        .ShapeBase1Q1.FillColor = mesOptionsGen.maCoulBase1
        .ShapeBase2Q1.FillColor = mesOptionsGen.maCoulBase2
        .ShapeFond1Q1.FillColor = mesOptionsGen.maCoulFond1
        .ShapeFond2Q1.FillColor = mesOptionsGen.maCoulFond2
        'Pour la qualité Q2
        .ShapeSurfQ2.FillColor = mesOptionsGen.maCoulSurf
        .ShapeBase1Q2.FillColor = mesOptionsGen.maCoulBase1
        .ShapeBase2Q2.FillColor = mesOptionsGen.maCoulBase2
        .ShapeFond1Q2.FillColor = mesOptionsGen.maCoulFond1
        .ShapeFond2Q2.FillColor = mesOptionsGen.maCoulFond2
    End With
End Sub

Public Sub AfficherFicheMat(uneStringFicheMat As String)
    'Affichage de la fiche matériau suivant le type du matériau
    'pour visulaiser ces caractéristiques
        
    'uneStringFicheMat = TypeMat + '/' + Abrégé
    'TypeMat = "Simple" ou "Composant" ou "FondBase" ou "Composé"
    
    'Chargement sans affichage
    Load FicheMat
    
    'Remplissage du tag de FicheMat avec le type de
    'de matériau et l'abrégé
    FicheMat.Tag = uneStringFicheMat
    
    'Centrage de la fiche matériau
    CentrerFenetreEcran FicheMat
    
    'Affichage modal
    FicheMat.Show vbModal
End Sub


Public Sub CalculerTraficCum(uneForm As Form)
    'Calcul du trafic cumulé dans l'onglet Trafic de la form
    'passée en paramètre
    Dim unTini As Integer, uneCA As Single, uneDS As Byte
    With uneForm
        If .TextTrafIni.Text = "" Then Exit Sub
        unTini = CInt(.TextTrafIni.Text)
        uneCA = DonnerCroissAn(uneForm) / 100
        uneDS = Val(.TextDuréeS)
        'On arrondi à l'entier supérieur par rapport à la formule
        'du cahier des charges
        .TextTrafCUM.Text = Format(CLng(365# * unTini * (uneDS + (uneCA * uneDS * (uneDS - 1)) / 2)))
        'Calcul du Nombre d'essieux équivalents si trafic cumulé connu
        uneForm.CalculerEtAfficherNE
    End With
End Sub

Public Function CalculerTraficIni(uneForm As Form) As Boolean
    'Calcul du trafic initial dans l'onglet Trafic de la form
    'passée en paramètre
    'Retourne VRAI si le trafic initial est entre les bonnes
    'valeurs min et max faux sinon
    Dim unTCum As Long, uneCA As Single
    Dim uneValMin As Long, uneValMax As Long
    Dim uneValMinTol As Long, uneValMaxTol As Long
    Dim unNomTypeVoie As String, unTypeVoie As Integer, unMsg As String
    Dim uneDS As Byte, unTextTrafIni As String
    
    With uneForm
        If .TextTrafCUM.Text = "" Then
            CalculerTraficIni = True
            Exit Function
        End If
        
        unTCum = CLng(.TextTrafCUM.Text)
        uneCA = DonnerCroissAn(uneForm) / 100
        uneDS = Val(.TextDuréeS)
        'On arrondi à l'entier supérieur par rapport à la formule
        'du cahier des charges
        unTextTrafIni = Format(CLng(unTCum / 365# / (uneDS + (uneCA * uneDS * (uneDS - 1)) / 2)))
        
        'Récup du domaine de validité suivant le type de voie
        DonnerMinMaxTraficIni uneForm, uneValMinTol, uneValMin, uneValMax, uneValMaxTol
        unMsg = ""
        unNomTypeVoie = DonnerNomTypeVoie(uneForm)
            
        'Test de la validité du trafic initial calculé
        If CLng(unTextTrafIni) < uneValMinTol Or CLng(unTextTrafIni) > uneValMaxTol Then
            'Cas d'erreur non admise dans le domaine de validité
            unMsg = MsgTraficCum + unTextTrafIni + Chr(13) + Chr(13)
            unMsg = unMsg + MsgTraficIni + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
            'Affichage du domaine de tolérance plus grande que le domaine de validité
            unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
            'Résultat invalide
            unTypeIcone = vbCritical
            CalculerTraficIni = False
        ElseIf (CLng(unTextTrafIni) >= uneValMinTol And CLng(unTextTrafIni) < uneValMin) Or (CLng(unTextTrafIni) > uneValMax And CLng(unTextTrafIni) <= uneValMaxTol) Then
            'Cas d'erreur tolérée dans le domaine de validité
            unMsg = MsgTraficCum + unTextTrafIni + Chr(13) + Chr(13)
            unMsg = unMsg + MsgTraficIni2 + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
            unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
            unTypeVoie = DonnerTypeVoie(uneForm)
            If CLng(unTextTrafIni) > uneValMax Then
                'Cas où on dépasse la valeur maxi
                If (unTypeVoie >= TypeVoieTraficLourd And unTypeVoie <= TypeVoieBus) Or (unTypeVoie = TypeGiratoireTraficLourd) Then
                    'Cas des voie bus, voie principale avec PL
                    'et Giratoire sur voie principale PL
                    'Si la valeur > au max et < max tolérée, conseiller de faire
                    'une étude en laboratoire
                    unMsg = unMsg + Chr(13) + Chr(13) + MsgValLabo + Format(uneValMax) + MsgIsLabo
                End If
            End If
            'Résultat OK
            unTypeIcone = vbInformation
            CalculerTraficIni = True
            'Mise à jour du Textbox trafic initial
            .TextTrafIni.Text = unTextTrafIni
        Else
            'Cas où on se trouve dans le domaine de validité
            CalculerTraficIni = True
            'Mise à jour du Textbox trafic initial
            .TextTrafIni.Text = unTextTrafIni
        End If
        
        If unMsg <> "" Then MsgBox unMsg, unTypeIcone
    End With
End Function

Public Function DonnerCroissAn(uneForm As Form) As Byte
    'Donner la croissance annuelle en cours d'une form (= une étude)
    'en scannant les valeurs des boutons options i%
    Dim i As Byte
    For i = 0 To 5
        'Les % vont de 0 à 5 % et le tableau de controle optionCA de 0 à 5
        If uneForm.OptionCA(i).Value Then
            DonnerCroissAn = i
            Exit For
        End If
    Next i
End Function
    
Public Sub DonnerMinMaxTraficIni(uneForm As Form, uneValMinTol As Long, uneValMin As Long, uneValMax As Long, uneValMaxTol As Long)
    'Récupération des valeurs min et max du trafic initial
    'et les min et max tolérées
    'à partir du type de voie d'une étude (= la form)
    If uneForm.OptionVoieDes.Value Then
        uneValMin = 0
        uneValMax = 25
        uneValMinTol = 0
        uneValMaxTol = 30
    ElseIf uneForm.OptionVoieDis.Value Or uneForm.OptionGirDis.Value Then
        uneValMin = 25
        uneValMax = 150
        uneValMinTol = 20
        uneValMaxTol = 170
    ElseIf uneForm.OptionVoiePL.Value Or uneForm.OptionGirPL.Value Then
        uneValMin = 150
        uneValMax = 750 '1000 modif en V2
        uneValMinTol = 120 ' 150 modif en V2
        uneValMaxTol = 1000
    Else
        'En v2 pour toutes lautres voies donc :
        'les voies réservés aux bus (aménagement standard) et parking
        uneValMin = 0
        uneValMax = 750 '1000 modif en V2
        uneValMinTol = 0
        uneValMaxTol = 1000
    End If
End Sub
    
Public Function DonnerPrecMinMaxCAM(uneForm As Form, uneValPrec As Single, uneValMin As Single, uneValMax As Single) As String
    'Récupération des valeurs préconisée, min et max du trafic initial
    'à partir du type de voie et du type de structure choisie d'une étude
    'Etude = la form et type structure = (Souple ou Bitu) ou (Hydro ou Béton)
        'CAM souple ou bitu  ==> TypeCAM = 1
        'CAM hydrau ou beton ==> TypeCAM = 2
    'et retourne le libellé du type de voie
    Dim uneStruct As Structure
    
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If uneForm.OptionVoieDes.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.1
        uneValMin = 0.05
        uneValMax = 0.4
        DonnerPrecMinMaxCAM = uneForm.OptionVoieDes.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionVoieDes.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.1
        uneValMin = 0.05
        uneValMax = 0.4
        DonnerPrecMinMaxCAM = uneForm.OptionVoieDes.Caption + MsgStructCAM_HB
    ElseIf uneForm.OptionVoieDis.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.1
        uneValMin = 0.05
        uneValMax = 0.6
        DonnerPrecMinMaxCAM = uneForm.OptionVoieDis.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionVoieDis.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.2
        uneValMin = 0.1
        uneValMax = 0.6
        DonnerPrecMinMaxCAM = uneForm.OptionVoieDis.Caption + MsgStructCAM_HB
    ElseIf uneForm.OptionVoiePL.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.2
        uneValMin = 0.1
        uneValMax = 0.8
        DonnerPrecMinMaxCAM = uneForm.OptionVoiePL.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionVoiePL.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.4
        uneValMin = 0.2
        uneValMax = 1
        DonnerPrecMinMaxCAM = uneForm.OptionVoiePL.Caption + MsgStructCAM_HB
    'Ajout pour la version 2
    ElseIf uneForm.OptionGirDis.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.5
        uneValMin = 0.2
        uneValMax = 1#
        DonnerPrecMinMaxCAM = uneForm.OptionGirDis.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionGirDis.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.5
        uneValMin = 0.2
        uneValMax = 1#
        DonnerPrecMinMaxCAM = uneForm.OptionGirDis.Caption + MsgStructCAM_HB
    ElseIf uneForm.OptionGirPL.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 1#
        uneValMin = 0.5
        uneValMax = 1.5
        DonnerPrecMinMaxCAM = uneForm.OptionGirPL.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionGirPL.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 1#
        uneValMin = 0.5
        uneValMax = 1.5
        DonnerPrecMinMaxCAM = uneForm.OptionGirPL.Caption + MsgStructCAM_HB
    ElseIf uneForm.OptionVoieParking.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.1
        uneValMin = 0.05
        uneValMax = 0.4
        DonnerPrecMinMaxCAM = uneForm.OptionVoieParking.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionVoieParking.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.1
        uneValMin = 0.05
        uneValMax = 0.4
        DonnerPrecMinMaxCAM = uneForm.OptionVoieParking.Caption + MsgStructCAM_HB
    'Modif pour la version 2
    ElseIf uneForm.OptionVoieBus.Value And uneStruct.monTypeCAM = 1 Then
        uneValPrec = 0.5
        uneValMin = 0.2
        uneValMax = 1#
        DonnerPrecMinMaxCAM = uneForm.OptionVoieBus.Caption + MsgStructCAM_SB
    ElseIf uneForm.OptionVoieBus.Value And uneStruct.monTypeCAM = 2 Then
        uneValPrec = 0.5
        uneValMin = 0.2
        uneValMax = 1#
        DonnerPrecMinMaxCAM = uneForm.OptionVoieBus.Caption + MsgStructCAM_HB
    End If
    
    'Cas particulier d'une structure de type souple, ou de type pavés/dallées avec une
    'couche de base GNT, le cam préconisé est multiplié par 2 par rapport à la valeur
    'souple/bitumineuse qui n'est valable que pour les bitumineuses
    If uneStruct.monTypeStructure = Souple Then
        uneValPrec = uneValPrec * 2
    ElseIf uneStruct.monTypeStructure = PavesDalles And UCase(uneStruct.maCoucheBase) = "GNT" Then
        uneValPrec = uneValPrec * 2
    End If
End Function


Public Sub CalculerIndiceGelRef(uneForm As Form)
    'Calcul de l'indice de gel de référence corrigé
    'pour les qualités Q1 et Q2
    Dim unIndGelRef As Integer, unCoefAgglo As Single
    Dim uneHStation As Integer, uneHAgglo As Integer
    
    'Initialisation de l'indice de gel et du coef agglo
    'pour avoir un indice de gel inconnu à la création d'une étude
    unIndGelRef = -1 'valeur d'indice de gel de référence inconnu
    unCoefAgglo = 1 'valeur neutre x * 1 = x
    
    'Récup des données de la station de référence
    uneHStation = monTabStation(uneForm.ComboStation.ListIndex + 1).monAltitude
    If uneForm.OptionHE.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHRE
    ElseIf uneForm.OptionHRNE.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHRNE
    ElseIf uneForm.OptionHC.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHC
    End If
    
    'Récup des données de l'agglo du projet
    If uneForm.ComboTailleAgglo.ListIndex = 0 Then
        unCoefAgglo = 1
    ElseIf uneForm.ComboTailleAgglo.ListIndex = 1 Then
        unCoefAgglo = 0.9
    ElseIf uneForm.ComboTailleAgglo.ListIndex = 2 Then
        unCoefAgglo = 0.8
    End If
    uneHAgglo = Format(uneForm.TextHAgglo.Text)
    
    'Calcul de l'indice de référence corrigé
    'Sans correction d'altitude pour l'instant
    'uneForm.monIndiceGelRefQ1 = CInt(unIndGelRef * unCoefAgglo * (1 - Abs(uneHStation - uneHAgglo) / uneHStation))
    uneForm.monIndiceGelRefQ1 = CInt(unIndGelRef * unCoefAgglo)
    uneForm.monIndiceGelRefQ2 = uneForm.monIndiceGelRefQ1
End Sub

Public Sub AfficherEtCalculerIndGelRef(uneForm As Form)
    'Calcul de l'indice de gel de référence
    'pour les qualités Q1 et Q2
    CalculerIndiceGelRef uneForm
    'Affichage dans la frame résultat de l'étude active
    ActualiserFrameVerifGel uneForm
End Sub

Public Sub MettreAJourOngletGel(uneForm As Form)
    With uneForm
        'Mise à jour de l'onglet gel
        
        'Partie Station de référence + Agglo du projet
        .TextHAgglo.Text = Format(.monHAgglo)
        .TextHAgglo.Tag = Format(.monHAgglo)
        .TextHAgglo.ForeColor = QBColor(0)
        RemplirLesStationsMétéo uneForm
        .ComboStation.ListIndex = .monIndStation - 1
        .ComboTailleAgglo.ListIndex = .monIndTailleAgglo
        
        'Partie hiver de référence
        If .monIndHiver = HE Then
            .OptionHE.Value = True
        ElseIf .monIndHiver = HRNE Then
            .OptionHRNE.Value = True
        ElseIf .monIndHiver = HC Then
            .OptionHC.Value = True
        End If
        
        'Partie couche de forme non gélive
        .TextEpaisseur.Text = Format(.monEpNonGel)
        .TextEpaisseur.Tag = Format(.monEpNonGel)
        .TextEpaisseur.ForeColor = QBColor(0)
        If .monCoefA = 0.12 Then
            .OptionANT.Value = True
        Else
            .OptionAT.Value = True
        End If
        
        'Partie Sol support
        If .monIndGelSol = 1 Then
            .OptionTGel.Value = True
        ElseIf .monIndGelSol = 2 Then
            .OptionPGel.Value = True
        ElseIf .monIndGelSol = 3 Then
            .OptionNGel.Value = True
        End If
        .TextPente.Text = .maPente
        .TextPente.Tag = .maPente
        If .maPente <> MsgInfinie Then
            'Modif de la borne 0.05 pour la pente, c'et maintenant non gélif
            'Avant non gélif c'était pour pente < 0.05
            If CSng(.maPente) = 0.05 Then
                .OptionNGel.Value = True
                'car la ligne précédente remet TextPente à 0
                .TextPente.Text = Format(.maPente)
            End If
        End If
        .TextPente.ForeColor = QBColor(0)
        
        'Affichage ou masquage de l'indice de gel perso
        .CheckIndGelPerso.Value = .monUtilIndGelPerso
        
        'Mise à jour du calcul de l'indice de gel admin
        AfficherEtCalculerIndGelAdm uneForm
    End With
End Sub

Public Sub RemplirLesStationsMétéo(uneForm As Form)
    'Remplissage du tableau de station de référence pour le gel
    'et Remplissage de la combobox ComboStation
        
    'Remise à zéro du nombre de stations
    monNbStation = 0

    RemplirUneStationMétéo uneForm, "Ambérieu", "01", 253, 270, 175, 65
    RemplirUneStationMétéo uneForm, "Saint Quentin", "02", 98, 225, 110, 45
    RemplirUneStationMétéo uneForm, "Vichy", "03", 249, 250, 115, 45
    RemplirUneStationMétéo uneForm, "St-Auban", "04", 459, 80, 35, 10
    RemplirUneStationMétéo uneForm, "Embrun", "05", 871, 165, 95, 60
    RemplirUneStationMétéo uneForm, "Nice", "06", 5, 0, 0, 0
    RemplirUneStationMétéo uneForm, "St-Girons", "09", 411, 120, 35, 15
    RemplirUneStationMétéo uneForm, "Romilly sur Seine", "10", 77, 210, 110, 35
    RemplirUneStationMétéo uneForm, "Carcassone", "11", 126, 85, 35, 10
    RemplirUneStationMétéo uneForm, "Millau", "12", 715, 140, 65, 40
    RemplirUneStationMétéo uneForm, "Marignane", "13", 4, 70, 15, 0
    RemplirUneStationMétéo uneForm, "Caen", "14", 64, 115, 60, 25
    RemplirUneStationMétéo uneForm, "Cognac", "16", 30, 100, 35, 15
    RemplirUneStationMétéo uneForm, "la Rochelle", "17", 4, 75, 30, 10
    RemplirUneStationMétéo uneForm, "Bourges", "18", 161, 160, 70, 30
    RemplirUneStationMétéo uneForm, "Ajaccio", "20", 4, 0, 0, 0
    RemplirUneStationMétéo uneForm, "Dijon", "21", 222, 200, 130, 65
    RemplirUneStationMétéo uneForm, "Rostrenen", "22", 262, 85, 50, 10
    RemplirUneStationMétéo uneForm, "Besançon", "25", 307, 220, 120, 70
    RemplirUneStationMétéo uneForm, "Lus-la-Croix-Haute", "26", 1059, 420, 275, 160
    RemplirUneStationMétéo uneForm, "Montélimar", "26", 73, 105, 40, 10
    RemplirUneStationMétéo uneForm, "Evreux", "27", 133, 195, 115, 60
    RemplirUneStationMétéo uneForm, "Chartres", "28", 155, 190, 100, 35
    RemplirUneStationMétéo uneForm, "Brest", "29", 96, 20, 10, 0
    RemplirUneStationMétéo uneForm, "Nîmes", "30", 59, 60, 20, 0
    RemplirUneStationMétéo uneForm, "Toulouse", "31", 148, 115, 40, 10
    RemplirUneStationMétéo uneForm, "Bordeaux", "33", 46, 95, 40, 10
    RemplirUneStationMétéo uneForm, "Montpellier", "34", 5, 55, 35, 0
    RemplirUneStationMétéo uneForm, "Dinard", "35", 58, 65, 25, 5
    RemplirUneStationMétéo uneForm, "Rennes", "35", 36, 80, 35, 10
    RemplirUneStationMétéo uneForm, "Chateauroux", "36", 156, 155, 75, 30
    RemplirUneStationMétéo uneForm, "Tours", "37", 108, 120, 75, 35
    RemplirUneStationMétéo uneForm, "Grenoble", "38", 384, 170, 145, 60
    RemplirUneStationMétéo uneForm, "Mont de Marsan", 40, 59, 100, 40, 10
    RemplirUneStationMétéo uneForm, "Romorantin", "41", 84, 135, 100, 30
    RemplirUneStationMétéo uneForm, "St-Etienne", "42", 400, 220, 110, 60
    RemplirUneStationMétéo uneForm, "le Puy", "43", 714, 240, 130, 65
    RemplirUneStationMétéo uneForm, "Nantes", "44", 26, 75, 55, 10
    RemplirUneStationMétéo uneForm, "Orléans", "45", 125, 170, 85, 45
    RemplirUneStationMétéo uneForm, "Gourdon", "46", 259, 120, 45, 20
    RemplirUneStationMétéo uneForm, "Agen", "47", 59, 110, 40, 15
    RemplirUneStationMétéo uneForm, "Angers", "49", 57, 100, 70, 15
    RemplirUneStationMétéo uneForm, "Cap de la Hague", "50", 3, 15, 5, 0
    RemplirUneStationMétéo uneForm, "Reims", "51", 94, 235, 105, 80
    RemplirUneStationMétéo uneForm, "Langres", "52", 464, 325, 170, 110
    RemplirUneStationMétéo uneForm, "St-Dizier", "52", 139, 235, 100, 65
    RemplirUneStationMétéo uneForm, "Nancy", "54", 212, 320, 155, 90
    RemplirUneStationMétéo uneForm, "Bar le Duc", "55", 279, 340, 290, 130
    RemplirUneStationMétéo uneForm, "Lorient", "56", 43, 40, 25, 10
    RemplirUneStationMétéo uneForm, "Metz", "57", 190, 290, 135, 75
    RemplirUneStationMétéo uneForm, "Château-Chinon", "58", 598, 225, 115, 80
    RemplirUneStationMétéo uneForm, "Nevers", "58", 175, 190, 110, 60
    RemplirUneStationMétéo uneForm, "Dunkerque", "59", 11, 165, 65, 20
    RemplirUneStationMétéo uneForm, "Lille", "59", 47, 250, 90, 55
    RemplirUneStationMétéo uneForm, "Beauvais", "60", 106, 215, 95, 40
    RemplirUneStationMétéo uneForm, "Alençon", "61", 144, 165, 70, 35
    RemplirUneStationMétéo uneForm, "Boulogne sur Mer", "62", 73, 165, 70, 30
    RemplirUneStationMétéo uneForm, "Clermont-Ferrand", "63", 320, 225, 115, 45
    RemplirUneStationMétéo uneForm, "Biarritz", "64", 69, 40, 10, 0
    RemplirUneStationMétéo uneForm, "Pau", "64", 183, 80, 30, 10
    RemplirUneStationMétéo uneForm, "Tarbes", "65", 360, 95, 35, 10
    RemplirUneStationMétéo uneForm, "Perpignan", "66", 42, 25, 0, 0
    RemplirUneStationMétéo uneForm, "Strasbourg", "67", 150, 410, 165, 100
    RemplirUneStationMétéo uneForm, "Mulhouse-Bâle", "68", 267, 415, 155, 105
    RemplirUneStationMétéo uneForm, "Lyon", "69", 200, 220, 110, 45
    RemplirUneStationMétéo uneForm, "Tarare", "69", 831, 275, 155, 95
    RemplirUneStationMétéo uneForm, "Luxeuil", "70", 272, 335, 165, 110
    RemplirUneStationMétéo uneForm, "Mâcon", "71", 216, 200, 115, 60
    RemplirUneStationMétéo uneForm, "Mont-St-Vincent", "71", 602, 270, 150, 95
    RemplirUneStationMétéo uneForm, "le Mans", "72", 51, 120, 70, 25
    RemplirUneStationMétéo uneForm, "Bourg-St-Maurice", "73", 865, 220, 190, 110
    RemplirUneStationMétéo uneForm, "Challes les Eaux", "73", 291, 225, 150, 60
    RemplirUneStationMétéo uneForm, "Cap de la Hève", "76", 100, 95, 60, 20
    RemplirUneStationMétéo uneForm, "Rouen", "76", 155, 130, 90, 30
    RemplirUneStationMétéo uneForm, "Melun", "77", 91, 185, 90, 50
    RemplirUneStationMétéo uneForm, "Abbeville", "80", 70, 165, 90, 50
    RemplirUneStationMétéo uneForm, "Saint-Raphaël", "83", 2, 25, 0, 0
    RemplirUneStationMétéo uneForm, "Toulon", "83", 24, 15, 0, 0
    RemplirUneStationMétéo uneForm, "Orange", "84", 83, 80, 45, 10
    RemplirUneStationMétéo uneForm, "Poitiers", "86", 117, 130, 65, 25
    RemplirUneStationMétéo uneForm, "Limoges", "87", 403, 160, 80, 30
    RemplirUneStationMétéo uneForm, "Auxerre", "89", 207, 200, 95, 55
    RemplirUneStationMétéo uneForm, "Belfort", "90", 422, 370, 175, 115
    RemplirUneStationMétéo uneForm, "Paris le Bourget", "93", 59, 160, 85, 35
End Sub

Public Sub RemplirUneStationMétéo(uneForm As Form, unNom As String, unNumDpt As String, uneAltitude As Integer, unHRE As Integer, unHRNE As Integer, unHC As Integer)
    'Remplissage de la station météo d'indice unInd dans le
    'tableau des stations météo
    'et Remplissage de la combobox ComboStation
    
    'Incrémentation du nombre de station
    'Le tableau des stations va de 1 à 84
    monNbStation = monNbStation + 1
    
    'Remplissage du tableau des stations. Il va de 1 à 84
    monTabStation(monNbStation).monNom = unNom
    monTabStation(monNbStation).monNumDpt = unNumDpt
    monTabStation(monNbStation).monAltitude = uneAltitude
    monTabStation(monNbStation).monHRE = unHRE
    monTabStation(monNbStation).monHRNE = unHRNE
    monTabStation(monNbStation).monHC = unHC
        
    'Remplissage de la combobox ComboStation. Elle va de 0 à 83
    uneForm.ComboStation.AddItem (unNom + " (" + unNumDpt + ")")
End Sub

Public Sub AfficherEtCalculerIndGelAdm(uneForm As Form)
    'Calcul de l'indice de gel admissibles
    'pour les qualités Q1 et Q2
    CalculerIndiceGelAdm uneForm
    'Affichage dans la frame résultat de l'étude active
    ActualiserFrameVerifGel uneForm
End Sub

Public Sub CalculerIndiceGelAdm(uneForm As Form)
    'Calcul de l'indice de gel admissible
    'pour les qualités Q1 et Q2
    Dim unIndGelRef As Integer, unCoefAgglo As Single
    Dim uneHStation As Integer, uneHAgglo As Integer
    Dim unTabEp As Variant, unHn As Integer
    Dim unQng As Single, unQg As Single, uneP As Single
    Dim unAH As Single, unBH As Single
    Dim unAcs As Single, unAcb As Single, unAcf As Single
    Dim unBcs As Single, unBcb As Single, unBcf As Single
    Dim uneStruct As Structure
    Dim uneColMatBF As Collection
    
    If uneForm.TextPente.Text = "" Then
        'Remise à zéro ==> Inconnu
        uneForm.monIndiceGelAdmQ1 = 0
        'Remise à zéro ==> Inconnu
        uneForm.monIndiceGelAdmQ2 = 0
        Exit Sub
    ElseIf uneForm.TextPente.Text <> MsgInfinie Then
        If Format(uneForm.TextPente.Text) <= 0.05 Then
            'Cas d'une chaussée hors gel pour les deux qualités
            'On affiche Chaussée hors gel que les épaisseurs soient
            'trouvées ou non ==> Indice de gel admissibles infini
            '==> On prend 1 milliard pour Struct-Urb
            uneForm.monIndiceGelAdmQ1 = HorsGel
            uneForm.monIndiceGelAdmQ2 = HorsGel
            Exit Sub
        End If
    End If
    
    'Coefficient a et b de gel des couches si une structure est choisie
    Set uneStruct = DonnerStructChoisie(uneForm)
    If uneStruct Is Nothing Then
        uneForm.monIndiceGelAdmQ1 = 0
        uneForm.monIndiceGelAdmQ2 = 0
        Exit Sub
    End If
    
    'Récup des bonnes listes de matériaux des couches
    If uneForm.CheckFichPerso.Value = 0 Or uneForm.LabelFichPerso.Caption = "" Then
        Set uneColMatBF = maColMatBFCERTU
    Else
        Set uneColMatBF = maColMatBFPerso
    End If
    
    'Récup de Agel
    If uneStruct.maCoucheSurface = "Aucune" Then
        unAcs = 0
    Else
        unAcs = maColMatSurf(uneStruct.maCoucheSurface).monAGel
    End If
    unAcb = uneColMatBF(uneStruct.maCoucheBase).monAGel
    If uneStruct.maCoucheFondation = "Aucune" Then
        unAcf = 0
    Else
        unAcf = uneColMatBF(uneStruct.maCoucheFondation).monAGel
    End If
    
    'Récup de Bgel
    If uneStruct.maCoucheSurface = "Aucune" Then
        unBcs = 0
    Else
        unBcs = maColMatSurf(uneStruct.maCoucheSurface).monBGel
    End If
    unBcb = uneColMatBF(uneStruct.maCoucheBase).monBGel
    If uneStruct.maCoucheFondation = "Aucune" Then
        unBcf = 0
    Else
        unBcf = uneColMatBF(uneStruct.maCoucheFondation).monBGel
    End If
    
    'Calcul de Qng
    unQng = CalculerQng(uneForm)
    
    'Calcul de Qg
    unQg = CalculerQg(uneForm)
        
    'Récup du tableau de variant contenant les épaisseurs
    unTabEp = uneForm.monTabEp
        
    If uneForm.monEpQ1Trouv Then
        'Si on a les épaisseurs pour la qualité Q1
        '===> Calcul de l'indice de gel admissible pour la qualité Q1
        'Valeur de 1 à 6 dans le tableau des épaisseurs
        'Calcul de ah
        unAH = 1 + (unAcs * (unTabEp(1) + unTabEp(2)) + unAcb * (unTabEp(3) + unTabEp(4)) + unAcf * (unTabEp(5) + unTabEp(6)))
        unBH = unBcs * (unTabEp(1) + unTabEp(2)) + unBcb * (unTabEp(3) + unTabEp(4)) + unBcf * (unTabEp(5) + unTabEp(6))
        uneForm.monIndiceGelAdmQ1 = 10 + (unAH * (uneForm.monQmQ1 + unQng + unQg) + unBH) * (unAH * (uneForm.monQmQ1 + unQng + unQg) + unBH) / 0.6
    Else
        'Remise à zéro ==> Inconnu
        uneForm.monIndiceGelAdmQ1 = 0
    End If
    
    If uneForm.monEpQ2Trouv Then
        'Si on a les épaisseurs pour la qualité Q
        '===> Calcul de l'indice de gel admissible pour la qualité Q2
        'Valeur de 7 à 12 dans le tableau des épaisseurs
        unAH = 1 + (unAcs * (unTabEp(7) + unTabEp(8)) + unAcb * (unTabEp(9) + unTabEp(10)) + unAcf * (unTabEp(11) + unTabEp(12)))
        unBH = unBcs * (unTabEp(7) + unTabEp(8)) + unBcb * (unTabEp(9) + unTabEp(10)) + unBcf * (unTabEp(11) + unTabEp(12))
        uneForm.monIndiceGelAdmQ2 = 10 + (unAH * (uneForm.monQmQ2 + unQng + unQg) + unBH) * (unAH * (uneForm.monQmQ2 + unQng + unQg) + unBH) / 0.6
    Else
        'Remise à zéro ==> Inconnu
        uneForm.monIndiceGelAdmQ2 = 0
    End If
End Sub


Public Function TrouverEpaisseurPossible(uneForm As Form)
    'Fonction retournant vrai si l'on peut trouver les épaisseurs pour Q1 et Q2
    '==> Tout est défini : le nombre d'essieux équivalent (donc le trafic cumulé
    'et le CAM), la structure et la classe de plate-forme
    Dim uneStruct As Structure
    
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If uneForm.OptionPF1.Value = False And uneForm.OptionPF2.Value = False And uneForm.OptionPF2Plus.Value = False And uneForm.OptionPF3.Value = False Then
        'Cas où la classe de plate-forme n'est pas renseignée
        TrouverEpaisseurPossible = False
    ElseIf uneStruct Is Nothing Then
        'Cas où la structure n'est pas renseignée
        TrouverEpaisseurPossible = False
    ElseIf uneForm.monNEEquiv < uneStruct.monNbEssieuxMin Or uneForm.monNEEquiv > uneStruct.monNbEssieuxMax Then
        'Cas où le NE calculé est supérieure aux NE min et max
        'de la structure choisie
        TrouverEpaisseurPossible = False
    ElseIf InStr(1, uneForm.LabelNEequiv.Caption, MsgInconnu) > 0 Then
        'Cas où le nombre d'essieux est inconnu
        TrouverEpaisseurPossible = False
    ElseIf InStr(1, uneForm.LabelNEequiv.Caption, MsgNEHorsLimite) > 0 Then
        'Cas où le nombre d'essieux est hors limite (> 10 millions)
        TrouverEpaisseurPossible = False
    Else
        'Cas où on peut trouver les épaisseurs
        TrouverEpaisseurPossible = True
    End If
End Function

Public Sub RechercherEpaisseur(uneForm As Form)
    'Recherche des épaisseurs de la structure choisie
    'correspondant au NE immédiatement supérieur
    'et affichage des carottes pour Q1 et Q2
    Dim uneStruct As Structure, unNEth As Long
    Dim unTabEp As Variant, unInd As Integer
    Dim uneColInfoPFQ1 As Collection
    Dim uneColInfoPFQ2 As Collection
    
    'Récup de la structure choisie
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If uneForm.monNEEquiv >= uneStruct.monNbEssieuxMin And uneForm.monNEEquiv <= uneStruct.monNbEssieuxMax Then
        'Cas où le NE est entre le min et le max de la structure choisie
        'Récup de la plateforme et affectation des bonnes collections d'épaisseurs
        If uneForm.OptionPF1.Value Then
            'Cas où PF = PF1
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF1Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF1Q2
        ElseIf uneForm.OptionPF2.Value Then
            'Cas où PF = PF2
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF2Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF2Q2
        ElseIf uneForm.OptionPF2Plus.Value Then
            'Cas où PF = PF2+
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF2PlusQ1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF2PlusQ2
        Else
            'Pour tous les autres cas, PF = PF3
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF3Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF3Q2
        End If
        
        'Récup du tableau de variant contenant les épaisseurs
        unTabEp = uneForm.monTabEp
        
        'Calcul du NE théorique immédiatement supérieur au NE calculé
        unNEth = uneForm.monNEEquiv
        TrouverNEthEtInd uneForm, uneStruct, unNEth, unInd
        uneForm.monNEth = unNEth
        
        'Affectation des épaisseurs dans le tableau des épaisseurs
        'De 1 à 6 pour Q1 et de 7 à 12 pour Q2
        'Avec un cas particulier pour la couche de surface où la 2ème
        'épaisseur est nulle (déterminé dans l'onglet Couche de surface)
        'unInd permet de trouver la bonne ligne d'épaisseur
        'celle correspondant au NE théorique
        For i = 1 To 6
            If i = 1 Then
                unTabEp(i) = uneColInfoPFQ1((unInd - 1) * 8 + i)
                'Affectation de l'épaisseur préconisé au cas où
                'le matériau de surface est composé
                uneForm.monEpPrecQ1 = uneColInfoPFQ1((unInd - 1) * 8 + i)
            ElseIf i = 2 Then
                unTabEp(i) = 0
            Else
                unTabEp(i) = uneColInfoPFQ1((unInd - 1) * 8 + i - 1)
            End If
        Next i
        For i = 1 To 6
            If i = 1 Then
                unTabEp(i + 6) = uneColInfoPFQ2((unInd - 1) * 8 + i)
                'Affectation de l'épaisseur préconisé au cas où
                'le matériau de surface est composé
                uneForm.monEpPrecQ2 = uneColInfoPFQ2((unInd - 1) * 8 + i)
            ElseIf i = 2 Then
                unTabEp(i + 6) = 0
            Else
                unTabEp(i + 6) = uneColInfoPFQ2((unInd - 1) * 8 + i - 1)
            End If
        Next i
        
        'Affectation des minimuns techno et maximuns pratiques
        'et les Qm
        With uneForm
            .monQmQ1 = uneColInfoPFQ1((unInd - 1) * 8 + 6)
            .monQmQ2 = uneColInfoPFQ2((unInd - 1) * 8 + 6)
            .monMinTecQ1 = uneColInfoPFQ1((unInd - 1) * 8 + 7)
            .monMinTecQ2 = uneColInfoPFQ2((unInd - 1) * 8 + 7)
            .monMaxPraQ1 = uneColInfoPFQ1((unInd - 1) * 8 + 8)
            .monMaxPraQ2 = uneColInfoPFQ2((unInd - 1) * 8 + 8)
            If UCase(.monMaxPraQ1) = "OUI" Then
                .monEpQ1Trouv = False
            Else
                .monEpQ1Trouv = True
            End If
            If UCase(.monMaxPraQ2) = "OUI" Then
                .monEpQ2Trouv = False
            Else
                .monEpQ2Trouv = True
            End If
        End With
        
        'Cas de couches de surfaces particulières
        If uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1 Then
            'Cas d'une couche de surface dont l'épaisseur n'a pas d'intêret
            'On affecte une couche de 3 cm pour une bonne visu à l'écran
            unTabEp(1) = 3 'épaisseur surface pour Q1
            unTabEp(7) = 3 'épaisseur surface pour Q2
        ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés" Then
            'Cas d'une couche de surface en dalles ou pavés
            'rajout du lit de pose
            unTabEp(2) = EpLitPose 'épaisseur lit de pose pour Q1
            unTabEp(8) = EpLitPose 'épaisseur lit de pose pour Q2
        End If
    
        'Affectation du tableau de variant contenant les épaisseurs
        uneForm.monTabEp = unTabEp
        
        'Mise à jour de l'affichage des carottes Q1 et Q2
        uneForm.monMSComp1Q1 = ""
        uneForm.monMSComp2Q1 = ""
        uneForm.monMSComp1Q2 = ""
        uneForm.monMSComp2Q2 = ""
        AfficherCarottes uneForm
    Else
        'Epaisseurs Q1 et Q2 non trouvées mais normalement on ne passe jamais là
        uneForm.monEpQ1Trouv = False
        uneForm.monEpQ2Trouv = False
        'Mise à jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes uneForm
    End If
    
    'Calcul de l'indice de gel admissible
    If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm uneForm
End Sub


Public Sub ValiderOngletCoucheSurface(uneForm As Form)
    'Procédure activant ou inhibant l'onglet de couche de surface
    'et le mettant à jour si il devient actif
    Dim uneStruct As Structure
    Dim unMatSurf As Object
    
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If Not (uneStruct Is Nothing) Then
        'Si matériau de surface composé et épaisseurs trouvées
        'pour Q1 et/ou Q2 ==> Activation de l'onglet Couche de surface
        If uneStruct.maCoucheSurface = "Aucune" Then
            uneForm.TabData.TabEnabled(OngletSurf) = False
        Else
            Set unMatSurf = maColMatSurf(uneStruct.maCoucheSurface)
            uneForm.TabData.TabEnabled(OngletSurf) = (TypeOf unMatSurf Is MatComposé) And (uneForm.monEpQ1Trouv Or uneForm.monEpQ2Trouv)
            If uneForm.TabData.TabEnabled(OngletSurf) Then MettreAJourOngletCoucheSurface uneForm, uneForm.monEpPrecQ1, uneForm.monEpPrecQ2
        End If
    End If
End Sub

Public Sub TrouverNEthEtInd(uneForm As Form, uneStruct As Structure, unNEth As Long, unInd As Integer)
    'Procédure trouvant le NE théorique, celui immédiatment supérieur
    'au NE calculé dans les NE de la structure choisie et son indice
    'dans les collections de données PFi-Qj
    'Ces valeurs sont retournées dans unNEth et unInd
    Dim unNEth0 As Long, unNETrouv As Boolean
    Dim unPas As Long
    
    unInd = 0
    unNEth0 = uneStruct.monNbEssieuxMin
    Do
        unInd = unInd + 1
        'Nombre de chiffres moins 1
        unM = Len(Format(unNEth0)) - 1
        unPas = Exp(unM * Log(10)) '10 exposant unM
        If unNEth > unNEth0 Then
            unNEth0 = unNEth0 + unPas
        Else
            unNEth = unNEth0
            unNETrouv = True
        End If
    Loop Until unNETrouv
End Sub

Public Sub MettreAjourOngletPlateForme(uneForm As Form)
    'Mise à jour de l'onglet PlateForme
    If uneForm.monIndicePF = 1 Then
        uneForm.OptionPF1.Value = True
    ElseIf uneForm.monIndicePF = 2 Then
        uneForm.OptionPF2.Value = True
    ElseIf uneForm.monIndicePF = 3 Then
        uneForm.OptionPF3.Value = True
    ElseIf uneForm.monIndicePF = 4 Then
        uneForm.OptionPF2Plus.Value = True
    End If
End Sub

Public Function EstNouvelleEtude(uneForm As Form) As Boolean
    'Fonction retournant vrai si l'étude active est une nouvelle
    'et faux si c'est une étude dèjà existante donc stockée dans un fichier URB
    If Val(Mid(uneForm.Caption, 7, 1)) > 0 Then
        'Cas d'une nouvelle étude ==> Titre de fenêtre = Etude N (N un entier > 0)
        EstNouvelleEtude = True
    Else
        'Cas d'une étude existante ==> Titre de fenêtre = Etude + nom du fichier
        EstNouvelleEtude = False
    End If
End Function

Public Function DonnerNomTypeVoie(uneForm As Form) As String
    'Retourne le nom du type de voie de l'étude active
    If uneForm.OptionVoieDes.Value Then
        DonnerNomTypeVoie = uneForm.OptionVoieDes.Caption
    ElseIf uneForm.OptionVoieDis.Value Then
        DonnerNomTypeVoie = uneForm.OptionVoieDis.Caption
    ElseIf uneForm.OptionVoiePL.Value Then
        DonnerNomTypeVoie = uneForm.OptionVoiePL.Caption
    ElseIf uneForm.OptionVoieBus.Value Then
        DonnerNomTypeVoie = uneForm.OptionVoieBus.Caption
    'Rajout de type de voies pour version 2
    ElseIf uneForm.OptionVoieParking.Value Then
        DonnerNomTypeVoie = uneForm.OptionVoieParking.Caption
    ElseIf uneForm.OptionGirDis.Value Then
        DonnerNomTypeVoie = uneForm.OptionGirDis.Caption
    ElseIf uneForm.OptionGirPL.Value Then
        DonnerNomTypeVoie = uneForm.OptionGirPL.Caption
    Else
        DonnerNomTypeVoie = ""
    End If
End Function


Public Function DonnerCouleurCouche(unInd As Byte) As Long
    'Retourne la couleur d'une couche suivant son index
    If unInd = 1 Or unInd = 2 Then
        DonnerCouleurCouche = mesOptionsGen.maCoulSurf
    ElseIf unInd = 3 Then
        DonnerCouleurCouche = mesOptionsGen.maCoulBase1
    ElseIf unInd = 4 Then
        DonnerCouleurCouche = mesOptionsGen.maCoulBase2
    ElseIf unInd = 5 Then
        DonnerCouleurCouche = mesOptionsGen.maCoulFond1
    Else
        'Sinon Couleur de la couche de fondation 2
        DonnerCouleurCouche = mesOptionsGen.maCoulFond2
    End If
End Function

Public Function CalculerQng(uneForm As Form) As Single
    'Calcul de Qng
    unHn = Val(uneForm.TextEpaisseur.Text)
    CalculerQng = DonnerCoefA(uneForm) * (unHn * unHn) / (unHn + 10)
End Function

Public Function CalculerQg(uneForm As Form) As Single
    'Calcul de Qg
    If uneForm.TextPente.Text = MsgInfinie Then
        uneP = 2
    Else
        uneP = Format(uneForm.TextPente.Text)
    End If
    
    If uneP <= 0.25 Then
        CalculerQg = 4
    ElseIf uneP <= 1 Then
        CalculerQg = 1 / uneP
    Else
        CalculerQg = 0
    End If
End Function

Public Sub ChangerHelpID(unNumOnglet As Integer)
    'Changer des contextes Id de l'aide
    'en fonction de l'onglet courant du site actif
    'de la MDI, de la fille active et de ses textbox
    'TitreEtude et DuréeCycle
    Dim unHelpId As Integer
    
    Select Case unNumOnglet
        Case OngletVoie ' Onglet voie
            unHelpId = IDhlp_OngletVoie
        Case OngletTrafic ' Onglet Trafic
            unHelpId = IDhlp_OngletTrafic
        Case OngletStruct ' Onglet Structure
            unHelpId = IDhlp_OngletStructure
        Case OngletCAM ' Onglet CAM
            unHelpId = IDhlp_OngletCAM
        Case OngletPF ' Onglet PlateForme
            unHelpId = IDhlp_OngletPlateForme
        Case OngletSurf ' Onglet Couche de surface
            unHelpId = IDhlp_OngletCoucheSurf
        Case OngletGel ' Onglet Gel
            unHelpId = IDhlp_OngletGel
    End Select
    
    'Affectation du nouveau contexte d'aide
    fMainForm.HelpContextID = unHelpId
    monEtude.HelpContextID = unHelpId
    monEtude.RichTextAide.HelpContextID = unHelpId
    If (monEtude.ActiveControl Is Nothing) = False Then
        monEtude.ActiveControl.HelpContextID = unHelpId
    Else
        'Pour l'onglet courant est le focus
        'et ainsi le F1 déclenche la bonne aide
        monEtude.TabData.SetFocus
    End If
End Sub
