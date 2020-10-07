Attribute VB_Name = "ModuleMain"
'Type pour les caract�ristiques d'une station m�t�o
'de r�f�rence pour la v�rif au gel
Type StationM�t�o
    monNom As String * 20
    monNumDpt As String * 3  'Num�ro de d�partement
    monAltitude As Integer
    monHRE As Integer       'Hiver Rigoureux Exceptionnel
    monHRNE As Integer      'Hiver Rigoureux Non Exceptionnel
    monHC As Integer        'Hiver Courant
End Type

'Type pour les options sur les mat�riaux de base
'ou de fondation (cf Fen�tre VB frmOptionsMat)
Type OptionsMat
    monFichPersoSTR As String 'Fichier de structures personnelles
    mesMatPersoNonAutoris�s As String 'String contenant les indices des
        'mat�riaux non autoris�s avec un blanc � la fin comme s�parateur
        'dans la collection de mat�riaux de base/fondation personnels
    mesMatCERTUNonAutoris�s As String 'String contenant les indices des
        'mat�riaux non autoris�s avec un blanc � la fin comme s�parateur
        'dans la collection de mat�riaux de base/fondation du CERTU
End Type

'Type pour les options g�n�rales (cf Fen�tre VB frmOptionsGen)
Type OptionsGen
    maDur�eService As Byte
    maCroisAnnuel As Byte
    maCoulSurf As Long
    maCoulBase1 As Long
    maCoulBase2 As Long
    maCoulFond1 As Long
    maCoulFond2 As Long
    maVerifGel As Byte
    monIndStationRef As Byte 'Indice de la station m�t�o de r�f�rence
    'Champs suivants pour les caract�ristiques d'une agglom�ration
    'o� l'on veut faire une v�rif au gel
    monAltiAgglo As Integer
    monTailleAgglo As Integer '0 ==> Inf � 100 000 Hab.
                              '1 ==> Entre 100 000 et 1 000 000 Hab.
                              '2 ==> Sup � 1 000 000 Hab.
    monCoefAgglo As Single    '1 ==> Inf � 100 000 Hab.
                              '0.9 ==> Entre 100 000 et 1 000 000 Hab.
                              '0.8 ==> Sup � 1 000 000 Hab.
End Type

'Collection contenant toutes les donn�es lues dans un fichier .urb
Public maColLectFich As New Collection

'Variable indiquant si l'ouverture de la fenetre
'a �t� OK (Form_Initialize event sans erreur)
Public monOuverture As Boolean

'Variable contenant les stations m�t�o de r�f�rences
Public monTabStation(1 To 84) As StationM�t�o
Public monNbStation As Integer 'Valeur initiale = 0

'Variable contenant les options g�n�rales
Public mesOptionsGen As OptionsGen
'Variable contenant les options mat�riaux
Public mesOptionsMat As OptionsMat

'Variables contenant les collections contenant
'les mat�riaux de surface
Public maColMatSurf As New Collection
Public maColMatComposant As New Collection

'Variables contenant les collections contenant les structures
'et les mat�riaux de base et de fondation du CERTU
Public maColStructCERTU As New Collection
Public maColMatBFCERTU As New Collection

'Variables contenant les collections contenant les structures
'et les mat�riaux de base et de fondation personnels
Public maColStructPerso As New Collection
Public maColMatBFPerso As New Collection

'Variable globale contenant la fen�tre m�re MDI
Public fMainForm As frmMain
'Variable globale contenant la fen�tre fille de l'�tude en cours
Public monEtude As Form

'Variable indiquant si on ouvre une nouvelle �tude
Public maNewEtude As Boolean

'Constante donnant l'indice de la station de LYON
'dans le tableau des stations
Public Const IndiceStationLYON As Integer = 65

'Constante donnant l'�paisseur totale maximale des carottes Q1 et Q2
Public Const EpTotMaxEcran As Integer = 3575
'c'est 65 *55, 65 twips pour 1 cm, �paisseur max totale 55 cm

'Constantes d'ent�tes de fichier MTS et STR
Public Const ENTETE_MTS As String = "Fichier de mat�riaux de surfaces"
Const ENTETE_STR_v100 As String = "Fichier de structures de chauss�es"
'Compatible avec les versions Struct-urb <= 1.00.0002
Const ENTETE_STR_v103 As String = "Fichier de structures de chauss�es pour Struct-Urb version >= 1.00.0003"
Const ENTETE_STR_v200 As String = "Fichier de structures de chauss�es pour Struct-urb version >= 2.0.0"
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

'Constantes pour l'�paisseur par d�faut en cm dans une carotte
Public Const EpParDefaut As Integer = 10

'Constantes pour l'�paisseur du lit de pose en cm dans une carotte
Public Const EpLitPose As Integer = 4

'Constantes les hivers de r�f�rence
Public Const HE As Integer = 1
Public Const HRNE As Integer = 2
Public Const HC As Integer = 3

'Constantes pour le type de condition de chantier
Public Const TypeChantierQ1 As Integer = 1
Public Const TypeChantierQ2 As Integer = 2

'Constantes pour le type d'�tude
Public Const TypeEtudeStandard As Integer = 1
Public Const TypeEtudeGiratoire As Integer = 2

'Constantes pour le sol support
Public Const TresGelif As Integer = 1
Public Const PeuGelif As Integer = 2
Public Const NonGelif As Integer = 3

'Constantes pour la couche de forme non g�live
Public Const NonTrait� As Integer = 1
Public Const Trait� As Integer = 2

'Constantes pour les formats de fichiers *.urb de Struct-Urb
'Pour la version finale, rajout d'une ligne contenant l'indice de gel perso
'et l'�tat de la case � cocher correspondante dans l'onglet gel
'Pour la version b�ta = sites pilotes cette ligne n'existe pas
Public Const FormatFichierVersionBeta As Byte = 1
Public Const FormatFichierVersionFinale As Byte = 2

'Constante donnant une valeur enti�re � rajouter au type de voie stock� dans
'le fichier *.urb lors de la sauvegarde pour savoir si on est en qualit� de
'chantier Standard(Q1) (typevoie < 100) ou difficile (Q2) (typevoie > 100)
Public Const ChantierDifficile As Byte = 100

'Constante pour marquer la fin du titre de l'�tude dans un fichier *.urb
Public Const FinTitre As String = "###FinTitre###"

'Constante pour indiquer chauss�e hors gel
Public Const HorsGel As Long = 1000000000

'Constantes indiquant le type de structures
Public Const ToutType As Byte = 0
Public Const Souple As Byte = 1
Public Const Bitumineuse As Byte = 2
Public Const GTLH As Byte = 3
Public Const Beton As Byte = 4
Public Const Mixte As Byte = 5
Public Const PavesDalles As Byte = 6

'Constante indiquant que l'on a chang� de type de structure et
'que la liste d�roulante de l'onglet structure permettant d'en choisir
'une. On prend -2 car les index vont de 0 � NbElementsListe-1 et -1 sert � dire
'qu'il n'y a rien de s�lectionner dans une combobox vb
Public Const ChangeTypeStruct As Integer = -2

'Constante pour les id de l'aide
Public Const IDhlp_VerifGel As Integer = 118 ' ch01s12.htm Partie Aide sur la v�rif au gel
Public Const IDhlp_OngletVoie As Integer = 215 'ch02s07s01 Onglet Voie
Public Const IDhlp_OngletTrafic As Integer = 216 'ch02s07s02 Onglet Trafic
Public Const IDhlp_OngletStructure As Integer = 218 'ch02s07s04 Onglet Structure
Public Const IDhlp_OngletCAM As Integer = 219 'ch02s07s05 Onglet CAM
Public Const IDhlp_OngletPlateForme As Integer = 217 'ch02s07s03 Onglet PlateForme
Public Const IDhlp_OngletCoucheSurf As Integer = 220 'ch02s07s06 Onglet Couche de surface
Public Const IDhlp_OngletGel As Integer = 221 'ch02s07s07 Onglet Gel

Public Const IDhlp_WinAbout As Integer = 236 'ch02s11s04 Fen�tre a propos
Public Const IDhlp_WinPrint As Integer = 210 'ch02s04s06 Fen�tre Impression
Public Const IDhlp_WinOptionsGen As Integer = 226 'ch02s09s01 Fen�tre Options g�n�rales
Public Const IDhlp_WinOptionsMat As Integer = 72 ' Fen�tre Options Mat�riaux

Public Const IDhlp_NewSite As Integer = 205 'ch02s04s01 menu nouveau
Public Const IDhlp_OpenSite As Integer = 206 'ch02s04s02 menu ouvrir
Public Const IDhlp_SaveSite As Integer = 208 'ch02s04s04 menu sauver
Public Const IDhlp_SaveAsSite As Integer = 209 'ch02s04s05 menu sauver sous
Public Const IDhlp_CloseSite As Integer = 207 'ch02s04s03 menu fermer

Sub Main()

    'R�cup du s�parateur d�cimale . ou ,
    'fix� dans les param�tres r�gionaux de Windows
    TrouverCaract�reD�cimalUtilis�
    
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
    
        ' V�rification de l'enregistrement
        If ProtectCheck("its00+-k") = "its00+-k" Then
        ' Affichage de la feuille principale
        'Lancement de la fenetre MDI m�re de l'application
            Set fMainForm = New frmMain
            fMainForm.Show
        Else 'la licence n'a pas �t� valid�e on ferme
            End
        End If
        '********************************
    End If
End Sub


Public Sub FermerFenetre(uneForm As Form)
    'Permet de fermer la fen�tre pass�e en param�tres
    'et de remettre au premier plan la fen�tre fille courante
    'Correction d'un bug dans la gestin des fen�tre actives
    'Windows si on ouvre un s�lectionneur de fichier, d'imprimante
    'ou de fontes.
    Unload uneForm
    fMainForm.Show
End Sub

Public Sub CentrerFenetreEcran(uneForm As Form)
    'Centrage d'une fenetre (= une Form VB) � l'�cran
    uneForm.Top = (Screen.Height - uneForm.Height) / 2
    uneForm.Left = (Screen.Width - uneForm.Width) / 2
End Sub
 
Public Sub TrouverCoefCorrecteurAgglo(unTypeTaille As Integer)
    'Affectation du coefficient correcteur de l'Agglo
    'pour la v�rif au gel suivant son type de taille
    If unTypeTaille = 0 Then
        '< � 100 000 Hab.
        mesOptionsGen.monCoefAgglo = 1
    ElseIf unTypeTaille = 1 Then
        'entre 100 000 et 1 000 000 Hab.
        mesOptionsGen.monCoefAgglo = 0.9
    ElseIf unTypeTaille = 2 Then
        '> � 1 000 000 Hab.
        mesOptionsGen.monCoefAgglo = 0.8
    Else
        MsgBox MsgErreurProg + MsgErreurTailleAgglo + MsgIn + "TrouverCoefCorrecteurAgglo", vbCritical
    End If
End Sub

Public Sub R�cup�rerOptionsGen()
    'R�cup�ration des options g�n�rales par lecture des valeurs
    'de ces options stock�es dans la base de registre
    
    mesOptionsGen.maDur�eService = GetSetting(App.Title, "OptionsGen", "Dur�eService", 20)
    mesOptionsGen.maCroisAnnuel = GetSetting(App.Title, "OptionsGen", "CroissAnnuel", 1)
    mesOptionsGen.maCoulSurf = GetSetting(App.Title, "OptionsGen", "CouleurSurf", QBColor(11))
    mesOptionsGen.maCoulBase1 = GetSetting(App.Title, "OptionsGen", "CouleurBase1", QBColor(2))
    mesOptionsGen.maCoulBase2 = GetSetting(App.Title, "OptionsGen", "CouleurBase2", QBColor(10))
    mesOptionsGen.maCoulFond1 = GetSetting(App.Title, "OptionsGen", "CouleurFond1", QBColor(12))
    mesOptionsGen.maCoulFond2 = GetSetting(App.Title, "OptionsGen", "CouleurFond2", QBColor(13))
    
    mesOptionsGen.maVerifGel = GetSetting(App.Title, "OptionsGen", "VerifGel", 1)
    
    'Affectation de l'Agglom�ration des �tudes
    mesOptionsGen.monTailleAgglo = GetSetting(App.Title, "OptionsGen", "TailleAgglo", 2)
    mesOptionsGen.monAltiAgglo = GetSetting(App.Title, "OptionsGen", "AltiAgglo", 200)
    TrouverCoefCorrecteurAgglo mesOptionsGen.monTailleAgglo
    
    'Affectation de la station de r�f�rence par d�faut
    mesOptionsGen.monIndStationRef = GetSetting(App.Title, "OptionsGen", "StationRef", IndiceStationLYON)
End Sub

Public Sub R�cup�rerOptionsMat()
    'R�cup�ration des options mat�riaux par lecture des valeurs
    'de ces options stock�es dans la base de registre
    
    'Affectation du nom de fichier de structures personnelles
    mesOptionsMat.monFichPersoSTR = GetSetting(App.Title, "OptionsMat", "FichierPersoSTR", "")
    
    'R�cup�ration de la chaine de caract�res contenant les indices
    'des mat�riaux de base et de fondation non autoris�s avec un blanc �
    'la fin comme s�parateur (une dizaine de mat�riaux au maximum)
    'Au d�but chaine vide car tous autoris�s
    mesOptionsMat.mesMatCERTUNonAutoris�s = GetSetting(App.Title, "OptionsMat", "MatCERTUNonAutoris�s", "")
    mesOptionsMat.mesMatPersoNonAutoris�s = GetSetting(App.Title, "OptionsMat", "MatPersoNonAutoris�s", "")
End Sub

Public Sub StockerOptionsMat()
    'Stockage des options mat�riaux dans la base de registre
    
    'Stockage du nom de fichier de structures personnelles
    SaveSetting App.Title, "OptionsMat", "FichierPersoSTR", mesOptionsMat.monFichPersoSTR
    
    'stockage de la chaine de caract�res contenant les indices
    'des mat�riaux de base et de fondation non autoris�s avec un blanc �
    'la fin comme s�parateur (une dizaine de mat�riaux au maximum)
    SaveSetting App.Title, "OptionsMat", "MatCERTUNonAutoris�s", mesOptionsMat.mesMatCERTUNonAutoris�s
    SaveSetting App.Title, "OptionsMat", "MatPersoNonAutoris�s", mesOptionsMat.mesMatPersoNonAutoris�s
End Sub

Public Sub StockerOptionsGen()
    'Stockage des options g�n�rales dans la base de registre
    With mesOptionsGen
        SaveSetting App.Title, "OptionsGen", "Dur�eService", .maDur�eService
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
    'Choix de la couleur parmi les couleurs syst�mes disponibles
    'pour la PictureBox pass�e en param�tre
    With fMainForm
          ' Attribue � CancelError la valeur True
          .dlgCommonDialog.CancelError = True
          On Error GoTo ErrHandler
          ' D�finit la propri�t� Flags
          .dlgCommonDialog.Flags = cdlCCRGBInit
          ' Affiche la bo�te de dialogue Couleur
          .dlgCommonDialog.ShowColor
          ' Attribue � l'arri�re-plan de la feuille la
          ' couleur s�lectionn�e
          unePicCouleur.BackColor = .dlgCommonDialog.Color
    End With
      
    Exit Sub

ErrHandler:
    ' L'utilisateur a cliqu� sur Annuler
    'On ne fait rien
End Sub

Public Function OuvrirFichierStructures(unFileName As String, uneColStruct As Collection, uneColMat As Collection) As Boolean
    'Ouverture d'un fichier structure et remplissage
    'de la collection de structures pass�e en param�tre
    'et de la collection de mat�riaux de couche de base ou fondation
    'pass�e en param�tre
    
    'Retourne :
    '       - TRUE si pas d'erreur � la lecture
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
    Dim uneSaisieCompl�te As Boolean, unNbTot As Long, unNumIndex As Integer
    Dim uneSizeFichSTR As Long, uneWinWait As frmWaitBar
    Dim unTypeChaussee As Byte, unTypeStructure As Byte, unUtilVParking As Integer
    Dim unUtilGDis As Integer, unUtilGPL As Integer
    
    'Correction de \\ par un seul \
    unFileName = CorrigerNomFichier(unFileName)
    'Calcul de la taille du fichier structure
    uneSizeFichSTR = FileLen(unFileName)
    
    'Ouverture de la fen�tre d'atente pour charger les structures
    'nouveau depuis la version 2, en r�seau cela permet de patienter
    'lors de l'ouverture de Struct-Urb
    Set uneWinWait = New frmWaitBar
    uneWinWait.LabelWait = LabelWaitLoadStructures
    uneWinWait.Show
    DoEvents
    
    'D�marrage de la gestion des erreurs
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
        'Avertissement que cette version de Struct-Urb >= � 2.0.0
        'n'est pas compatible avec le fichier structure des versions <= 2.00.0000
        uneWinWait.Hide 'on cache la fen�tre d'attente
        unMsg = "Le fichier structure " + unFileName + ", de version 1.0, n'est plus compatible avec " + fMainForm.Caption
        unMsg = unMsg + Chr(13) + "Il faut r�installer " + fMainForm.Caption + " ou contacter le Certu."
        MsgBox unMsg, vbCritical
        'Fonction retourne faux car erreur
        OuvrirFichierStructures = False
    ElseIf unEntete = ENTETE_STR_v200 Then
    'ElseIf unEntete = ENTETE_STR_v103 Then
        'Remplissage des structures et des mat�riaux de couches
        'de base ou de fondation pour un format de fichier structure
        'compatible avec Struct-Urb version >= 2.0.0000
        Do
            'Affichage de la progression de la lecture du fichier structure
            uneWinWait.ProgressBar1.Value = Int(Seek(unFichSTR) / (uneSizeFichSTR) * 100)
            'Lecture d 'un �l�ment du fichier structure
            LireString unFichSTR, unTypeMat
            If unTypeMat = "Structure" Then
                'Lecture fichier pour une structure de chauss�es
                Get #unFichSTR, , unNumIndex
                LireString unFichSTR, unAbrege
                LireString unFichSTR, uneCouSurf
                LireString unFichSTR, uneCouBase
                LireString unFichSTR, uneCouFond
                Get #unFichSTR, , uneCoucheSurfSansEp
                Get #unFichSTR, , uneSaisieCompl�te
                Get #unFichSTR, , unUtilVDes
                Get #unFichSTR, , unUtilVDis
                Get #unFichSTR, , unUtilVPL
                Get #unFichSTR, , unUtilVBus
                
                'Lecture de donn�es suppl�mentaires pour le fichier structure v2
                'les nouveaux types de voies
                Get #unFichSTR, , unUtilVParking
                Get #unFichSTR, , unUtilGDis
                Get #unFichSTR, , unUtilGPL
                Get #unFichSTR, , unTypeChaussee
                
                Get #unFichSTR, , unTauxRisque
                Get #unFichSTR, , unTypeCAM
                
                'Lecture de donn�es suppl�mentaires pour le fichier structure v2
                'le type de structure (souple, bitumineuse, GTLH,...)
                Get #unFichSTR, , unTypeStructure
                
                Get #unFichSTR, , unNbEssieuxMin
                Get #unFichSTR, , unNbEssieuxMax
                LireString unFichSTR, unComment
                LireString unFichSTR, unCommentSuite
                Do While unCommentSuite <> FIN_COMMENT
                    'Corrrection du probl�me li� � la pr�sence
                    'de " dans le texte RTF
                    '==> d�but ou fin de string en lecture
                    unComment = unComment + unCommentSuite
                    LireString unFichSTR, unCommentSuite
                Loop
                
                'Cr�ation des structures
                Set uneStruct = New Structure
                uneStruct.monNumIndex = unNumIndex
                uneStruct.monAbr�g� = unAbrege
                uneStruct.SetPropsInfo uneSaisieCompl�te, unUtilVDes, unUtilVDis, unUtilVPL, unUtilVBus, unComment, unTauxRisque, unNbEssieuxMin, unNbEssieuxMax, unTypeCAM
                uneStruct.SetComposition uneCouSurf, uneCouBase, uneCouFond, uneCoucheSurfSansEp
                
                'Stockage des infos rajout�s en version 2
                uneStruct.SetPropsInfoV2 unTypeChaussee, unTypeStructure, unUtilVParking, unUtilGDis, unUtilGPL
                
                'Lecture des donn�es par plate-forme
                'et alimentation des col PF de la nouvelle structure
                ' on passe de 6 � 8 en version 2
                For i = 1 To 8
                    Set uneColPF = DonnerColPF(uneStruct, i)
                    Get #unFichSTR, , unNbTot
                    For j = 1 To (unNbTot \ 8) '8 colonnes de donn�es par plate-forme
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
            
            ElseIf unTypeMat = "Mat�riauFondBase" Then
                'Lecture fichier pour un mat�riau de couche de
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
                    'Corrrection du probl�me li� � la pr�sence
                    'de " dans le texte RTF
                    '==> d�but ou fin de string en lecture
                    unComment = unComment + unCommentSuite
                    LireString unFichSTR, unCommentSuite
                Loop
                'Cr�ation des mat�riaux
                Set unMat = New Mat�riau
                unMat.SetProps unNom, unAbrege, uneNorme, unComment
                unMat.SetPropsPhysic unYoung, unPoisson, unEpsilon, unSigma
                RemplirQualit�Gel unMat, uneQual, unAGel, unBGel
                'Alimentation de la collection mat�riau base/fondation
                uneColMat.Add unMat, unMat.monAbr�g�
            ElseIf EOF(unFichSTR) = False Then
                'Cas o� ce n'est pas la fin du fichier
                uneWinWait.Hide 'on cache la fen�tre d'attente
                MsgBox MsgErreurMat�riauInconnu + ": " + unTypeMat + Chr(13) + MsgFichStruct + unFileName + MsgIncorrect, vbCritical
                'Sortie du programme
                OuvrirFichierStructures = False
                Close unFichSTR
                Exit Function
            End If
        Loop Until EOF(unFichSTR)
        'Valeur de retour mis � vrai car tout est OK
        OuvrirFichierStructures = True
    Else
        uneWinWait.Hide 'on cache la fen�tre d'attente
        MsgBox MsgFichStruct + unFileName + MsgIncorrect, vbCritical
        'Fonction retourne faux car erreur
        OuvrirFichierStructures = False
    End If
            
    'Fermeture du fichier structure
    Close unFichSTR
    
    'Fermeture de la fen�tre d'attente du chargement des structures
    Unload uneWinWait
    DoEvents
    
    'Sortie du programme et fin de la gestion d'erreur
    On Error GoTo 0
    Exit Function
    
erreurfichier_str:
    'Gestion d'erreur de lecture des fichiers structures *.STR
    unMsg = unMsg + Chr(13) + MsgFichStruct + unFileName + MsgRunError + Format(Err.Number) + " - " + Err.Description
    If Err.Number <> 53 Then
        'Cas d'une erreur diff�rente de fichier introuvable
        unMsg = unMsg + Chr(13) + Chr(13) + MsgFichStruct + MsgIncorrect
    End If
    OuvrirFichierStructures = False
    'Fermeture de la fen�tre d'attente du chargement des structures
    Unload uneWinWait
    DoEvents
    'Affichage de l'erreur survenue
    MsgBox unMsg, vbCritical
    'Sortie du programme
    Close unFichSTR
    
    On Error GoTo 0
    Exit Function
End Function


Public Sub RemplirSpreadMat(unSpreadMat As vaSpread, uneColMat As Collection, uneChaineIndicesNonAutoris�s As String)
    'Remplir un Spread avec les mat�riaux base/fondation
    'avec leur autorisation d'utilisation
    unSpreadMat.MaxRows = uneColMat.Count
    For i = 1 To uneColMat.Count
        unSpreadMat.Row = i
        'Remplissage de l'abr�g� du mat�riau
        unSpreadMat.Col = 1
        unSpreadMat.Text = uneColMat(i).monAbr�g�
        'Remplissage de son autorisation d'utilisation
        'par recherche dans une chaine contenant les indices
        'des mat�riaux non autoris�s s�par�s par des blancs,
        'm�me le dernier est suivi d'un blanc
        unSpreadMat.Col = 2
        If uneColMat(i).monUtilisationAutoris�e Then
            'Cas d'un indice de mat�riau autoris�
            unSpreadMat.Value = 1
        Else
            'Cas d'un indice de mat�riau non autoris�
            unSpreadMat.Value = 0
        End If
    Next i
End Sub

Public Sub AlimenterAutorisation(unFichierCERTU As Boolean)
    'Proc�dure alimentant les autorisations d'utilisation
    'des mat�riaux de base/fondation
    Dim uneColMat As Collection, unRes As Boolean
    Dim uneChaineIndicesNonAutoris�s As String
    
    'en version 2, on ne fait plus rien ici
    Exit Sub
    
    If unFichierCERTU Then
        'Cas du fichier de structures CERTU
        Set uneColMat = maColMatBFCERTU
        uneChaineIndicesNonAutoris�s = mesOptionsMat.mesMatCERTUNonAutoris�s
    Else
        'Cas du fichier de structures personnelles
        Set uneColMat = maColMatBFPerso
        uneChaineIndicesNonAutoris�s = mesOptionsMat.mesMatPersoNonAutoris�s
    End If
    
    For i = 1 To uneColMat.Count
        unRes = InStr(1, uneChaineIndicesNonAutoris�s, " " + Format(i) + " ")
        uneColMat(i).monUtilisationAutoris�e = (unRes = 0)
        'En effet si unRes = 0 ==> i n'est pas dans la string uneChaineIndicesNonAutoris�s
        'donc i n'est pas un indice de mat�riau non autoris�
    Next i
End Sub

Public Sub ActualiserFrameVerifGel(uneFrmDoc As Form)
    'Affichage ou non de la v�rif au gel dans la frame
    'de la fen�tre fille pass�e en param�tre
    Dim uneVisu As Boolean
    
    'En version 2 b�tatest on affiche que les indices de gel Ref et Admis
    'de la qualit� choisie
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
            'l'onglet gel ayant �t� gris� avant
            uneFrmDoc.TabData.Tab = OngletVoie
        End If
    End If
    
    'Affichage �ventuel des valeurs des valeurs
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
    'Positionnement des diff�rents labels d'affichage
    uneFrmDoc.LabelIGelAdmin.Caption = LabelIGelAdminCaption
    uneFrmDoc.LabelIGelAdmin.Left = uneFrmDoc.BtnGelQ1.Width + uneFrmDoc.BtnGelQ1.Left * 2
    'uneFrmDoc.LabelIGelAdmin.Left = (uneFrmDoc.FrameGel.Width - uneFrmDoc.LabelIGelAdmin.Width) / 2
    uneFrmDoc.LabelIGAdmQ1.Caption = uneString1
    uneFrmDoc.LabelIGAdmQ1.Left = uneFrmDoc.LabelIGelAdmin.Left + uneFrmDoc.LabelIGelAdmin.Width
    uneFrmDoc.LabelIGAdmQ2.Caption = uneString2
    uneFrmDoc.LabelIGAdmQ2.Left = uneFrmDoc.LabelIGelAdmin.Left + uneFrmDoc.LabelIGelAdmin.Width
    
    'Affichage �ventuel des valeurs des valeurs
    'de l'indice de gel de r�f�rence corrig� pour Q1 et Q2
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
    
    'Positionnement des diff�rents labels d'affichage
    uneFrmDoc.LabelIGelRef.Caption = LabelIGelRefCaption
    uneFrmDoc.LabelIGelRef.Left = uneFrmDoc.BtnGelQ1.Width + uneFrmDoc.BtnGelQ1.Left * 2
    'uneFrmDoc.LabelIGelRef.Left = (uneFrmDoc.FrameGel.Width - uneFrmDoc.LabelIGelRef.Width) / 2
    uneFrmDoc.LabelIGRefQ1.Caption = uneString1
    uneFrmDoc.LabelIGRefQ1.Left = uneFrmDoc.LabelIGelRef.Left + uneFrmDoc.LabelIGelRef.Width
    uneFrmDoc.LabelIGRefQ2.Caption = uneString2
    uneFrmDoc.LabelIGRefQ2.Left = uneFrmDoc.LabelIGelRef.Left + uneFrmDoc.LabelIGelRef.Width
    
    'Indication dans le cas o� la chauss�e n'est pas prot�g�e au gel pour Q1
    If uneFrmDoc.monIndiceGelRefQ1 > uneFrmDoc.monIndiceGelAdmQ1 And uneFrmDoc.monIndiceGelAdmQ1 > 0 Then
        uneFrmDoc.BtnGelQ1.Visible = True And mesOptionsGen.maVerifGel And uneVisu
        'Indication que la v�rif au gel n'est pas ok
        uneVerifGel = False
    Else
        uneFrmDoc.BtnGelQ1.Visible = False
        'Indication que la v�rif au gel est ok
        uneVerifGel = uneFrmDoc.monIndiceGelAdmQ1 > 0 And uneFrmDoc.monIndiceGelAdmQ1 <> HorsGel
    End If
    
    'Indication dans le cas o� la chauss�e n'est pas prot�g�e au gel pour Q2
    If uneFrmDoc.monIndiceGelRefQ2 > uneFrmDoc.monIndiceGelAdmQ2 And uneFrmDoc.monIndiceGelAdmQ2 > 0 Then
        uneFrmDoc.BtnGelQ2.Visible = True And mesOptionsGen.maVerifGel And Not uneVisu
        'Indication que la v�rif au gel n'est pas ok si on est en Q2
        If Not uneVisu Then uneVerifGel = False
    Else
        uneFrmDoc.BtnGelQ2.Visible = False
        'Indication que la v�rif au gel est ok si on est en Q2
        If Not uneVisu Then uneVerifGel = uneFrmDoc.monIndiceGelAdmQ2 > 0 And uneFrmDoc.monIndiceGelAdmQ2 <> HorsGel
    End If
    'Pour la version 2 b�tatest, le bouton gel Q2 est mis au m�me endroit
    'que le bouton gel Q1 (les top sont d�j� �gaux)
    uneFrmDoc.BtnGelQ2.Left = uneFrmDoc.BtnGelQ1.Left
    'Alignement du bouton OK gel
    uneFrmDoc.BtnOKGel.Left = uneFrmDoc.BtnGelQ1.Left
    uneFrmDoc.BtnOKGel.Top = uneFrmDoc.BtnGelQ1.Top
    'Affichage du bouton indiquant que la chauss�e est prot�g�e au gel en Q1 ou Q2
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
    
    uneListEp = InputBox("Entrez les �paisseurs de la carotte de qualit� Q1, puis celle de qualit� Q2 s�par�es par des blancs, �paisseur nulle pour une couche absente :", , "2 2 11 12 13 14 oui oui oui 3 0 15 0 20 0 non non oui")
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
            'Si unDecInd = 0 ==> Affichage pour v�rif saisie
            'sinon on met undecInd = 1 ==> pas d'affichage de v�rif
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
    'Dessin des carottes de qualit� Q1 et Q2 d'une fenetre fille en
    'respectant les proportions des diff�rentes �paisseurs en commen�ant
    '� partir de la couche symbolisant la plateforme et en remontant les couches
    'Echelle 1 cm repr�sent�e par EpTotMaxEcran / max Ep totale r�elle (Q1,Q2)
    'et les Carottes totales doit �tre compris entre 0 et 55 cm
    Dim unCmEcran As Single, uneEchelle As Single
    Dim uneEpTotQ1 As Integer, uneEpTotQ2 As Integer
    
    'Calcul de l'�chelle
    For i = 1 To 6
        'Calcul de l'�paisseur totale de la carotte Q1
        uneEpTotQ1 = uneEpTotQ1 + unTabEp(i)
    Next i
    For i = 7 To 12
        'Calcul de l'�paisseur totale de la carotte Q2
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
    '� partir de la plateforme
    DessinerCarotteQ1 uneFrmDoc, CInt(unTabEp(1)), CInt(unTabEp(2)), CInt(unTabEp(3)), CInt(unTabEp(4)), CInt(unTabEp(5)), CInt(unTabEp(6)), uneEchelle, unMinTecQ1, unMaxPraQ1, unTrouvEpQ1
    'Dessin de la carotte Q2 de la fenetre fille
    '� partir de la plateforme
    DessinerCarotteQ2 uneFrmDoc, CInt(unTabEp(7)), CInt(unTabEp(8)), CInt(unTabEp(9)), CInt(unTabEp(10)), CInt(unTabEp(11)), CInt(unTabEp(12)), uneEchelle, unMinTecQ2, unMaxPraQ2, unTrouvEpQ2
End Sub

Public Sub DessinerCarotteQ1(uneFrmDoc As Form, uneEpS1 As Integer, uneEpS2 As Integer, uneEpB1 As Integer, uneEpB2 As Integer, uneEpF1 As Integer, uneEpF2 As Integer, uneEchelle As Single, unMinTec As String, unMaxPra As String, unTrouvEp As Boolean)
    'Dessin de la carotte de qualit� Q1 en respectant les proportions
    'des diff�rentes �paisseurs en commen�ant � partir de la
    'couche symbolisant la plateforme et en remontant les couches
    Dim unTop As Long, uneSurfSansEpaisseur As Boolean
    Dim uneStruct As Structure
    Dim uneExistence As Boolean, unMatSurfCompos� As Boolean
    Dim uneStringEpS1 As String, uneStringEpS2 As String
    Dim uneStringEpB1 As String, uneStringEpB2 As String
    Dim uneStringEpF1 As String, uneStringEpF2 As String
    
    unMatSurfCompos� = False
    'En version 2 b�tatest on n'affiche que la carotte de la qualit� choisie Q1 ou Q2
    uneVisu = (uneFrmDoc.monTypeChantier = TypeChantierQ1)
    uneFrmDoc.LabelPFQ1.Visible = uneVisu
    uneFrmDoc.ShapePFQ1.Visible = uneVisu
    
    'Affectation des libell�s des diff�rentes couches
    Set uneStruct = DonnerStructChoisie(uneFrmDoc)
    If Not (uneStruct Is Nothing) Then
        'Cas d'une structure choisie, on affiche les mat�riaux
        'de ses couches
        With uneFrmDoc
            .LabelCFond2Q1.Caption = uneStruct.maCoucheFondation
            .LabelCFond1Q1.Caption = uneStruct.maCoucheFondation
            .LabelCBase2Q1.Caption = uneStruct.maCoucheBase
            .LabelCBase1Q1.Caption = uneStruct.maCoucheBase
            .LabelCSurfQ1.Caption = uneStruct.maCoucheSurface
        End With
        
        If uneStruct.maCoucheSurface <> "Aucune" Then
            unMatSurfCompos� = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatCompos�)
        End If
    End If
    
    'Recherche si on a une structure avec une couche de surface sans �paisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Affectation des valeurs d'�paisseurs � afficher suivant
    'le r�sultat du dimensionnement
    If UCase(unMaxPra) = "OUI" Or unTrouvEp = False Then
        'Maximun pratique atteint ou pas d'�paisseur trouv�e
        '(qui fait qussi affichage apr�s premier choix de structure)
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
        'Affichages �ventuels du min techno et du max pratique
        .LabelInfoMaxPQ1.Visible = (UCase(unMaxPra) = "OUI") And uneVisu
        .LabelInfoMaxPQ1.Top = .FrameCarotte.Height / 2
        .LabelInfoMinTechnoQ1.Visible = (UCase(unMinTec) = "OUI") And uneVisu
        .LabelEpTotQ1.Visible = UCase(unMaxPra) = "NON" And uneVisu
    
        'Initialisation du d�but du dessin
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
        'uneEpS1 = Ep totale de la surface et uneEpS2 = 0 si mat�riau simple
        'sinon uneEpS1 et uneEpS2 non nulles pour un mat�riau de surface compos�
        uneExistence = (uneEpS1 > 0)
        uneCSurfEnDallesOuPaves = (uneStruct.maCoucheSurface = "Dalles") Or (uneStruct.maCoucheSurface = "Pav�s")
        .ShapeLitPoseQ1.Visible = uneCSurfEnDallesOuPaves And uneVisu
        .LabelLitPoseQ1.Visible = uneCSurfEnDallesOuPaves And uneVisu
        If uneExistence Then
            If uneCSurfEnDallesOuPaves Then
                'Cas d'une couche de surface en dalles ou pav�s
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
            If unMatSurfCompos� Then
                'Cas d'un mat�riau de surface compos�
                .LabelSurf1Q1.Top = .ShapeSurfQ1.Top + .ShapeSurfQ1.Height / 2 - .LabelSurf2Q1.Height
                .LabelSurf1Q1.Caption = .monMSComp1Q1 + " " + uneStringEpS1 + " cm"
                If uneEpS2 > 0 Then .LabelSurf1Q1.Caption = .LabelSurf1Q1.Caption + " +"
                .LabelSurf2Q1.Caption = .monMSComp2Q1 + " " + uneStringEpS2 + " cm"
                .LabelSurf2Q1.Top = .LabelSurf1Q1.Top + .LabelSurf1Q1.Height
            Else
                'Cas d'un mat�riau de surface simple
                .LabelSurf1Q1.Top = .LabelCSurfQ1.Top
                If uneSurfSansEpaisseur Then
                    'Cas d'une couche de surface sans �paisseur
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
    'Dessin de la carotte de qualit� Q2 en respectant les proportions
    'des diff�rentes �paisseurs en commen�ant � partir de la
    'couche symbolisant la plateforme et en remontant les couches
    Dim unTop As Long, uneVisu As Boolean
    Dim uneStruct As Structure
    Dim uneExistence As Boolean, unMatSurfCompos� As Boolean
    Dim uneStringEpS1 As String, uneStringEpS2 As String
    Dim uneStringEpB1 As String, uneStringEpB2 As String
    Dim uneStringEpF1 As String, uneStringEpF2 As String
    
    unMatSurfCompos� = False
    'En version 2 b�tatest on n'affiche que la carotte de la qualit� choisie Q1 ou Q2
    uneVisu = (uneFrmDoc.monTypeChantier = TypeChantierQ2)
    uneFrmDoc.LabelPFQ2.Visible = uneVisu
    uneFrmDoc.ShapePFQ2.Visible = uneVisu
    
    'Affectation des libell�s des diff�rentes couches
    Set uneStruct = DonnerStructChoisie(uneFrmDoc)
    If Not (uneStruct Is Nothing) Then
        'Cas d'une structure choisie, on affiche les mat�riaux
        'de ses couches
        With uneFrmDoc
            .LabelCFond2Q2.Caption = uneStruct.maCoucheFondation
            .LabelCFond1Q2.Caption = uneStruct.maCoucheFondation
            .LabelCBase2Q2.Caption = uneStruct.maCoucheBase
            .LabelCBase1Q2.Caption = uneStruct.maCoucheBase
            .LabelCSurfQ2.Caption = uneStruct.maCoucheSurface
        End With
        
        If uneStruct.maCoucheSurface <> "Aucune" Then
            unMatSurfCompos� = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatCompos�)
        End If
    End If
    
    'Recherche si on a une structure avec une couche de surface sans �paisseur
    uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
    
    'Affectation des valeurs d'�paisseurs � afficher suivant
    'le r�sultat du dimensionnement
    If UCase(unMaxPra) = "OUI" Or unTrouvEp = False Then
        'Maximun pratique atteint ou pas d'�paisseur trouv�e
        '(qui fait qussi affichage apr�s premier choix de structure)
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
        'Affichages �ventuels du min techno et du max pratique
        .LabelInfoMaxPQ2.Visible = (UCase(unMaxPra) = "OUI") And uneVisu
        .LabelInfoMaxPQ2.Top = .FrameCarotte.Height / 2
        .LabelInfoMinTechnoQ2.Visible = (UCase(unMinTec) = "OUI") And uneVisu
        uneFrmDoc.LabelEpTotQ2.Visible = UCase(unMaxPra) = "NON" And uneVisu
        
        'Initialisation du d�but du dessin
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
        'uneEpS1 = Ep totale de la surface et uneEpS2 = 0 si mat�riau simple
        'sinon uneEpS1 et uneEpS2 non nulles pour un mat�riau de surface compos�
        uneExistence = (uneEpS1 > 0)
        uneCSurfEnDallesOuPaves = (uneStruct.maCoucheSurface = "Dalles") Or (uneStruct.maCoucheSurface = "Pav�s")
        .ShapeLitPoseQ2.Visible = uneCSurfEnDallesOuPaves And uneVisu
        .LabelLitPoseQ2.Visible = uneCSurfEnDallesOuPaves And uneVisu
        If uneExistence Then
            If uneCSurfEnDallesOuPaves Then
                'Cas d'une couche de surface en dalles ou pav�s
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
            If unMatSurfCompos� Then
                'Cas d'un mat�riau de surface compos� (deux �paisseurs)
                .LabelSurf1Q2.Top = .ShapeSurfQ2.Top + .ShapeSurfQ2.Height / 2 - .LabelSurf2Q2.Height
                .LabelSurf1Q2.Caption = .monMSComp1Q2 + " " + uneStringEpS1 + " cm"
                If uneEpS2 > 0 Then .LabelSurf1Q2.Caption = .LabelSurf1Q2.Caption + " +"
                .LabelSurf2Q2.Caption = .monMSComp2Q2 + " " + uneStringEpS2 + " cm"
                .LabelSurf2Q2.Top = .LabelSurf1Q2.Top + .LabelSurf1Q2.Height + 10
            Else
                'Cas d'un mat�riau de surface simple
                .LabelSurf1Q2.Top = .LabelCSurfQ2.Top
                If uneSurfSansEpaisseur Then
                    'Cas d'une couche de surface sans �paisseur
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
    'Afficher les messages d'erreur si une saisie est erron�e dans
    'l'onglet courant de la fenetre fille active
    Dim uneForm As Form
    
    'Initialisation du code de retour
    VerifierSaisieOngletCourant = True
    
    If Forms.Count > 1 Then
        'Cas o� il y a au moins une fenetre fille
        Set uneForm = fMainForm.ActiveForm
        If uneForm.TabData.Tab = OngletTrafic Then
            'V�rification dans l'onglet Trafic
            VerifierSaisieOngletCourant = VerifierMinMaxTraficIni(uneForm, uneForm.TextTrafIni.Text)
            If VerifierSaisieOngletCourant Then VerifierSaisieOngletCourant = VerifierMinMaxDur�eService(uneForm)
        End If
    End If
End Function

Public Function VerifierFinSaisie() As Boolean
    'V�rifier si une saisie est en cours dans un textbox,
    'il faut valider avant pour une impression, save, new, open,
    'quitter ou fermer
    Dim unMaskEdBox As MaskEdBox
    
    If monEtude Is Nothing Then
        VerifierFinSaisie = True
        Exit Function
    End If
    
    'V�rification que l'on ne soit pas encore de saisie dans un textbox
    'demandant une validation par sortie du champ ou par retour chariot
    If Not (monEtude.ActiveControl Is Nothing) Then
        Set unMaskEdBox = monEtude.MaskCAM
    
        'Cas o� il y a un control actif
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
    ElseIf monEtude.ActiveControl.Name = "TextDur�eS" Then
        DonnerNomTextBox = MsgTextBoxDur�eS
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

Public Sub MettreAJourFrameR�sultat(uneForm As Form)
    'Mise � jour de l'affichage de la frame visualisant
    'les r�sultats et les carottes de qualit� Q1 et Q2
    With uneForm
        If .monTraficCumul� = 0 Then
            uneString = MsgInconnu
        Else
            uneString = Format(.monTraficCumul�, "### ### ###")
        End If
        .LabelTraficCum.Caption = LabelTraficCumCaption + uneString
        
        If .monCAM = "" Then
            uneString = MsgInconnu
        Else
            'uneString = .monCAM
            uneString = Mid(.monCAM, 1, 1) + monCarDeci + Mid(.monCAM, 3)
            'On prend de part et d'autre du s�parateur d�cimale
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
    'Mise � jour de l'affichage de l'onglet voie
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
    'Mise � jour de l'affichage de l'onglet trafic
    With uneForm
        If .monTraficIni = 0 Then
            .TextTrafIni.Text = ""
        Else
            .TextTrafIni.Text = Format(.monTraficIni)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextTrafIni.Tag = .TextTrafIni.Text
        .TextTrafIni.ForeColor = QBColor(0) 'Car pas de modif
        
        If .maDur�eService = 0 Then
            .TextDur�eS.Text = Format(mesOptionsGen.maDur�eService)
        Else
            .TextDur�eS.Text = Format(.maDur�eService)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextDur�eS.Tag = .TextDur�eS.Text
        .TextDur�eS.ForeColor = QBColor(0) 'Car pas de modif
        
        If .maCroisAnnuel = 100 Then
            'Cas d'une nouvelle �tude la croiss annuel est inconnu (= 100)
            'Les % vont de 0 � 5 % et
            'le tableau de controle optionCA de 0 � 5
            .OptionCA(mesOptionsGen.maCroisAnnuel).Value = True
        Else
            .OptionCA(.maCroisAnnuel).Value = True
        End If
        
        If .monTraficCumul� = 0 Then
            .TextTrafCUM.Text = ""
        Else
            .TextTrafCUM.Text = Format(.monTraficCumul�)
        End If
        'Sauvegarde dans le tag pour la touche echappement
        .TextTrafCUM.Tag = .TextTrafCUM.Text
        .TextTrafCUM.ForeColor = QBColor(0) 'Car pas de modif
    End With
End Sub

Public Sub MettreAJourOngletCoucheSurface(uneForm As Form, uneEpPrecTrouvQ1 As Integer, uneEpPrecTrouvQ2 As Integer)
    'Mise � jour du contenu Onglet couche de surface
    Dim unNomMat1 As String, unNomMat2 As String, unNomMat3 As String
    Dim uneEpMat1 As Integer, uneEpMat2 As Integer, uneEpMat3 As Integer
    Dim uneEpPrec As Integer, unN As Integer
    Dim lesComp As Collection
    Dim uneStruct As Structure
    Dim unMatCompos� As Object
    
    With uneForm
        .LabelValEpPrecQ1.Caption = Format(uneEpPrecTrouvQ1)
        .LabelValEpPrecQ2.Caption = Format(uneEpPrecTrouvQ2)
        
        'R�cup de la structure choisie
        Set uneStruct = DonnerStructChoisie(uneForm)
        If uneStruct Is Nothing Then Exit Sub
        
        'R�cup du mat�riau compos� en couche de surface
        Set unMatCompos� = maColMatSurf(uneStruct.maCoucheSurface)
        If Not (TypeOf unMatCompos� Is MatCompos�) Then Exit Sub
        
        'On vide la liste des mat�riaux composants possibles
        'et la combobox des compositions possibles
        .ListViewMat.ListItems.Clear
        .ComboCompQ1.Clear
        .ComboCompQ2.Clear
        
        'Remplissage de la combobox combocomp des compositions possibles
        Set lesComp = unMatCompos�.mesCompositions
        unNbComp = 3
        '3 = Nb de colonnes de composant
        unNbComposition = lesComp.Count \ (2 * unNbComp + 1)
        For j = 1 To unNbComposition
            unN = (j - 1) * (2 * unNbComp + 1) + 1
            'R�cup de l'�paisseur pr�conis�e de la composition j
            uneEpPrec = CInt(lesComp(unN))
            
            'Recup premier composant
            uneEpMat1 = CInt(lesComp(unN + 1))
            unNomMat1 = Format(lesComp(unN + 2))
            
            'Recup deuxi�me composant
            uneEpMat2 = CInt(lesComp(unN + 3))
            unNomMat2 = Format(lesComp(unN + 4))
                        
            'Recup troisi�me composant
            uneEpMat3 = CInt(lesComp(unN + 5))
            unNomMat3 = Format(lesComp(unN + 6))
                            
            If uneEpPrec = uneEpPrecTrouvQ1 Then
                'Cas o� l'�paisseur pr�conis�e trouv� Q1
                'correspond � celle de la composition j
                '===> Ajout aux compositions possibles
                AjouterComposition uneForm, .ComboCompQ1, uneEpMat1, unNomMat1, uneEpMat2, unNomMat2, uneEpMat3, unNomMat3, unN
            End If
             
             If uneEpPrec = uneEpPrecTrouvQ2 Then
                'Cas o� l'�paisseur pr�conis�e trouv� Q2
                'correspond � celle de la composition j
                '===> Ajout aux compositions possibles
                AjouterComposition uneForm, .ComboCompQ2, uneEpMat1, unNomMat1, uneEpMat2, unNomMat2, uneEpMat3, unNomMat3, unN
           End If
        Next j
        
        'Affichage des compositions Q1 et Q2
        If uneForm.monJustOpen Then
            'Cas o� on vient d'ouvrir l'�tude on met les indices stock�s
            'et on n'y passe qu'une fois
            .ComboCompQ1.ListIndex = (.monIndCompQ1 - 1)
            .ComboCompQ2.ListIndex = (.monIndCompQ2 - 1)
            uneForm.monJustOpen = False
        Else
            'tous les autres cas ==> on remet � vide
            .ComboCompQ1.ListIndex = -1
            .ComboCompQ2.ListIndex = -1
        End If
    End With
End Sub

Public Sub MettreAJourOngletStructure(uneForm As Form, unUtilFichPerso As Byte, unIndStructChoisie As Integer)
    'Mise � jour du contenu Onglet Structure
    Dim uneColStruct As Collection
    Dim uneStruct As Structure
    
    'Changement de valeur de la case � cocher Utiliser un fichier perso
    'Comme elle vaut Grayed = 2 toutes les valeurs 0 ou 1 d�clenchent
    'sont click event (cf FrmDocument code, checkfichperso_click)
    
    uneForm.CheckFichPerso.Value = unUtilFichPerso
    
    'Affichage de la case � cocher Utiliser Fichier personnel si
    'il y  un fichier personnel donn� dans les options g�n�rales
    'ou si l'�tude en utilise un.
    If uneForm.monFichPersoSTR = "" And mesOptionsMat.monFichPersoSTR = "" Then
        uneForm.CheckFichPerso.Visible = False
    End If
    
    'Affichage de la structure choisie dans la combobox combostruct
    If unIndStructChoisie > 0 Then
        'Cas du chargement d'une �tude ayant d�j� choisie une structure
        'R�cup de la bonne collection de structures
        If uneForm.CheckFichPerso.Value = 0 Then
            Set uneColStruct = maColStructCERTU
        Else
            Set uneColStruct = maColStructPerso
        End If
        
        'En version 2, D�clenchement du radio bouton correspondant
        'au type de la strucutre chosie �ventuelle
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
                'qui a �t� rempli avant par le CheckFichPerso_click
                'gr�ce au CheckFichPerso.value = du d�but
                
                'On cherche la structure de m�me abr�g� et se trouvant
                '� la position i (tout �a car les abr�g�s de structures
                'ne sont pas uniques)
                unInd = uneForm.ComboStruct.ItemData(j)
                If uneForm.ComboStruct.List(j) = uneStruct.monAbr�g� And unInd = unIndStructChoisie Then
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
    'Fonction rajout�e en version 2
    'Activation ou d�activation des radio boutons � cliquer
    'de l'�tude active, donc de la form pass�e en param�tre,
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
        
    'Mise en d�sactiv� de tous les radio boutons et d�cochage
    For i = Souple To PavesDalles
        uneForm.OptionTypeStruct(i).Enabled = False
        uneForm.OptionTypeStruct(i).Value = False
    Next i
    
    'Mise en activation �ventuelle des radio boutons
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
    'donc ayant des mat�riaux de base et de fondation
    'autoris�s et ayant un type de voie compatibles avec
    'celui de l'�tude active, donc de la form pass�e en param�tre
    
    'EN PLUS EN VERSION 2, IL FAUT LES STRUCTURES DU BON TYPE
    Dim uneColStruct As Collection
    Dim uneColMatBF As Collection
    Dim unTypeVoieOK As Boolean
    Dim uneStruct As Structure, unMat As Mat�riau
    
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
        
        'V�rification que cette structure a un type de voie
        'compatible avec celui de l'�tude (= la form) active
        unTypeVoieOK = (uneForm.monTypeVoie = TypeVoieDesserte) And (uneStruct.monUtilVDes = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieDistribution) And (uneStruct.monUtilVDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieTraficLourd) And (uneStruct.monUtilVPL = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieBus) And (uneStruct.monUtilVBus = 1)
        
        'Pour la version 2 b�tatest, on rajoute les VoieParking
        'et pour les voies giratoires
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeVoieParking) And (uneStruct.monUtilVParking = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireDistribution) And (uneStruct.monUtilGDis = 1)
        unTypeVoieOK = unTypeVoieOK Or (uneForm.monTypeVoie = TypeGiratoireTraficLourd) And (uneStruct.monUtilGPL = 1)
        
        'EN PLUS EN VERSION 2, IL FAUT LES STRUCTURES DU BON TYPE
        unTypeVoieOK = unTypeVoieOK And (uneStruct.monTypeStructure = unTypeStruct)
        
        If unTypeVoieOK Then
            'Cas o� la structure i est utilisable
            'pour le type de voie de l'etude active
            
            'V�rification que le mat�riau de base
            's'il existe est autoris�
            If uneStruct.maCoucheBase = "Aucune" Then
                unMatBaseOK = True
            Else
                Set unMat = uneColMatBF.Item(uneStruct.maCoucheBase)
                unMatBaseOK = unMat.monUtilisationAutoris�e
            End If
            'V�rification que le mat�riau de fondation
            's'il existe est autoris�
            If uneStruct.maCoucheFondation = "Aucune" Then
                unMatFondOK = True
            Else
                Set unMat = uneColMatBF.Item(uneStruct.maCoucheFondation)
                unMatFondOK = unMat.monUtilisationAutoris�e
            End If
            
            If unMatBaseOK And unMatFondOK Then
                'Ajout dans la liste des structures possibles
                uneForm.ComboStruct.AddItem uneColStruct(i).monAbr�g�
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
        'Cas o� aucune structure n'a �t� choisie d'o�
        'Inhibition des boutons de visu des mat�riaux surface simple,
        'base et fondation
        'Et affichage d'un libell� de boutons sans abr�g� mat�riau
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
            'Cas o� le fichier perso utilis� est celui des options mat�riaux
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
            'Cas o� le fichier perso existe sans �tre celui des options mat�riaux
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
        'R�cup de la strucutre choisie
        Set uneStruct = DonnerStructChoisie(uneForm)
        If uneStruct Is Nothing Then Exit Sub
        
        'Indication des couches existantes
        unMatSurf1Exist = Not (uneStruct.maCoucheSurface = "Aucune")
        If unMatSurf1Exist Then
            'Cas avec couche de surface
            unMatSurf2Exist = (TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatCompos�)
        Else
            'Cas sans couche de surface
            unMatSurf2Exist = unMatSurf1Exist
        End If
        unMatBaseExist = Not (uneStruct.maCoucheBase = "Aucune")
        unMatFondExist = Not (uneStruct.maCoucheFondation = "Aucune")
        
        'Recherche si on a une structure avec une couche de surface sans �paisseur
        uneSurfSansEpaisseur = (unMatSurf1Exist And uneStruct.maCoucheSurfSansEp = 1)
        
        'R�cup du tableau de variant contenant les �paisseurs
        unTabEp = .monTabEp
        
        If .monEpQ1Trouv = False Then
            'Pas d'�paisseur trouv�e pour Q1
            '==> on prend �paisseur par d�faut
            'si aucune �paisseur trouv�e pour Q2
            'sinon on prend les �paisseurs de Q2 d'o� m�me carotte
            'sauf si on est en �tude giratoire car pas d'�paisseur en Q2 dans ces structures
            If .monEpQ2Trouv And .OptionEtudeGiratoire.Value = False Then
                For i = 1 To 6
                    unTabEp(i) = unTabEp(i + 6)
                Next i
            Else
                'Les booleens False = 0 et True = -1
                'd'ou abs pour la valeur absolue pour avoir 0 ou 1
                If uneSurfSansEpaisseur Then
                    'Trois cm pour une bonne visu � l'�cran d'une couche
                    'de surface sans �paisseur
                    unTabEp(1) = 3
                    unTabEp(2) = 0
                ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pav�s" Then
                    'EpLitPose cm pour le lit de pose � l'�cran d'une couche
                    'de surface en dalles ou pav�s
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
            'Pas d'�paisseur trouv�e pour Q2
            '==> on prend �paisseur par d�faut
            'si aucune �paisseur trouv�e pour Q1
            'sinon on prend les �paisseurs de Q1 d'o� m�me carotte
            If .monEpQ1Trouv Then
                For i = 7 To 12
                    unTabEp(i) = unTabEp(i - 6)
                Next i
            Else
                'Les booleens False = 0 et True = -1
                'd'ou abs pour la valeur absolue
                If uneSurfSansEpaisseur Then
                    'Trois cm pour une bonne visu � l'�cran d'une couche
                    'de surface sans �paisseur
                    unTabEp(7) = 3
                    unTabEp(8) = 0
                ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pav�s" Then
                    'EpLitPose cm pour le lit de pose � l'�cran d'une couche
                    'de surface en dalles ou pav�s
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
        
        'Affectation du tableau de variant contenant les �paisseurs
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
    'Ouverture du fichier de mat�riaux de surface
    'Test de l'existence des fichiers CERTU.mts et CERTU.str
    'dans le r�pertoire de l'application GestionStructure
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
        'Test si fichier mat�riaux surface valide
        'Ouverture du fichier Mat�riaux de surface CERTU.MTS
        Open unFileName For Binary As #unFichMTS
        LireString unFichMTS, unEntete
        If unEntete = ENTETE_MTS Then
            'Remplissage des mat�riaux de surface
            Do
                LireString unFichMTS, unTypeMat
                If unTypeMat = "Mat�riauSimple" Or unTypeMat = "Mat�riau" Then
                    'Lecture fichier pour un mat�riau simple ou composant
                    LireString unFichMTS, unNom
                    LireString unFichMTS, unAbrege
                    LireString unFichMTS, uneNorme
                    LireString unFichMTS, uneQual
                    Get #unFichMTS, , unAGel
                    Get #unFichMTS, , unBGel
                    If unTypeMat = "Mat�riau" Then
                        Get #unFichMTS, , unYoung
                        Get #unFichMTS, , unPoisson
                        Get #unFichMTS, , unEpsilon
                        Get #unFichMTS, , unSigma
                    End If
                    LireString unFichMTS, unComment
                    LireString unFichMTS, unCommentSuite
                    Do While unCommentSuite <> FIN_COMMENT
                        'Corrrection du probl�me li� � la pr�sence
                        'de " dans le texte RTF
                        '==> d�but ou fin de string en lecture
                        unComment = unComment + unCommentSuite
                        LireString unFichMTS, unCommentSuite
                    Loop
                    'Cr�ation des mat�riaux suivant leur type
                    If unTypeMat = "Mat�riauSimple" Then
                        Set unMat = New MatSimple
                        unMat.SetProps unNom, unAbrege, uneNorme, unComment
                    ElseIf unTypeMat = "Mat�riau" Then
                        Set unMat = New Mat�riau
                        unMat.SetProps unNom, unAbrege, uneNorme, unComment
                        unMat.SetPropsPhysic unYoung, unPoisson, unEpsilon, unSigma
                        'Alimentation de la collection des mat�riaux composants
                        maColMatComposant.Add unMat, unMat.monAbr�g�
                    End If
                    RemplirQualit�Gel unMat, uneQual, unAGel, unBGel
                    'Alimentation de la collection des mat�riaux de surface
                    maColMatSurf.Add unMat, unMat.monAbr�g�
                ElseIf unTypeMat = "Mat�riauCompos�" Then
                    'Lecture fichier pour un mat�riau compos�
                    LireString unFichMTS, unNom
                    LireString unFichMTS, unAbrege
                    Get #unFichMTS, , unNbComposition
                    Get #unFichMTS, , unAGel
                    Get #unFichMTS, , unBGel
                    Set uneColComp = New Collection
                    unNbComp = 3
                    '3 = Nb de colonnes de composant et alimentation d'une
                    'collection contenant unNbcomp couples
                    '(�paisseur, composant) + �paisseur pr�conis�e
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
                    'Cr�ation des mat�riaux
                    Set unMat = New MatCompos�
                    unMat.SetProps unNom, unAbrege
                    Set unMat.mesCompositions = uneColComp
                    RemplirQualit�Gel unMat, "", unAGel, unBGel
                    'Alimentation de la collection des mat surface
                    maColMatSurf.Add unMat, unMat.monAbr�g�
                ElseIf EOF(unFichMTS) = False Then
                    'Car unTypeMat ="" apr�s la derni�re lecture
                    'en fin de fichier
                    MsgBox MsgErreurMat�riauInconnu + ": " + unTypeMat + Chr(13) + MsgFich + " " + unFileName + MsgIncorrect, vbCritical
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
    'Ajout dans le listview listviewmat d'un nom de mat�riau
    's'il n'y est pas d�j�
    Dim unItemX As ListItem
    
    'Test si d�j� pr�sent dans les items du listview
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
    'Ajouter dans la combobox de composition pass�e en param�tre
    'une composition possible par rapport � l'�paisseur pr�conis�e
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
    'les couleurs des options g�n�rales d'une Form = une �tude
    With uneForm
        'Pour la qualit� Q1
        .ShapeSurfQ1.FillColor = mesOptionsGen.maCoulSurf
        .ShapeBase1Q1.FillColor = mesOptionsGen.maCoulBase1
        .ShapeBase2Q1.FillColor = mesOptionsGen.maCoulBase2
        .ShapeFond1Q1.FillColor = mesOptionsGen.maCoulFond1
        .ShapeFond2Q1.FillColor = mesOptionsGen.maCoulFond2
        'Pour la qualit� Q2
        .ShapeSurfQ2.FillColor = mesOptionsGen.maCoulSurf
        .ShapeBase1Q2.FillColor = mesOptionsGen.maCoulBase1
        .ShapeBase2Q2.FillColor = mesOptionsGen.maCoulBase2
        .ShapeFond1Q2.FillColor = mesOptionsGen.maCoulFond1
        .ShapeFond2Q2.FillColor = mesOptionsGen.maCoulFond2
    End With
End Sub

Public Sub AfficherFicheMat(uneStringFicheMat As String)
    'Affichage de la fiche mat�riau suivant le type du mat�riau
    'pour visulaiser ces caract�ristiques
        
    'uneStringFicheMat = TypeMat + '/' + Abr�g�
    'TypeMat = "Simple" ou "Composant" ou "FondBase" ou "Compos�"
    
    'Chargement sans affichage
    Load FicheMat
    
    'Remplissage du tag de FicheMat avec le type de
    'de mat�riau et l'abr�g�
    FicheMat.Tag = uneStringFicheMat
    
    'Centrage de la fiche mat�riau
    CentrerFenetreEcran FicheMat
    
    'Affichage modal
    FicheMat.Show vbModal
End Sub


Public Sub CalculerTraficCum(uneForm As Form)
    'Calcul du trafic cumul� dans l'onglet Trafic de la form
    'pass�e en param�tre
    Dim unTini As Integer, uneCA As Single, uneDS As Byte
    With uneForm
        If .TextTrafIni.Text = "" Then Exit Sub
        unTini = CInt(.TextTrafIni.Text)
        uneCA = DonnerCroissAn(uneForm) / 100
        uneDS = Val(.TextDur�eS)
        'On arrondi � l'entier sup�rieur par rapport � la formule
        'du cahier des charges
        .TextTrafCUM.Text = Format(CLng(365# * unTini * (uneDS + (uneCA * uneDS * (uneDS - 1)) / 2)))
        'Calcul du Nombre d'essieux �quivalents si trafic cumul� connu
        uneForm.CalculerEtAfficherNE
    End With
End Sub

Public Function CalculerTraficIni(uneForm As Form) As Boolean
    'Calcul du trafic initial dans l'onglet Trafic de la form
    'pass�e en param�tre
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
        uneDS = Val(.TextDur�eS)
        'On arrondi � l'entier sup�rieur par rapport � la formule
        'du cahier des charges
        unTextTrafIni = Format(CLng(unTCum / 365# / (uneDS + (uneCA * uneDS * (uneDS - 1)) / 2)))
        
        'R�cup du domaine de validit� suivant le type de voie
        DonnerMinMaxTraficIni uneForm, uneValMinTol, uneValMin, uneValMax, uneValMaxTol
        unMsg = ""
        unNomTypeVoie = DonnerNomTypeVoie(uneForm)
            
        'Test de la validit� du trafic initial calcul�
        If CLng(unTextTrafIni) < uneValMinTol Or CLng(unTextTrafIni) > uneValMaxTol Then
            'Cas d'erreur non admise dans le domaine de validit�
            unMsg = MsgTraficCum + unTextTrafIni + Chr(13) + Chr(13)
            unMsg = unMsg + MsgTraficIni + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
            'Affichage du domaine de tol�rance plus grande que le domaine de validit�
            unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
            'R�sultat invalide
            unTypeIcone = vbCritical
            CalculerTraficIni = False
        ElseIf (CLng(unTextTrafIni) >= uneValMinTol And CLng(unTextTrafIni) < uneValMin) Or (CLng(unTextTrafIni) > uneValMax And CLng(unTextTrafIni) <= uneValMaxTol) Then
            'Cas d'erreur tol�r�e dans le domaine de validit�
            unMsg = MsgTraficCum + unTextTrafIni + Chr(13) + Chr(13)
            unMsg = unMsg + MsgTraficIni2 + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
            unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
            unTypeVoie = DonnerTypeVoie(uneForm)
            If CLng(unTextTrafIni) > uneValMax Then
                'Cas o� on d�passe la valeur maxi
                If (unTypeVoie >= TypeVoieTraficLourd And unTypeVoie <= TypeVoieBus) Or (unTypeVoie = TypeGiratoireTraficLourd) Then
                    'Cas des voie bus, voie principale avec PL
                    'et Giratoire sur voie principale PL
                    'Si la valeur > au max et < max tol�r�e, conseiller de faire
                    'une �tude en laboratoire
                    unMsg = unMsg + Chr(13) + Chr(13) + MsgValLabo + Format(uneValMax) + MsgIsLabo
                End If
            End If
            'R�sultat OK
            unTypeIcone = vbInformation
            CalculerTraficIni = True
            'Mise � jour du Textbox trafic initial
            .TextTrafIni.Text = unTextTrafIni
        Else
            'Cas o� on se trouve dans le domaine de validit�
            CalculerTraficIni = True
            'Mise � jour du Textbox trafic initial
            .TextTrafIni.Text = unTextTrafIni
        End If
        
        If unMsg <> "" Then MsgBox unMsg, unTypeIcone
    End With
End Function

Public Function DonnerCroissAn(uneForm As Form) As Byte
    'Donner la croissance annuelle en cours d'une form (= une �tude)
    'en scannant les valeurs des boutons options i%
    Dim i As Byte
    For i = 0 To 5
        'Les % vont de 0 � 5 % et le tableau de controle optionCA de 0 � 5
        If uneForm.OptionCA(i).Value Then
            DonnerCroissAn = i
            Exit For
        End If
    Next i
End Function
    
Public Sub DonnerMinMaxTraficIni(uneForm As Form, uneValMinTol As Long, uneValMin As Long, uneValMax As Long, uneValMaxTol As Long)
    'R�cup�ration des valeurs min et max du trafic initial
    'et les min et max tol�r�es
    '� partir du type de voie d'une �tude (= la form)
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
        'les voies r�serv�s aux bus (am�nagement standard) et parking
        uneValMin = 0
        uneValMax = 750 '1000 modif en V2
        uneValMinTol = 0
        uneValMaxTol = 1000
    End If
End Sub
    
Public Function DonnerPrecMinMaxCAM(uneForm As Form, uneValPrec As Single, uneValMin As Single, uneValMax As Single) As String
    'R�cup�ration des valeurs pr�conis�e, min et max du trafic initial
    '� partir du type de voie et du type de structure choisie d'une �tude
    'Etude = la form et type structure = (Souple ou Bitu) ou (Hydro ou B�ton)
        'CAM souple ou bitu  ==> TypeCAM = 1
        'CAM hydrau ou beton ==> TypeCAM = 2
    'et retourne le libell� du type de voie
    Dim uneStruct As Structure
    
    'R�cup de la structure choisie
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
    
    'Cas particulier d'une structure de type souple, ou de type pav�s/dall�es avec une
    'couche de base GNT, le cam pr�conis� est multipli� par 2 par rapport � la valeur
    'souple/bitumineuse qui n'est valable que pour les bitumineuses
    If uneStruct.monTypeStructure = Souple Then
        uneValPrec = uneValPrec * 2
    ElseIf uneStruct.monTypeStructure = PavesDalles And UCase(uneStruct.maCoucheBase) = "GNT" Then
        uneValPrec = uneValPrec * 2
    End If
End Function


Public Sub CalculerIndiceGelRef(uneForm As Form)
    'Calcul de l'indice de gel de r�f�rence corrig�
    'pour les qualit�s Q1 et Q2
    Dim unIndGelRef As Integer, unCoefAgglo As Single
    Dim uneHStation As Integer, uneHAgglo As Integer
    
    'Initialisation de l'indice de gel et du coef agglo
    'pour avoir un indice de gel inconnu � la cr�ation d'une �tude
    unIndGelRef = -1 'valeur d'indice de gel de r�f�rence inconnu
    unCoefAgglo = 1 'valeur neutre x * 1 = x
    
    'R�cup des donn�es de la station de r�f�rence
    uneHStation = monTabStation(uneForm.ComboStation.ListIndex + 1).monAltitude
    If uneForm.OptionHE.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHRE
    ElseIf uneForm.OptionHRNE.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHRNE
    ElseIf uneForm.OptionHC.Value Then
        unIndGelRef = monTabStation(uneForm.ComboStation.ListIndex + 1).monHC
    End If
    
    'R�cup des donn�es de l'agglo du projet
    If uneForm.ComboTailleAgglo.ListIndex = 0 Then
        unCoefAgglo = 1
    ElseIf uneForm.ComboTailleAgglo.ListIndex = 1 Then
        unCoefAgglo = 0.9
    ElseIf uneForm.ComboTailleAgglo.ListIndex = 2 Then
        unCoefAgglo = 0.8
    End If
    uneHAgglo = Format(uneForm.TextHAgglo.Text)
    
    'Calcul de l'indice de r�f�rence corrig�
    'Sans correction d'altitude pour l'instant
    'uneForm.monIndiceGelRefQ1 = CInt(unIndGelRef * unCoefAgglo * (1 - Abs(uneHStation - uneHAgglo) / uneHStation))
    uneForm.monIndiceGelRefQ1 = CInt(unIndGelRef * unCoefAgglo)
    uneForm.monIndiceGelRefQ2 = uneForm.monIndiceGelRefQ1
End Sub

Public Sub AfficherEtCalculerIndGelRef(uneForm As Form)
    'Calcul de l'indice de gel de r�f�rence
    'pour les qualit�s Q1 et Q2
    CalculerIndiceGelRef uneForm
    'Affichage dans la frame r�sultat de l'�tude active
    ActualiserFrameVerifGel uneForm
End Sub

Public Sub MettreAJourOngletGel(uneForm As Form)
    With uneForm
        'Mise � jour de l'onglet gel
        
        'Partie Station de r�f�rence + Agglo du projet
        .TextHAgglo.Text = Format(.monHAgglo)
        .TextHAgglo.Tag = Format(.monHAgglo)
        .TextHAgglo.ForeColor = QBColor(0)
        RemplirLesStationsM�t�o uneForm
        .ComboStation.ListIndex = .monIndStation - 1
        .ComboTailleAgglo.ListIndex = .monIndTailleAgglo
        
        'Partie hiver de r�f�rence
        If .monIndHiver = HE Then
            .OptionHE.Value = True
        ElseIf .monIndHiver = HRNE Then
            .OptionHRNE.Value = True
        ElseIf .monIndHiver = HC Then
            .OptionHC.Value = True
        End If
        
        'Partie couche de forme non g�live
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
            'Modif de la borne 0.05 pour la pente, c'et maintenant non g�lif
            'Avant non g�lif c'�tait pour pente < 0.05
            If CSng(.maPente) = 0.05 Then
                .OptionNGel.Value = True
                'car la ligne pr�c�dente remet TextPente � 0
                .TextPente.Text = Format(.maPente)
            End If
        End If
        .TextPente.ForeColor = QBColor(0)
        
        'Affichage ou masquage de l'indice de gel perso
        .CheckIndGelPerso.Value = .monUtilIndGelPerso
        
        'Mise � jour du calcul de l'indice de gel admin
        AfficherEtCalculerIndGelAdm uneForm
    End With
End Sub

Public Sub RemplirLesStationsM�t�o(uneForm As Form)
    'Remplissage du tableau de station de r�f�rence pour le gel
    'et Remplissage de la combobox ComboStation
        
    'Remise � z�ro du nombre de stations
    monNbStation = 0

    RemplirUneStationM�t�o uneForm, "Amb�rieu", "01", 253, 270, 175, 65
    RemplirUneStationM�t�o uneForm, "Saint Quentin", "02", 98, 225, 110, 45
    RemplirUneStationM�t�o uneForm, "Vichy", "03", 249, 250, 115, 45
    RemplirUneStationM�t�o uneForm, "St-Auban", "04", 459, 80, 35, 10
    RemplirUneStationM�t�o uneForm, "Embrun", "05", 871, 165, 95, 60
    RemplirUneStationM�t�o uneForm, "Nice", "06", 5, 0, 0, 0
    RemplirUneStationM�t�o uneForm, "St-Girons", "09", 411, 120, 35, 15
    RemplirUneStationM�t�o uneForm, "Romilly sur Seine", "10", 77, 210, 110, 35
    RemplirUneStationM�t�o uneForm, "Carcassone", "11", 126, 85, 35, 10
    RemplirUneStationM�t�o uneForm, "Millau", "12", 715, 140, 65, 40
    RemplirUneStationM�t�o uneForm, "Marignane", "13", 4, 70, 15, 0
    RemplirUneStationM�t�o uneForm, "Caen", "14", 64, 115, 60, 25
    RemplirUneStationM�t�o uneForm, "Cognac", "16", 30, 100, 35, 15
    RemplirUneStationM�t�o uneForm, "la Rochelle", "17", 4, 75, 30, 10
    RemplirUneStationM�t�o uneForm, "Bourges", "18", 161, 160, 70, 30
    RemplirUneStationM�t�o uneForm, "Ajaccio", "20", 4, 0, 0, 0
    RemplirUneStationM�t�o uneForm, "Dijon", "21", 222, 200, 130, 65
    RemplirUneStationM�t�o uneForm, "Rostrenen", "22", 262, 85, 50, 10
    RemplirUneStationM�t�o uneForm, "Besan�on", "25", 307, 220, 120, 70
    RemplirUneStationM�t�o uneForm, "Lus-la-Croix-Haute", "26", 1059, 420, 275, 160
    RemplirUneStationM�t�o uneForm, "Mont�limar", "26", 73, 105, 40, 10
    RemplirUneStationM�t�o uneForm, "Evreux", "27", 133, 195, 115, 60
    RemplirUneStationM�t�o uneForm, "Chartres", "28", 155, 190, 100, 35
    RemplirUneStationM�t�o uneForm, "Brest", "29", 96, 20, 10, 0
    RemplirUneStationM�t�o uneForm, "N�mes", "30", 59, 60, 20, 0
    RemplirUneStationM�t�o uneForm, "Toulouse", "31", 148, 115, 40, 10
    RemplirUneStationM�t�o uneForm, "Bordeaux", "33", 46, 95, 40, 10
    RemplirUneStationM�t�o uneForm, "Montpellier", "34", 5, 55, 35, 0
    RemplirUneStationM�t�o uneForm, "Dinard", "35", 58, 65, 25, 5
    RemplirUneStationM�t�o uneForm, "Rennes", "35", 36, 80, 35, 10
    RemplirUneStationM�t�o uneForm, "Chateauroux", "36", 156, 155, 75, 30
    RemplirUneStationM�t�o uneForm, "Tours", "37", 108, 120, 75, 35
    RemplirUneStationM�t�o uneForm, "Grenoble", "38", 384, 170, 145, 60
    RemplirUneStationM�t�o uneForm, "Mont de Marsan", 40, 59, 100, 40, 10
    RemplirUneStationM�t�o uneForm, "Romorantin", "41", 84, 135, 100, 30
    RemplirUneStationM�t�o uneForm, "St-Etienne", "42", 400, 220, 110, 60
    RemplirUneStationM�t�o uneForm, "le Puy", "43", 714, 240, 130, 65
    RemplirUneStationM�t�o uneForm, "Nantes", "44", 26, 75, 55, 10
    RemplirUneStationM�t�o uneForm, "Orl�ans", "45", 125, 170, 85, 45
    RemplirUneStationM�t�o uneForm, "Gourdon", "46", 259, 120, 45, 20
    RemplirUneStationM�t�o uneForm, "Agen", "47", 59, 110, 40, 15
    RemplirUneStationM�t�o uneForm, "Angers", "49", 57, 100, 70, 15
    RemplirUneStationM�t�o uneForm, "Cap de la Hague", "50", 3, 15, 5, 0
    RemplirUneStationM�t�o uneForm, "Reims", "51", 94, 235, 105, 80
    RemplirUneStationM�t�o uneForm, "Langres", "52", 464, 325, 170, 110
    RemplirUneStationM�t�o uneForm, "St-Dizier", "52", 139, 235, 100, 65
    RemplirUneStationM�t�o uneForm, "Nancy", "54", 212, 320, 155, 90
    RemplirUneStationM�t�o uneForm, "Bar le Duc", "55", 279, 340, 290, 130
    RemplirUneStationM�t�o uneForm, "Lorient", "56", 43, 40, 25, 10
    RemplirUneStationM�t�o uneForm, "Metz", "57", 190, 290, 135, 75
    RemplirUneStationM�t�o uneForm, "Ch�teau-Chinon", "58", 598, 225, 115, 80
    RemplirUneStationM�t�o uneForm, "Nevers", "58", 175, 190, 110, 60
    RemplirUneStationM�t�o uneForm, "Dunkerque", "59", 11, 165, 65, 20
    RemplirUneStationM�t�o uneForm, "Lille", "59", 47, 250, 90, 55
    RemplirUneStationM�t�o uneForm, "Beauvais", "60", 106, 215, 95, 40
    RemplirUneStationM�t�o uneForm, "Alen�on", "61", 144, 165, 70, 35
    RemplirUneStationM�t�o uneForm, "Boulogne sur Mer", "62", 73, 165, 70, 30
    RemplirUneStationM�t�o uneForm, "Clermont-Ferrand", "63", 320, 225, 115, 45
    RemplirUneStationM�t�o uneForm, "Biarritz", "64", 69, 40, 10, 0
    RemplirUneStationM�t�o uneForm, "Pau", "64", 183, 80, 30, 10
    RemplirUneStationM�t�o uneForm, "Tarbes", "65", 360, 95, 35, 10
    RemplirUneStationM�t�o uneForm, "Perpignan", "66", 42, 25, 0, 0
    RemplirUneStationM�t�o uneForm, "Strasbourg", "67", 150, 410, 165, 100
    RemplirUneStationM�t�o uneForm, "Mulhouse-B�le", "68", 267, 415, 155, 105
    RemplirUneStationM�t�o uneForm, "Lyon", "69", 200, 220, 110, 45
    RemplirUneStationM�t�o uneForm, "Tarare", "69", 831, 275, 155, 95
    RemplirUneStationM�t�o uneForm, "Luxeuil", "70", 272, 335, 165, 110
    RemplirUneStationM�t�o uneForm, "M�con", "71", 216, 200, 115, 60
    RemplirUneStationM�t�o uneForm, "Mont-St-Vincent", "71", 602, 270, 150, 95
    RemplirUneStationM�t�o uneForm, "le Mans", "72", 51, 120, 70, 25
    RemplirUneStationM�t�o uneForm, "Bourg-St-Maurice", "73", 865, 220, 190, 110
    RemplirUneStationM�t�o uneForm, "Challes les Eaux", "73", 291, 225, 150, 60
    RemplirUneStationM�t�o uneForm, "Cap de la H�ve", "76", 100, 95, 60, 20
    RemplirUneStationM�t�o uneForm, "Rouen", "76", 155, 130, 90, 30
    RemplirUneStationM�t�o uneForm, "Melun", "77", 91, 185, 90, 50
    RemplirUneStationM�t�o uneForm, "Abbeville", "80", 70, 165, 90, 50
    RemplirUneStationM�t�o uneForm, "Saint-Rapha�l", "83", 2, 25, 0, 0
    RemplirUneStationM�t�o uneForm, "Toulon", "83", 24, 15, 0, 0
    RemplirUneStationM�t�o uneForm, "Orange", "84", 83, 80, 45, 10
    RemplirUneStationM�t�o uneForm, "Poitiers", "86", 117, 130, 65, 25
    RemplirUneStationM�t�o uneForm, "Limoges", "87", 403, 160, 80, 30
    RemplirUneStationM�t�o uneForm, "Auxerre", "89", 207, 200, 95, 55
    RemplirUneStationM�t�o uneForm, "Belfort", "90", 422, 370, 175, 115
    RemplirUneStationM�t�o uneForm, "Paris le Bourget", "93", 59, 160, 85, 35
End Sub

Public Sub RemplirUneStationM�t�o(uneForm As Form, unNom As String, unNumDpt As String, uneAltitude As Integer, unHRE As Integer, unHRNE As Integer, unHC As Integer)
    'Remplissage de la station m�t�o d'indice unInd dans le
    'tableau des stations m�t�o
    'et Remplissage de la combobox ComboStation
    
    'Incr�mentation du nombre de station
    'Le tableau des stations va de 1 � 84
    monNbStation = monNbStation + 1
    
    'Remplissage du tableau des stations. Il va de 1 � 84
    monTabStation(monNbStation).monNom = unNom
    monTabStation(monNbStation).monNumDpt = unNumDpt
    monTabStation(monNbStation).monAltitude = uneAltitude
    monTabStation(monNbStation).monHRE = unHRE
    monTabStation(monNbStation).monHRNE = unHRNE
    monTabStation(monNbStation).monHC = unHC
        
    'Remplissage de la combobox ComboStation. Elle va de 0 � 83
    uneForm.ComboStation.AddItem (unNom + " (" + unNumDpt + ")")
End Sub

Public Sub AfficherEtCalculerIndGelAdm(uneForm As Form)
    'Calcul de l'indice de gel admissibles
    'pour les qualit�s Q1 et Q2
    CalculerIndiceGelAdm uneForm
    'Affichage dans la frame r�sultat de l'�tude active
    ActualiserFrameVerifGel uneForm
End Sub

Public Sub CalculerIndiceGelAdm(uneForm As Form)
    'Calcul de l'indice de gel admissible
    'pour les qualit�s Q1 et Q2
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
        'Remise � z�ro ==> Inconnu
        uneForm.monIndiceGelAdmQ1 = 0
        'Remise � z�ro ==> Inconnu
        uneForm.monIndiceGelAdmQ2 = 0
        Exit Sub
    ElseIf uneForm.TextPente.Text <> MsgInfinie Then
        If Format(uneForm.TextPente.Text) <= 0.05 Then
            'Cas d'une chauss�e hors gel pour les deux qualit�s
            'On affiche Chauss�e hors gel que les �paisseurs soient
            'trouv�es ou non ==> Indice de gel admissibles infini
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
    
    'R�cup des bonnes listes de mat�riaux des couches
    If uneForm.CheckFichPerso.Value = 0 Or uneForm.LabelFichPerso.Caption = "" Then
        Set uneColMatBF = maColMatBFCERTU
    Else
        Set uneColMatBF = maColMatBFPerso
    End If
    
    'R�cup de Agel
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
    
    'R�cup de Bgel
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
        
    'R�cup du tableau de variant contenant les �paisseurs
    unTabEp = uneForm.monTabEp
        
    If uneForm.monEpQ1Trouv Then
        'Si on a les �paisseurs pour la qualit� Q1
        '===> Calcul de l'indice de gel admissible pour la qualit� Q1
        'Valeur de 1 � 6 dans le tableau des �paisseurs
        'Calcul de ah
        unAH = 1 + (unAcs * (unTabEp(1) + unTabEp(2)) + unAcb * (unTabEp(3) + unTabEp(4)) + unAcf * (unTabEp(5) + unTabEp(6)))
        unBH = unBcs * (unTabEp(1) + unTabEp(2)) + unBcb * (unTabEp(3) + unTabEp(4)) + unBcf * (unTabEp(5) + unTabEp(6))
        uneForm.monIndiceGelAdmQ1 = 10 + (unAH * (uneForm.monQmQ1 + unQng + unQg) + unBH) * (unAH * (uneForm.monQmQ1 + unQng + unQg) + unBH) / 0.6
    Else
        'Remise � z�ro ==> Inconnu
        uneForm.monIndiceGelAdmQ1 = 0
    End If
    
    If uneForm.monEpQ2Trouv Then
        'Si on a les �paisseurs pour la qualit� Q
        '===> Calcul de l'indice de gel admissible pour la qualit� Q2
        'Valeur de 7 � 12 dans le tableau des �paisseurs
        unAH = 1 + (unAcs * (unTabEp(7) + unTabEp(8)) + unAcb * (unTabEp(9) + unTabEp(10)) + unAcf * (unTabEp(11) + unTabEp(12)))
        unBH = unBcs * (unTabEp(7) + unTabEp(8)) + unBcb * (unTabEp(9) + unTabEp(10)) + unBcf * (unTabEp(11) + unTabEp(12))
        uneForm.monIndiceGelAdmQ2 = 10 + (unAH * (uneForm.monQmQ2 + unQng + unQg) + unBH) * (unAH * (uneForm.monQmQ2 + unQng + unQg) + unBH) / 0.6
    Else
        'Remise � z�ro ==> Inconnu
        uneForm.monIndiceGelAdmQ2 = 0
    End If
End Sub


Public Function TrouverEpaisseurPossible(uneForm As Form)
    'Fonction retournant vrai si l'on peut trouver les �paisseurs pour Q1 et Q2
    '==> Tout est d�fini : le nombre d'essieux �quivalent (donc le trafic cumul�
    'et le CAM), la structure et la classe de plate-forme
    Dim uneStruct As Structure
    
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If uneForm.OptionPF1.Value = False And uneForm.OptionPF2.Value = False And uneForm.OptionPF2Plus.Value = False And uneForm.OptionPF3.Value = False Then
        'Cas o� la classe de plate-forme n'est pas renseign�e
        TrouverEpaisseurPossible = False
    ElseIf uneStruct Is Nothing Then
        'Cas o� la structure n'est pas renseign�e
        TrouverEpaisseurPossible = False
    ElseIf uneForm.monNEEquiv < uneStruct.monNbEssieuxMin Or uneForm.monNEEquiv > uneStruct.monNbEssieuxMax Then
        'Cas o� le NE calcul� est sup�rieure aux NE min et max
        'de la structure choisie
        TrouverEpaisseurPossible = False
    ElseIf InStr(1, uneForm.LabelNEequiv.Caption, MsgInconnu) > 0 Then
        'Cas o� le nombre d'essieux est inconnu
        TrouverEpaisseurPossible = False
    ElseIf InStr(1, uneForm.LabelNEequiv.Caption, MsgNEHorsLimite) > 0 Then
        'Cas o� le nombre d'essieux est hors limite (> 10 millions)
        TrouverEpaisseurPossible = False
    Else
        'Cas o� on peut trouver les �paisseurs
        TrouverEpaisseurPossible = True
    End If
End Function

Public Sub RechercherEpaisseur(uneForm As Form)
    'Recherche des �paisseurs de la structure choisie
    'correspondant au NE imm�diatement sup�rieur
    'et affichage des carottes pour Q1 et Q2
    Dim uneStruct As Structure, unNEth As Long
    Dim unTabEp As Variant, unInd As Integer
    Dim uneColInfoPFQ1 As Collection
    Dim uneColInfoPFQ2 As Collection
    
    'R�cup de la structure choisie
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If uneForm.monNEEquiv >= uneStruct.monNbEssieuxMin And uneForm.monNEEquiv <= uneStruct.monNbEssieuxMax Then
        'Cas o� le NE est entre le min et le max de la structure choisie
        'R�cup de la plateforme et affectation des bonnes collections d'�paisseurs
        If uneForm.OptionPF1.Value Then
            'Cas o� PF = PF1
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF1Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF1Q2
        ElseIf uneForm.OptionPF2.Value Then
            'Cas o� PF = PF2
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF2Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF2Q2
        ElseIf uneForm.OptionPF2Plus.Value Then
            'Cas o� PF = PF2+
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF2PlusQ1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF2PlusQ2
        Else
            'Pour tous les autres cas, PF = PF3
            Set uneColInfoPFQ1 = uneStruct.mesInfoPF3Q1
            Set uneColInfoPFQ2 = uneStruct.mesInfoPF3Q2
        End If
        
        'R�cup du tableau de variant contenant les �paisseurs
        unTabEp = uneForm.monTabEp
        
        'Calcul du NE th�orique imm�diatement sup�rieur au NE calcul�
        unNEth = uneForm.monNEEquiv
        TrouverNEthEtInd uneForm, uneStruct, unNEth, unInd
        uneForm.monNEth = unNEth
        
        'Affectation des �paisseurs dans le tableau des �paisseurs
        'De 1 � 6 pour Q1 et de 7 � 12 pour Q2
        'Avec un cas particulier pour la couche de surface o� la 2�me
        '�paisseur est nulle (d�termin� dans l'onglet Couche de surface)
        'unInd permet de trouver la bonne ligne d'�paisseur
        'celle correspondant au NE th�orique
        For i = 1 To 6
            If i = 1 Then
                unTabEp(i) = uneColInfoPFQ1((unInd - 1) * 8 + i)
                'Affectation de l'�paisseur pr�conis� au cas o�
                'le mat�riau de surface est compos�
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
                'Affectation de l'�paisseur pr�conis� au cas o�
                'le mat�riau de surface est compos�
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
        
        'Cas de couches de surfaces particuli�res
        If uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1 Then
            'Cas d'une couche de surface dont l'�paisseur n'a pas d'int�ret
            'On affecte une couche de 3 cm pour une bonne visu � l'�cran
            unTabEp(1) = 3 '�paisseur surface pour Q1
            unTabEp(7) = 3 '�paisseur surface pour Q2
        ElseIf uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pav�s" Then
            'Cas d'une couche de surface en dalles ou pav�s
            'rajout du lit de pose
            unTabEp(2) = EpLitPose '�paisseur lit de pose pour Q1
            unTabEp(8) = EpLitPose '�paisseur lit de pose pour Q2
        End If
    
        'Affectation du tableau de variant contenant les �paisseurs
        uneForm.monTabEp = unTabEp
        
        'Mise � jour de l'affichage des carottes Q1 et Q2
        uneForm.monMSComp1Q1 = ""
        uneForm.monMSComp2Q1 = ""
        uneForm.monMSComp1Q2 = ""
        uneForm.monMSComp2Q2 = ""
        AfficherCarottes uneForm
    Else
        'Epaisseurs Q1 et Q2 non trouv�es mais normalement on ne passe jamais l�
        uneForm.monEpQ1Trouv = False
        uneForm.monEpQ2Trouv = False
        'Mise � jour de l'affichage des carottes Q1 et Q2
        AfficherCarottes uneForm
    End If
    
    'Calcul de l'indice de gel admissible
    If mesOptionsGen.maVerifGel Then AfficherEtCalculerIndGelAdm uneForm
End Sub


Public Sub ValiderOngletCoucheSurface(uneForm As Form)
    'Proc�dure activant ou inhibant l'onglet de couche de surface
    'et le mettant � jour si il devient actif
    Dim uneStruct As Structure
    Dim unMatSurf As Object
    
    Set uneStruct = DonnerStructChoisie(uneForm)
    
    If Not (uneStruct Is Nothing) Then
        'Si mat�riau de surface compos� et �paisseurs trouv�es
        'pour Q1 et/ou Q2 ==> Activation de l'onglet Couche de surface
        If uneStruct.maCoucheSurface = "Aucune" Then
            uneForm.TabData.TabEnabled(OngletSurf) = False
        Else
            Set unMatSurf = maColMatSurf(uneStruct.maCoucheSurface)
            uneForm.TabData.TabEnabled(OngletSurf) = (TypeOf unMatSurf Is MatCompos�) And (uneForm.monEpQ1Trouv Or uneForm.monEpQ2Trouv)
            If uneForm.TabData.TabEnabled(OngletSurf) Then MettreAJourOngletCoucheSurface uneForm, uneForm.monEpPrecQ1, uneForm.monEpPrecQ2
        End If
    End If
End Sub

Public Sub TrouverNEthEtInd(uneForm As Form, uneStruct As Structure, unNEth As Long, unInd As Integer)
    'Proc�dure trouvant le NE th�orique, celui imm�diatment sup�rieur
    'au NE calcul� dans les NE de la structure choisie et son indice
    'dans les collections de donn�es PFi-Qj
    'Ces valeurs sont retourn�es dans unNEth et unInd
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
    'Mise � jour de l'onglet PlateForme
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
    'Fonction retournant vrai si l'�tude active est une nouvelle
    'et faux si c'est une �tude d�j� existante donc stock�e dans un fichier URB
    If Val(Mid(uneForm.Caption, 7, 1)) > 0 Then
        'Cas d'une nouvelle �tude ==> Titre de fen�tre = Etude N (N un entier > 0)
        EstNouvelleEtude = True
    Else
        'Cas d'une �tude existante ==> Titre de fen�tre = Etude + nom du fichier
        EstNouvelleEtude = False
    End If
End Function

Public Function DonnerNomTypeVoie(uneForm As Form) As String
    'Retourne le nom du type de voie de l'�tude active
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
    'TitreEtude et Dur�eCycle
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
        'et ainsi le F1 d�clenche la bonne aide
        monEtude.TabData.SetFocus
    End If
End Sub
