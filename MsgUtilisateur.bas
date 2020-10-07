Attribute VB_Name = "MsgUtilisateur"
'Constantes contenant les messages d'erreurs possibles
Public Const MsgErreurProg As String = "ERREUR de programmation : "
Public Const MsgErreur As String = "ERREUR : "
Public Const MsgFileNotFile As String = "Ce fichier n'est pas un fichier "
Public Const MsgAltiAgglo As String = "L'altitude de l'agglom�ration"
Public Const MsgErreurCollectionInconnue As String = "Collection inconnue "
Public Const MsgErreurMat�riauInconnu As String = "Type de mat�riau inconnu "
Public Const MsgErreurTypeVoieInconnu As String = "Type de voie inconnu "
Public Const MsgErreurTypeEtudeInconnue As String = "Type d'�tude inconnu "
Public Const MsgErreurTypeShowWinInconnu As String = "Type de fen�tre ShowOpen ou ShowSave inconnu "
Public Const MsgErreurTailleAgglo As String = "Type de taille d'agglom�ration inconnu "
Public Const MsgIn As String * 5 = "dans "
Public Const MsgMaskCAM As String = "Coefficient d'Agressivit� Moyen"
Public Const MsgFinirSaisie As String = "Vous devez valider la saisie par la touche ENTREE ou remettre la valeur initiale par la touche ECHAP dans la zone de saisie en ROUGE : "
Public Const MsgDur�eService As String = "La dur�e de service "
Public Const MsgCroissAnn As String = "La croissance annuelle "
Public Const MsgTraficIni As String = "Le trafic initial � la mise en service pour un(e) "
Public Const MsgTraficCum As String = "La valeur du trafic cumul� a g�n�r�e un trafic initial � la mise en service de "
Public Const MsgTraficIni2 As String = "Mais le trafic initial � la mise en service pour un(e) "
Public Const MsgCoefCAM As String = "Le CAM "
Public Const MsgSupA As String = "doit �tre >= � "
Public Const MsgSupStrictA As String = "doit �tre > � "
Public Const MsgEtInfA As String = "et <= � "
Public Const MsgAnd As String * 3 = "et "
Public Const MsgSaisieEntierPositif As String = "Saisie d'entiers positifs uniquement"
Public Const MsgSaisieR�elPositif As String = "Saisie de nombres r�els positifs uniquement"
Public Const MsgFichStruct As String = "Le fichier de structures "
Public Const MsgRunError As String = " a d�clench� l'erreur : "
Public Const MsgIncorrect As String = " est incorrect "
Public Const MsgInexistant As String = " n'existe pas "
Public Const MsgStructCAM_SB As String = " et une structure Souple ou Bitumineuse"
Public Const MsgStructCAM_HB As String = " et une structure Hydraulique ou B�ton"
Public Const MsgFich As String = "Le fichier"
Public Const MsgFichStrKO As String = " est incorrect ou introuvable."
Public Const MsgOfThisEtude As String = " de cette �tude"
Public Const MsgMatAutoris� As String = "Il faut au moins un mat�riau CERTU et au moins un mat�riau Personnel autoris�"
Public Const MsgChngOptMatImp As String = "Changement des options mat�riaux impossibles"
Public Const MsgEtude0 As String = "Etude "
Public Const MsgEtude As String = "L'�tude "
Public Const MsgAyantStructChoix As String = "de structure"
Public Const MsgAvoirMatNonAuto0 As String = "contient des mat�riaux non autoris�s"
Public Const MsgAvoirMatNonAuto As String = "contiennent des mat�riaux non autoris�s"
Public Const MsgEtudesSuivOuvert As String = "Les �tudes ouvertes suivantes"
Public Const MsgFichPersoDiff As String = "utilisent un fichier de structures personnelles diff�rent"
Public Const MsgTextBoxTrafIni As String = "Trafic initial � la mise en service"
Public Const MsgTextBoxTrafCum As String = "Trafic cumul� PL"
Public Const MsgTextBoxDur�eS As String = "Dur�e de service"
Public Const MsgTextBoxHAgglo As String = "H Agglom�ration"
Public Const MsgTextIndGelPerso As String = "Indice de gel"
Public Const MsgTextBoxEpais As String = "Epaisseur Couche de forme non g�live"
Public Const MsgTextBoxPente As String = "Pente Sol Support"
Public Const MsgInconnu As String = "(Inconnu)"
Public Const MsgInfinie As String = "Infinie"
Public Const MsgChausseeHorsGel As String = "Hors Gel"
Public Const MsgNoPrintNoDim As String = "Impression impossible car la structure choisie n'est dimensionn�e ni pour la qualit� Q1 et ni pour la qualit� Q2."
Public Const MsgNoPrintNoStruct As String = "Impression impossible car aucune structure n'a �t� choisie."
Public Const MsgWaitConfig As String = "Configuration de l'impression en cours, Patientez SVP..."
Public Const MsgTitreFrmImp As String = "Imprimer"
Public Const MsgNEHorsLimite As String = "Erreur (> 10 millions)"
Public Const MsgNED�passement1 As String = "Vos hypoth�ses de trafic pour le calcul de NE conduisent � une valeur sup�rieure � 10 millions, limite au del� de laquelle l'application n'est plus adapt�e."
Public Const MsgNED�passement2 As String = "R�duisez NE (trafic, dur�e de service, taux de croissance, ...) ou reportez-vous aux m�thodes habituelles de dimensionnement (guide et catalogue SETRA-LCPC, documents CERTU, ...)"
Public Const MsgNEInfMin1 As String = "Vos hypoth�ses de trafic pour le calcul de NE conduisent � une valeur inf�rieure au minimun calcul� pour la structure choisie."
Public Const MsgNEInfMin2 As String = "Vous devez augmenter NE (trafic, dur�e de service, taux de croissance, ...) ou changer de structure."
Public Const MsgNESupMax1 As String = "Vos hypoth�ses de trafic pour le calcul de NE conduisent � une valeur sup�rieure au maximun calcul� pour la structure choisie."
Public Const MsgNESupMax2 As String = "Vous devez r�duire NE (trafic, dur�e de service, taux de croissance, ...) ou changer de structure."
Public Const MsgNECalcul� As String = "NE calcul�"
Public Const MsgMinTechno1 As String = "L'�paisseur indiqu�e est sup�rieure aux r�sultats du dimensionnement m�canique."
Public Const MsgMinTechno2 As String = "Elle correspond au minimun technologique de mise en oeuvre."
Public Const MsgMaxPra1 As String = "L'application des hypoth�ses prises pour les mat�riaux de qualit� "
Public Const MsgMaxPra2 As String = " conduit � des �paisseurs excessives (> 50 cm)."
Public Const MsgMaxPra3 As String = "Vous pouvez :"
Public Const MsgMaxPra4 As String = "   - Choisir des conditions de mise en oeuvre conduisant � une qualit� Q1"
Public Const MsgMaxPra5 As String = "   - Renforcer la plate-forme"
Public Const MsgMaxPra6 As String = "   - Changer de mat�riaux"
Public Const MsgFichPersoKO1 As String = "Le fichier de structures personnelles utilis� n'est plus celui des options mat�riaux."
Public Const MsgFichPersoKO2 As String = "Utiliser le menu Options / Param�tres mat�riaux pour changer le fichier de structures personnelles."
Public Const MsgFichPersoKO3 As String = "Fichier de structures personnelles utilis� par cette �tude : "
Public Const MsgFichPersoKO4 As String = "Fichier de structures personnelles utilis� des options mat�riaux : "
Public Const MsgMatAutoKO As String = "Utiliser le menu Options / Param�tres mat�riaux pour changer les mat�riaux autoris�s."
Public Const MsgUseFichPerso As String = "utilisant le fichier de structures personnelles "
Public Const MsgUseFichCERTU As String = "utilisant le fichier de structures standards "
Public Const MsgOpenError As String = "ERREUR d'ouverture de "
Public Const MsgOpen As String = "Ouvrir"
Public Const MsgDejaOpen As String = "Fichier d�j� ouvert"
Public Const MsgSaveAs As String = "Enregistrer sous"
Public Const MsgPrintInFile As String = "Imprimer dans un fichier"
Public Const MsgUrbFile As String = "Tous les fichiers (*.urb)|*.urb"
Public Const MsgTxtFile As String = "Tous les fichiers (*.txt)|*.txt"
Public Const MsgRTFFile As String = "Tous les fichiers (*.rtf)|*.rtf"
Public Const MsgStrFile As String = "Tous les fichiers (*.str)|*.str"
Public Const MsgSaveFile As String = "Voulez-vous enregistrer les modifications apport�es � "
Public Const MsgNoVerifGel As String = "Pas de v�rification au gel"
Public Const MsgValTol As String = "Cependant une valeur comprise entre "
Public Const MsgIsTol As String = " est possible." '" est tol�r�e."
Public Const MsgValLabo As String = "Mais pour une valeur sup�rieure � "
Public Const MsgIsLabo As String = " PL/J, il est fortement recommand� pour de tels trafics (T0) de faire appel � un bureau d'�tudes sp�cialis� et/ou d'utiliser des m�thodes de dimensionnement traditionnelles. Cependant Struct-Urb permet le calcul des �paisseurs pour un trafic jusqu'� 1000 PL/J."
Public Const MsgPassageEnQ2_1 As String = "Vous avez choisi un dimensionnement �tabli pour des conditions de chantier difficiles."
Public Const MsgPassageEnQ2_2 As String = "Les �paisseurs de mat�riaux ont �t� calcul�es en prenant comme param�tres m�caniques"
Public Const MsgPassageEnQ2_3 As String = "des valeurs diminu�es de 15 % pour repr�senter forfaitairement des conditions d�grad�es"
Public Const MsgPassageEnQ2_4 As String = "de mise en oeuvre. Il s'agit ainsi d'une marge de s�curit�."
Public Const MsgOpenStructFileFailed As String = "OpenStructFileFailed"
Public Const MsgMoreInfoForGiratoire As String = "Les �paisseurs d'assises calcul�es ont �t� major�es de 15 % pour tenir compte de mani�re forfaitaire des difficult�s inh�rentes aux chantiers de giratoires, � savoir, r�alisation sous circulation, mise en oeuvre en faibles quantit�s, mauvaise adaptation des mat�riels d'�pandage et de compactage et surtout dispersion des �paisseurs de mise en oeuvre."

'Constantes pour les labels de la frame r�sultats de gauche
Public Const LabelTypeVoieCaption As String = "Type de Voie : "
Public Const LabelTraficCumCaption As String = "Trafic Cumul� : "
Public Const LabelCAMCaption As String = "CAM : "
Public Const LabelIGelCaption As String = "Indice de Gel en �C.J"
Public Const LabelNEequivCaption As String = "Nombre d'Essieux Equivalents : "
Public Const LabelIGelAdminCaption As String = " Admissible = "
Public Const LabelIGelRefCaption As String = " R�f Corrig� = "
Public Const LabelCmdInfoMB As String = "Informations Mat�riau couche de base : "
Public Const LabelCmdInfoMF As String = "Informations Mat�riau couche de fondation : "
Public Const LabelCmdInfoMS As String = "Informations Mat�riau couche de surface : "
Public Const LabelValPrecCAMCaption As String = "Valeur pr�conis�e du CAM  ==>  "
Public Const LabelValMinCAMCaption As String = "Valeur minimale du CAM      ==>  "
Public Const LabelValMaxCAMCaption As String = "Valeur maximale du CAM     ==>  "
Public Const LabelInfo1Caption As String = "R�sultats calcul�s"
Public Const LabelInfo2Caption As String = "+ 15 %" '"augment�s de 15 %"
Public Const LabelEpPrecQ1Caption As String = "Epaisseur pr�conis�e Q1 :"
Public Const LabelCompChoisieQ1Caption As String = "Composition choisie Q1 :"
Public Const LabelWaitLoadStructures As String = "Chargement des structures en cours..."
