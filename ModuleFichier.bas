Attribute VB_Name = "ModuleFichier"
Public Const EnteteFichUrbV1 As String = "Fichier Struct-Urb 1.0"
Public Const EnteteFichUrbV2 As String = "Fichier Struct-Urb 2.0"


Public Function EcrireDansFichier(unNomFich As String, uneForm As Form) As Boolean
    'Ecriture dans le fichier unNomFich du contenu de l'�tude uneForm
    'Retour Vrai si tout ok, faux sinon
    Dim uneStructChoisie As Structure, uneColMatBF As Collection
    Dim unIndStructChoisie As Integer, unTypeVoieEtudeChantier As Integer
    Dim unIndMatBase As Integer, unIndMatFond As Integer
    
    EcrireDansFichier = True
    Set uneStructChoisie = DonnerStructChoisie(uneForm)
    If uneStructChoisie Is Nothing Then
        unIndMatBase = 0
        unIndMatFond = 0
    Else
        'R�cup de la bonne collection de mat�riaux base/fondation
        If uneForm.CheckFichPerso = 1 Then
            Set uneColMatBF = maColMatBFPerso
        Else
            Set uneColMatBF = maColMatBFCERTU
        End If
        unIndMatBase = TrouverIndexInCol(uneColMatBF, uneStructChoisie.maCoucheBase)
        unIndMatFond = TrouverIndexInCol(uneColMatBF, uneStructChoisie.maCoucheFondation)
    End If
    
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurEcriture
    
    ' Fermeture du fichier pour d�lock� et ainsi pouvoir �crire dedans.
    If uneForm.monFichId <> 0 Then
        'Cas d'un Site qui n'est pas Sans Nom (Titre Etude + unNum�ro)
        unFichId = uneForm.monFichId
        Close #unFichId
    End If
        
    'Ouvre le fichier en �criture.
    unFichId = FreeFile(0)
    uneForm.monFichId = unFichId
    Open unNomFich For Output As #unFichId
    
    'Mettre � jour la date de derni�re modif = derni�re sauvegarde
    uneForm.LabelDate.Caption = Format(Date, "dd/mm/yyyy")
    
    'Remplissage du fichier � partir des donn�es du site (=uneForm)
    '(cf Format de fichier Struct-Urb .urb)
    With uneForm
        'Ecriture de l'ent�te des fichiers *.urb
        Write #unFichId, EnteteFichUrbV2
        'Ecriture des donn�es de l'onglet Voie
        Write #unFichId, .TextTitre.Text
        Write #unFichId, FinTitre
        
        'Codage � partir de la version 2, de l'entier lu en quatri�me
        'position dans le fichier *.urb qui repr�sentait en V1 le type de voie
        'maintenant on y stocke la qualit� de chantier et le type d'�tude
        'cf l'event Form_Initialize
        unTypeVoieEtudeChantier = DonnerTypeVoie(uneForm)
        If .monTypeChantier = TypeChantierQ2 Then
            'Cas d'un chantier difficile = qualit� Q2
            monTypeChantier = TypeChantierQ2
            'Rajout forfaitaire de la valeur ChantierDifficile
            unTypeVoieEtudeChantier = unTypeVoieEtudeChantier + ChantierDifficile
        End If
        'Ecriture dans le fichier urb
        Write #unFichId, unTypeVoieEtudeChantier, .TextVar.Text, .LabelDate.Caption
        
        'Ecriture des donn�es de l'onglet Trafic
        Write #unFichId, .TextTrafIni.Text, CInt(.TextDur�eS.Text), DonnerCroissAn(uneForm), .TextTrafCUM.Text
        'Ecriture des donn�es de l'onglet Structure
        'Indice dans la liste de la combobox combostruct, num�ro d'index de la structure (perso ou CERTU)
        If .ComboStruct.ListIndex = -1 Then
            unIndStructChoisie = 0
        Else
            'R�cup de la structure choisie, transformation de la position dans
            'la collection des structures perso ou Certu par le num�ro d'index
            'de la structure choisie qui est son ordre de cr�ation
            If .CheckFichPerso.Value = 1 Then
                Set uneColStruct = maColStructPerso
            Else
                Set uneColStruct = maColStructCERTU
            End If
            unIndStructChoisie = uneColStruct(.ComboStruct.ItemData(.ComboStruct.ListIndex)).monNumIndex
        End If
        Write #unFichId, .ComboStruct.ListIndex, unIndStructChoisie, .CheckFichPerso.Value, .LabelFichPerso.Caption
        Write #unFichId, " " + Format(unIndMatBase) + " ", " " + Format(unIndMatFond) + " "
        'Ecriture des donn�es des onglet CAM et plateforme
        Write #unFichId, .MaskCAM.Text, DonnerIndicePF(uneForm)
        'Ecriture des donn�es de l'onglet couche surface
        '(+1 car les combobox vont de 0 � n-1)
        Write #unFichId, .ComboCompQ1.ListIndex + 1, .ComboCompQ2.ListIndex + 1
        'Ecriture des donn�es de l'onglet gel
        Write #unFichId, DonnerTypeHiver(uneForm), .ComboStation.ListIndex + 1, CInt(.TextHAgglo.Text), .ComboTailleAgglo.ListIndex
        Write #unFichId, DonnerTypeGelSol(uneForm), .TextPente.Text, DonnerCoefA(uneForm), CInt(.TextEpaisseur.Text)
        'Ecriture des donn�es de l'indice de gel perso, si elle n'existent
        'pas on met les valeurs par d�faut
        '==> une ligne de plus dans les *.urb
        'C'est le format des fichiers en version finale
        Write #unFichId, FormatFichierVersionFinale, uneForm.monIndGelPerso, uneForm.monUtilIndGelPerso
    End With
    
    'Mise � jour du titre de la fenetre �tude courante
    uneForm.Caption = MsgEtude0 + unNomFich
    
    'Fermeture du fichier.
    Close #unFichId
        
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    ' Quitte pour �viter le gestionnaire d'erreur.
    Exit Function
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurEcriture:
    
    EcrireDansFichier = False
    ' Traite les autres situations ici...
    unMsg = MsgErreur + Format(Err.Number) + " : " + Err.Description
    MsgBox unMsg, vbCritical
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'fermeture du fichier
    Close #unFichId
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    'On remet � jour la date de derni�re modif = derni�re sauvegarde
    uneForm.LabelDate.Caption = Format(uneForm.maDate, "dd/mm/yyyy")
    Exit Function
End Function

Public Sub OuvrirEtude(unNomFich As String)
    'Ouvre l'�tude contenue dans le fichier pass� en param�tre
    Dim uneString As String, unInt As Integer, unLong As Long
    Dim unReel As Single, uneString2 As String, unByte As Byte
    Dim frmD As frmDocument, unInt2 As Integer, unInt3 As Integer
    Dim unMatBase As String, unMatFond As String, uneColStruct As Collection
    Dim unUtilFichPerso As Byte, unFichPerso As String
    Dim unIndStructChoisie As Integer, unByte0 As Byte
    
    'suppression protection
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Sub
    'fin suppression protection
    
    'Indication de l'ouverture d'une �tude existante
    maNewEtude = False
    
    'Lecture du fichier .urb
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurLecture
    
    'Ouverture du fichier en lecture lock�e pour �viter deux ouvertures
    unFichId = FreeFile(0)
    Open unNomFich For Input Lock Read Write As #unFichId
    
    ViderCollection maColLectFich
    'Ajout du fichier id
    maColLectFich.Add unFichId
    
    'Remplissage de la collection contenant les donn�es du .urb
    '(cf Format de fichier Struct-Urb .urb)
    'Lecture de l'ent�te des fichiers *.urb
    Input #unFichId, uneString
    If uneString <> EnteteFichUrbV1 And uneString <> EnteteFichUrbV2 Then
        'Cas d'un fichier qui n'est pas un .urb
        '===> Fermeture du fichier.
        Close #unFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + " version 1 ou 2", vbCritical
    Else
        maColLectFich.Add uneString
        'Lecture du titre de l'�tude dans l'onglet Voie
        uneString2 = ""
        Input #unFichId, uneString
        Do While uneString <> FinTitre
            uneString2 = uneString2 + uneString
            Input #unFichId, uneString
        Loop
        maColLectFich.Add uneString2
        'Lecture des autres donn�es de l'onglet Voie
        Input #unFichId, unInt, uneString, uneString2
        maColLectFich.Add unInt
        'Cet entier lu ci-dessous permet en version 2 de trouver le type de voie,
        'le type d'�tudes et la qualit� du chantier
        maColLectFich.Add uneString
        maColLectFich.Add uneString2
        'Lecture des donn�es de l'onglet Trafic
        Input #unFichId, uneString, unInt, unByte, uneString2
        maColLectFich.Add uneString
        maColLectFich.Add unInt
        maColLectFich.Add unByte
        maColLectFich.Add uneString2
        'Lecture des donn�es de l'onglet Structure
        Input #unFichId, unInt, unInt2, unInt3, uneString
        maColLectFich.Add unInt
        maColLectFich.Add unInt2
        unIndStructChoisie = unInt2
        maColLectFich.Add unInt3
        unUtilFichPerso = unInt3
        maColLectFich.Add uneString
        unFichPerso = uneString
        'Lecture des index des mat�riaux de base et de fondation �ventuel
        'pour tester la coh�rence de l'�tude en cours d'ouverture
        'avec les options g�n�rales
        Input #unFichId, unMatBase, unMatFond
        'Lecture des donn�es des onglets CAM et plateforme
        Input #unFichId, uneString, unByte
        maColLectFich.Add uneString
        maColLectFich.Add unByte
        'Lecture des donn�es de l'onglet couche surface
        Input #unFichId, unInt, unInt2
        maColLectFich.Add unInt
        maColLectFich.Add unInt2
        'Lecture des donn�es de l'onglet gel
        Input #unFichId, unInt, unInt2, unInt3, unByte
        maColLectFich.Add unInt
        maColLectFich.Add unInt2
        maColLectFich.Add unInt3
        maColLectFich.Add unByte
        Input #unFichId, unInt, uneString, unReel, unInt2
        maColLectFich.Add unInt
        maColLectFich.Add uneString
        maColLectFich.Add unReel
        maColLectFich.Add unInt2
        
        'Stockage du titre de la fenetre d'�tude � ouvrir
        'en derni�re position
        maColLectFich.Add (MsgEtude0 + unNomFich)
            
        'Lecture des donn�es de l'indice de gel perso
        If EOF(unFichId) Then
            'Si Fin de fichier ==> format de fichier version 1.0 beta des
            'sites pilotes o� il n'y avait pas d'indice de gel perso
            maColLectFich.Add FormatFichierVersionBeta
        Else
            'Si pas fin de fichier on lit les info de l'indice du gel
            'perso sur la derni�re ligne, c'est le format de fichier
            'de la version 1.0 finale.
            'Cette ligne contient l'indicateur de version de format
            'de fichier urb, l'indice de gel perso et
            'l'�tat de la checkbox correspondnate
            Input #unFichId, unByte0, unInt, unByte
            maColLectFich.Add unByte0
            maColLectFich.Add unInt
            maColLectFich.Add unByte
        End If
       
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
        'R�cup de la structure choisie, transformation du num�ro d'index,
        'qui est son ordre de cr�ation,
        'par la position dans la collection des structures perso ou Certu
        If unUtilFichPerso = 1 Then
            Set uneColStruct = maColStructPerso
        Else
            Set uneColStruct = maColStructCERTU
        End If
        For i = 1 To uneColStruct.Count
            If uneColStruct(i).monNumIndex = unIndStructChoisie Then
                unIndStructChoisie = i
                Exit For
            End If
        Next i
        'Modif de l'indice structure choisie dans la collection des valeurs lues
        'car elle contient pour l'instant le num�ro d'index et pas la position
        'dans la liste des structures en position 12
        maColLectFich.Remove 12
        maColLectFich.Add unIndStructChoisie, , 12
        'Mettre � jour liste des fichiers r�cents
        ActualiserListeFichiersRecents unNomFich
        
        'OF cas o� l'utilisateur n'a pas choisi de structure
        If iunIndStructChoisie = 0 Then
            i = 0
        End If
        If i > uneColStruct.Count Then
            'Cas d'une structure n'existant plus dans le fichier de structures (.str)
            MsgBox "L'�tude (" + unNomFich + ") n'est pas compatible avec Struct-Urb version 2 et sup�rieure, car la structure, qui a �t� utilis�e en version 1, dans cette �tude n'existe plus. Vous devez utiliser la version 1 pour r�cup�rer les donn�es de cette �tude.", vbCritical
            ViderCollection maColLectFich
            'fermeture du fichier
            Close #unFichId
        ElseIf TesterCoh�renceEtude(unNomFich, unUtilFichPerso, unFichPerso, unIndStructChoisie, unMatBase, unMatFond) Then
            'Affichage de la fen�tre de l'�tude
            'si cette �tude est coh�rente avec les options g�n�rales
            'des mat�riaux
            Set frmD = New frmDocument
            If monOuverture Then
                frmD.Show
                unFileName = CorrigerNomFichier(App.Path + "\OngletVoie.rtf")
                frmD.RichTextAide.LoadFile unFileName, rtfRTF
                AfficherCarottes frmD
            Else
                'Cas d'erreur � l'ouverture
                monOuverture = True
                Close #unFichId
            End If
        Else
            ViderCollection maColLectFich
            'fermeture du fichier
            Close #unFichId
        End If
    End If
    ' Quitte pour �viter le gestionnaire d'erreur.
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurLecture:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'fermeture du fichier
    Close #unFichId
    Exit Sub
End Sub

Public Function SauverEtude(uneForm As Form, unNomFich As String, unSaveAs As Boolean) As String
    'Sauve l'�tude courante dans son fichier .urb si elle existe
    'ou demande un nom de fichier par s�lecteur si c'est une nouvelle �tude
    
    'suppresion protection
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then
    '    SauverEtude = ""
    '    Exit Function
    'End If
    'fin suppression protection
    
    If EstNouvelleEtude(uneForm) Or unSaveAs Then
        'Cas d'une nouvelle �tude ou d'un enregistrer sous d'une �tude existante
        unNomFich = fMainForm.ChoisirFichier(MsgSaveAs, MsgUrbFile, CurDir)
    End If
    
    If unNomFich <> "" Then
        'Cas o� l'utilisateur n'a pas fait annuler
        'dans le s�lecteur de fichiers
        'ou Cas d'une �tude existante (d�j� stock�e dans un fichier .URB)
        '==> unNomFich pas vide
        If EcrireDansFichier(unNomFich, uneForm) Then
            'Mettre � jour liste des fichiers r�cents
            ActualiserListeFichiersRecents unNomFich
            'Mettre � jour la date de derni�re modif = derni�re sauvegarde
            uneForm.LabelDate.Caption = Format(Date, "dd/mm/yyyy")
            uneForm.maDate = uneForm.LabelDate.Caption
            'Mettre � jour les donn�es sauv�es de l'�tude active
            'pour ne pas demander une sauvegarde lors de la fermeture
            'apr�s un Save ou un SaveAs
            uneForm.maModif = False
            uneForm.monTitreEtude = uneForm.TextTitre
            uneForm.maVariante = uneForm.TextVar
            uneForm.maDur�eService = uneForm.TextDur�eS.Text
            uneForm.maCroisAnnuel = DonnerCroissAn(uneForm)
            If uneForm.TextTrafIni.Text = "" Then
                uneForm.monTraficIni = 0
            Else
                uneForm.monTraficIni = CInt(uneForm.TextTrafIni.Text)
            End If
            If uneForm.TextTrafCUM.Text = "" Then
                uneForm.monTraficCumul� = 0
            Else
                uneForm.monTraficCumul� = CLng(uneForm.TextTrafCUM.Text)
            End If
            uneForm.monCAM = Format(uneForm.MaskCAM.Text, "fixed")
            uneForm.monIndicePF = DonnerIndicePF(uneForm)
            If uneForm.ComboStruct.ListIndex = -1 Then
                uneForm.monIndStructChoisie = 0
            Else
                uneForm.monIndStructChoisie = uneForm.ComboStruct.ItemData(uneForm.ComboStruct.ListIndex)
            End If
            uneForm.monUtilFichPerso = uneForm.CheckFichPerso.Value
            uneForm.monIndCompQ1 = uneForm.ComboCompQ1.ListIndex + 1
            uneForm.monIndCompQ2 = uneForm.ComboCompQ2.ListIndex + 1
            uneForm.monIndHiver = DonnerTypeHiver(uneForm)
            uneForm.monIndStation = uneForm.ComboStation.ListIndex + 1
            uneForm.monCoefA = DonnerCoefA(uneForm)
            uneForm.monIndTailleAgglo = uneForm.ComboTailleAgglo.ListIndex
        End If
    End If
    
    SauverEtude = unNomFich
End Function


Public Function TesterCoh�renceEtude(unNomFich As String, unUtilFichPerso As Byte, unFichPersoSTR As String, unIndStructChoisie As Integer, unMatBase As String, unMatFond As String) As Boolean
    'Fonction retournant vrai si l'�tude que l'on ouvre
    '(fichier = unNomFich) est coh�rente
    'avec les options mat�riaux (m�me fichier perso de structure et
    'structure choisie ne contenant pas mat�riaux non autoris�s
    Dim uneStruct As Structure, unMsg As String
    Dim uneStringListMatBFAuto As String
    Dim unMSurfKO As Boolean, unMBaseKO As Boolean, unMFondKO As Boolean
    
    If unUtilFichPerso = 1 And unFichPersoSTR <> mesOptionsMat.monFichPersoSTR Then
        'Cas o� le fichier perso utilis� n'est plus celui des options mat�riaux
        TesterCoh�renceEtude = False
        unMsg = MsgFichPersoKO1 + Chr(13) + Chr(13) + MsgFichPersoKO3 + Chr(13)
        unMsg = unMsg + "    " + unFichPersoSTR + Chr(13) + Chr(13)
        unMsg = unMsg + MsgFichPersoKO4 + Chr(13)
        unMsg = unMsg + "    " + mesOptionsMat.monFichPersoSTR + Chr(13) + Chr(13)
        unMsg = unMsg + MsgFichPersoKO2
        MsgBox unMsg, vbcritcal, MsgOpenError + unNomFich
    ElseIf unIndStructChoisie > 0 Then
        'Recherche dans le fichier de structures utilis� (CERTU ou PERSO)
        'si la structure choisie ne contient aucun mat�riau Base/Fondation
        'non autoris�
        
        'R�cup de la structure choisie
        If unUtilFichPerso = 1 Then
            Set uneStruct = maColStructPerso(unIndStructChoisie)
            uneStringListMatBFAuto = mesOptionsMat.mesMatPersoNonAutoris�s
            unMsg1 = MsgUseFichPerso + unFichPersoSTR
        Else
            Set uneStruct = maColStructCERTU(unIndStructChoisie)
            uneStringListMatBFAuto = mesOptionsMat.mesMatCERTUNonAutoris�s
            unMsg1 = MsgUseFichCERTU
        End If
        'Tester de contenance de mat�riau non autoris�e
        unMBaseKO = InStr(1, uneStringListMatBFAuto, unMatBase)
        unMFondKO = InStr(1, uneStringListMatBFAuto, unMatFond)
        If unMBaseKO Or unMFondKO Then
            unMsg = MsgEtude + unNomFich + Chr(13) + Chr(13) + MsgAyantStructChoix
            unMsg = unMsg + " " + uneStruct.monAbr�g�
            unMsg = unMsg + Chr(13) + Chr(13) + unMsg1
            unMsg = unMsg + Chr(13) + Chr(13) + MsgAvoirMatNonAuto0
            unMsg = unMsg + Chr(13) + Chr(13) + MsgMatAutoKO
            MsgBox unMsg, vbCritical, MsgOpenError + unNomFich
            TesterCoh�renceEtude = False
        Else
            TesterCoh�renceEtude = True
        End If
    Else
        TesterCoh�renceEtude = True
    End If
End Function


Public Function DonnerTypeVoie(uneForm As Form) As Integer
    'Retourne le type de voie de l'�tude active
    If uneForm.OptionVoieDes.Value Then
        DonnerTypeVoie = TypeVoieDesserte
    ElseIf uneForm.OptionVoieDis.Value Then
        DonnerTypeVoie = TypeVoieDistribution
    ElseIf uneForm.OptionVoiePL.Value Then
        DonnerTypeVoie = TypeVoieTraficLourd
    ElseIf uneForm.OptionVoieBus.Value Then
        DonnerTypeVoie = TypeVoieBus
    'Rajout de voie pour la version 2
    ElseIf uneForm.OptionVoieParking.Value Then
        DonnerTypeVoie = TypeVoieParking
    ElseIf uneForm.OptionGirDis.Value Then
        DonnerTypeVoie = TypeGiratoireDistribution
    ElseIf uneForm.OptionGirPL.Value Then
        DonnerTypeVoie = TypeGiratoireTraficLourd
    Else
        DonnerTypeVoie = TypeVoieInconnu
    End If
End Function

Public Function DonnerIndicePF(uneForm As Form) As Byte
    'Retourne le type de plateforme de l'�tude active
    DonnerIndicePF = Abs(uneForm.OptionPF1.Value * 1 + uneForm.OptionPF2.Value * 2 + uneForm.OptionPF3.Value * 3 + uneForm.OptionPF2Plus.Value * 4)
    'car True = -1 et False = 0
End Function

Public Function DonnerTypeHiver(uneForm As Form) As Integer
    'Retourne le type d'hiver de r�f�rence de l'�tude active
    DonnerTypeHiver = Abs(uneForm.OptionHE.Value * 1 + uneForm.OptionHRNE.Value * 2 + uneForm.OptionHC.Value * 3)
End Function

Public Function DonnerTypeGelSol(uneForm As Form) As Integer
    'Retourne le type de gel du sol support de l'�tude active
    DonnerTypeGelSol = Abs(uneForm.OptionTGel.Value * 1 + uneForm.OptionPGel.Value * 2 + uneForm.OptionNGel.Value * 3)
End Function

Public Function DonnerCoefA(uneForm As Form) As Single
    'Retourne le coefficient A de la couche
    'de forme non g�live de l'�tude active
    If uneForm.OptionANT.Value Then
        DonnerCoefA = 0.12
    Else
        DonnerCoefA = 0.14
    End If
End Function


Public Sub ActualiserListeFichiersRecents(unNomFich As String)
    'Mise � jour de la liste des fichiers r�cents (4 maximun)
    'avec le nom de fichier pass� en param�tre
    'Si ce nom n'est pas dans la liste des fichiers r�cents,
    'il devient num�ro 1, donc passe en t�te et le dernier est supprim�
    'de la liste et les autres d�cal�s de 1
    'S'il est dans la liste, il devient num�ro 1, donc passe en t�te et
    'les autres entre l'ancien 1 et nouveau 1 sont d�cal�s de 1
    
    'Recherche s'il est d�j� pr�sent dans les MRU
    'Dans les mnuFileMRU la chaine est du type "&i Nomfichier"
    For i = 0 To 3
        If fMainForm.mnuFileMRU(i).Visible Then unePos = i + 1
        If StrComp(unNomFich, Mid(fMainForm.mnuFileMRU(i).Caption, 4), vbTextCompare) = 0 Then
            'Comparaison de texte sans distinguer minuscule et majuscule
            unePos = i
            Exit For
        End If
    Next i
    
    'Cas o� le fichier �tait d�j� dans les MRU files et pas en t�te
    'ou absent (traitement idem que s'il �tait en dernier)
    'D�calage de 1 des MRU files entre les num�ros 0 et unePos-1
    If unePos = 4 Then unePos = 3
    For i = unePos To 1 Step -1
        fMainForm.mnuFileMRU(i).Caption = "&" + Format(i + 1) + Mid(fMainForm.mnuFileMRU(i - 1).Caption, 3)
        fMainForm.mnuFileMRU(i).Visible = True
    Next i
    
    'Mise en t�te du fichier en cours
    fMainForm.mnuFileMRU(0).Caption = "&1 " + unNomFich
    fMainForm.mnuFileMRU(0).Visible = True
    fMainForm.mnuFileBar6.Visible = True
End Sub

Public Function ModifierEtude(uneForm As Form) As Boolean
    'Fonction retournant si l'�tude (uneForm) a �t� modifi�e
    If uneForm.maModif Then
        'Permet de savoir si le type de voie a �t� chang� pour
        'une �tude existante, ou si modif de l'altim�trie de l'Agglo,
        'ou si modif l'�paisseur de couche non g�live ou si modif
        'de la pente du sol support
        ModifierEtude = True
    ElseIf uneForm.monTitreEtude <> uneForm.TextTitre Or uneForm.maVariante <> uneForm.TextVar Then
        'Cas o� le titre de l'�tude ou le texte de variante change
        ModifierEtude = True
    ElseIf Format(uneForm.maDur�eService) <> uneForm.TextDur�eS.Text Or DonnerCroissAn(uneForm) <> uneForm.maCroisAnnuel Then
        'Cas o� la dur�e de service ou la croissance annuelle change
        ModifierEtude = True
    ElseIf Format(uneForm.monTraficIni, "#,###") <> uneForm.TextTrafIni.Text Or Format(uneForm.monTraficCumul�, "###,###,###") <> uneForm.TextTrafCUM.Text Then
        'Cas o� le trafic initial ou cumul� change
        ModifierEtude = True
    ElseIf Format(uneForm.monCAM, "fixed") <> uneForm.MaskCAM.Text Or uneForm.monIndicePF <> DonnerIndicePF(uneForm) Then
        'Cas o� le CAM ou la plateforme change
        ModifierEtude = True
    ElseIf uneForm.monUtilFichPerso <> uneForm.CheckFichPerso.Value Then
        'Cas o� le fichier de structure (CERTU ou perso) change
        ModifierEtude = True
    ElseIf uneForm.ComboCompQ1.ListIndex <> (uneForm.monIndCompQ1 - 1) Or uneForm.ComboCompQ2.ListIndex <> (uneForm.monIndCompQ2 - 1) Then
        'Cas o� les compositions en couche de surface
        'pour les qualit�s Q1 et/ou Q2 change
        ModifierEtude = True
    ElseIf uneForm.monIndHiver <> DonnerTypeHiver(uneForm) Or uneForm.monIndStation <> uneForm.ComboStation.ListIndex + 1 Then
        'Cas o� l'hiver de r�f�rence ou la station de r�f�rence change
        ModifierEtude = True
    ElseIf uneForm.monCoefA <> DonnerCoefA(uneForm) Or uneForm.monIndTailleAgglo <> uneForm.ComboTailleAgglo.ListIndex Then
        'Cas o� la taille de l'agglo ou le coef A change
        ModifierEtude = True
    Else
        'Tous les autres cas ==> Etude non modifi�e
        ModifierEtude = False
    End If
    
    If uneForm.ComboStruct.ListIndex = -1 Then
        If uneForm.monIndStructChoisie <> 0 Then ModifierEtude = True
    Else
        If uneForm.monIndStructChoisie <> uneForm.ComboStruct.ItemData(uneForm.ComboStruct.ListIndex) Then
            'Cas o� la structure choisie change
            ModifierEtude = True
        End If
    End If
    
End Function
