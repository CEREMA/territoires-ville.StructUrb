VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmOptionsMat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramètres d'utilisation des matériaux de couche de base et de fondation"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "frmOptionsMat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramefichMatPerso 
      Caption         =   "Fichier de Matériaux Personnels :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton btnParcourir 
         Caption         =   "Parcourir..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LabelFichMatPerso 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame FrameMatPerso 
      Caption         =   "Matériaux Personnels :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      Begin FPSpread.vaSpread SpreadMatPerso 
         Height          =   3495
         Left            =   240
         OleObjectBlob   =   "frmOptionsMat.frx":030A
         TabIndex        =   6
         Top             =   360
         Width           =   3300
      End
   End
   Begin VB.Frame FrameMatCERTU 
      Caption         =   "Matériaux Standards :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
      Begin FPSpread.vaSpread SpreadMatCERTU 
         Height          =   3495
         Left            =   240
         OleObjectBlob   =   "frmOptionsMat.frx":1490
         TabIndex        =   4
         Top             =   360
         Width           =   3300
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Tag             =   "Annuler"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Tag             =   "OK"
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptionsMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnParcourir_Click()
    Dim unFich As String
    
    unFich = fMainForm.ChoisirFichier(MsgOpen, MsgStrFile, CurDir)
    
    If unFich <> "" Then
        'Cas d'un fichier personnel sélectionné différent
        'de celui déjà existant
        'On vide les structures et matériaux perso
        'avant le remplissage par lecture du fichier
        ViderCollection maColStructPerso
        ViderCollection maColMatBFPerso
        If OuvrirFichierStructures(unFich, maColStructPerso, maColMatBFPerso) Then
            'Cas de lecture complète du fichier de structures
            'Remplir un SpreadMat avec les matériaux base/fondation
            'avec leur autorisation d'utilisation
            FrameMatPerso.Enabled = True
            SpreadMatPerso.Visible = True
            RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutorisés
            LabelFichMatPerso.Caption = unFich
        Else
            'Cas d'erreur lors de la lecture du fichier de structures
            '==> on ne l'utilise pas
            'On vide les structures et matériaux perso
            'partiellement remplies
            ViderCollection maColStructPerso
            ViderCollection maColMatBFPerso
            LabelFichMatPerso.Caption = ""
            RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutorisés
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim unFich As String
    
    'Récupération des anciennes valeurs d'options matériaux
    'dans la base de registre
    RécupérerOptionsMat
    unFich = mesOptionsMat.monFichPersoSTR
    
    'On vide les structures et matériaux perso
    'avant le remplissage par lecture du fichier perso précédent
    ViderCollection maColStructPerso
    ViderCollection maColMatBFPerso
    If unFich <> "" Then
        'Remplissage des structures et matériaux personnels
        Call OuvrirFichierStructures(unFich, maColStructPerso, maColMatBFPerso)
    End If
    
    'Restauration des autorisations d'utilisation à partir des indices
    'des matériaux PERSO non autorisés stockées dans les options matériaux
    AlimenterAutorisation False
    
    'Idem pour le fichier structures CERTU
    AlimenterAutorisation True
    
    FermerFenetre Me
End Sub

Private Sub cmdOK_Click()
    'Sauvegarde dans les options matériaux pour avoir
    'les mêmes valeurs pendant la session ouverte
    Dim uneString As String, unMsgErr As String
    Dim unNbMatCERTUAutorisés As Integer, unListIndex As Integer
    Dim unNbMatPersoAutorisés As Integer
    Dim unOKpossible As Boolean, uneStruct As Structure
    
    'Vérification que les études ouvertes utilisent le
    'fichier personnel nouvellement choisi
    unOKpossible = True
    unMsgErr = ""
    For i = 1 To Forms.Count - 2
        'La 0 = Mdi mère ==> pas une étude
        'et la der est la boite d'options ==> pas une étude non plus
        unFichPerso = Forms(i).LabelFichPerso.Caption
        If unFichPerso <> LabelFichMatPerso.Caption And unFichPerso <> "" Then
            unMsgErr = unMsgErr + Chr(13) + "- " + Forms(i).Caption
            unOKpossible = False
        End If
    Next i
    
    If unOKpossible = False Then
        unMsgErr = MsgChngOptMatImp + Chr(13) + Chr(13) + MsgEtudesSuivOuvert + " " + MsgFichPersoDiff + " : " + unMsgErr
        MsgBox unMsgErr, vbCritical
        Exit Sub
    End If
    
    'Vérification qu'un matériau non autorisé n'appartient
    'à aucune structure choisie d'une étude ouverte
    unOKpossible = True
    uneListMatCERTU = " "
    unMsgErr = ""
    For i = 1 To maColMatBFCERTU.Count
        SpreadMatCERTU.Col = 2
        SpreadMatCERTU.Row = i
        If SpreadMatCERTU.Value = 0 Then
            'Cas d'un matériau que l'on ne veut pas utiliser
            SpreadMatCERTU.Col = 1
            uneListMatCERTU = uneListMatCERTU + SpreadMatCERTU.Text + " "
        End If
    Next i
    uneListMatPerso = " "
    For i = 1 To maColMatBFPerso.Count
        SpreadMatPerso.Col = 2
        SpreadMatPerso.Row = i
        If SpreadMatPerso.Value = 0 Then
            'Cas d'un matériau que l'on ne veut pas utiliser
            SpreadMatPerso.Col = 1
            uneListMatPerso = uneListMatPerso + SpreadMatPerso.Text + " "
        End If
    Next i
    For i = 1 To Forms.Count - 2
        'La 0 = Mdi mère ==> pas une étude
        'et la der est la boite d'options ==> pas une étude non plus
        If Forms(i).ComboStruct.ListIndex = -1 Then
            'Pas de structure choisie, donc on ne fait rien
        ElseIf Forms(i).LabelFichPerso.Caption = mesOptionsMat.monFichPersoSTR And Forms(i).CheckFichPerso.Value = 1 Then
            'Cas où fichier perso = fichier perso des options matériaux
            Set uneStruct = maColStructPerso(Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex))
            unMatB = " " + uneStruct.maCoucheBase + " "
            unMatF = " " + uneStruct.maCoucheFondation + " "
            If InStr(1, uneListMatPerso, unMatB) > 0 Or InStr(1, uneListMatPerso, unMatF) > 0 Then
                unOKpossible = False
                unMsgErr = unMsgErr + Chr(13) + "- " + Forms(i).Caption + " " + MsgAyantStructChoix + " " + uneStruct.monAbrégé
            End If
        ElseIf Forms(i).CheckFichPerso.Value = 0 Then
            'Cas où fichier structures = celui du CERTU
            Set uneStruct = maColStructCERTU(Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex))
            unMatB = uneStruct.maCoucheBase
            unMatF = uneStruct.maCoucheFondation
            If InStr(1, uneListMatCERTU, unMatB) > 0 Or InStr(1, uneListMatCERTU, unMatF) > 0 Then
                unOKpossible = False
                unMsgErr = unMsgErr + Chr(13) + "- " + Forms(i).Caption + " " + MsgAyantStructChoix + " " + uneStruct.monAbrégé
            End If
        End If
    Next i
    If unOKpossible = False Then
        unMsgErr = MsgChngOptMatImp + Chr(13) + Chr(13) + MsgEtudesSuivOuvert + " " + MsgAvoirMatNonAuto + " : " + unMsgErr
        MsgBox unMsgErr, vbCritical
        Exit Sub
    End If
    
    'Initialisation
    unNbMatCERTUAutorisés = 0
    unNbMatPersoAutorisés = 0
    
    mesOptionsMat.monFichPersoSTR = LabelFichMatPerso.Caption
    
    'Stockage des autorisations d'utilisation des matériaux CERTU
    SpreadMatCERTU.Col = 2
    uneString = " "
    For i = 1 To maColMatBFCERTU.Count
        SpreadMatCERTU.Row = i
        If SpreadMatCERTU.Value = 0 Then
            'Cas d'un matériau que l'on ne veut pas utiliser
            uneString = uneString + Format(i) + " "
        Else
            unNbMatCERTUAutorisés = unNbMatCERTUAutorisés + 1
        End If
        maColMatBFCERTU(i).monUtilisationAutorisée = (SpreadMatCERTU.Value = 1)
        'En effet 1 = cochée et 0 = non cochée
    Next i
    mesOptionsMat.mesMatCERTUNonAutorisés = uneString
    
    'Stockage des autorisations d'utilisation des matériaux PERSO
    SpreadMatPerso.Col = 2
    uneString = " "
    For i = 1 To maColMatBFPerso.Count
        SpreadMatPerso.Row = i
        If SpreadMatPerso.Value = 0 Then
            'Cas d'un matériau que l'on ne veut pas utiliser
            uneString = uneString + Format(i) + " "
        Else
            unNbMatPersoAutorisés = unNbMatPersoAutorisés + 1
        End If
        maColMatBFPerso(i).monUtilisationAutorisée = (SpreadMatPerso.Value = 1)
        'En effet 1 = cochée et 0 = non cochée
    Next i
    mesOptionsMat.mesMatPersoNonAutorisés = uneString
    
    If (maColMatBFPerso.Count = 0 Or unNbMatPersoAutorisés > 0) And unNbMatCERTUAutorisés > 0 Then
        'Sauvegarde dans la base de registre pour récupérer
        'les mêmes valeurs à la prochaine
        StockerOptionsMat
                
        FermerFenetre Me
        
        'Mettre à jour toutes les fenetres d'études ouvertes
        If mesOptionsMat.monFichPersoSTR <> "" Then
            For i = 1 To Forms.Count - 1 'La 0 = Mdi mère ==> pas une étude
                'Affichage de la case à cocher Utiliser fichier perso
                Forms(i).CheckFichPerso.Visible = True
            Next i
        End If
        
        For i = 1 To Forms.Count - 1 'La 0 = Mdi mère ==> pas une étude
            'Mise à jour du contenu de la combobox listant
            'les structures possibles de l'onglet Structure
            If Forms(i).LabelFichPerso.Caption <> mesOptionsMat.monFichPersoSTR And Forms(i).CheckFichPerso.Value = 1 Then Exit For
            'Cas d'une utilisation d'un fichier perso qui n'est pas
            'celui des options générales matériaux
            
            If Forms(i).ComboStruct.ListIndex = -1 Then
                'Pas de structure choisie
                unIndStruct = -1
            Else
                unIndStruct = Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex)
            End If
            
            RemplirComboStructures Forms(i)
            
            unListIndex = -1
            For j = 0 To Forms(i).ComboStruct.ListCount - 1
                'recherche de la structure  sélectionné avant
                'elle ne doit pas changé
                If Forms(i).ComboStruct.ItemData(j) = unIndStruct Then
                    unListIndex = j
                    Exit For
                End If
            Next j
            Forms(i).ComboStruct.Tag = "NoClickEvent"
            '==> pour ne pas déclencher le click event de la combobox
            Forms(i).ComboStruct.ListIndex = unListIndex
            Forms(i).ComboStruct.Tag = ""
        Next i
    Else
        'Tous les matériaux CERTU ou PERSO ne sont pas autorisés
        MsgBox MsgMatAutorisé, vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    CentrerFenetreEcran Me
    HelpContextID = IDhlp_WinOptionsMat
        
    'Affichage du fichier de structures personnelles éventuel
    LabelFichMatPerso.Caption = mesOptionsMat.monFichPersoSTR
    
    'Affectation de la couleur du fond de cette fenêtre
    'aux cellules du SpreadMatCERTU et SpreadMatPerso
    SpreadMatCERTU.BackColor = vbInfoBackground
    SpreadMatPerso.BackColor = vbInfoBackground
    
    If mesOptionsMat.monFichPersoSTR = "" Then
        FrameMatPerso.Enabled = False
        SpreadMatPerso.Visible = False
    Else
        'Remplir SpreadMatPerso avec les matériaux base/fondation
        'personnels avec leur autorisation d'utilisation
        RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutorisés
    End If
    
    'Remplir SpreadMatCERTU avec les matériaux base/fondation
    'du CERTU avec leur autorisation d'utilisation
    RemplirSpreadMat SpreadMatCERTU, maColMatBFCERTU, mesOptionsMat.mesMatCERTUNonAutorisés
End Sub

