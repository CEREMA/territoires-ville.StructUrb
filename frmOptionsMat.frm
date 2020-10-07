VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmOptionsMat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Param�tres d'utilisation des mat�riaux de couche de base et de fondation"
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
      Caption         =   "Fichier de Mat�riaux Personnels :"
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
      Caption         =   "Mat�riaux Personnels :"
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
      Caption         =   "Mat�riaux Standards :"
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
        'Cas d'un fichier personnel s�lectionn� diff�rent
        'de celui d�j� existant
        'On vide les structures et mat�riaux perso
        'avant le remplissage par lecture du fichier
        ViderCollection maColStructPerso
        ViderCollection maColMatBFPerso
        If OuvrirFichierStructures(unFich, maColStructPerso, maColMatBFPerso) Then
            'Cas de lecture compl�te du fichier de structures
            'Remplir un SpreadMat avec les mat�riaux base/fondation
            'avec leur autorisation d'utilisation
            FrameMatPerso.Enabled = True
            SpreadMatPerso.Visible = True
            RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutoris�s
            LabelFichMatPerso.Caption = unFich
        Else
            'Cas d'erreur lors de la lecture du fichier de structures
            '==> on ne l'utilise pas
            'On vide les structures et mat�riaux perso
            'partiellement remplies
            ViderCollection maColStructPerso
            ViderCollection maColMatBFPerso
            LabelFichMatPerso.Caption = ""
            RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutoris�s
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim unFich As String
    
    'R�cup�ration des anciennes valeurs d'options mat�riaux
    'dans la base de registre
    R�cup�rerOptionsMat
    unFich = mesOptionsMat.monFichPersoSTR
    
    'On vide les structures et mat�riaux perso
    'avant le remplissage par lecture du fichier perso pr�c�dent
    ViderCollection maColStructPerso
    ViderCollection maColMatBFPerso
    If unFich <> "" Then
        'Remplissage des structures et mat�riaux personnels
        Call OuvrirFichierStructures(unFich, maColStructPerso, maColMatBFPerso)
    End If
    
    'Restauration des autorisations d'utilisation � partir des indices
    'des mat�riaux PERSO non autoris�s stock�es dans les options mat�riaux
    AlimenterAutorisation False
    
    'Idem pour le fichier structures CERTU
    AlimenterAutorisation True
    
    FermerFenetre Me
End Sub

Private Sub cmdOK_Click()
    'Sauvegarde dans les options mat�riaux pour avoir
    'les m�mes valeurs pendant la session ouverte
    Dim uneString As String, unMsgErr As String
    Dim unNbMatCERTUAutoris�s As Integer, unListIndex As Integer
    Dim unNbMatPersoAutoris�s As Integer
    Dim unOKpossible As Boolean, uneStruct As Structure
    
    'V�rification que les �tudes ouvertes utilisent le
    'fichier personnel nouvellement choisi
    unOKpossible = True
    unMsgErr = ""
    For i = 1 To Forms.Count - 2
        'La 0 = Mdi m�re ==> pas une �tude
        'et la der est la boite d'options ==> pas une �tude non plus
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
    
    'V�rification qu'un mat�riau non autoris� n'appartient
    '� aucune structure choisie d'une �tude ouverte
    unOKpossible = True
    uneListMatCERTU = " "
    unMsgErr = ""
    For i = 1 To maColMatBFCERTU.Count
        SpreadMatCERTU.Col = 2
        SpreadMatCERTU.Row = i
        If SpreadMatCERTU.Value = 0 Then
            'Cas d'un mat�riau que l'on ne veut pas utiliser
            SpreadMatCERTU.Col = 1
            uneListMatCERTU = uneListMatCERTU + SpreadMatCERTU.Text + " "
        End If
    Next i
    uneListMatPerso = " "
    For i = 1 To maColMatBFPerso.Count
        SpreadMatPerso.Col = 2
        SpreadMatPerso.Row = i
        If SpreadMatPerso.Value = 0 Then
            'Cas d'un mat�riau que l'on ne veut pas utiliser
            SpreadMatPerso.Col = 1
            uneListMatPerso = uneListMatPerso + SpreadMatPerso.Text + " "
        End If
    Next i
    For i = 1 To Forms.Count - 2
        'La 0 = Mdi m�re ==> pas une �tude
        'et la der est la boite d'options ==> pas une �tude non plus
        If Forms(i).ComboStruct.ListIndex = -1 Then
            'Pas de structure choisie, donc on ne fait rien
        ElseIf Forms(i).LabelFichPerso.Caption = mesOptionsMat.monFichPersoSTR And Forms(i).CheckFichPerso.Value = 1 Then
            'Cas o� fichier perso = fichier perso des options mat�riaux
            Set uneStruct = maColStructPerso(Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex))
            unMatB = " " + uneStruct.maCoucheBase + " "
            unMatF = " " + uneStruct.maCoucheFondation + " "
            If InStr(1, uneListMatPerso, unMatB) > 0 Or InStr(1, uneListMatPerso, unMatF) > 0 Then
                unOKpossible = False
                unMsgErr = unMsgErr + Chr(13) + "- " + Forms(i).Caption + " " + MsgAyantStructChoix + " " + uneStruct.monAbr�g�
            End If
        ElseIf Forms(i).CheckFichPerso.Value = 0 Then
            'Cas o� fichier structures = celui du CERTU
            Set uneStruct = maColStructCERTU(Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex))
            unMatB = uneStruct.maCoucheBase
            unMatF = uneStruct.maCoucheFondation
            If InStr(1, uneListMatCERTU, unMatB) > 0 Or InStr(1, uneListMatCERTU, unMatF) > 0 Then
                unOKpossible = False
                unMsgErr = unMsgErr + Chr(13) + "- " + Forms(i).Caption + " " + MsgAyantStructChoix + " " + uneStruct.monAbr�g�
            End If
        End If
    Next i
    If unOKpossible = False Then
        unMsgErr = MsgChngOptMatImp + Chr(13) + Chr(13) + MsgEtudesSuivOuvert + " " + MsgAvoirMatNonAuto + " : " + unMsgErr
        MsgBox unMsgErr, vbCritical
        Exit Sub
    End If
    
    'Initialisation
    unNbMatCERTUAutoris�s = 0
    unNbMatPersoAutoris�s = 0
    
    mesOptionsMat.monFichPersoSTR = LabelFichMatPerso.Caption
    
    'Stockage des autorisations d'utilisation des mat�riaux CERTU
    SpreadMatCERTU.Col = 2
    uneString = " "
    For i = 1 To maColMatBFCERTU.Count
        SpreadMatCERTU.Row = i
        If SpreadMatCERTU.Value = 0 Then
            'Cas d'un mat�riau que l'on ne veut pas utiliser
            uneString = uneString + Format(i) + " "
        Else
            unNbMatCERTUAutoris�s = unNbMatCERTUAutoris�s + 1
        End If
        maColMatBFCERTU(i).monUtilisationAutoris�e = (SpreadMatCERTU.Value = 1)
        'En effet 1 = coch�e et 0 = non coch�e
    Next i
    mesOptionsMat.mesMatCERTUNonAutoris�s = uneString
    
    'Stockage des autorisations d'utilisation des mat�riaux PERSO
    SpreadMatPerso.Col = 2
    uneString = " "
    For i = 1 To maColMatBFPerso.Count
        SpreadMatPerso.Row = i
        If SpreadMatPerso.Value = 0 Then
            'Cas d'un mat�riau que l'on ne veut pas utiliser
            uneString = uneString + Format(i) + " "
        Else
            unNbMatPersoAutoris�s = unNbMatPersoAutoris�s + 1
        End If
        maColMatBFPerso(i).monUtilisationAutoris�e = (SpreadMatPerso.Value = 1)
        'En effet 1 = coch�e et 0 = non coch�e
    Next i
    mesOptionsMat.mesMatPersoNonAutoris�s = uneString
    
    If (maColMatBFPerso.Count = 0 Or unNbMatPersoAutoris�s > 0) And unNbMatCERTUAutoris�s > 0 Then
        'Sauvegarde dans la base de registre pour r�cup�rer
        'les m�mes valeurs � la prochaine
        StockerOptionsMat
                
        FermerFenetre Me
        
        'Mettre � jour toutes les fenetres d'�tudes ouvertes
        If mesOptionsMat.monFichPersoSTR <> "" Then
            For i = 1 To Forms.Count - 1 'La 0 = Mdi m�re ==> pas une �tude
                'Affichage de la case � cocher Utiliser fichier perso
                Forms(i).CheckFichPerso.Visible = True
            Next i
        End If
        
        For i = 1 To Forms.Count - 1 'La 0 = Mdi m�re ==> pas une �tude
            'Mise � jour du contenu de la combobox listant
            'les structures possibles de l'onglet Structure
            If Forms(i).LabelFichPerso.Caption <> mesOptionsMat.monFichPersoSTR And Forms(i).CheckFichPerso.Value = 1 Then Exit For
            'Cas d'une utilisation d'un fichier perso qui n'est pas
            'celui des options g�n�rales mat�riaux
            
            If Forms(i).ComboStruct.ListIndex = -1 Then
                'Pas de structure choisie
                unIndStruct = -1
            Else
                unIndStruct = Forms(i).ComboStruct.ItemData(Forms(i).ComboStruct.ListIndex)
            End If
            
            RemplirComboStructures Forms(i)
            
            unListIndex = -1
            For j = 0 To Forms(i).ComboStruct.ListCount - 1
                'recherche de la structure  s�lectionn� avant
                'elle ne doit pas chang�
                If Forms(i).ComboStruct.ItemData(j) = unIndStruct Then
                    unListIndex = j
                    Exit For
                End If
            Next j
            Forms(i).ComboStruct.Tag = "NoClickEvent"
            '==> pour ne pas d�clencher le click event de la combobox
            Forms(i).ComboStruct.ListIndex = unListIndex
            Forms(i).ComboStruct.Tag = ""
        Next i
    Else
        'Tous les mat�riaux CERTU ou PERSO ne sont pas autoris�s
        MsgBox MsgMatAutoris�, vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    CentrerFenetreEcran Me
    HelpContextID = IDhlp_WinOptionsMat
        
    'Affichage du fichier de structures personnelles �ventuel
    LabelFichMatPerso.Caption = mesOptionsMat.monFichPersoSTR
    
    'Affectation de la couleur du fond de cette fen�tre
    'aux cellules du SpreadMatCERTU et SpreadMatPerso
    SpreadMatCERTU.BackColor = vbInfoBackground
    SpreadMatPerso.BackColor = vbInfoBackground
    
    If mesOptionsMat.monFichPersoSTR = "" Then
        FrameMatPerso.Enabled = False
        SpreadMatPerso.Visible = False
    Else
        'Remplir SpreadMatPerso avec les mat�riaux base/fondation
        'personnels avec leur autorisation d'utilisation
        RemplirSpreadMat SpreadMatPerso, maColMatBFPerso, mesOptionsMat.mesMatPersoNonAutoris�s
    End If
    
    'Remplir SpreadMatCERTU avec les mat�riaux base/fondation
    'du CERTU avec leur autorisation d'utilisation
    RemplirSpreadMat SpreadMatCERTU, maColMatBFCERTU, mesOptionsMat.mesMatCERTUNonAutoris�s
End Sub

