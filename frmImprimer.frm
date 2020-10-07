VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmImprimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimer"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmImprimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckPrintFichier 
      Caption         =   "Imprimer dans un fichier au format RTF"
      Height          =   405
      Left            =   5160
      TabIndex        =   11
      Top             =   1800
      Width           =   2000
   End
   Begin RichTextLib.RichTextBox RichTextTmp 
      Height          =   615
      Left            =   5160
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmImprimer.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameInfoImp 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Image ImagePortrait 
         Height          =   600
         Left            =   1320
         Picture         =   "frmImprimer.frx":0386
         Stretch         =   -1  'True
         Top             =   720
         Width           =   615
      End
      Begin VB.Image ImagePaysage 
         Height          =   495
         Left            =   1320
         Picture         =   "frmImprimer.frx":0AC8
         Stretch         =   -1  'True
         Top             =   720
         Width           =   600
      End
      Begin VB.Label NomImp 
         AutoSize        =   -1  'True
         Caption         =   "Imprimante courante : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orientation :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Configurer l'imprimante..."
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Tag             =   "Annuler"
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Tag             =   "Annuler"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame FrameOptionImp 
      Caption         =   "Informations complémentaires à imprimer : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4935
      Begin VB.CheckBox CheckInfoGel 
         Caption         =   "Informations complémentaires sur le gel"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CheckBox CheckCommentStruct 
         Caption         =   "Commentaire de la structure choisie"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   4455
      End
      Begin VB.CheckBox CheckCommentMat 
         Caption         =   "Commentaires de tous les matériaux de la structure choisie"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmImprimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Constante pour la marge haut, bas, droite et gauche
'en twips sachant que 567 twips = 1 cm
Const maMarge As Single = 567

'Constante pour le facteur d'échelle de représentation
'sur l'imprimante de 1 cm d'épaisseur
'On veut que 10 cm à l'imprimante = 50 cm réelle (567 twips = 1 cm)
'Epaisseur imprimante = epaisseur réelle *567 / 5
Const monEch As Single = 567 / 5

'Variable permettant de dire si les ordres d'impressions
'vont sur l'imprimante (unePrintZone) ou dans une picture box (PictureCarotte)
'de la fenêtre de l'étude en cours que 'lon collera grâce au presse-papier
'dans un controle richtextbox pour faire une sortie en RTF
Private unePrintZone As Object


Private Sub cmdCancel_Click()
    FermerFenetre Me
End Sub

Private Sub cmdConfig_Click()
    Caption = UCase(MsgWaitConfig)
    ' Active la routine de gestion d'erreur.
    On Error GoTo CancelPress
    
    'Cas où l'on travaille sur la zone d'impression qui est une imprimante
    Set unePrintZone = Printer
    'Affichage de la fenetre de configuration d'imprimante
    fMainForm.dlgCommonDialog.CancelError = True
    fMainForm.dlgCommonDialog.Flags = cdlPDPrintSetup
    fMainForm.dlgCommonDialog.ShowPrinter
    Caption = MsgTitreFrmImp
    'Mise à jour du nom de l'imprimante courante
    NomImp.Caption = "Imprimante courante : " + Chr(13) + unePrintZone.DeviceName
    'Mise à jour de l'orientation
    If unePrintZone.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
    End If
    
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub 'Pour éviter le traitement des erreurs s'il n'y a pas eu
    
    'Gestion des erreurs
CancelPress:
    
    Caption = MsgTitreFrmImp
    Select Case Err.Number
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

   
Private Sub cmdOK_Click()
    Dim unFileId As Integer
    
    If CheckPrintFichier.Value = 0 Then
        'Cas d'une impression directe sur l'imprimante
        Set unePrintZone = Printer
        'Utilisation de la fonte arial pour les sorties
        UtiliserFonteArial
        
        'Impression en ligne pleine
        Printer.DrawStyle = vbSolid
        
        'Impression générale
        ImprimerTitre
        ImprimerDonnées
        ImprimerValInter
        ImprimerGel
        ImprimerCarottes
        
        'Impressions complémentaires
        If CheckCommentMat.Value = 1 Or CheckCommentStruct.Value = 1 Or CheckInfoGel.Value = 1 Then
            'Si on imprime des données complémentaires on passe à la page suivante
            Printer.NewPage
            ImprimerEntête
        End If
        
        If CheckCommentMat.Value = 1 Then ImprimerInfoMat
        If CheckCommentStruct.Value = 1 Then ImprimerInfoStruct
        If CheckInfoGel.Value = 1 Then ImprimerInfoGel
        
        'Envoi à l'imprimante
        Printer.EndDoc
    Else
        'Cas d'une impression dans un fichier texte en version 1
        'ImpressionInFile
        'Cas d'une impression dans un fichier RTF en version 2
        ImpressionInFileRTF
    End If
    
    FermerFenetre Me
End Sub

Private Sub Form_Load()
    Caption = MsgTitreFrmImp
    CentrerFenetreEcran Me
    HelpContextID = IDhlp_WinPrint
    NomImp.Caption = "Imprimante courante : " + Chr(13) + Printer.DeviceName
    
    CheckInfoGel.Enabled = mesOptionsGen.maVerifGel
    
    'Mise à jour de l'orientation
    If Printer.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
    End If
End Sub

Private Sub ImprimerTitre()
    Dim unePosRC As Integer
    Dim uneString As String
    Dim unCurX As Single, unCurY As Single
                
    'Imprimer Entête de page avec struct-urb + numéro de version
    ImprimerEntête
    
    'Impression du titre de l'étude
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.CurrentY = maMarge * 1.1
    unePrintZone.Font.Size = 12
    unePrintZone.Font.Bold = True
    unePrintZone.Print "TITRE DE L'ETUDE :"
    
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Bold = False
    unePrintZone.Font.Underline = False
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec le titre en gras
        
    uneString = monEtude.TextTitre
    unePosRC = InStr(1, uneString, Chr(13))
    Do While unePosRC > 0
        unePrintZone.CurrentX = maMarge * 1.1
        unePrintZone.Print Mid(uneString, 1, unePosRC - 1)
        uneString = Mid(uneString, unePosRC + 2)
        '+2 pour repartir aprés le retour chariot et le saut de ligne
        unePosRC = InStr(1, uneString, Chr(13))
    Loop
    
    If unePosRC = 0 And uneString <> "" Then
        'Cas où plus de retour chariot et
        'affichage du reste du titre s'il en reste
        unePrintZone.CurrentX = maMarge * 1.1
        unePrintZone.Print uneString
    End If
    
    unePrintZone.Font.Bold = True
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.CurrentY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE")
    unCurX = unePrintZone.CurrentX + unePrintZone.TextWidth("Date : ")
    unCurY = unePrintZone.CurrentY
    unePrintZone.Print "Date : "
    unePrintZone.CurrentX = unCurX
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print monEtude.LabelDate
    
    unePrintZone.Font.Bold = True
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX + unePrintZone.TextWidth("Variante : ")
    unCurY = unePrintZone.CurrentY
    unePrintZone.Print "Variante : "
    unePrintZone.CurrentX = unCurX
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print monEtude.TextVar
    
    unePrintZone.Font.Bold = True
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX + unePrintZone.TextWidth("Enregistrée sous : ")
    unCurY = unePrintZone.CurrentY
    unePrintZone.Print "Enregistrée sous : "
    unePrintZone.CurrentX = unCurX
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    If EstNouvelleEtude(monEtude) = False Then
        'Cas où ce n'est pas une étude nouvelle Etude N
        unePrintZone.Print Mid(monEtude.Caption, 7)
    Else
        unePrintZone.Print "Etude pas encore enregistrée"
    End If
    
    'Dessin d'un cadre autour de la partie titre
    unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE") / 2
    unePrintZone.CurrentX = maMarge
    unePrintZone.CurrentY = unCurY
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY), QBColor(0)
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, maMarge), QBColor(0)
    unePrintZone.Line -(maMarge, maMarge), QBColor(0)
    unePrintZone.Line -(maMarge, unCurY), QBColor(0)
End Sub
    
Private Sub ImprimerDonnées()
    Dim unCurY As Single, unCurX As Single, unCurY0 As Single
    Dim unTypePL As String, uneStrTmp As String
    
    'Impression des données de l'étude
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.CurrentY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE")
    unCurY0 = unePrintZone.CurrentY - unePrintZone.TextHeight("TITRE") / 2
    unePrintZone.Font.Size = 12
    unePrintZone.Font.Bold = True
    unePrintZone.Print "DONNEES :"
    
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Underline = False
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec le titre en gras
    
    'Impression du type de voie
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Print "Type de voie : "
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Type de voie : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print DonnerNomTypeVoie(monEtude)
        
    'Impression du type d'étude
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Type d'aménagement : "
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Type d'aménagement : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    If monEtude.monTypeEtude = TypeEtudeStandard Then
        uneStrTmp = monEtude.OptionEtudeStandard.Caption
    Else
        uneStrTmp = monEtude.OptionEtudeGiratoire.Caption
    End If
    unePrintZone.Print uneStrTmp
    
    'Impression du type de conditions de chantier
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Chantier : "
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Chantier : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    If monEtude.monTypeEtude = TypeEtudeGiratoire Then
        uneStrTmp = monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption
    ElseIf monEtude.monTypeChantier = TypeChantierQ1 Then
        uneStrTmp = monEtude.OptionChoixQ1.Caption
    Else
        uneStrTmp = monEtude.OptionChoixQ2.Caption
    End If
    unePrintZone.Print uneStrTmp
    
    'Impression des données trafic ini, durée service
    'et croisance annuelle et de la classe de plateforme
    If DonnerTypeVoie(monEtude) = 4 Then
        unTypePL = "BUS"
    Else
        unTypePL = "Poids Lourds"
    End If

    'modification OF le 21/07/2005 : pas d'affichage des données de trafic si type de voie=parking
    Dim unTrafic, uneDureeS, unTauxC As String
    
    If DonnerTypeVoie(monEtude) = 5 Then '5=type de voie parking
        unTypePL = "Poids Lourds"
        unTrafic = "12"
        uneDureeS = "20 ans"
        unTauxC = "1 % par an"
    Else
        unTrafic = monEtude.TextTrafIni.Text
        uneDureeS = monEtude.TextDuréeS.Text + " ans"
        unTauxC = Format(DonnerCroissAn(monEtude)) + " % par an"
    End If
    'fin modification OF le 21/07/2005
    
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Trafic initial à la mise en service (par sens, par voie et par jour) : "
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Trafic initial à la mise en service (par sens, par voie et par jour) : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print unTrafic + " " + unTypePL
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Durée de service : "
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Durée de service : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print uneDureeS
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Taux de croissance : "
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Taux de croissance : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print unTauxC
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Plate-forme : "
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Plate-forme : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unIndPF = DonnerIndicePF(monEtude)
    If unIndPF = 4 Then
        unePrintZone.Print "PF2+"
    Else
        unePrintZone.Print "PF" + Format(unIndPF)
    End If
    
    'Dessin d'un cadre autour de la partie titre
    unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE") / 2
    unePrintZone.CurrentX = maMarge
    unePrintZone.CurrentY = unCurY
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY), QBColor(0)
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY0), QBColor(0)
    unePrintZone.Line -(maMarge, unCurY0), QBColor(0)
    unePrintZone.Line -(maMarge, unCurY), QBColor(0)
End Sub


    
Private Sub ImprimerValInter()
    Dim unCurY As Single, unCurX As Single, unCurY0 As Single
    Dim unNbTypePL As String, uneStruct As Structure
    
    'Impression des valeurs intermédiaires de l'étude
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.CurrentY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE")
    unCurY0 = unePrintZone.CurrentY - unePrintZone.TextHeight("TITRE") / 2
    unePrintZone.Font.Size = 12
    unePrintZone.Font.Bold = True
    unePrintZone.Print "VALEURS INTERMEDIAIRES :"
    
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Underline = False
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec le titre en gras
           
    'Impression des données trafic cumulé, risque de calcul
    'si valeur présente, le CAM et le NE arrondi
    If DonnerTypeVoie(monEtude) = 4 Then
        unNbTypePL = "Nombre Cumulé de BUS : "
    Else
        unNbTypePL = "Nombre Cumulé de Poids Lourds : "
    End If
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print unNbTypePL
    
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth(unNbTypePL)
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print monEtude.TextTrafCUM.Text
    
    'Si une structure est choisie
    'Affichage du risque de calcul si non nul
    Set uneStruct = DonnerStructChoisie(monEtude)
    If Not (uneStruct Is Nothing) Then
        If uneStruct.monTauxRisque <> 0 Then
            unePrintZone.CurrentX = maMarge * 1.1
            unCurX = unePrintZone.CurrentX
            unCurY = unePrintZone.CurrentY
            unePrintZone.Font.Bold = True
            unePrintZone.Print "Risque de calcul : "
            
            unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Risque de calcul : ")
            unePrintZone.CurrentY = unCurY
            unePrintZone.Font.Bold = False
            unePrintZone.Print Format(uneStruct.monTauxRisque) + " %"
        End If
    End If
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "CAM : "
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("CAM : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print monEtude.MaskCAM.Text
    
    unePrintZone.CurrentX = maMarge * 1.1
    unCurX = unePrintZone.CurrentX
    unCurY = unePrintZone.CurrentY
    unePrintZone.Font.Bold = True
    unePrintZone.Print "NE arrondi : "
    unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("NE arrondi : ")
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Bold = False
    unePrintZone.Print Format(monEtude.monNEth, "##,###,###")
    
    'Dessin d'un cadre autour de la partie titre
    unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE") / 2
    unePrintZone.CurrentX = maMarge
    unePrintZone.CurrentY = unCurY
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY), QBColor(0)
    unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY0), QBColor(0)
    unePrintZone.Line -(maMarge, unCurY0), QBColor(0)
    unePrintZone.Line -(maMarge, unCurY), QBColor(0)
End Sub
    


Private Sub ImprimerGel()
    Dim unCurY As Single, unCurX As Single, unCurY0 As Single
    Dim uneInfoGelQ1 As String, uneInfoGelQ2 As String
    
    If mesOptionsGen.maVerifGel Then
        'Impression des valeurs de vérif au gel de l'étude
        unePrintZone.CurrentX = maMarge * 1.1
        unePrintZone.CurrentY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE")
        unCurY0 = unePrintZone.CurrentY - unePrintZone.TextHeight("TITRE") / 2
        unePrintZone.Font.Size = 12
        unePrintZone.Font.Bold = True
        unePrintZone.Print "GEL :"
        
        unePrintZone.Font.Size = 10
        unePrintZone.Font.Underline = False
        unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
        'Pour sauter une demi-ligne pour espacer avec le titre en gras
        
        'Affichage des résultats gel pour la qualité Q1
        'en version 2, on n'affiche qu'une seule carotte
        If monEtude.monEpQ1Trouv And monEtude.monTypeChantier = TypeChantierQ1 Then
            unePrintZone.CurrentX = maMarge * 1.1
            unePrintZone.Font.Bold = True
            If monEtude.monTypeEtude = TypeEtudeGiratoire Then
                unePrintZone.Print "En condition de chantier : " + monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption
            Else
                unePrintZone.Print "En condition de chantier standard (qualité Q1) : "
            End If
            unePrintZone.CurrentX = maMarge * 2
            unCurX = unePrintZone.CurrentX
            unCurY = unePrintZone.CurrentY
            unePrintZone.Print "Indice de Gel de Référence corrigé : "
            unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Indice de Gel de Référence corrigé : ")
            unePrintZone.CurrentY = unCurY
            unePrintZone.Font.Bold = False
            unePrintZone.Print Format(monEtude.monIndiceGelRefQ1) + " °C.j"
            
            unePrintZone.Font.Bold = True
            unePrintZone.CurrentX = maMarge * 2
            unCurX = unePrintZone.CurrentX
            unCurY = unePrintZone.CurrentY
            unePrintZone.Print "Indice de Gel Admissible : "
            unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Indice de Gel Admissible : ")
            unePrintZone.CurrentY = unCurY
            unePrintZone.Font.Bold = False
            If monEtude.monIndiceGelRefQ1 > monEtude.monIndiceGelAdmQ1 Then
                uneInfoGelQ1 = " ============> Chaussée non protégée au gel"
            Else
                uneInfoGelQ1 = " ============> Chaussée protégée au gel"
            End If
            If monEtude.monIndiceGelAdmQ1 = HorsGel Then
                unePrintZone.Print "Chaussée Hors Gel"
            Else
                unePrintZone.Print Format(monEtude.monIndiceGelAdmQ1) + " °C.j" + uneInfoGelQ1
            End If
        End If
        
        'Affichage des résultats gel pour la qualité Q2
        'en version 2, on n'affiche qu'une seule carotte
        If monEtude.monEpQ2Trouv And monEtude.monTypeChantier = TypeChantierQ2 Then
            unePrintZone.CurrentX = maMarge * 1.1
            unePrintZone.Font.Bold = True
            unePrintZone.Print "En condition de chantier standard (qualité Q2) : "
            unePrintZone.CurrentX = maMarge * 2
            unCurX = unePrintZone.CurrentX
            unCurY = unePrintZone.CurrentY
            unePrintZone.Print "Indice de Gel de Référence corrigé : "
            unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Indice de Gel de Référence corrigé : ")
            unePrintZone.CurrentY = unCurY
            unePrintZone.Font.Bold = False
            unePrintZone.Print Format(monEtude.monIndiceGelRefQ2) + " °C.j"
            
            unePrintZone.Font.Bold = True
            unePrintZone.CurrentX = maMarge * 2
            unCurX = unePrintZone.CurrentX
            unCurY = unePrintZone.CurrentY
            unePrintZone.Print "Indice de Gel Admissible : "
            unePrintZone.CurrentX = unCurX + unePrintZone.TextWidth("Indice de Gel Admissible : ")
            unePrintZone.CurrentY = unCurY
            unePrintZone.Font.Bold = False
            If monEtude.monIndiceGelRefQ2 > monEtude.monIndiceGelAdmQ2 Then
                uneInfoGelQ2 = " ============> Chaussée non protégée au gel"
            Else
                uneInfoGelQ2 = " ============> Chaussée protégée au gel"
            End If
            If monEtude.monIndiceGelAdmQ2 = HorsGel Then
                unePrintZone.Print "Chaussée Hors Gel"
            Else
                unePrintZone.Print Format(monEtude.monIndiceGelAdmQ2) + " °C.j" + uneInfoGelQ2
            End If
        End If
        
        'Dessin d'un cadre autour de la partie titre
        unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE") / 2
        unePrintZone.CurrentX = maMarge
        unePrintZone.CurrentY = unCurY
        unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY), QBColor(0)
        unePrintZone.Line -(unePrintZone.ScaleWidth - maMarge, unCurY0), QBColor(0)
        unePrintZone.Line -(maMarge, unCurY0), QBColor(0)
        unePrintZone.Line -(maMarge, unCurY), QBColor(0)
    End If
End Sub
    


Private Sub ImprimerCarottes()
    Dim unCurY As Single, unCurY0 As Single
    Dim unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single
    Dim unDec As Single, i As Byte
    
    'Pour espacer
    unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("TITRE")
    
    If unePrintZone.CurrentY >= unePrintZone.ScaleHeight And (unePrintZone Is Printer) Then
        'On passe à la page suivante
        unePrintZone.NewPage
        ImprimerEntête
        unCurY = maMarge
    End If
    
    If unCurY + 5670 >= (unePrintZone.ScaleHeight - maMarge) And (monEtude.monEpQ1Trouv Or monEtude.monEpQ2Trouv) And (unePrintZone Is Printer) Then
        'On passe à la page suivante s'il reste moins de
        '5670 twips = 10 cm et qu'il n'y pas d'épaisseur
        'trouvée ni pour Q1 et ni pour Q2
        unePrintZone.NewPage
        ImprimerEntête
        unCurY = maMarge
    End If
    
    unePrintZone.CurrentX = maMarge
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Bold = False
    
    'Stockage du dernier current y
    unCurY0 = unCurY
    
    'Dessin des carottes Q1 et Q2 éventuelles
    'For i = 1 To 2
    'En version 2, on n'imprime qu'une seule carotte celle définie
    'par les conditions de chantier
    'If i = 1 Then
    If monEtude.monTypeChantier = TypeChantierQ1 Then
        'Cas de la carotte Q1 à imprimer
        unEpQTrouv = monEtude.monEpQ1Trouv
        unMinTecQ = monEtude.monMinTecQ1
        unMaxPraQ = monEtude.monMaxPraQ1
        'Décalage de la carotte Q1 ==> aucun, justifiée à droite
        unDec = 0
        'en version 2, on met manuel la qualité dans i
        i = 1
    Else
        'Cas de la carotte Q2 à  imprimer
        unEpQTrouv = monEtude.monEpQ2Trouv
        unMinTecQ = monEtude.monMinTecQ2
        unMaxPraQ = monEtude.monMaxPraQ2
        'Décalage de la carotte Q2 ==> aucun, mise à gauche à coté de Q1
        unDec = 9 * 567 'décalage de 9 cm + la marge
        'Modification Olivier FOREL 21/07/2005 : Suppression du décalage pour Q2
        unDec = 0
        'fin Modification Olivier FOREL
        'en version 2, on met manuel la qualité dans i
        i = 2
    End If
    
    unCurY = unCurY0
    
    If unEpQTrouv Then
        'Cas d'épaisseurs trouvées de qualité Q
        '==> Dessin de la carotte Q
        PrintCarotQ i, unCurY
        
        unePrintZone.CurrentX = maMarge + unDec
        unePrintZone.CurrentY = unCurY + unePrintZone.TextHeight("TITRE") / 2
        'unCurY a été mis à jour par PrintCarotQ
        
        If UCase(unMinTecQ) = "OUI" Then
            'Affichage d'un message signalant le minimum techno atteint
            unePrintZone.Font.Bold = True
            unePosEspMilieu = InStr(Len(MsgMinTechno1) / 2, MsgMinTechno1, " ")
            unePrintZone.Print Mid(MsgMinTechno1, 1, unePosEspMilieu)
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print Mid(MsgMinTechno1, unePosEspMilieu + 1)
            unePosEspMilieu = InStr(Len(MsgMinTechno2) / 2, MsgMinTechno2, " ")
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print Mid(MsgMinTechno2, 1, unePosEspMilieu)
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print Mid(MsgMinTechno2, unePosEspMilieu + 1)
        End If
    Else
        If UCase(unMaxPraQ) = "OUI" Then
            'Cas où maximun pratique atteint pour la qualité Q
            '==> Affichage du message l'indiquant
            unePrintZone.Font.Size = 10
            unePrintZone.Font.Bold = True
            unePosEspMilieu = InStr(Len(MsgMaxPra1) / 2, MsgMaxPra1, " ")
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.CurrentY = unCurY + unePrintZone.TextHeight("TITRE") / 2
            unePrintZone.Print Mid(MsgMaxPra1, 1, unePosEspMilieu)
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print Mid(MsgMaxPra1, unePosEspMilieu + 1) + "Q" + Format(i)
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print Trim(MsgMaxPra2)
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print MsgMaxPra3
            If i = 2 Then
                unePosEspMilieu = InStr(Len(MsgMaxPra4) / 2, MsgMaxPra4, " ")
                unePrintZone.CurrentX = maMarge + unDec
                unePrintZone.Print Mid(MsgMaxPra4, 1, unePosEspMilieu)
                unePrintZone.CurrentX = maMarge + unDec
                unePrintZone.Print "     " + Mid(MsgMaxPra4, unePosEspMilieu + 1)
            End If
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print MsgMaxPra5
            unePrintZone.CurrentX = maMarge + unDec
            unePrintZone.Print MsgMaxPra6
        End If
    End If
    'Next i
End Sub

Private Sub ImprimerCarottesInFile(unFileId As Integer)
    Dim unCurY As Single, unCurY0 As Single
    Dim unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single
    Dim unDec As Single, i As Byte
    
    'Dessin des carottes Q1 et Q2 éventuelles
    Print #unFileId, "DESCRIPTION DES CAROTTES : "
    Print #unFileId, "---------------------------"
    Print #unFileId,
    
    For i = 1 To 2
        If i = 1 Then
            'Cas de la carotte Q1
            unEpQTrouv = monEtude.monEpQ1Trouv
            unMinTecQ = monEtude.monMinTecQ1
            unMaxPraQ = monEtude.monMaxPraQ1
        Else
            'Cas de la carotte Q2
            unEpQTrouv = monEtude.monEpQ2Trouv
            unMinTecQ = monEtude.monMinTecQ2
            unMaxPraQ = monEtude.monMaxPraQ2
        End If
        If unEpQTrouv Then
            'Cas d'épaisseurs trouvées de qualité Q
            '==> Affichage de la carotte Q
            PrintCarotQInFile unFileId, i
            
            Print #unFileId,
            
            If UCase(unMinTecQ) = "OUI" Then
                'Affichage d'un message signalant le minimum techno atteint
                Print #unFileId, MsgMinTechno1
                Print #unFileId, MsgMinTechno2
                Print #unFileId,
            End If
        Else
            If UCase(unMaxPraQ) = "OUI" Then
                'Cas où maximun pratique atteint pour la qualité Q
                '==> Affichage du message l'indiquant
                Print #unFileId, MsgMaxPra1 + "Q" + Format(i)
                Print #unFileId, MsgMaxPra2
                Print #unFileId, MsgMaxPra3
                If i = 2 Then
                    Print #unFileId, MsgMaxPra4
                End If
                Print #unFileId, MsgMaxPra5
                Print #unFileId, MsgMaxPra6
                Print #unFileId,
            End If
        End If
    Next i
End Sub

Private Sub UtiliserFonteArial()
    'Prend la fonte Arial ou la première fonte
    'de l'imprimante par défaut pour imprimer
    Printer.Font.Name = Printer.Fonts(0)
    For i = 0 To Printer.FontCount - 1
        If LCase(Printer.Fonts(i)) = "arial" Then
            Printer.Font.Name = Printer.Fonts(i)
            Exit For
        End If
    Next i
End Sub
    
Private Sub ImprimerInfoMatInFile(unFileId As Integer)
    Dim unNumErr As Byte
    Dim uneStruct As Structure, uneColMatBF As Collection
    Dim uneColMat As New Collection, uneColComment As New Collection
    
    Set uneStruct = DonnerStructChoisie(monEtude)
    'Récup de la bonne collection de matériaux base/fondation
    If monEtude.CheckFichPerso = 1 Then
        Set uneColMatBF = maColMatBFPerso
    Else
        Set uneColMatBF = maColMatBFCERTU
    End If
    
    'Gestion d'erreur
    On Error GoTo ErreurPrintInfoMat0
    
    'Recherche des matériaux présents dans les carottes Q1/Q2
    If uneStruct.maCoucheSurface <> "Aucune" Then
        If TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé Then
            If monEtude.monMSComp1Q1 <> "" And monEtude.monTypeChantier = TypeChantierQ1 Then
                uneColMat.Add monEtude.monMSComp1Q1, monEtude.monMSComp1Q1
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp1Q1).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp2Q1 <> "" And monEtude.monTypeChantier = TypeChantierQ1 Then
                uneColMat.Add monEtude.monMSComp2Q1, monEtude.monMSComp2Q1
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp2Q1).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp1Q2 <> "" And monEtude.monTypeChantier = TypeChantierQ2 Then
                uneColMat.Add monEtude.monMSComp1Q2, monEtude.monMSComp1Q2
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp1Q2).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp2Q2 <> "" And monEtude.monTypeChantier = TypeChantierQ2 Then
                uneColMat.Add monEtude.monMSComp2Q2, monEtude.monMSComp2Q2
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp2Q2).monCommentaire
                uneColComment.Add RichTextTmp.Text
           End If
        Else
            uneColMat.Add uneStruct.maCoucheSurface, uneStruct.maCoucheSurface
            RichTextTmp.TextRTF = maColMatSurf(uneStruct.maCoucheSurface).monCommentaire
            uneColComment.Add RichTextTmp.Text
        End If
    End If
    
    If uneStruct.maCoucheBase <> "Aucune" Then
        uneColMat.Add uneStruct.maCoucheBase, uneStruct.maCoucheBase
        RichTextTmp.TextRTF = uneColMatBF(uneStruct.maCoucheBase).monCommentaire
        uneColComment.Add RichTextTmp.Text
    End If
    
    If uneStruct.maCoucheFondation <> "Aucune" Then
        uneColMat.Add uneStruct.maCoucheFondation, uneStruct.maCoucheFondation
        RichTextTmp.TextRTF = uneColMatBF(uneStruct.maCoucheFondation).monCommentaire
        uneColComment.Add RichTextTmp.Text
    End If
    
    'Impression des commentaires des matériaux présents dans les carottes Q1/Q2
    For i = 1 To uneColMat.Count
        If InStr(1, uneColMat(i), "Erreur#") = 0 Then
            Print #unFileId, "Commentaire du matériau : " + uneColMat(i)
            Print #unFileId, "-------------------------"
            Print #unFileId,
            
            'Découpage en lignes de 50 caractères (au blanc le plus proche)
            'pour afficher 50 cractères maxi
            uneString = uneColComment(i)
            Do
                j = 0
                If Len(uneString) <= 50 Then
                    unePosCR = InStr(1, uneString, Chr(13))
                    If unePosCR > 0 Then
                        j = unePosCR + 2
                        uneLgTxt = unePosCR - 1
                    Else
                        j = Len(uneString) + 1
                        uneLgTxt = Len(uneString)
                    End If
                Else
                    unePosEsp = InStr(50, uneString, " ")
                    unePosCR = InStr(1, uneString, Chr(13))
                    If unePosEsp > 0 And unePosCR > 0 Then
                        If unePosEsp > unePosCR Then
                            j = unePosCR + 2
                            uneLgTxt = unePosCR - 1
                        Else
                            j = unePosEsp + 1
                            uneLgTxt = unePosEsp - 1
                        End If
                    ElseIf unePosEsp = 0 And unePosCR > 0 Then
                        j = unePosCR + 2
                        uneLgTxt = unePosCR - 1
                    ElseIf unePosEsp > 0 And unePosCR = 0 Then
                        j = unePosEsp + 1
                        uneLgTxt = unePosEsp - 1
                    Else
                        j = 52
                        uneLgTxt = 51
                    End If
                End If
                
                If j > 0 Then
                    Print #unFileId, Mid(uneString, 1, uneLgTxt)
                    uneString = Mid(uneString, j)
                End If
            Loop Until uneString = ""
            Print #unFileId,
        End If
    Next i
    
    'Vidange des collections
    ViderCollection uneColMat
    ViderCollection uneColComment
    
    'Sortie pour éviter la gestion d'erreur
    On Error GoTo 0
    Exit Sub
    
ErreurPrintInfoMat0:
    'Gestion d'erreur
    If Err.Number = 457 Then
        'Le matériau composant est déjà présent dans la
        'collection des abrégés de matériau (exemple même composition en Q1 et Q2)
        unNumErr = unNumErr + 1
        uneColMat.Add "Erreur#" + Format(unNumErr), "Erreur#" + Format(unNumErr)
        Resume Next
    End If
    'Vidange des collections
    ViderCollection uneColMat
    ViderCollection uneColComment
End Sub

Private Sub ImprimerInfoMat()
    Dim unCurX As Single, unCurY As Single, unNumErr As Byte
    Dim uneStruct As Structure, uneColMatBF As Collection
    Dim uneColMat As New Collection, uneColComment As New Collection
    
    unCurX0 = maMarge
    unCurY0 = maMarge
    Set uneStruct = DonnerStructChoisie(monEtude)
    'Récup de la bonne collection de matériaux base/fondation
    If monEtude.CheckFichPerso = 1 Then
        Set uneColMatBF = maColMatBFPerso
    Else
        Set uneColMatBF = maColMatBFCERTU
    End If
    
    'Gestion d'erreur
    On Error GoTo ErreurPrintInfoMat
    
    'Recherche des matériaux présents dans les carottes Q1/Q2
    If uneStruct.maCoucheSurface <> "Aucune" Then
        If TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé Then
            If monEtude.monMSComp1Q1 <> "" And monEtude.monTypeChantier = TypeChantierQ1 Then
                uneColMat.Add monEtude.monMSComp1Q1, monEtude.monMSComp1Q1
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp1Q1).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp2Q1 <> "" And monEtude.monTypeChantier = TypeChantierQ1 Then
                uneColMat.Add monEtude.monMSComp2Q1, monEtude.monMSComp2Q1
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp2Q1).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp1Q2 <> "" And monEtude.monTypeChantier = TypeChantierQ2 Then
                uneColMat.Add monEtude.monMSComp1Q2, monEtude.monMSComp1Q2
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp1Q2).monCommentaire
                uneColComment.Add RichTextTmp.Text
            End If
            If monEtude.monMSComp2Q2 <> "" And monEtude.monTypeChantier = TypeChantierQ2 Then
                uneColMat.Add monEtude.monMSComp2Q2, monEtude.monMSComp2Q2
                RichTextTmp.TextRTF = maColMatSurf(monEtude.monMSComp2Q2).monCommentaire
                uneColComment.Add RichTextTmp.Text
           End If
        Else
            uneColMat.Add uneStruct.maCoucheSurface, uneStruct.maCoucheSurface
            RichTextTmp.TextRTF = maColMatSurf(uneStruct.maCoucheSurface).monCommentaire
            uneColComment.Add RichTextTmp.Text
        End If
    End If
    
    If uneStruct.maCoucheBase <> "Aucune" Then
        uneColMat.Add uneStruct.maCoucheBase, uneStruct.maCoucheBase
        RichTextTmp.TextRTF = uneColMatBF(uneStruct.maCoucheBase).monCommentaire
        uneColComment.Add RichTextTmp.Text
    End If
    
    If uneStruct.maCoucheFondation <> "Aucune" Then
        uneColMat.Add uneStruct.maCoucheFondation, uneStruct.maCoucheFondation
        RichTextTmp.TextRTF = uneColMatBF(uneStruct.maCoucheFondation).monCommentaire
        uneColComment.Add RichTextTmp.Text
    End If
    
    'Impression des commentaires des matériaux présents dans les carottes Q1/Q2
    unePrintZone.CurrentY = maMarge
    For i = 1 To uneColMat.Count
        If InStr(1, uneColMat(i), "Erreur#") = 0 Then
            unePrintZone.CurrentX = maMarge * 1.1
            unePrintZone.CurrentY = unePrintZone.CurrentY + maMarge * 0.1
            unePrintZone.Font.Size = 10
            unePrintZone.Font.Bold = True
            unePrintZone.Print "Commentaire du matériau : " + uneColMat(i)
            unePrintZone.Font.Bold = False
            
            'Découpage en lignes de 80 caractères (au blanc le plus proche)
            'pour afficher 80 cractères maxi
            uneString = uneColComment(i)
            Do
                j = 0
                If Len(uneString) <= 80 Then
                    unePosCR = InStr(1, uneString, Chr(13))
                    If unePosCR > 0 Then
                        j = unePosCR + 2
                        uneLgTxt = unePosCR - 1
                    Else
                        j = Len(uneString) + 1
                        uneLgTxt = Len(uneString)
                    End If
                Else
                    unePosEsp = InStr(80, uneString, " ")
                    unePosCR = InStr(1, uneString, Chr(13))
                    If unePosEsp > 0 And unePosCR > 0 Then
                        If unePosEsp > unePosCR Then
                            j = unePosCR + 2
                            uneLgTxt = unePosCR - 1
                        Else
                            j = unePosEsp + 1
                            uneLgTxt = unePosEsp - 1
                        End If
                    ElseIf unePosEsp = 0 And unePosCR > 0 Then
                        j = unePosCR + 2
                        uneLgTxt = unePosCR - 1
                    ElseIf unePosEsp > 0 And unePosCR = 0 Then
                        j = unePosEsp + 1
                        uneLgTxt = unePosEsp - 1
                    Else
                        j = 82
                        uneLgTxt = 81
                    End If
                End If
                
                If j > 0 Then
                    unePrintZone.CurrentX = maMarge * 1.1
                    unePrintZone.Print Mid(uneString, 1, uneLgTxt)
                    uneString = Mid(uneString, j)
                End If
            Loop Until uneString = ""
            
            unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
            'Pour sauter une demi-ligne pour espacer avec le titre en gras
        End If
    Next i
    
    'Dessin d'un cadre autour
    unCurX = unePrintZone.ScaleWidth - maMarge
    unCurY = unePrintZone.CurrentY
    unePrintZone.Line (unCurX0, unCurY0)-(unCurX, unCurY), QBColor(0), B
    
    'Vidange des collections
    ViderCollection uneColMat
    ViderCollection uneColComment
    
    'Sortie pour éviter la gestion d'erreur
    On Error GoTo 0
    Exit Sub
    
ErreurPrintInfoMat:
    'Gestion d'erreur
    If Err.Number = 457 Then
        'Le matériau composant est déjà présent dans la
        'collection des abrégés de matériau (exemple même composition en Q1 et Q2)
        unNumErr = unNumErr + 1
        uneColMat.Add "Erreur#" + Format(unNumErr), "Erreur#" + Format(unNumErr)
        Resume Next
    End If
    'Vidange des collections
    ViderCollection uneColMat
    ViderCollection uneColComment
End Sub
    
Private Sub ImprimerInfoStructInFile(unFileId As Integer)
    'Impression du commentaire de la structure choisie
    Dim uneStruct As Structure
        
    Set uneStruct = DonnerStructChoisie(monEtude)
    
    Print #unFileId, "Commentaire de la structure : " + uneStruct.monAbrégé
    Print #unFileId, "-----------------------------"
    Print #unFileId,
    
    'Découpage en lignes de 50 caractères (au blanc le plus proche)
    'pour afficher 50 caractères maxi
    RichTextTmp.TextRTF = uneStruct.monComment
    uneString = RichTextTmp.Text
    Do
        j = 0
        If Len(uneString) <= 50 Then
            unePosCR = InStr(1, uneString, Chr(13))
            If unePosCR > 0 Then
                j = unePosCR + 2
                uneLgTxt = unePosCR - 1
            Else
                j = Len(uneString) + 1
                uneLgTxt = Len(uneString)
            End If
        Else
            unePosEsp = InStr(50, uneString, " ")
            unePosCR = InStr(1, uneString, Chr(13))
            If unePosEsp > 0 And unePosCR > 0 Then
                If unePosEsp > unePosCR Then
                    j = unePosCR + 2
                    uneLgTxt = unePosCR - 1
                Else
                    j = unePosEsp + 1
                    uneLgTxt = unePosEsp - 1
                End If
            ElseIf unePosEsp = 0 And unePosCR > 0 Then
                j = unePosCR + 2
                uneLgTxt = unePosCR - 1
            ElseIf unePosEsp > 0 And unePosCR = 0 Then
                j = unePosEsp + 1
                uneLgTxt = unePosEsp - 1
            Else
                j = 52
                uneLgTxt = 51
            End If
        End If
        
        If j > 0 Then
            Print #unFileId, Mid(uneString, 1, uneLgTxt)
            uneString = Mid(uneString, j)
        End If
    Loop Until uneString = ""
    Print #unFileId,
End Sub

Private Sub ImprimerInfoStruct()
    'Impression du commentaire de la structure choisie
    Dim uneStruct As Structure
    Dim unCurX As Single, unCurY As Single
    
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec les info précédentes
    
    Set uneStruct = DonnerStructChoisie(monEtude)
    unCurX0 = maMarge
    unCurY0 = unePrintZone.CurrentY
   
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.CurrentY = unePrintZone.CurrentY + maMarge * 0.1
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Bold = True
    unePrintZone.Print "Commentaire de la structure : " + uneStruct.monAbrégé
    
    unePrintZone.Font.Bold = False
        
    'Découpage en lignes de 80 caractères (au blanc le plus proche)
    'pour afficher 80 cractères maxi
    RichTextTmp.TextRTF = uneStruct.monComment
    uneString = RichTextTmp.Text
    Do
        j = 0
        If Len(uneString) <= 80 Then
            unePosCR = InStr(1, uneString, Chr(13))
            If unePosCR > 0 Then
                j = unePosCR + 2
                uneLgTxt = unePosCR - 1
            Else
                j = Len(uneString) + 1
                uneLgTxt = Len(uneString)
            End If
        Else
            unePosEsp = InStr(80, uneString, " ")
            unePosCR = InStr(1, uneString, Chr(13))
            If unePosEsp > 0 And unePosCR > 0 Then
                If unePosEsp > unePosCR Then
                    j = unePosCR + 2
                    uneLgTxt = unePosCR - 1
                Else
                    j = unePosEsp + 1
                    uneLgTxt = unePosEsp - 1
                End If
            ElseIf unePosEsp = 0 And unePosCR > 0 Then
                j = unePosCR + 2
                uneLgTxt = unePosCR - 1
            ElseIf unePosEsp > 0 And unePosCR = 0 Then
                j = unePosEsp + 1
                uneLgTxt = unePosEsp - 1
            Else
                j = 82
                uneLgTxt = 81
            End If
        End If
        
        If j > 0 Then
            unePrintZone.CurrentX = maMarge * 1.1
            unePrintZone.Print Mid(uneString, 1, uneLgTxt)
            uneString = Mid(uneString, j)
        End If
    Loop Until uneString = ""
    
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec le titre en gras
    
    'Dessin d'un cadre autour
    unCurX = unePrintZone.ScaleWidth - maMarge
    unCurY = unePrintZone.CurrentY
    unePrintZone.Line (unCurX0, unCurY0)-(unCurX, unCurY), QBColor(0), B
End Sub
    
Private Sub ImprimerInfoGelInFile(unFileId As Integer)
    'Impression des données complémentaires du gel
    Dim uneStruct As Structure, uneString As String
    Dim unCoefAgglo As String
    
    Set uneStruct = DonnerStructChoisie(monEtude)
        
    If monEtude.monUtilIndGelPerso = 0 Then
        'Cas d'un indice de gel trouvé à partir des stations de références
        Print #unFileId, "Informations complémentaires sur l'indice de gel"
        Print #unFileId, "------------------------------------------------"
        Print #unFileId, "Station météo de référence : "; monEtude.ComboStation.Text
         
        If DonnerTypeHiver(monEtude) = HE Then
            uneString = "Hiver Exceptionnel"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHRE
        ElseIf DonnerTypeHiver(monEtude) = HRNE Then
            uneString = "Hiver Rigoureux Non Exceptionnel"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHRNE
        Else
            uneString = "Hiver Courant"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHC
        End If
        
        Print #unFileId, "Type d'hiver : "; uneString
        Print #unFileId, "Indice de Gel brut : "; Format(unIndGelRef) + " °C.j"
            
        'Récup des données de l'agglo du projet
        TrouverCaractèreDécimalUtilisé
        If monEtude.ComboTailleAgglo.ListIndex = 0 Then
            unCoefAgglo = "1 (< à 100 000 Habitants)"
        ElseIf monEtude.ComboTailleAgglo.ListIndex = 1 Then
            unCoefAgglo = "0" + monCarDeci + "9 (entre 100 000 et 1 000 000 Habitants)"
        ElseIf monEtude.ComboTailleAgglo.ListIndex = 2 Then
            unCoefAgglo = "0" + monCarDeci + "8 (> 1 000 000 Habitants)"
        End If
        
        Print #unFileId, "Correction taille d'agglomération : "; unCoefAgglo
    Else
        'Cas d'un indice de gel personnel
        Print #unFileId, "Information complémentaire sur l'indice de gel"
        Print #unFileId, "----------------------------------------------"
        Print #unFileId, "Indice de Gel personnel : "; Format(monEtude.monIndGelPerso) + " °C.j"
    End If
    
    Print #unFileId,
    Print #unFileId, "Sol support"
    Print #unFileId, "-----------"
    
    If DonnerTypeGelSol(monEtude) = TresGelif Then
        uneString = "Très Gélif"
    ElseIf DonnerTypeGelSol(monEtude) = PeuGelif Then
        uneString = "Peu Gélif"
    Else
        uneString = "Non Gélif"
    End If
    
    Print #unFileId, "Gélivité : "; uneString
    Print #unFileId, "Pente de la courbe de gonflement : "; monEtude.TextPente.Text
    If monEtude.TextPente.Text = MsgInfinie Then
        uneString = Format(CalculerQg(monEtude))
    ElseIf CSng(monEtude.TextPente.Text) <= 0.05 Then
        uneString = "Infinie"
    Else
        uneString = Format(CalculerQg(monEtude))
    End If
    Print #unFileId, "Quantité de gel admis par le sol support : "; uneString
    
    Print #unFileId,
    Print #unFileId, "Plateforme"
    Print #unFileId, "----------"
    
    Print #unFileId, "Epaisseur : "; monEtude.TextEpaisseur.Text + " cm"
    If monEtude.OptionANT.Value Then
        uneString = "Non Traitée"
    Else
        uneString = "Traitée"
    End If
    Print #unFileId, "Couche de forme : "; uneString
    Print #unFileId, "Quantité de gel admis par la partie non gélive de la plateforme : "; Format(CalculerQng(monEtude))

    Print #unFileId,
    Print #unFileId, "Apport mécanique de la chaussée"
    Print #unFileId, "-------------------------------"
    
    If monEtude.monEpQ1Trouv And monEtude.monTypeChantier = TypeChantierQ1 Then
        uneString = monEtude.monQmQ1
        If monEtude.monTypeEtude = TypeEtudeGiratoire Then
            uneString = "En condition de chantier (" + monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption + ") : " + uneString
        Else
            uneString = "En condition de chantier standard (qualité Q1) : " + uneString
        End If
    ElseIf monEtude.monEpQ2Trouv And monEtude.monTypeChantier = TypeChantierQ2 Then
        uneString = monEtude.monQmQ2
        uneString = "En condition de chantier difficile (qualité Q2) : " + uneString
    Else
        uneString = "Structure invalide"
        uneString = "Erreur : " + uneString
    End If
    
    Print #unFileId, uneString
End Sub

Private Sub ImprimerInfoGel()
    'Impression des données complémentaires du gel
    Dim uneStruct As Structure, uneString As String
    Dim unCoefAgglo As String
    Dim unCurX As Single, unCurY As Single
    
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec les info précédentes
    
    Set uneStruct = DonnerStructChoisie(monEtude)
    unCurX0 = maMarge
    unCurY0 = unePrintZone.CurrentY
      
    unePrintZone.CurrentY = unePrintZone.CurrentY + maMarge * 0.1
    If monEtude.monUtilIndGelPerso = 0 Then
        'Cas d'un indice de gel trouvé à partir des stations de références
        PrintString "Station météo de référence : ", monEtude.ComboStation.Text
         
        If DonnerTypeHiver(monEtude) = HE Then
            uneString = "Hiver Exceptionnel"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHRE
        ElseIf DonnerTypeHiver(monEtude) = HRNE Then
            uneString = "Hiver Rigoureux Non Exceptionnel"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHRNE
        Else
            uneString = "Hiver Courant"
            unIndGelRef = monTabStation(monEtude.ComboStation.ListIndex + 1).monHC
        End If
        
        PrintString "Type d'hiver : ", uneString
        PrintString "Indice de Gel brut : ", Format(unIndGelRef) + " °C.j"
            
        'Récup des données de l'agglo du projet
        TrouverCaractèreDécimalUtilisé
        If monEtude.ComboTailleAgglo.ListIndex = 0 Then
            unCoefAgglo = "1 (< à 100 000 Habitants)"
        ElseIf monEtude.ComboTailleAgglo.ListIndex = 1 Then
            unCoefAgglo = "0" + monCarDeci + "9 (entre 100 000 et 1 000 000 Habitants)"
        ElseIf monEtude.ComboTailleAgglo.ListIndex = 2 Then
            unCoefAgglo = "0" + monCarDeci + "8 (> 1 000 000 Habitants)"
        End If
        
        PrintString "Correction taille d'agglomération : ", unCoefAgglo
    Else
        'Cas d'un indice de gel personnel
        PrintString "Indice de Gel personnel : ", Format(monEtude.monIndGelPerso) + " °C.j"
    End If
    
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec les info précédentes
    
    unePrintZone.Font.Underline = True
    PrintString "Sol support", ""
    unePrintZone.Font.Underline = False
    
    If DonnerTypeGelSol(monEtude) = TresGelif Then
        uneString = "Très Gélif"
    ElseIf DonnerTypeGelSol(monEtude) = PeuGelif Then
        uneString = "Peu Gélif"
    Else
        uneString = "Non Gélif"
    End If
    
    PrintString "Gélivité : ", uneString
    PrintString "Pente de la courbe de gonflement : ", monEtude.TextPente.Text
    If monEtude.TextPente.Text = MsgInfinie Then
        uneString = Format(CalculerQg(monEtude))
    ElseIf CSng(monEtude.TextPente.Text) <= 0.05 Then
        uneString = "Infinie"
    Else
        uneString = Format(CalculerQg(monEtude))
    End If
    PrintString "Quantité de gel admis par le sol support : ", uneString
    
    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec les info précédentes
    
    unePrintZone.Font.Underline = True
    PrintString "Plateforme", ""
    unePrintZone.Font.Underline = False
    PrintString "Epaisseur : ", monEtude.TextEpaisseur.Text + " cm"
    If monEtude.OptionANT.Value Then
        uneString = "Non Traitée"
    Else
        uneString = "Traitée"
    End If
    PrintString "Couche de forme : ", uneString
    PrintString "Quantité de gel admis par la partie non gélive de la plateforme : ", Format(CalculerQng(monEtude))

    unePrintZone.CurrentY = unePrintZone.CurrentY + TextHeight("TITRE") / 2
    'Pour sauter une demi-ligne pour espacer avec les info précédentes
    
    unePrintZone.Font.Underline = True
    PrintString "Apport mécanique de la chaussée", ""
    unePrintZone.Font.Underline = False
    If monEtude.monEpQ1Trouv And monEtude.monTypeChantier = TypeChantierQ1 Then
        uneString = monEtude.monQmQ1
        If monEtude.monTypeEtude = TypeEtudeGiratoire Then
            PrintString "En condition de chantier (" + monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption + ") : ", uneString
        Else
            PrintString "En condition de chantier standard (qualité Q1) : ", uneString
        End If
    ElseIf monEtude.monEpQ2Trouv And monEtude.monTypeChantier = TypeChantierQ2 Then
        uneString = monEtude.monQmQ2
        PrintString "En condition de chantier difficile (qualité Q2) : ", uneString
    Else
        uneString = "Structure invalide"
        PrintString "Erreur : ", uneString
    End If

    'Dessin d'un cadre autour
    unCurX = unePrintZone.ScaleWidth - maMarge
    unCurY = unePrintZone.CurrentY + unePrintZone.TextHeight("Titre") / 2
    unePrintZone.Line (unCurX0, unCurY0)-(unCurX, unCurY), QBColor(0), B
End Sub


Private Sub ImprimerEntête()
    'Imprimer Entête de page avec struct-urb + numéro de version
    unePrintZone.Font.Size = 10
    unePrintZone.Font.Bold = True
    unePrintZone.Font.Italic = True
    unePrintZone.CurrentX = maMarge
    unePrintZone.CurrentY = maMarge / 2 - TextHeight("TITRE") / 2
    unePrintZone.Print App.Title + " version " + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
    unePrintZone.Font.Italic = False
End Sub



Private Sub PrintCarotQ(uneQ As Byte, unYini As Single)
    'Impression de la carotte de la qualité Q (1 ou 2)
    'Position de démarrage du dessin donnée par unYini
    'Au final unYini est mis à jour pour donner un Y à partir
    'Duquel on peut dessiner sans écraser la carotte de la qualité Q
    
    Dim uneStruct As Structure, uneColMatBF As Collection
    Dim unTabLargCol(1 To 4) As Single
    Dim unTabImpQ(0 To 12, 1 To 4) As String
    Dim unMSComp1 As String, unMSComp2 As String
    Dim unDecInd As Byte, unXini As Single
    Dim i As Byte, j As Byte, unTabEp As Variant
    Dim unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single
    
    Set uneStruct = DonnerStructChoisie(monEtude)
    unTabEp = monEtude.monTabEp
    
    'Récup de la bonne collection de matériaux base/fondation
    If monEtude.CheckFichPerso = 1 Then
        Set uneColMatBF = maColMatBFPerso
    Else
        Set uneColMatBF = maColMatBFCERTU
    End If
    
    If uneQ = 2 Then
        unDecInd = 6
        'unXini = maMarge + 9 * 567 'décalage de 9 cm + la marge
        unXini = maMarge 'en version 2, on n'imprime qu'une seule carotte
        'donc pas besoin de décaler la carotte Q2 sur la droite
        unMSComp1 = monEtude.monMSComp1Q2
        unMSComp2 = monEtude.monMSComp2Q2
    Else
        'Pour Q1
        unDecInd = 0
        unXini = maMarge
        unMSComp1 = monEtude.monMSComp1Q1
        unMSComp2 = monEtude.monMSComp2Q1
    End If
    
    'Remplissage des largeurs de colonnes en twips (1 cm = 567 twips)
    unTabLargCol(1) = 2 * 567 '1.5 * 567
    unTabLargCol(2) = 2.5 * 567
    unTabLargCol(3) = 1.8 * 567
    unTabLargCol(4) = 2.4 * 567 '2.9 * 567
    
    'Remplissage du contenu du tableau à imprimer
    unIndPF = DonnerIndicePF(monEtude)
    If unIndPF = 4 Then
        unStringIndPF = "2+"
    Else
        unStringIndPF = Format(unIndPF)
    End If
    If monEtude.monTypeEtude = TypeEtudeGiratoire Then
        unTabImpQ(0, 1) = "PF" + unStringIndPF
    Else
        unTabImpQ(0, 1) = "Q" + Format(uneQ) + " / PF" + unStringIndPF
    End If
    unTabImpQ(0, 2) = "Norme"
    unTabImpQ(0, 3) = "Classe"
    unTabImpQ(0, 4) = "Epaisseur"
    
    For i = 1 + unDecInd To 6 + unDecInd
        If i = 1 + unDecInd And unMSComp1 <> "" Then
            unTabImpQ(i, 1) = unMSComp1
            unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
            unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
        ElseIf i = 2 + unDecInd And unMSComp2 <> "" Then
            unTabImpQ(i, 1) = unMSComp2
            unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
            unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
        ElseIf i = 2 + unDecInd And (uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés") Then
            'Cas des couches de surfaces en dalles ou pavés
            'Il faut afficher le lit de pose en 2ème couche de surface
            unTabImpQ(i, 1) = "Lit de pose"
            unTabImpQ(i, 2) = ""
            unTabImpQ(i, 3) = ""
        ElseIf i < 3 + unDecInd Then
            unTabImpQ(i, 1) = uneStruct.maCoucheSurface
            If uneStruct.maCoucheSurface = "Aucune" Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            ElseIf TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            Else
                'Pour tous les autres cas on a une norme et une classe (=qualité)
                unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
                unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
            End If
        Else
            If i = 3 + unDecInd Or i = 4 + unDecInd Then
                unTabImpQ(i, 1) = uneStruct.maCoucheBase
            Else
                'i = 5 ou 6 pour Q1 et 11 ou 12 pour Q2
                unTabImpQ(i, 1) = uneStruct.maCoucheFondation
            End If
            If unTabImpQ(i, 1) = "Aucune" Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            Else
                unTabImpQ(i, 2) = uneColMatBF(unTabImpQ(i, 1)).maNorme
                unTabImpQ(i, 3) = uneColMatBF(unTabImpQ(i, 1)).maQualité
            End If
        End If
        
        'Recherche si on a une structure avec une couche de surface sans épaisseur
        uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
        If uneSurfSansEpaisseur And i = 1 + unDecInd Then
            unTabImpQ(i, 4) = ""
        'ElseIf (uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés") And i = 1 + unDecInd Then
        '    unTabImpQ(i, 4) = Format(unTabEp(i)) + " cm + lit de pose"
        Else
            unTabImpQ(i, 4) = Format(unTabEp(i)) + " cm"
        End If
    Next i
    
    'Diminution de la fonte des textes
    unePrintZone.Font.Size = 9
    
    'Impression du tableau avec carottes
    unX1 = unXini
    unY1 = unYini
    For i = 0 To 6
        If uneQ = 2 And i > 0 Then
            unDecInd = 6
        Else
            unDecInd = 0
        End If
        If unTabEp(i + unDecInd) > 0 Then
            For j = 1 To 4
                unX2 = unX1 + unTabLargCol(j)
                unY2 = unY1 + unTabEp(i + unDecInd) * monEch
                'Pour que Epaisseur imprimante = epaisseur réelle * monEch
                
                uneLgTxt = unePrintZone.TextWidth(unTabImpQ(i + unDecInd, j))
                uneHtTxt = unePrintZone.TextHeight(unTabImpQ(i + unDecInd, j))
                
                'Dessin et remplissage ou non des cellules
                If j = 1 And i > 0 Then
                    'Première colonne c'est la carotte donc cellule remplie
                    'et sans la ligne 0 des entêtes
                    unePrintZone.Font.Bold = True
                    unePrintZone.DrawStyle = vbInvisible
                    If i = 2 And (uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés") Then
                        'Cas des couches de surfaces en dalles ou pavés
                        'Il faut afficher le lit de pose en 2ème couche de surface
                        'Fond et couleur fixé en dur
                        unePrintZone.FillColor = monEtude.ShapeLitPoseQ1.FillColor
                        unePrintZone.FillStyle = monEtude.ShapeLitPoseQ1.FillStyle
                    Else
                        unePrintZone.FillColor = DonnerCouleurCouche(i)
                        unePrintZone.FillStyle = vbFSSolid
                    End If
                    'Dessin avec la couleur de couches autour du nom matériau
                    uneH = (unY2 - unY1 - uneHtTxt) / 2
                    uneL = (unX2 - unX1 - uneLgTxt) / 2
                    unePrintZone.Line (unX1, unY1)-(unX2, unY1 + uneH), QBColor(0), B
                    unePrintZone.Line (unX1, unY1)-(unX1 + uneL, unY2), QBColor(0), B
                    unePrintZone.Line (unX2 - uneL, unY1)-(unX2, unY2), QBColor(0), B
                    unePrintZone.Line (unX1, unY2 - uneH)-(unX2, unY2), QBColor(0), B
                    'Dessin du cadre de cette cellule
                    unePrintZone.DrawStyle = vbSolid
                    unePrintZone.FillStyle = vbFSTransparent
                    unePrintZone.Line (unX1, unY1)-(unX2, unY2), QBColor(0), B
                Else
                    'Autres colonnes ===> pas remplies
                    If i = 0 Then
                        unePrintZone.Font.Bold = True
                    Else
                        unePrintZone.Font.Bold = False
                    End If
                    unePrintZone.Line (unX1, unY1)-(unX2, unY2), QBColor(0), B
                End If
                
                'Affichage des textes centrés dans les cellules
                unePrintZone.CurrentX = unX1 + (unX2 - unX1 - unePrintZone.TextWidth(unTabImpQ(i + unDecInd, j))) / 2
                unePrintZone.CurrentY = unY1 + (unY2 - unY1 - unePrintZone.TextHeight(unTabImpQ(i + unDecInd, j))) / 2
                unePrintZone.Print unTabImpQ(i + unDecInd, j)
               
               'Stockage pour la colonne suivante
                unX1 = unX2
            Next j
            
            'Stockage du x1 et y1 pour les lignes suivantes
            unX1 = unXini
            unY1 = unY2
        End If
    Next i
    
    'Affichage en bas à droite dans la carotte de l'épaisseur totale
    unePrintZone.Font.Bold = True
    unePrintZone.CurrentY = unY2 - unePrintZone.TextHeight("Epaisseur")
    uneEpTot = monEtude.DonnerEpaisseurTotale(uneQ)
    unePrintZone.CurrentX = unX2 - unTabLargCol(4) + (unTabLargCol(4) - unePrintZone.TextWidth("Total = " + Format(uneEpTot) + " cm")) / 2
    unePrintZone.Print "Total = " + Format(uneEpTot) + " cm"
    
    'Mis à jour du Y pour ne pas écraser les dessins
    'A utiliser pour mettre à jour unePrintZone.CurrentY
    unYini = unY2
End Sub

Private Sub PrintCarotQInFile(unFileId As Integer, uneQ As Byte)
    'Impression dans un fichier de la carotte de la qualité Q (1 ou 2)
    
    Dim uneStruct As Structure, uneColMatBF As Collection
    Dim unTabImpQ(0 To 12, 1 To 4) As String
    Dim unMSComp1 As String, unMSComp2 As String
    Dim unDecInd As Byte, uneString As String
    Dim i As Byte, j As Byte, unTabEp As Variant
    
    Set uneStruct = DonnerStructChoisie(monEtude)
    unTabEp = monEtude.monTabEp
    
    'Récup de la bonne collection de matériaux base/fondation
    If monEtude.CheckFichPerso = 1 Then
        Set uneColMatBF = maColMatBFPerso
    Else
        Set uneColMatBF = maColMatBFCERTU
    End If
    
    If uneQ = 2 Then
        unDecInd = 6
        unMSComp1 = monEtude.monMSComp1Q2
        unMSComp2 = monEtude.monMSComp2Q2
    Else
        'Pour Q1
        unDecInd = 0
        unMSComp1 = monEtude.monMSComp1Q1
        unMSComp2 = monEtude.monMSComp2Q1
    End If
        
    'Remplissage du contenu du tableau à imprimer
    unIndPF = DonnerIndicePF(monEtude)
    If unIndPF = 4 Then
        unStringIndPF = "2+"
    Else
        unStringIndPF = Format(unIndPF)
    End If
    unTabImpQ(0, 1) = "Q" + Format(uneQ) + " / PF" + unStringIndPF
    unTabImpQ(0, 2) = "Norme"
    unTabImpQ(0, 3) = "Classe"
    unTabImpQ(0, 4) = "Epaisseur"
    
    For i = 1 + unDecInd To 6 + unDecInd
        If i = 1 + unDecInd And unMSComp1 <> "" Then
            unTabImpQ(i, 1) = unMSComp1
            unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
            unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
        ElseIf i = 2 + unDecInd And unMSComp2 <> "" Then
            unTabImpQ(i, 1) = unMSComp2
            unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
            unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
        ElseIf i = 2 + unDecInd And (uneStruct.maCoucheSurface = "Dalles" Or uneStruct.maCoucheSurface = "Pavés") Then
            'Cas des couches de surfaces en dalles ou pavés
            'Il faut afficher le lit de pose en 2ème couche de surface
            unTabImpQ(i, 1) = "Lit de pose"
            unTabImpQ(i, 2) = ""
            unTabImpQ(i, 3) = ""
        ElseIf i < 3 + unDecInd Then
            unTabImpQ(i, 1) = uneStruct.maCoucheSurface
            If uneStruct.maCoucheSurface = "Aucune" Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            ElseIf TypeOf maColMatSurf(uneStruct.maCoucheSurface) Is MatComposé Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            Else
                'Pour tous les autres cas on a une norme et une classe (=qualité)
                unTabImpQ(i, 2) = maColMatSurf(unTabImpQ(i, 1)).maNorme
                unTabImpQ(i, 3) = maColMatSurf(unTabImpQ(i, 1)).maQualité
            End If
        Else
            If i = 3 + unDecInd Or i = 4 + unDecInd Then
                unTabImpQ(i, 1) = uneStruct.maCoucheBase
            Else
                'i = 5 ou 6 pour Q1 et 11 ou 12 pour Q2
                unTabImpQ(i, 1) = uneStruct.maCoucheFondation
            End If
            If unTabImpQ(i, 1) = "Aucune" Then
                unTabImpQ(i, 2) = ""
                unTabImpQ(i, 3) = ""
            Else
                unTabImpQ(i, 2) = uneColMatBF(unTabImpQ(i, 1)).maNorme
                unTabImpQ(i, 3) = uneColMatBF(unTabImpQ(i, 1)).maQualité
            End If
        End If
        
        'Recherche si on a une structure avec une couche de surface sans épaisseur
        uneSurfSansEpaisseur = (uneStruct.maCoucheSurface <> "Aucune" And uneStruct.maCoucheSurfSansEp = 1)
        If uneSurfSansEpaisseur And i = 1 + unDecInd Then
            unTabImpQ(i, 4) = ""
        Else
            unTabImpQ(i, 4) = Format(unTabEp(i)) + " cm"
        End If
    Next i
    
    'Impression du tableau avec carottes
    For i = 0 To 6
        If uneQ = 2 And i > 0 Then
            unDecInd = 6
        Else
            unDecInd = 0
        End If
        If unTabEp(i + unDecInd) > 0 Then
            uneString = ""
            For j = 1 To 3
                uneString = uneString + unTabImpQ(i + unDecInd, j) + Space(5 * Abs(j = 2) + 11 - Len(unTabImpQ(i + unDecInd, j))) + " | "
            Next j
            'Pour j = 4 les épaisseurs, on met un blanc de plus
            'si épaisseur a 4 caractères (moins de 2 chiffres + un blanc + cm)
            If Len(unTabImpQ(i + unDecInd, 4)) = 4 Then
                unTabImpQ(i + unDecInd, 4) = " " + unTabImpQ(i + unDecInd, 4)
            End If
            uneString = uneString + unTabImpQ(i + unDecInd, j) + Space(5 * Abs(j = 2) + 11 - Len(unTabImpQ(i + unDecInd, j))) + " | "
            'Affichage des textes
            Print #unFileId, uneString
        End If
        If i = 0 Then Print #unFileId, String(11, "-") + "-|-" + String(16, "-") + "-|-" + String(11, "-") + "-|-" + String(11, "-") + "-|"
    Next i
    
    'Affichage en bas à droite dans la carotte de l'épaisseur totale
    Print #unFileId,
    uneEpTot = monEtude.DonnerEpaisseurTotale(uneQ)
    Print #unFileId, Spc(28); "Epaisseur Totale = " + Format(uneEpTot) + " cm"
End Sub

Private Sub PrintString(unTitre As String, uneVal As String, Optional uneSizeFontBold As Byte = 10, Optional uneSizeFont As Byte = 10)
    Dim unCurY As Single
    
    unCurY = unePrintZone.CurrentY
    unePrintZone.CurrentX = maMarge * 1.1
    unePrintZone.Font.Size = uneSizeFontBold
    unePrintZone.Font.Bold = True
    unePrintZone.Print unTitre
    unePrintZone.CurrentX = maMarge * 1.1 + unePrintZone.TextWidth(unTitre)
    unePrintZone.CurrentY = unCurY
    unePrintZone.Font.Size = uneSizeFont
    unePrintZone.Font.Bold = False
    unePrintZone.Print uneVal
End Sub

Private Sub ImpressionInFile()
    'Impression dans un fichier texte en version 1
    Dim unPrintFile As String
    Dim unFileId As Integer
    Dim unePosRC As Integer
    Dim uneString As String
    Dim unTypePL As String
    Dim unNbTypePL As String, uneStruct As Structure
    Dim uneInfoGelQ1 As String, uneInfoGelQ2 As String
        
    unPrintFile = fMainForm.ChoisirFichier(MsgPrintInFile, MsgTxtFile, CurDir)
    If unPrintFile <> "" Then
        unFileId = FreeFile(0)
        Open unPrintFile For Output As #unFileId
                    
        'Impression générale
        'Imprimer Entête de page avec struct-urb + numéro de version
        uneString = App.Title + " version " + Format(App.Major) + "." + Format(App.Minor)
        Print #unFileId, uneString
        Print #unFileId, String(Len(uneString), "-")
        Print #unFileId,
    
        'Impression du titre de l'étude
        Print #unFileId, "TITRE DE L'ETUDE :"
        Print #unFileId, "------------------"
        Print #unFileId,
        uneString = monEtude.TextTitre
        unePosRC = InStr(1, uneString, Chr(13))
        Do While unePosRC > 0
            Print #unFileId, Spc(5); Mid(uneString, 1, unePosRC - 1)
            uneString = Mid(uneString, unePosRC + 2)
            '+2 pour repartir aprés le retour chariot et le saut de ligne
            unePosRC = InStr(1, uneString, Chr(13))
        Loop
        If unePosRC = 0 And uneString <> "" Then
            'Cas où plus de retour chariot et
            'affichage du reste du titre s'il en reste
            Print #unFileId, Spc(5); uneString
        End If
        Print #unFileId,
        Print #unFileId, "Date : "; monEtude.LabelDate
        Print #unFileId, "Variante : "; monEtude.TextVar
        If EstNouvelleEtude(monEtude) = False Then
            'Cas où ce n'est pas une étude nouvelle Etude N
            uneString = Mid(monEtude.Caption, 7)
        Else
            uneString = "Etude pas encore enregistrée"
        End If
        Print #unFileId, "Enregistrée sous : "; uneString
        Print #unFileId,
        
        'Impression des données de l'étude
        Print #unFileId, "DONNEES :"
        Print #unFileId, "---------"
        Print #unFileId,
        'Impression du type de voie
        Print #unFileId, "Type de voie : "; DonnerNomTypeVoie(monEtude)
        'Impression des données trafic ini, durée service
        'et croisance annuelle et de la classe de plateforme
        If DonnerTypeVoie(monEtude) = 4 Then
            unTypePL = "BUS"
        Else
            unTypePL = "Poids Lourds"
        End If
        Print #unFileId, "Trafic initial à la mise en service (par sens, par voie et par jour) : "; monEtude.TextTrafIni.Text + " " + unTypePL
        Print #unFileId, "Durée de service : "; monEtude.TextDuréeS.Text + " ans"
        Print #unFileId, "Taux de croissance : "; Format(DonnerCroissAn(monEtude)) + " % par an"
        unIndPF = DonnerIndicePF(monEtude)
        If unIndPF = 4 Then
            unStringIndPF = "2+"
        Else
            unStringIndPF = Format(unIndPF)
        End If
        Print #unFileId, "Plate-forme : "; "PF" + unStringIndPF
        Print #unFileId,
    
        'Impression des valeurs intermédiaires de l'étude
        Print #unFileId, "VALEURS INTERMEDIAIRES :"
        Print #unFileId, "------------------------"
        Print #unFileId,
        'Impression des données trafic cumulé, risque de calcul
        'si valeur présente, le CAM et le NE arrondi
        If DonnerTypeVoie(monEtude) = 4 Then
            unNbTypePL = "Nombre Cumulé de BUS : "
        Else
            unNbTypePL = "Nombre Cumulé de Poids Lourds : "
        End If
        Print #unFileId, unNbTypePL; monEtude.TextTrafCUM.Text
        'Si une structure est choisie
        'Affichage du risque de calcul si non nul
        Set uneStruct = DonnerStructChoisie(monEtude)
        If Not (uneStruct Is Nothing) Then
            If uneStruct.monTauxRisque <> 0 Then
                Print #unFileId, "Risque de calcul : "; Format(uneStruct.monTauxRisque) + " %"
            End If
        End If
        Print #unFileId, "CAM : "; monEtude.MaskCAM.Text
        Print #unFileId, "NE arrondi : "; Format(monEtude.monNEth, "##,###,###")
        Print #unFileId,

        'Impression des valeurs de vérif au gel de l'étude si vérif au gel voulue
        If mesOptionsGen.maVerifGel Then
            Print #unFileId, "GEL :"
            Print #unFileId, "-----"
            Print #unFileId,
            
            'Affichage des résultats gel pour la qualité Q1
            If monEtude.monEpQ1Trouv Then
                Print #unFileId, "Pour la qualité Q1 : "
                Print #unFileId, Spc(5); "Indice de Gel de Référence corrigé : "; Format(monEtude.monIndiceGelRefQ1) + " °C.j"
                If monEtude.monIndiceGelRefQ1 > monEtude.monIndiceGelAdmQ1 Then
                    uneInfoGelQ1 = " ============> Chaussée non protégée au gel"
                Else
                    uneInfoGelQ1 = ""
                End If
                If monEtude.monIndiceGelAdmQ1 = HorsGel Then
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; "Chaussée Hors Gel"
                Else
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; Format(monEtude.monIndiceGelAdmQ1) + " °C.j" + uneInfoGelQ1
                End If
            End If
            
            'Affichage des résultats gel pour la qualité Q2
            If monEtude.monEpQ2Trouv Then
                Print #unFileId, "Pour la qualité Q2 : "
                Print #unFileId, Spc(5); "Indice de Gel de Référence corrigé : "; Format(monEtude.monIndiceGelRefQ2) + " °C.j"
                If monEtude.monIndiceGelRefQ2 > monEtude.monIndiceGelAdmQ2 Then
                    uneInfoGelQ2 = " ============> Chaussée non protégée au gel"
                Else
                    uneInfoGelQ2 = ""
                End If
                If monEtude.monIndiceGelAdmQ2 = HorsGel Then
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; "Chaussée Hors Gel"
                Else
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; Format(monEtude.monIndiceGelAdmQ2) + " °C.j" + uneInfoGelQ2
                End If
            End If
            Print #unFileId,
        End If
        
        'Impression des carottes
        ImprimerCarottesInFile unFileId
        
        'Impressions complémentaires
        If CheckCommentMat.Value = 1 Then
            ImprimerInfoMatInFile unFileId
        End If
        If CheckCommentStruct.Value = 1 Then
            ImprimerInfoStructInFile unFileId
        End If
        If CheckInfoGel.Value = 1 Then
            ImprimerInfoGelInFile unFileId
        End If
        
        Close #unFileId
    End If
End Sub

Private Sub ImpressionInFileRTF()
    'Impression dans un fichier texte en version 1
    Dim unPrintFile As String
    Dim unFileId As Integer
    Dim unePosRC As Integer
    Dim uneString As String
    Dim unTypePL As String
    Dim unNbTypePL As String, uneStruct As Structure
    Dim uneInfoGelQ1 As String, uneInfoGelQ2 As String
        
    unPrintFile = fMainForm.ChoisirFichier(MsgPrintInFile, MsgRTFFile, CurDir)
    If unPrintFile <> "" Then
        unFileId = FreeFile(0)
        Open unPrintFile For Output As #unFileId
                    
        'Impression générale
        'Imprimer Entête de page avec struct-urb + numéro de version
        uneString = App.Title + " version " + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
        Print #unFileId, uneString
        Print #unFileId, String(Len(uneString), "-")
        Print #unFileId,
    
        'Impression du titre de l'étude
        Print #unFileId, "TITRE DE L'ETUDE :"
        Print #unFileId, "------------------"
        Print #unFileId,
        uneString = monEtude.TextTitre
        unePosRC = InStr(1, uneString, Chr(13))
        Do While unePosRC > 0
            Print #unFileId, Spc(5); Mid(uneString, 1, unePosRC - 1)
            uneString = Mid(uneString, unePosRC + 2)
            '+2 pour repartir aprés le retour chariot et le saut de ligne
            unePosRC = InStr(1, uneString, Chr(13))
        Loop
        If unePosRC = 0 And uneString <> "" Then
            'Cas où plus de retour chariot et
            'affichage du reste du titre s'il en reste
            Print #unFileId, Spc(5); uneString
        End If
        
        Print #unFileId,
        Print #unFileId, "Date : "; monEtude.LabelDate
        Print #unFileId, "Variante : "; monEtude.TextVar
        If EstNouvelleEtude(monEtude) = False Then
            'Cas où ce n'est pas une étude nouvelle Etude N
            uneString = Mid(monEtude.Caption, 7)
        Else
            uneString = "Etude pas encore enregistrée"
        End If
        Print #unFileId, "Enregistrée sous : "; uneString
        Print #unFileId,
        
        'Impression des données de l'étude
        Print #unFileId, "DONNEES :"
        Print #unFileId, "---------"
        Print #unFileId,
        'Impression du type de voie
        Print #unFileId, "Type de voie : "; DonnerNomTypeVoie(monEtude)
        'Impression du type d'étude
        If monEtude.monTypeEtude = TypeEtudeStandard Then
            uneStrTmp = monEtude.OptionEtudeStandard.Caption
        Else
            uneStrTmp = monEtude.OptionEtudeGiratoire.Caption
        End If
        Print #unFileId, "Type d'aménagement : "; uneStrTmp
        'Impression type de chantier
        If monEtude.monTypeEtude = TypeEtudeGiratoire Then
            uneStrTmp = monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption
        ElseIf monEtude.monTypeChantier = TypeChantierQ1 Then
            uneStrTmp = monEtude.OptionChoixQ1.Caption
        Else
            uneStrTmp = monEtude.OptionChoixQ2.Caption
        End If
        Print #unFileId, "Chantier : "; uneStrTmp
        'Impression des données trafic ini, durée service
        'et croisance annuelle et de la classe de plateforme
        If DonnerTypeVoie(monEtude) = 4 Then
            unTypePL = "BUS"
        Else
            unTypePL = "Poids Lourds" '"PL"
        End If
        
        'modification OF le 21/07/2005 : pas d'affichage des données de trafic si type de voie=parking
        Dim unTrafic, uneDureeS, unTauxC As String
        
        If DonnerTypeVoie(monEtude) = 5 Then '5=type de voie parking
            unTypePL = "Poids Lourds"
            unTrafic = "12"
            uneDureeS = "20 ans"
            unTauxC = "1 % par an"
        Else
            unTrafic = monEtude.TextTrafIni.Text
            uneDureeS = monEtude.TextDuréeS.Text + " ans"
            unTauxC = Format(DonnerCroissAn(monEtude)) + " % par an"
        End If
        'fin modification OF le 21/07/2005
        
        Print #unFileId, "Trafic initial à la mise en service (par sens, par voie et par jour) : "; unTrafic + " " + unTypePL
        Print #unFileId, "Durée de service : "; uneDureeS
        Print #unFileId, "Taux de croissance : "; unTauxC
        unIndPF = DonnerIndicePF(monEtude)
        If unIndPF = 4 Then
            unStringIndPF = "2+"
        Else
            unStringIndPF = Format(unIndPF)
        End If
        Print #unFileId, "Plate-forme : "; "PF" + unStringIndPF
        Print #unFileId,
    
        'Impression des valeurs intermédiaires de l'étude
        Print #unFileId, "VALEURS INTERMEDIAIRES :"
        Print #unFileId, "------------------------"
        Print #unFileId,
        'Impression des données trafic cumulé, risque de calcul
        'si valeur présente, le CAM et le NE arrondi
        If DonnerTypeVoie(monEtude) = 4 Then
            unNbTypePL = "Nombre Cumulé de BUS : "
        Else
            unNbTypePL = "Nombre Cumulé de Poids Lourds : "
        End If
        Print #unFileId, unNbTypePL; monEtude.TextTrafCUM.Text
        'Si une structure est choisie
        'Affichage du risque de calcul si non nul
        Set uneStruct = DonnerStructChoisie(monEtude)
        If Not (uneStruct Is Nothing) Then
            If uneStruct.monTauxRisque <> 0 Then
                Print #unFileId, "Risque de calcul : "; Format(uneStruct.monTauxRisque) + " %"
            End If
        End If
        Print #unFileId, "CAM : "; monEtude.MaskCAM.Text
        Print #unFileId, "NE arrondi : "; Format(monEtude.monNEth, "##,###,###")
        Print #unFileId,

        'Impression des valeurs de vérif au gel de l'étude si vérif au gel voulue
        If mesOptionsGen.maVerifGel Then
            Print #unFileId, "GEL :"
            Print #unFileId, "-----"
            Print #unFileId,
            
            'Affichage des résultats gel pour la qualité Q1
            If monEtude.monEpQ1Trouv And monEtude.monTypeChantier = TypeChantierQ1 Then
                If monEtude.monTypeEtude = TypeEtudeGiratoire Then
                    Print #unFileId, "En condition de chantier : " + monEtude.LabelInfo1.Caption + " " + monEtude.LabelInfo2.Caption
                Else
                    Print #unFileId, "En condition de chantier standard (qualité Q1) : "
                End If
                Print #unFileId, Spc(5); "Indice de Gel de Référence corrigé : "; Format(monEtude.monIndiceGelRefQ1) + " °C.j"
                If monEtude.monIndiceGelRefQ1 > monEtude.monIndiceGelAdmQ1 Then
                    uneInfoGelQ1 = " ======> Chaussée non protégée au gel"
                Else
                    uneInfoGelQ1 = " ======> Chaussée protégée au gel"
                End If
                If monEtude.monIndiceGelAdmQ1 = HorsGel Then
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; "Chaussée Hors Gel"
                Else
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; Format(monEtude.monIndiceGelAdmQ1) + " °C.j" + uneInfoGelQ1
                End If
            End If
            
            'Affichage des résultats gel pour la qualité Q2
            If monEtude.monEpQ2Trouv And monEtude.monTypeChantier = TypeChantierQ2 Then
                Print #unFileId, "En condition de chantier standard (qualité Q2) : "
                Print #unFileId, Spc(5); "Indice de Gel de Référence corrigé : "; Format(monEtude.monIndiceGelRefQ2) + " °C.j"
                If monEtude.monIndiceGelRefQ2 > monEtude.monIndiceGelAdmQ2 Then
                    uneInfoGelQ2 = " ======> Chaussée non protégée au gel"
                Else
                    uneInfoGelQ2 = " ======> Chaussée protégée au gel"
                End If
                If monEtude.monIndiceGelAdmQ2 = HorsGel Then
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; "Chaussée Hors Gel"
                Else
                    Print #unFileId, Spc(5); "Indice de Gel Admissible : "; Format(monEtude.monIndiceGelAdmQ2) + " °C.j" + uneInfoGelQ2
                End If
            End If
            Print #unFileId,
        End If
        
        'Dessin de la carotte Q1 ou Q2 une seule visible en version 2
        Print #unFileId, "DESCRIPTION DE LA CAROTTE : "
        Print #unFileId, "---------------------------"
        Print #unFileId,
        
       'Fermeture du rtf en cours d'écriture
        Close #unFileId
        'Stockage du point d'insertion pour l'image de la carotte
        'Stockage dans le richtextbox
        RichTextTmp.LoadFile unPrintFile, rtfText
        unePosImag = Len(RichTextTmp.Text)
        
        'On le ré-ouvre en mode ajout pour la suite des infos
        Open unPrintFile For Append As #unFileId
         
         'Impressions complémentaires
        Print #unFileId,
        If CheckCommentMat.Value = 1 Then
            ImprimerInfoMatInFile unFileId
        End If
        If CheckCommentStruct.Value = 1 Then
            ImprimerInfoStructInFile unFileId
        End If
        If CheckInfoGel.Value = 1 Then
            ImprimerInfoGelInFile unFileId
        End If
       
       'Fermeture du rtf en cours d'écriture
        Close #unFileId
        'Stockage dans le richtextbox
        RichTextTmp.LoadFile unPrintFile, rtfText
        'Point d'insertion mis en fin de fichier
        RichTextTmp.SelText = ""
        RichTextTmp.SelStart = unePosImag
        
        'Dessin des carottes dans la picture box de l'étude en cours
        Set unePrintZone = monEtude.PictureCarotte
        'Retaillage et placement de la picture box
        monEtude.PictureCarotte.Top = monEtude.TabData.Top
        monEtude.PictureCarotte.Left = monEtude.TabData.Left
        monEtude.PictureCarotte.Width = Printer.ScaleWidth * 0.5
        monEtude.PictureCarotte.Height = Printer.ScaleHeight * 0.5
        monEtude.PictureCarotte.Visible = True
        DoEvents
        'Dessin
        ImprimerCarottes
        'Impression des carottes par copier/coller dans presse-papier
        'et rajouter dans le richtextbox
        RichTextTmp.Top = 0
        RichTextTmp.Left = 0
        RichTextTmp.Height = Me.ScaleHeight
        RichTextTmp.Width = Me.ScaleWidth
        RichTextTmp.Visible = True
        RichTextTmp.ZOrder 0
        RichTextTmp.SelStart = unePosImag
        DoEvents
        
        CollerDansRichTextBoxRTF Printer.ScaleHeight * 0.5, Printer.ScaleWidth * 0.5
        'On sauve le fichier RTF
        RichTextTmp.Font.Name = "Courier"
        RichTextTmp.Font.Size = 10
        DoEvents
        RichTextTmp.SaveFile unPrintFile, rtfRTF
        DoEvents
        'On vide le richtextbox
        RichTextTmp.Text = ""
        'ImprimerCarottesInFile unFileId
        'Close #unFileId
    End If
End Sub

Private Sub CollerDansRichTextBoxRTF(uneH As Long, uneW As Long)
    'Retaillage de la picture box
    monEtude.PictureCarotte.Width = uneW
    monEtude.PictureCarotte.Height = uneH
    monEtude.PictureCarotte.Visible = True
    DoEvents
    'Vidage du presse-papier
    Clipboard.Clear
    DoEvents
    'Copier dans le presse-papier de l'image de la picture box
    Clipboard.SetData monEtude.PictureCarotte.Image, vbCFBitmap
    DoEvents
    'Copie du contenu du presse-papier dans le rich text box
    'en simulant la touche ctrlv dans le richtextbox
    RichTextTmp.SetFocus
    DoEvents
    Call keybd_event(vbKeyControl, 0, 0, 0)
    Call keybd_event(vbKeyV, 0, 0, 0)
    Call keybd_event(vbKeyV, 0, KEYEVENTF_KEYUP, 0) 'release V
    Call keybd_event(vbKeyControl, 0, KEYEVENTF_KEYUP, 0) 'release Ctrl
    'SendMessage RichTextTmp.hWnd, WM_PASTE, 0, 0
    'ci-dessus marche en Win9x mais pas en XP
    DoEvents
    'On vide le presse-papier et la picture box pour pouvoir
    'y mettre les infos supplémentaires éventuelles
    Clipboard.Clear
    monEtude.PictureCarotte.Cls
    monEtude.PictureCarotte.Visible = False
    DoEvents
End Sub



Private Sub RichTextTmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Caption = Caption + "/" + Format(KeyCode) + "/" + Format(Shift)
End Sub
