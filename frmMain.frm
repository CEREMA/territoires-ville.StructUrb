VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Struct-Urb"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nouveau"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Ouvrir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Enregistrer"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4740
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6853
            Text            =   "Struct-Urb version ?.?.?"
            TextSave        =   "Struct-Urb version ?.?.?"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "31/08/2005"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:33"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   600
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   1560
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1052
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":243E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2790
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2AE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Etude"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nouvelle"
         HelpContextID   =   100
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Fermer"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Enre&gistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Enregistrer &sous..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Barre d'&outils"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barre d'&état"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsGen 
         Caption         =   "Paramètres &généraux..."
      End
      Begin VB.Menu mnuOptionsMat 
         Caption         =   "Paramètres &matériaux..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Fenêtre"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&Nouvelle fenêtre"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Réorganiser les icônes"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpSommaire 
         Caption         =   "&Sommaire"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "A&ide sur..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Rechercher..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "À &propos de Struct-Urb..."
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Sub MDIForm_Activate()
    'On sort si ouverture fichier structure s'est mal passée
    If Tag = MsgOpenStructFileFailed Then Unload Me
End Sub

Private Sub MDIForm_Load()
    Dim unMRUSettings As Variant, unNomFich As String
    
    'Mise à jour de l'ihm du à QLM
     Call InitQlm
    
    'Affichage de la version dnas le titre et dans la barre d'état
    'Caption = App.Title + " version " & App.Major & "." & App.Minor & "." & App.Revision
    sbStatusBar.Panels(1).Text = Caption
    
    'Affectation du fichier d'aide
    'modif O.FOREL du 14/01/2005 :insertion du fichier struct-chm
    unNomFich = CorrigerNomFichier(App.Path + "\Struct-Urb.chm")
    App.HelpFile = unNomFich
    dlgCommonDialog.HelpFile = App.HelpFile
    
    'Index des aides pour les menus (constantes definies dans sub main)
    mnuFileNew.HelpContextID = IDhlp_NewSite
    mnuFileOpen.HelpContextID = IDhlp_OpenSite
    mnuFileSave.HelpContextID = IDhlp_SaveSite
    mnuFileSaveAs.HelpContextID = IDhlp_SaveAsSite
    mnuFileClose.HelpContextID = IDhlp_CloseSite
    mnuFilePrint.HelpContextID = IDhlp_WinPrint
    mnuOptionsGen.HelpContextID = IDhlp_WinOptionsGen
    mnuHelpAbout.HelpContextID = IDhlp_WinAbout
    
    'Récupération de la liste des fichiers récents
    unMRUSettings = GetAllSettings(App.Title, "Recent Files")
    If IsEmpty(unMRUSettings) = False Then
        'Cas où la liste des fichiers récents (MRU Files) n'est pas vide
        'getallsettings renvoit un variant non initialisé = Empty
        'On alimente les menus mnuFileMRU
        For i = UBound(unMRUSettings, 1) To 0 Step -1
            'A l'envers car on met le nom de fichier toujours en tête
            unNomFich = unMRUSettings(i, 1)
            ActualiserListeFichiersRecents unNomFich
        Next i
    End If
    
    'Récupération des options générales par lecture des valeurs
    'de ces options stockées dans la base de registre
    RécupérerOptionsGen
    
    'Récupération des options matériaux par lecture des valeurs
    'de ces options stockées dans la base de registre
    RécupérerOptionsMat
    
    'Récupération de la position et de la taille dans la base de registre
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 8500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    If (Screen.Width / Screen.TwipsPerPixelX) > 800 And (Screen.Height / Screen.TwipsPerPixelY) > 600 Then
        'Si résolution supérieure à 800 x 600
        Me.WindowState = vbNormal
        Left = 0
        Top = 0
        Width = Screen.Width * 0.8
        Height = Screen.Height * 0.8
    End If
    
    'Mise à jour des boutons dans la toolbar permettant l'impression
    'et la sauvegarde car il n'y a pas de fenêtre fille ouverte
    '==> Impression et sauvegarde impossible
    tbToolBar.Buttons("Print").Visible = False
    tbToolBar.Buttons("Save").Visible = False
    
    'Chargement du fichier structure du CERTU
    If OuvrirFichierStructures(App.Path + "\CERTU.STR", maColStructCERTU, maColMatBFCERTU) = False Then
        'Cas d'erreur lors de la lecture du fichier de structures CERTU
        '==> on ne l'utilise pas et on vide les collections partiellement
        'remplies
        ViderCollection maColStructCERTU
        ViderCollection maColMatBFCERTU
        
        'On indique qu'il faut fermer la MDI mère
        'car on a un mauvais fichier de structures
        'Ce sera fait dans le MDIForm_Activate
        Me.Tag = MsgOpenStructFileFailed
    End If
    'Affectation des autorisations d'utilisation à partir des indices
    'des matériaux CERTU non autorisés stockées dans les options matériaux
    AlimenterAutorisation True

    'Chargement du fichier structure personnel éventuel
    If mesOptionsMat.monFichPersoSTR <> "" Then
        If OuvrirFichierStructures(mesOptionsMat.monFichPersoSTR, maColStructPerso, maColMatBFPerso) = False Then
            'Cas d'erreur lors de la lecture du fichier de structures perso
            '==> on ne l'utilise pas et on vide les collections partiellement
            'remplies
            ViderCollection maColStructPerso
            ViderCollection maColMatBFPerso
            mesOptionsMat.monFichPersoSTR = ""
        End If
    End If
    'Affectation des autorisations d'utilisation à partir des indices
    'des matériaux PERSO non autorisés stockées dans les options matériaux
    AlimenterAutorisation False
    
    'Lancement d'une nouvelle étude si pas de fichier de démarrage
    'l'argument de la ligne de commande
    If Command <> "" Then
        'Ouvrir Struct-Urb avec le paramètre de la ligne commande
        '= Nom complet du fichier sur lequel on a double-cliqué
        OuvrirEtude Command
    End If
End Sub


Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument, uneForm As Form

    'suppression protection
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Sub
    'fin suppression protection
    
    'Indication de l'ouverture d'une nouvelle étude
    maNewEtude = True
    
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = MsgEtude0 + Format(lDocumentCount)
    
    If frmD.Tag = "A_Fermer" Then
        'Fermeture de la fenêtre car erreur survenue à son load
        Unload frmD
    Else
        'Ouverture si pas d'erreur lors du load
        frmD.Show
        'Redessin des carottes à cause d'un bug VB
        Set uneForm = fMainForm.ActiveForm
        If (uneForm.ComboStruct.ListIndex > -1) Then
            AfficherCarottes uneForm
        End If
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not VerifierFinSaisie
    'si Cancel = true on ne sort pas d l'application
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    'Stockage en base de registre des fichiers récents
    For i = 1 To 4
        If mnuFileMRU(i - 1).Visible Then
            unFileMRU = "File" + Format(i)
            SaveSetting App.Title, "Recent Files", unFileMRU, Mid(mnuFileMRU(i - 1).Caption, 4)
        End If
    Next
End Sub


Private Sub mnuFile_Click()
    Dim uneMiseEnGrisé As Boolean, uneMiseEnVisible As Boolean
    
    'Stockage dans le tag de la vérif de la saisie
    If VerifierFinSaisie Then
        mnuFile.Tag = "Vrai"
    Else
        mnuFile.Tag = "Faux"
    End If
    
    If Forms.Count = 1 Then
        'Aucune fenêtre fille ouverte
        'La seul fenetre ouverte la MDI mère
        uneMiseEnGrisé = False
        uneMiseEnVisible = False
    Else
        'Des fenêtres filles ouvertes
        uneMiseEnGrisé = True
        uneMiseEnVisible = True
    End If
    
    'Mise à jour des items du menu Site (= mnuFile)
    mnuFileClose.Enabled = uneMiseEnGrisé
    mnuFileSave.Enabled = uneMiseEnGrisé
    mnuFileSaveAs.Enabled = uneMiseEnGrisé
    mnuFilePrint.Enabled = uneMiseEnGrisé
End Sub

Private Sub mnuHelp_Click()
    VerifierFinSaisie
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuLicence_Click()
    frmKey.Show 1
    'Mise à jour de l'ihm
    Call InitQlm
End Sub

Private Sub mnuOptions_Click()
    VerifierFinSaisie
End Sub

Private Sub mnuOptionsGen_Click()
    frmOptionsGen.Show vbModal, Me
End Sub

Private Sub mnuOptionsMat_Click()
    frmOptionsMat.Show vbModal, Me
End Sub

Private Sub mnuView_Click()
    VerifierFinSaisie
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub


Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbToolBar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbToolBar.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub

Private Sub mnuWindow_Click()
    VerifierFinSaisie
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    If VerifierFinSaisie = False Then Exit Sub

    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

'Ajout O.Forel 14/01/2005 : modification du menu aide (méthode décrites dans Chelp.bas)
Private Sub mnuHelpIndex_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    'Modif fait par Frank Trifiletti on utilise le contextid de la fenêtre étude en cours
    'qui est dans la globale monetude dont son helpcontextid est mis à jour dans la sub ChangerHelpId
    'qui est appellé à chaque Form_Activate et dans le TabData_Click de frmDocument.frm
    'car le contextid était toujours nulle avec showindex normal on ne le passe pas en argument.
    If monEtude Is Nothing Then
        'Cas d'appel  de F1 si aucun étude ouverte sinon plantage
        'Onglet Index supprimé!!!
        'Call objHelp.ShowIndex(App.HelpFile, "Main")
        Call objHelp.Show(App.HelpFile, "Main")
    Else
        Call objHelp.Show(App.HelpFile, "Main", monEtude.HelpContextID)
    End If
    'Fin modif F.Trifiletti
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpSearch_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.ShowSearch(App.HelpFile, "Main")
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpSommaire_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.Show(App.HelpFile, "Main")
    Set objHelp = Nothing
End Sub
'fin ajout o.Forel


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuWindowNewWindow_Click()
    'Création d'une nouvelle étude
    LoadNewDoc
End Sub

Private Sub mnuFileOpen_Click()
    Dim unFich As String
    
    'Si pas fin de saisie on ne fait rien
    If mnuFile.Tag = "Faux" Then Exit Sub
    
    unFich = ChoisirFichier(MsgOpen, MsgUrbFile, CurDir)
    If unFich <> "" Then OuvrirEtude unFich
End Sub


Private Sub mnuFileClose_Click()
    Unload Screen.ActiveForm
End Sub


Private Sub mnuFileSave_Click()
    'Si pas fin de saisie on ne fait rien
    If mnuFile.Tag = "Faux" Then Exit Sub
    
    'Sauvegarde de l'étude active
    'Le nom du fichier ne sert que si c'est une étude existante
    'Titre fenetre = "Etude " + numéro ou nom fichier
    SauverEtude monEtude, Mid(monEtude.Caption, 7), False
End Sub


Private Sub mnuFileSaveAs_Click()
    'Sauvegarde de l'étude active
    'Le nom du fichier est vide car on fait un enregistrer sous
    'le nom de fichier est choisi dans SauverEtude
    SauverEtude monEtude, "", True
End Sub

Private Sub mnuFilePrint_Click()

    'suppression protection
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Sub
    'fin suppression protection
    
    'Si pas fin de saisie on ne fait rien
    If mnuFile.Tag = "Faux" Then Exit Sub
    
    'Si la structure n'est pas choisie ou n'est dimensionnée
    'ni pour Q1 et ni pour Q2 ==> aucune impression
    If DonnerStructChoisie(Screen.ActiveForm) Is Nothing Then
        MsgBox MsgNoPrintNoStruct, vbInformation
    ElseIf Screen.ActiveForm.monEpQ1Trouv Or Screen.ActiveForm.monEpQ2Trouv Then
        If Printers.Count = 0 Then
            MsgBox "Aucune imprimante n'est connectée à ce poste.", vbCritical
        Else
            frmImprimer.Show vbModal
        End If
    Else
        MsgBox MsgNoPrintNoDim, vbInformation
    End If
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
    OuvrirEtude Mid(mnuFileMRU(Index).Caption, 4)
End Sub


Private Sub mnuFileExit_Click()
    'décharger la feuille
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    'Si pas fin de saisie on ne fait rien
    If mnuFile.Tag = "Faux" Then Exit Sub
    
    LoadNewDoc
End Sub


Public Function ChoisirFichier(unTitre As String, uneExtension As String, unInitDir As String)
    'Fonction ouvrant un sélectionneur de fichier avec l'extension passée
    'en paramètre et retournant le nom complet du fichier choisi ou une
    'chaine vide si rien de choisir ou click sur Annuler
    'Le sélectionneur de fichier s'ouvre dans le répertoire unInitDir
    
    With dlgCommonDialog
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurChoix
        
        'Ouverture d'une fenêtre Ouvrir fichier
        
        'définir les indicateurs et attributs
        'du contrôle des dialogues communs
        .CancelError = True
        .DialogTitle = unTitre
        .InitDir = unInitDir
        .Filter = uneExtension
        .Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        .FileName = ""
        If unTitre = MsgOpen Then
            .ShowOpen
        ElseIf unTitre = MsgSaveAs Or unTitre = MsgPrintInFile Then
            .ShowSave
        Else
            MsgBox MsgErreurProg + MsgErreurTypeShowWinInconnu + MsgIn + "frmMain:ChoisirFichier", vbCritical
        End If
        
        If Len(.FileName) = 0 Then
            'Cas où aucun fichier choisi
            ChoisirFichier = ""
        Else
            'Affectation du fichier à ouvrir
            ChoisirFichier = .FileName
        End If
        
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
        'Sortie de la procédure pour éviter le passage
        'dans la gestion d'erreur
        Exit Function
    End With
    
ErreurChoix:
    'Cas où click sur Annuler
    ChoisirFichier = ""
    Exit Function
End Function

'Code pour modifier l'ihm suite à l'implémentation de Qlm
Private Sub InitQlm()
    'Initialisation des menus modifiés par QLM
    'les variables globales sont maj par protection.bas
    'ATTENTION : vérifier les noms des menus!!!
    Me.mnuHelpBar2.Visible = GvisibiliteMnuBarre
    Me.mnuLicence.Visible = GvisibiliteMnuLicence
    'a adapter en fonction du clogiciel
    Me.Caption = "Struct-Urb v" + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision) + GmodifTitreApplication
    'fin initialisation qlm
End Sub
