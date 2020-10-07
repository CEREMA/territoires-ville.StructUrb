VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmOptionsGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramètres généraux"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmOptionsGen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.Frame Frame1 
      Caption         =   "Trafic :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   3495
      Begin ComCtl2.UpDown UpDownDS 
         Height          =   395
         Left            =   2320
         TabIndex        =   38
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   327680
         Value           =   5
         BuddyControl    =   "TextDuréeS"
         BuddyDispid     =   196612
         OrigLeft        =   2320
         OrigTop         =   450
         OrigRight       =   2560
         OrigBottom      =   825
         Max             =   50
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox ComboCroisAnn 
         Height          =   315
         ItemData        =   "frmOptionsGen.frx":030A
         Left            =   1920
         List            =   "frmOptionsGen.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   640
      End
      Begin VB.TextBox TextDuréeS 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Durée de service : "
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ans"
         Height          =   195
         Left            =   2640
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Croissance annuelle : "
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "% par an"
         Height          =   195
         Left            =   2640
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame CadreCoul 
      Caption         =   "Couleurs : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   3495
      Begin VB.PictureBox CoulFond1 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.PictureBox CoulFond2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   6
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox CoulBase2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox CoulBase1 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.PictureBox CoulSurf 
         BackColor       =   &H00808000&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Couche de Fondation 2 :"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Couche de Fondation 1 : "
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Couche de Surface : "
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Couche de Base 1 : "
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Couche de Base 2 :"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1410
      End
   End
   Begin VB.Frame CadreGel 
      Caption         =   "Gel :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3720
      TabIndex        =   19
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox ComboTailleAgglo 
         Height          =   315
         ItemData        =   "frmOptionsGen.frx":0336
         Left            =   2280
         List            =   "frmOptionsGen.frx":0343
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   2655
      End
      Begin VB.ComboBox ComboStation 
         Height          =   315
         ItemData        =   "frmOptionsGen.frx":03A0
         Left            =   2280
         List            =   "frmOptionsGen.frx":03A2
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox CheckVerifGel 
         Caption         =   "Vérification au Gel"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox AltiAgglo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label HStation 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "mètres"
         Height          =   195
         Left            =   3240
         TabIndex        =   34
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label LabGel 
         AutoSize        =   -1  'True
         Caption         =   "Station de Référence : "
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label LabAlt 
         AutoSize        =   -1  'True
         Caption         =   "Altitude de l'agglomération : "
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Altitude de la Station : "
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1590
      End
      Begin VB.Label LabAgglo 
         AutoSize        =   -1  'True
         Caption         =   "Taille de l'agglomération : "
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1830
      End
      Begin VB.Label LabAltUnit 
         AutoSize        =   -1  'True
         Caption         =   "mètres"
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   465
      End
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
      Left            =   6240
      TabIndex        =   11
      Tag             =   "OK"
      Top             =   3240
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   12
      Tag             =   "Annuler"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   18
         Tag             =   "Exemple 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   17
         Tag             =   "Exemple 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   15
         Tag             =   "Exemple 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptionsGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AltiAgglo_KeyUp(KeyCode As Integer, Shift As Integer)
    SaisieEntierPositifEntreMinMax KeyCode, AltiAgglo, 200, 0, 9000, MsgAltiAgglo
End Sub

Private Sub AltiAgglo_LostFocus()
    If ActiveControl.Name <> "cmdCancel" Then
        SaisieEntierPositifEntreMinMax 8, AltiAgglo, 200, 0, 9000, MsgAltiAgglo
    End If
End Sub

Private Sub cmdCancel_Click()
    FermerFenetre Me
End Sub


Private Sub cmdOK_Click()
    'Sauvegarde dans les options générales pour avoir
    'les mêmes valeurs pendant la session ouverte
    With mesOptionsGen
        .maDuréeService = Val(TextDuréeS)
        .maCroisAnnuel = Val(ComboCroisAnn.Text)
        
        .maCoulSurf = CoulSurf.BackColor
        .maCoulBase1 = CoulBase1.BackColor
        .maCoulBase2 = CoulBase2.BackColor
        .maCoulFond1 = CoulFond1.BackColor
        .maCoulFond2 = CoulFond2.BackColor
        
        .maVerifGel = CheckVerifGel.Value
        .monTailleAgglo = ComboTailleAgglo.ListIndex
        .monAltiAgglo = Val(AltiAgglo.Text)
        .monIndStationRef = ComboStation.ListIndex + 1
    
        'Sauvegarde dans la base de registre pour récupérer
        'les mêmes valeurs à la prochaine
        StockerOptionsGen
    End With
    
    'Mise à jour de la frame de vérif au gel
    'et des couleurs des couches des carottes Q1 et Q2
    'de toutes les fenêtres filles
    For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmDocument Then
            ActualiserFrameVerifGel Forms(i)
            ChangerCouleurCouches Forms(i)
        End If
    Next i

    FermerFenetre Me
End Sub


Private Sub ComboStation_Click()
    'Affichage de l'altitude de la station de référence choisie
    HStation.Caption = Format(monTabStation(ComboStation.ListIndex + 1).monAltitude)
End Sub

Private Sub CoulBase1_Click()
    ChoisirCouleur CoulBase1
End Sub

Private Sub CoulBase2_Click()
    ChoisirCouleur CoulBase2
End Sub

Private Sub CoulFond1_Click()
    ChoisirCouleur CoulFond1
End Sub

Private Sub CoulFond2_Click()
    ChoisirCouleur CoulFond2
End Sub

Private Sub CoulSurf_Click()
    ChoisirCouleur CoulSurf
End Sub

Private Sub Form_Load()
    CentrerFenetreEcran Me
    HelpContextID = IDhlp_WinOptionsGen
    
    'Remplissage du tableau de station de référence pour le gel
    'et Remplissage de la combobox ComboStation
    RemplirLesStationsMétéo Me
    
    'Affichage des valeurs par défaut lue dans les options générales
    TextDuréeS.Text = Format(mesOptionsGen.maDuréeService)
    ComboCroisAnn.Text = Format(mesOptionsGen.maCroisAnnuel)
    CoulSurf.BackColor = mesOptionsGen.maCoulSurf
    CoulBase1.BackColor = mesOptionsGen.maCoulBase1
    CoulBase2.BackColor = mesOptionsGen.maCoulBase2
    CoulFond1.BackColor = mesOptionsGen.maCoulFond1
    CoulFond2.BackColor = mesOptionsGen.maCoulFond2
    
    CheckVerifGel.Value = mesOptionsGen.maVerifGel
    
    'Affectation de l'Agglomération des études
    ComboTailleAgglo.ListIndex = mesOptionsGen.monTailleAgglo
    AltiAgglo.Text = mesOptionsGen.monAltiAgglo
    
    'Affectation de la station de référence par défaut
    ComboStation.ListIndex = mesOptionsGen.monIndStationRef - 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Remise à zéro du nombre de stations
    monNbStation = 0
End Sub


Private Sub TextDuréeS_KeyUp(KeyCode As Integer, Shift As Integer)
    Call VerifSaisieEntierPositif(KeyCode, TextDuréeS, "")
End Sub

Private Sub TextDuréeS_LostFocus()
    If ActiveControl.Name <> "cmdCancel" Then VerifierSortieSaisieEntierPositif TextDuréeS, 5, 50
End Sub



