VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de Struct-Urb"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "À propos de StrucUrb"
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   120
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   360
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   6
      Top             =   3000
      Width           =   6975
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4560
      Width           =   1260
   End
   Begin VB.Image Image2 
      Height          =   1260
      Left            =   3840
      Picture         =   "frmAbout.frx":0614
      Top             =   120
      Width           =   3540
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   5760
      Picture         =   "frmAbout.frx":EEA6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label NumLicence 
      AutoSize        =   -1  'True
      Caption         =   "Licence N° "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Tag             =   "Version"
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lblDescription 
      Caption         =   "Aide au choix, au dimensionnement et à la vérification au gel d'une structure de chaussée urbaine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   960
      TabIndex        =   5
      Tag             =   "Description de l'application"
      Top             =   1680
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Titre de l'application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Tag             =   "Titre de l'application"
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   240
      Width           =   930
   End
   Begin VB.Label WarningLabel 
      Caption         =   "Avertissement:                            Logiciel protégé                      Toute reproduction interdite  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   360
      TabIndex        =   2
      Tag             =   "Avertissement: ..."
      Top             =   4320
      Width           =   2670
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key - Options de sécurité ...
Const KEY_ALL_ACCESS = &H2003F

' Reg Key - Types de ROOT...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' chaîne Unicode terminée par 0
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Const SRCCOPY = &HCC0020
Const ShowText$ = "Frank TRIFILETTI"
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nheight As Integer, ByVal hSrcDC As Long, ByVal Xsrc As Integer, ByVal Ysrc As Integer, ByVal dwRop As Long) As Integer
Dim ShowIt%, monIndMsg%
Dim monTabString(10) As String


Private Sub Form_Load()
    lblVersion.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
    lblVersion.Left = lblTitle.Left + lblTitle.Width + 60
    lblVersion.Top = lblTitle.Top + lblTitle.Height - lblVersion.Height
    
    CentrerFenetreEcran Me
    
    'Affectation du contexte d'aide
    HelpContextID = IDhlp_WinAbout
    
    'Affichage du numéro de licence
    NumLicence.Caption = LBLICENCE & NumeroLicence
     
    'Traitement permettant de lister la boucle des intervenants
    unDecalage = "     "
    WarningLabel.Caption = "Avertissement :" + Chr(13) + unDecalage + "Logiciel protégé"
    WarningLabel.Caption = WarningLabel.Caption + Chr(13) + unDecalage + "Toute reproduction interdite"
    WarningLabel.Font.Bold = True
    'Initialisation de l'indice des messages listant les participants
    monIndMsg% = 0
    'Initialisation des noms des participants
    monTabString(0) = "Production du cahier des charges"
    monTabString(1) = "    CEREMA / TV / VOI / CGR"
    monTabString(2) = "    CERTU / Département SYSTEMES / Groupe Génie Urbain"
    monTabString(3) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
    monTabString(4) = "Réalisation du développement du logiciel"
    monTabString(5) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
    monTabString(6) = "Diffusion et Assistance au logiciel"
    monTabString(7) = "    CEREMA / TV / EREN / Digit@l"
    monTabString(8) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
End Sub




Private Sub cmdOK_Click()
    FermerFenetre Me
End Sub





Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Compteur de boucle
        Dim rc As Long                                          ' Code de retour
        Dim hKey As Long                                        ' Pointeur vers une clé de registre ouvert
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Type de données d'une clé de registre
        Dim tmpVal As String                                    ' Stockage temp. pour une valeur de clé de registre
        Dim KeyValSize As Long                                  ' Taille de la variable clé de registre
        '------------------------------------------------------------
        ' Ouvrir RegKey sous KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvrir clé de registre
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gérer les erreurs...
        

        tmpVal = String$(1024, 0)                               ' Allouer l'espace pour la variable
        KeyValSize = 1024                                       ' Marquer la taille de la variable
        

        '------------------------------------------------------------
        ' Extraire la valeur de clé de registre...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Lire/créer validation de clé
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gérer les erreurs
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 termine les chaînes par 0...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null atteint, extraire de la chaîne
        Else                                                    ' WinNT ne termine pas les chaînes par 0...
                tmpVal = Left(tmpVal, KeyValSize)                   ' 0 non trouvé, extraire chaîne uniquement
        End If
        '---------------------------------------------------------------
        ' Determiner le type de la valeur de la clé pour la convertir...
        '---------------------------------------------------------------
        Select Case KeyValType                                  ' Rechercher types de données...
        Case REG_SZ                                             ' Type de données de clé de registre String
                KeyVal = tmpVal                                     ' Copier valeur de la chaîne
        Case REG_DWORD                                          ' Type de données de clé de registre Double Word
                For i = Len(tmpVal) To 1 Step -1                    ' Convertir chaque bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construire valeur caractère par caractère
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word en String
        End Select
        

        GetKeyValue = True                                      ' Renvoyer Réussite
        rc = RegCloseKey(hKey)                                  ' Fermer la clé de registre
        Exit Function                                           ' Sortir

GetKeyError:    ' Nettoyage si erreur...
        KeyVal = ""                                             ' Affecter chaîne vide à la valeur de retour
        GetKeyValue = False                                     ' Renvoyer Échec
        rc = RegCloseKey(hKey)                                  ' Fermer la clé de registre
End Function



Private Sub Timer1_Timer()
    Dim i As Integer
    Dim uneString As String
    
    If (ShowIt% Mod 20 = 0) Then
        Picture1.CurrentX = 20
        Picture1.CurrentY = Picture1.ScaleHeight - 20
        'Affichage du participant d'indice monIndMsg%
        Picture1.Print monTabString(monIndMsg% Mod 9)
        ShowIt% = 1
        If monIndMsg% = 9 Then
            'Pour éviter un débordement de capacité des entiers
            monIndMsg% = 1
        Else
            'Permettra l'affichage du message suivant
            monIndMsg% = monIndMsg% + 1
        End If
    Else
        i = BitBlt(Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight - 1, Picture1.hDC, 0, 1, SRCCOPY)
        ShowIt% = ShowIt% + 1
    End If
End Sub
