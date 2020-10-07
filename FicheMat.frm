VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FicheMat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fiche de mat�riau"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "FicheMat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread SpreadComposant 
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "FicheMat.frx":030A
      TabIndex        =   14
      Top             =   2760
      Width           =   8535
   End
   Begin FPSpread.vaSpread SpreadInfo 
      Height          =   735
      Left            =   120
      OleObjectBlob   =   "FicheMat.frx":08D2
      TabIndex        =   11
      Top             =   1560
      Width           =   5415
   End
   Begin FPSpread.vaSpread SpreadParamGel 
      Height          =   525
      Left            =   6720
      OleObjectBlob   =   "FicheMat.frx":0DFD
      TabIndex        =   6
      Top             =   600
      Width           =   1950
   End
   Begin VB.CommandButton BtnFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer la fiche"
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   0
      Top             =   4020
      Width           =   8535
   End
   Begin VB.TextBox TextNorme 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox TextQualit� 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox TextNomMat 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   2
      Top             =   120
      Width           =   7815
   End
   Begin VB.TextBox TextAbregMat 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox Comment 
      Height          =   1335
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2355
      _Version        =   327680
      BackColor       =   -2147483624
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FicheMat.frx":121B
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Norme : "
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Commentaire : "
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Classe : "
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nom : "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Abr�g� : "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Param�tres de transmission du gel : "
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "FicheMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private monNbActivate As Byte 'Indicateur du premier activate

Private Sub BtnFermer_Click()
    Unload Me
End Sub

Public Sub RemplirFicheMat(unMat As Object)
    'Remplissage de la fiche mat�riau
    'du mat�riau pass� en param�tre
    'Tous les cellules de tous les spread sont lock�es
    '==> non-modifiable
    
    'Stockage du nom du mat�riau pour v�rif unicit�
    monNom = unMat.monNom
    
    'Propri�t�s communes � tous les mat�riaux
    TextNomMat.Text = unMat.monNom
    TextAbregMat.Text = unMat.monAbr�g�
    'Remplissage des param�tres de gel A et B
    SpreadParamGel.Row = 1
    SpreadParamGel.Col = 1
    SpreadParamGel.Text = Format(unMat.monAGel)
    SpreadParamGel.Lock = True
    SpreadParamGel.Col = 2
    SpreadParamGel.Text = Format(unMat.monBGel)
    SpreadParamGel.Lock = True
        
    'Propri�t�s diff�rentes
    If TypeOf unMat Is MatSimple Or TypeOf unMat Is Mat�riau Then
        TextQualit�.Text = unMat.maQualit�
        TextNorme.Text = unMat.maNorme
        Comment.TextRTF = unMat.monCommentaire
        If TypeOf unMat Is Mat�riau Then
            SpreadInfo.Row = 1
            SpreadInfo.Col = 1
            SpreadInfo.Text = Format(unMat.monYoung)
            SpreadInfo.Lock = True
            SpreadInfo.Col = 2
            SpreadInfo.Text = Format(unMat.monPoisson)
            SpreadInfo.Lock = True
            SpreadInfo.Col = 3
            If unMat.monEpsilon = 0 Then
                SpreadInfo.Text = ""
            Else
                SpreadInfo.Text = Format(unMat.monEpsilon)
            End If
            SpreadInfo.Lock = True
            SpreadInfo.Col = 4
            If unMat.monSigma = 0 Then
                SpreadInfo.Text = ""
            Else
                SpreadInfo.Text = Format(unMat.monSigma)
            End If
            SpreadInfo.Lock = True
        End If
    ElseIf TypeOf unMat Is MatCompos� Then
        'Remplissage du spread � partir la collection mesCompositions
        'qui contient unNbComp+1 valeurs par composition
        'Eprec(i),e(i,1),comp(i,1),...,e(i,N),Comp(i,N) � la suite dans
        'la collection mesCompositions
        
        'Calcul du nombre de composants max fix� par l'IHM
        unNbComp = (SpreadComposant.MaxCols - 1) / 2
        
        'Calcul du nombre de lignes du spread (grace � la division enti�re)
        SpreadComposant.MaxRows = unMat.mesCompositions.Count \ (2 * unNbComp + 1)
        
        For i = 1 To SpreadComposant.MaxRows
            SpreadComposant.Row = i
            SpreadComposant.Col = 1
            unN = (i - 1) * (2 * unNbComp + 1) + 1
            SpreadComposant.Text = Format(unMat.mesCompositions(unN))
            SpreadComposant.Lock = True
            For j = 1 To unNbComp
                'Affichage �paisseur du composant
                SpreadComposant.Col = j * 2
                uneVal = unMat.mesCompositions(unN + 2 * j - 1)
                If uneVal = 0 Then
                    SpreadComposant.Text = ""
                Else
                    SpreadComposant.Text = Format(uneVal)
                End If
                SpreadComposant.Lock = True
                'Passage � une colonne de composant
                SpreadComposant.Col = j * 2 + 1
                'Affichage de l'abr�g� du composant
                SpreadComposant.Text = unMat.mesCompositions(unN + 2 * j)
                SpreadComposant.Lock = True
            Next j
        Next i
    Else
        MsgBox MsgErreurProg + MsgErreurMat�riauInconnu + MsgIn + "FicheMat:RemplirFicheMat", vbCritical
    End If
End Sub

Private Sub Form_Activate()
    'D�codage du tag pour visualier les bons controles
    'lors du permier activate
    Dim unePos As Long
    Dim unTypeMat As String, unAbr�g� As String
    Dim unMatSel As Object
    
    'Stockage du premier activate
    If monNbActivate = 1 Then
        Exit Sub
    Else
        monNbActivate = 1
    End If
    
    'R�cup du s�parateur d�cimale . ou ,
    'fix� dans les param�tres r�gionaux de Windows
    TrouverCaract�reD�cimalUtilis�
    
    'Couleur de fond pour les cellules lock�es
    'on prend la couleur des info-bulles
    SpreadParamGel.LockBackColor = vbInfoBackground
    SpreadInfo.LockBackColor = vbInfoBackground
    SpreadComposant.LockBackColor = vbInfoBackground
    
    'Utilisation du param�tre d�cimal en cours dans les spreads
    InitSpreadCaract�reD�cimal SpreadParamGel, monCarDeci
    InitSpreadCaract�reD�cimal SpreadInfo, monCarDeci
    
    'Limitation en largeur du richTextBox
    Comment.RightMargin = Comment.Width - 120
    
    'R�cup�ration du type de mat�riau
    unePos = InStr(1, Tag, "/")
    unTypeMat = Mid(Tag, 1, unePos - 1)
    
    If unTypeMat = "Simple" Then
        SpreadInfo.Visible = False
        SpreadComposant.Visible = False
        Caption = Caption + " " + unTypeMat
    ElseIf unTypeMat = "Composant" Or unTypeMat = "FondBase" Then
        SpreadComposant.Visible = False
    ElseIf unTypeMat = "Compos�" Then
        SpreadInfo.Visible = False
        SpreadComposant.Visible = True
        Caption = Caption + " " + unTypeMat
        'Repositionnement du SpreadComposant
        SpreadComposant.Top = SpreadInfo.Top
        SpreadComposant.Height = BtnFermer.Top - SpreadComposant.Top - 60
    Else
        MsgBox MsgErreurProg + MsgErreurMat�riauInconnu + MsgIn + "FicheMat:Form_Activate", vbCritical
    End If
    
    'R�cup de l'abr�g�
    'unePos = position du 1er et dernier /
    monAbr�g� = Mid(Tag, unePos + 1)
    
    'R�cup du mat�riau s�lectionn�
    If unTypeMat = "Simple" Or unTypeMat = "Compos�" Or unTypeMat = "Composant" Then
        Set unMatSel = maColMatSurf.Item(monAbr�g�)
        RemplirFicheMat unMatSel
    ElseIf unTypeMat = "FondBase" Then
        If monEtude.CheckFichPerso.Value = 0 Then
            'Cas o� Fichier de structure utilis� est celui du CERTU
            Set unMatSel = maColMatBFCERTU.Item(monAbr�g�)
        Else
            'Cas o� Fichier de structure utilis� est celui personnel
            Set unMatSel = maColMatBFPerso.Item(monAbr�g�)
        End If
        RemplirFicheMat unMatSel
    Else
        MsgBox MsgErreurProg + MsgErreurMat�riauInconnu + MsgIn + "FicheMat:Form_Activate", vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Remise � z�ro du compteur signalant le premier activate event
    monNbActivate = 0
End Sub

