VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mat�riau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonNom As String 'copie locale
Private mvarmonAbr�g� As String 'copie locale
Private mvarmaNorme As String 'copie locale
Private mvarmonCommentaire As String 'copie locale
Private mvarmonYoung As Single 'copie locale
Private mvarmonPoisson As Single 'copie locale
Private mvarmonEpsilon As Single 'copie locale
Private mvarmonSigma As Single 'copie locale
Private mvarmonAGel As Single 'copie locale
Private mvarmonBGel As Single 'copie locale
Private mvarmaQualit� As String 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonUtilisationAutoris�e As Boolean 'copie locale
Public Property Let monUtilisationAutoris�e(ByVal vData As Boolean)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monUtilisationAutoris�e = 5
    mvarmonUtilisationAutoris�e = vData
End Property


Public Property Get monUtilisationAutoris�e() As Boolean
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monUtilisationAutoris�e
    monUtilisationAutoris�e = mvarmonUtilisationAutoris�e
End Property



Public Property Let maQualit�(ByVal vData As String)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maQualit� = 5
    mvarmaQualit� = vData
End Property


Public Property Get maQualit�() As String
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maQualit�
    maQualit� = mvarmaQualit�
End Property



Public Property Let monBGel(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monBGel = 5
    mvarmonBGel = vData
End Property


Public Property Get monBGel() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monBGel
    monBGel = mvarmonBGel
End Property



Public Property Let monAGel(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monAGel = 5
    mvarmonAGel = vData
End Property


Public Property Get monAGel() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monAGel
    monAGel = mvarmonAGel
End Property




Public Sub SetPropsPhysic(unYoung As Single, unPoisson As Single, unEpsilon As Single, unSigma As Single)
    monYoung = unYoung
    monPoisson = unPoisson
    monEpsilon = unEpsilon
    monSigma = unSigma
End Sub

Public Sub SetProps(unNom As String, unAbr�g� As String, uneNorme As String, unCommentaire As String)
    monNom = unNom
    monAbr�g� = unAbr�g�
    maNorme = uneNorme
    monCommentaire = unCommentaire
    monUtilisationAutoris�e = True
End Sub

Public Property Let monSigma(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monSigma = 5
    mvarmonSigma = vData
End Property


Public Property Get monSigma() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monSigma
    monSigma = mvarmonSigma
End Property



Public Property Let monEpsilon(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monEpsilon = 5
    mvarmonEpsilon = vData
End Property


Public Property Get monEpsilon() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monEpsilon
    monEpsilon = mvarmonEpsilon
End Property



Public Property Let monPoisson(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monPoisson = 5
    mvarmonPoisson = vData
End Property


Public Property Get monPoisson() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monPoisson
    monPoisson = mvarmonPoisson
End Property



Public Property Let monYoung(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monYoung = 5
    mvarmonYoung = vData
End Property


Public Property Get monYoung() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monYoung
    monYoung = mvarmonYoung
End Property



Public Property Let monCommentaire(ByVal vData As String)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monCommentaire = 5
    mvarmonCommentaire = vData
End Property


Public Property Get monCommentaire() As String
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monCommentaire
    monCommentaire = mvarmonCommentaire
End Property



Public Property Let maNorme(ByVal vData As String)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maNorme = 5
    mvarmaNorme = vData
End Property


Public Property Get maNorme() As String
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maNorme
    maNorme = mvarmaNorme
End Property



Public Property Let monAbr�g�(ByVal vData As String)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monAbr�g� = 5
    mvarmonAbr�g� = vData
End Property


Public Property Get monAbr�g�() As String
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monAbr�g�
    monAbr�g� = mvarmonAbr�g�
End Property



Public Property Let monNom(ByVal vData As String)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monNom = 5
    mvarmonNom = vData
End Property


Public Property Get monNom() As String
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monNom
    monNom = mvarmonNom
End Property



