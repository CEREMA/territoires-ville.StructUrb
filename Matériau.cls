VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matériau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonNom As String 'copie locale
Private mvarmonAbrégé As String 'copie locale
Private mvarmaNorme As String 'copie locale
Private mvarmonCommentaire As String 'copie locale
Private mvarmonYoung As Single 'copie locale
Private mvarmonPoisson As Single 'copie locale
Private mvarmonEpsilon As Single 'copie locale
Private mvarmonSigma As Single 'copie locale
Private mvarmonAGel As Single 'copie locale
Private mvarmonBGel As Single 'copie locale
Private mvarmaQualité As String 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonUtilisationAutorisée As Boolean 'copie locale
Public Property Let monUtilisationAutorisée(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilisationAutorisée = 5
    mvarmonUtilisationAutorisée = vData
End Property


Public Property Get monUtilisationAutorisée() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilisationAutorisée
    monUtilisationAutorisée = mvarmonUtilisationAutorisée
End Property



Public Property Let maQualité(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maQualité = 5
    mvarmaQualité = vData
End Property


Public Property Get maQualité() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maQualité
    maQualité = mvarmaQualité
End Property



Public Property Let monBGel(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monBGel = 5
    mvarmonBGel = vData
End Property


Public Property Get monBGel() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monBGel
    monBGel = mvarmonBGel
End Property



Public Property Let monAGel(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monAGel = 5
    mvarmonAGel = vData
End Property


Public Property Get monAGel() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monAGel
    monAGel = mvarmonAGel
End Property




Public Sub SetPropsPhysic(unYoung As Single, unPoisson As Single, unEpsilon As Single, unSigma As Single)
    monYoung = unYoung
    monPoisson = unPoisson
    monEpsilon = unEpsilon
    monSigma = unSigma
End Sub

Public Sub SetProps(unNom As String, unAbrégé As String, uneNorme As String, unCommentaire As String)
    monNom = unNom
    monAbrégé = unAbrégé
    maNorme = uneNorme
    monCommentaire = unCommentaire
    monUtilisationAutorisée = True
End Sub

Public Property Let monSigma(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monSigma = 5
    mvarmonSigma = vData
End Property


Public Property Get monSigma() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monSigma
    monSigma = mvarmonSigma
End Property



Public Property Let monEpsilon(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monEpsilon = 5
    mvarmonEpsilon = vData
End Property


Public Property Get monEpsilon() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monEpsilon
    monEpsilon = mvarmonEpsilon
End Property



Public Property Let monPoisson(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monPoisson = 5
    mvarmonPoisson = vData
End Property


Public Property Get monPoisson() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monPoisson
    monPoisson = mvarmonPoisson
End Property



Public Property Let monYoung(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monYoung = 5
    mvarmonYoung = vData
End Property


Public Property Get monYoung() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monYoung
    monYoung = mvarmonYoung
End Property



Public Property Let monCommentaire(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monCommentaire = 5
    mvarmonCommentaire = vData
End Property


Public Property Get monCommentaire() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCommentaire
    monCommentaire = mvarmonCommentaire
End Property



Public Property Let maNorme(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maNorme = 5
    mvarmaNorme = vData
End Property


Public Property Get maNorme() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maNorme
    maNorme = mvarmaNorme
End Property



Public Property Let monAbrégé(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monAbrégé = 5
    mvarmonAbrégé = vData
End Property


Public Property Get monAbrégé() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monAbrégé
    monAbrégé = mvarmonAbrégé
End Property



Public Property Let monNom(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNom = 5
    mvarmonNom = vData
End Property


Public Property Get monNom() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNom
    monNom = mvarmonNom
End Property



