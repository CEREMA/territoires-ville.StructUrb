VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmonNom As String 'copie locale
Private mvarmonAbrégé As String 'copie locale
Private mvarmaNorme As String 'copie locale
Private mvarmonCommentaire As String 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaQualité As String 'copie locale
Private mvarmonAGel As Single 'copie locale
Private mvarmonBGel As Single 'copie locale
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




Public Sub SetProps(unNom As String, unAbrégé As String, uneNorme As String, unCommentaire As String)
    monNom = unNom
    monAbrégé = unAbrégé
    maNorme = uneNorme
    monCommentaire = unCommentaire
End Sub

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



