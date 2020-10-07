VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatComposé"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonAbrégé As String 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonNom As String 'copie locale
Private mvarmesCompositions As Collection 'copie locale
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



Public Property Set mesCompositions(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesCompositions = Form1
    Set mvarmesCompositions = vData
End Property


Public Property Get mesCompositions() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesCompositions
    Set mesCompositions = mvarmesCompositions
End Property

Public Sub SetProps(unNom As String, unAbrégé As String)
    monNom = unNom
    monAbrégé = unAbrégé
End Sub

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

