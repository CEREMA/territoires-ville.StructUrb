VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatCompos�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonAbr�g� As String 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonNom As String 'copie locale
Private mvarmesCompositions As Collection 'copie locale
Private mvarmonAGel As Single 'copie locale
Private mvarmonBGel As Single 'copie locale
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



Public Property Set mesCompositions(ByVal vData As Object)
'utilis� lors de l'affectation d'un Object � la propri�t�, du cot� gauche de l'instruction Set
'Syntax: Set x.mesCompositions = Form1
    Set mvarmesCompositions = vData
End Property


Public Property Get mesCompositions() As Collection
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.mesCompositions
    Set mesCompositions = mvarmesCompositions
End Property

Public Sub SetProps(unNom As String, unAbr�g� As String)
    monNom = unNom
    monAbr�g� = unAbr�g�
End Sub

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

