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
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonNom As String 'copie locale
Private mvarmonAbr�g� As String 'copie locale
Private mvarmaNorme As String 'copie locale
Private mvarmonCommentaire As String 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmaQualit� As String 'copie locale
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




Public Sub SetProps(unNom As String, unAbr�g� As String, uneNorme As String, unCommentaire As String)
    monNom = unNom
    monAbr�g� = unAbr�g�
    maNorme = uneNorme
    monCommentaire = unCommentaire
End Sub

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



