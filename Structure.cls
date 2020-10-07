VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmonUtilVDes As Integer 'copie locale
Private mvarmonUtilVDis As Integer 'copie locale
Private mvarmonUtilVPL As Integer 'copie locale
Private mvarmonUtilVBus As Integer 'copie locale
Private mvarmonComment As String 'copie locale
Private mvarmonTauxRisque As Integer 'copie locale
Private mvarmonNbEssieuxMin As Long 'copie locale
Private mvarmonNbEssieuxMax As Long 'copie locale
Private mvarmonTypeCAM As Integer 'copie locale
Private mvarmaCoucheSurface As String 'copie locale
Private mvarmaCoucheBase As String 'copie locale
Private mvarmaCoucheFondation As String 'copie locale
Private mvarmesInfoPF1Q1 As Collection 'copie locale
Private mvarmesInfoPF1Q2 As Collection 'copie locale
Private mvarmesInfoPF2Q1 As Collection 'copie locale
Private mvarmesInfoPF2Q2 As Collection 'copie locale
Private mvarmesInfoPF3Q1 As Collection 'copie locale
Private mvarmesInfoPF3Q2 As Collection 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonAbrégé As String 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaCoucheSurfSansEp As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaSaisieComplète As Boolean 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonNumIndex As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonUtilVParking As Integer 'copie locale
Private mvarmonUtilGDis As Integer 'copie locale
Private mvarmonUtilGPL As Integer 'copie locale
Private mvarmesInfoPF2PlusQ2 As Collection 'copie locale
Private mvarmesInfoPF2PlusQ1 As Collection 'copie locale
Private mvarmonTypeChaussee As Byte 'copie locale
Private mvarmonTypeStructure As Byte 'copie locale
Public Property Let monTypeStructure(ByVal vData As Byte)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTypeStructure = 5
    mvarmonTypeStructure = vData
End Property


Public Property Get monTypeStructure() As Byte
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTypeStructure
    monTypeStructure = mvarmonTypeStructure
End Property



Public Property Let monTypeChaussee(ByVal vData As Byte)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTypeChaussee = 5
    mvarmonTypeChaussee = vData
End Property


Public Property Get monTypeChaussee() As Byte
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTypeChaussee
    monTypeChaussee = mvarmonTypeChaussee
End Property



Public Property Set mesInfoPF2PlusQ1(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF2PlusQ1 = Form1
    Set mvarmesInfoPF2PlusQ1 = vData
End Property


Public Property Get mesInfoPF2PlusQ1() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF2PlusQ1
    Set mesInfoPF2PlusQ1 = mvarmesInfoPF2PlusQ1
End Property



Public Property Set mesInfoPF2PlusQ2(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF2PlusQ2 = Form1
    Set mvarmesInfoPF2PlusQ2 = vData
End Property


Public Property Get mesInfoPF2PlusQ2() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF2PlusQ2
    Set mesInfoPF2PlusQ2 = mvarmesInfoPF2PlusQ2
End Property



Public Property Let monUtilGPL(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilGPL = 5
    mvarmonUtilGPL = vData
End Property


Public Property Get monUtilGPL() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilGPL
    monUtilGPL = mvarmonUtilGPL
End Property



Public Property Let monUtilGDis(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilGDis = 5
    mvarmonUtilGDis = vData
End Property


Public Property Get monUtilGDis() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilGDis
    monUtilGDis = mvarmonUtilGDis
End Property



Public Property Let monUtilVParking(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilVParking = 5
    mvarmonUtilVParking = vData
End Property


Public Property Get monUtilVParking() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilVParking
    monUtilVParking = mvarmonUtilVParking
End Property



Public Property Let monNumIndex(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNumIndex = 5
    mvarmonNumIndex = vData
End Property


Public Property Get monNumIndex() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNumIndex
    monNumIndex = mvarmonNumIndex
End Property



Public Property Let maSaisieComplète(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maSaisieComplète = 5
    mvarmaSaisieComplète = vData
End Property


Public Property Get maSaisieComplète() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maSaisieComplète
    maSaisieComplète = mvarmaSaisieComplète
End Property



Public Property Let maCoucheSurfSansEp(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoucheSurfSansEp = 5
    mvarmaCoucheSurfSansEp = vData
End Property


Public Property Get maCoucheSurfSansEp() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoucheSurfSansEp
    maCoucheSurfSansEp = mvarmaCoucheSurfSansEp
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



Public Property Set mesInfoPF3Q2(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF3Q2 = Form1
    Set mvarmesInfoPF3Q2 = vData
End Property


Public Property Get mesInfoPF3Q2() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF3Q2
    Set mesInfoPF3Q2 = mvarmesInfoPF3Q2
End Property



Public Property Set mesInfoPF3Q1(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF3Q1 = Form1
    Set mvarmesInfoPF3Q1 = vData
End Property


Public Property Get mesInfoPF3Q1() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF3Q1
    Set mesInfoPF3Q1 = mvarmesInfoPF3Q1
End Property



Public Property Set mesInfoPF2Q2(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF2Q2 = Form1
    Set mvarmesInfoPF2Q2 = vData
End Property


Public Property Get mesInfoPF2Q2() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF2Q2
    Set mesInfoPF2Q2 = mvarmesInfoPF2Q2
End Property



Public Property Set mesInfoPF2Q1(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF2Q1 = Form1
    Set mvarmesInfoPF2Q1 = vData
End Property


Public Property Get mesInfoPF2Q1() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF2Q1
    Set mesInfoPF2Q1 = mvarmesInfoPF2Q1
End Property



Public Property Set mesInfoPF1Q2(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF1Q2 = Form1
    Set mvarmesInfoPF1Q2 = vData
End Property


Public Property Get mesInfoPF1Q2() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF1Q2
    Set mesInfoPF1Q2 = mvarmesInfoPF1Q2
End Property



Public Property Set mesInfoPF1Q1(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesInfoPF1Q1 = Form1
    Set mvarmesInfoPF1Q1 = vData
End Property


Public Property Get mesInfoPF1Q1() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesInfoPF1Q1
    Set mesInfoPF1Q1 = mvarmesInfoPF1Q1
End Property



Public Property Let maCoucheFondation(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoucheFondation = 5
    mvarmaCoucheFondation = vData
End Property


Public Property Get maCoucheFondation() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoucheFondation
    maCoucheFondation = mvarmaCoucheFondation
End Property



Public Property Let maCoucheBase(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoucheBase = 5
    mvarmaCoucheBase = vData
End Property


Public Property Get maCoucheBase() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoucheBase
    maCoucheBase = mvarmaCoucheBase
End Property



Public Property Let maCoucheSurface(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoucheSurface = 5
    mvarmaCoucheSurface = vData
End Property


Public Property Get maCoucheSurface() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoucheSurface
    maCoucheSurface = mvarmaCoucheSurface
End Property



Public Property Let monTypeCAM(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTypeCAM = 5
    mvarmonTypeCAM = vData
End Property


Public Property Get monTypeCAM() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTypeCAM
    monTypeCAM = mvarmonTypeCAM
End Property



Public Property Let monNbEssieuxMax(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNbEssieuxMax = 5
    mvarmonNbEssieuxMax = vData
End Property


Public Property Get monNbEssieuxMax() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNbEssieuxMax
    monNbEssieuxMax = mvarmonNbEssieuxMax
End Property



Public Property Let monNbEssieuxMin(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNbEssieuxMin = 5
    mvarmonNbEssieuxMin = vData
End Property


Public Property Get monNbEssieuxMin() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNbEssieuxMin
    monNbEssieuxMin = mvarmonNbEssieuxMin
End Property



Public Property Let monTauxRisque(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTauxRisque = 5
    mvarmonTauxRisque = vData
End Property


Public Property Get monTauxRisque() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTauxRisque
    monTauxRisque = mvarmonTauxRisque
End Property



Public Property Let monComment(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monComment = 5
    mvarmonComment = vData
End Property


Public Property Get monComment() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monComment
    monComment = mvarmonComment
End Property



Public Property Let monUtilVBus(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilVBus = 5
    mvarmonUtilVBus = vData
End Property


Public Property Get monUtilVBus() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilVBus
    monUtilVBus = mvarmonUtilVBus
End Property



Public Property Let monUtilVPL(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilVPL = 5
    mvarmonUtilVPL = vData
End Property


Public Property Get monUtilVPL() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilVPL
    monUtilVPL = mvarmonUtilVPL
End Property



Public Property Let monUtilVDis(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilVDis = 5
    mvarmonUtilVDis = vData
End Property


Public Property Get monUtilVDis() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilVDis
    monUtilVDis = mvarmonUtilVDis
End Property



Public Property Let monUtilVDes(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monUtilVDes = 5
    mvarmonUtilVDes = vData
End Property


Public Property Get monUtilVDes() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monUtilVDes
    monUtilVDes = mvarmonUtilVDes
End Property


Public Sub SetPropsInfo(uneSaisieComplète As Boolean, unUtilVDes As Integer, unUtilVDis As Integer, unUtilVPL As Integer, unUtilVBus As Integer, unComment As String, unTauxRisque As Integer, unNbEssMin As Long, unNbEssMax As Long, unTypeCAM As Integer)
    'Remplissage des attributs correspondant au contenu
    'de l'onglet Informations de la fiche Structure
    monUtilVDes = unUtilVDes
    monUtilVDis = unUtilVDis
    monUtilVPL = unUtilVPL
    monUtilVBus = unUtilVBus
    monComment = unComment
    monTauxRisque = unTauxRisque
    monNbEssieuxMin = unNbEssMin
    monNbEssieuxMax = unNbEssMax
    monTypeCAM = unTypeCAM
    'Indication que toutes les données de la structure ont été saisies
    maSaisieComplète = uneSaisieComplète
End Sub

Public Sub SetPropsTabPF(uneCoucheSurf As String, uneCoucheBase As String, uneCoucheFond As String, uneColPF1Q1 As Collection, uneColPF1Q2 As Collection, uneColPF2Q1 As Collection, uneColPF2Q2 As Collection, uneColPF3Q1 As Collection, uneColPF3Q2 As Collection)
    'Remplissage des attributs correspondant au contenu
    'de l'onglet Tableaux et graphiques de la fiche Structure
    maCoucheSurface = uneCoucheSurf
    maCoucheBase = uneCoucheBase
    maCoucheFondation = uneCoucheFond
    Set mesInfoPF1Q1 = uneColPF1Q1
    Set mesInfoPF1Q2 = uneColPF1Q2
    Set mesInfoPF2Q1 = uneColPF2Q1
    Set mesInfoPF2Q2 = uneColPF2Q2
    Set mesInfoPF3Q1 = uneColPF3Q1
    Set mesInfoPF3Q2 = uneColPF3Q2
End Sub

Public Sub SetComposition(uneCoucheSurf As String, uneCoucheBase As String, uneCoucheFond As String, uneCoucheSurfSansEp As Integer)
    'Remplissage des attributs correspondant à la
    'composition de la structure dans la fiche Structure
    maCoucheSurface = uneCoucheSurf
    maCoucheBase = uneCoucheBase
    maCoucheFondation = uneCoucheFond
    maCoucheSurfSansEp = uneCoucheSurfSansEp
End Sub

Private Sub Class_Initialize()
    'Allocation des tableaux d'épaisseurs de plateforme
    Set mesInfoPF1Q1 = New Collection
    Set mesInfoPF1Q2 = New Collection
    Set mesInfoPF2Q1 = New Collection
    Set mesInfoPF2Q2 = New Collection
    Set mesInfoPF3Q1 = New Collection
    Set mesInfoPF3Q2 = New Collection
    'Ajout en version 2
    Set mesInfoPF2PlusQ1 = New Collection
    Set mesInfoPF2PlusQ2 = New Collection
End Sub

Private Sub Class_Terminate()
    ViderMémoire
    Set mesInfoPF1Q1 = Nothing
    Set mesInfoPF1Q2 = Nothing
    Set mesInfoPF2Q1 = Nothing
    Set mesInfoPF2Q2 = Nothing
    Set mesInfoPF3Q1 = Nothing
    Set mesInfoPF3Q2 = Nothing
    'Ajout en version 2
    Set mesInfoPF2PlusQ1 = Nothing
    Set mesInfoPF2PlusQ2 = Nothing
End Sub

Public Sub ViderMémoire()
    'Libération mémoire des tableaux d'épaisseurs de plateforme
    For i = 1 To mesInfoPF1Q1.Count
        mesInfoPF1Q1.Remove 1
    Next i
    
    For i = 1 To mesInfoPF1Q2.Count
        mesInfoPF1Q2.Remove 1
    Next i
    
    For i = 1 To mesInfoPF2Q1.Count
        mesInfoPF2Q1.Remove 1
    Next i
    
    For i = 1 To mesInfoPF2Q2.Count
        mesInfoPF2Q2.Remove 1
    Next i
    
    For i = 1 To mesInfoPF3Q1.Count
        mesInfoPF3Q1.Remove 1
    Next i
    
    For i = 1 To mesInfoPF3Q2.Count
        mesInfoPF3Q2.Remove 1
    Next i
End Sub

Public Sub SetPropsInfoV2(unTypeChaussee As Byte, unTypeStructure As Byte, unUtilVParking As Integer, unUtilGDis As Integer, unUtilGPL As Integer)
    'Remplissage des attributs correspondant au contenu
    'de l'onglet Informations de la fiche Structure rajoutés en version 2
    monUtilGDis = unUtilGDis
    monUtilGPL = unUtilGPL
    monTypeChaussee = unTypeChaussee
    monTypeStructure = unTypeStructure
    monUtilVParking = unUtilVParking
End Sub
