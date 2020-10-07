Attribute VB_Name = "Utilitaires"
'Fonction API windows pour cr�er des copies d'�cran
Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nheight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
'Constantes allant de apire avec cette fonction BitBlt
Const BLACKNESS = &H42
' Const CAPTUREBLT = ???
Const DSTINVERT = &H550009
Const MERGECOPY = &HC000CA
Const MERGEPAINT = &HBB0226
' Const NOMIRRORBITMAP = ???
Const NOTSRCCOPY = &H330008
Const NOTSRCERASE = &H1100A6
Const PATCOPY = &HF00021
Const PATINVERT = &H5A0049
Const PATPAINT = &HFB0A09
Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCERASE = &H440328
Const SRCINVERT = &H660046
Const SRCPAINT = &HEE0086
Const WHITENESS = &HFF0062
'API pour retrouver le Device context, propri�t� hDC sur certains controls
Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
'# D�claration pour copie �cran #
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
'#define HWND_BROADCAST  ((HWND)0xffff)
'Const HWND_BROADCAST = &HFFFF
Public Const WM_PASTE = &H302
Public Const KEYEVENTF_KEYUP = &H2
'Variable stockant le caract�re d�cimale utilis�
Public monCarDeci As String

Public Sub TrouverCaract�reD�cimalUtilis�()
    'Recherche du caract�re d�cimale utilis� . ou ,
    'et stockage dans la variable globale monCarDeci
    
    'If Format(CDbl("1,1")) = "1.1" Then
        'MsgBox "Caract�re d�cimal = le point"
    '    monCarDeci = "."
    'Else
        'MsgBox "Caract�re d�cimal = la virgule"
    '    monCarDeci = ","
    'End If
    
    If IsNumeric("1.1") = True Then
        'MsgBox "Caract�re d�cimal = le point"
        monCarDeci = "."
    Else
        'MsgBox "Caract�re d�cimal = la virgule"
        monCarDeci = ","
    End If
End Sub

Public Sub InitSpreadCaract�reD�cimal(unSpread As vaSpread, unCarDeci As String)
    'Initialisation du spread avec le caract�re d�cimal choisi
    'S�lection de tout le spread
    unSpread.BlockMode = True
    unSpread.Row = 1
    unSpread.Col = 1
    unSpread.Row2 = unSpread.MaxRows
    unSpread.Col2 = unSpread.MaxCols
    'utilisation du caract�re d�cimal choisi
    unSpread.TypeFloatDecimalChar = Asc(unCarDeci)
    unSpread.FloatDefDecimalChar = Asc(unCarDeci)
    'Remise du block mode � faux
    unSpread.BlockMode = False
End Sub

Public Function DonnerColPF(uneStruct As Structure, unTabInd As Integer) As Collection
    'R�cup�ration des tableaux d'�paisseur de plate-forme de
    'la structure courante
    
    Set DonnerColPF = Nothing
    
    If unTabInd = 1 Then
        Set DonnerColPF = uneStruct.mesInfoPF1Q1
    ElseIf unTabInd = 2 Then
        Set DonnerColPF = uneStruct.mesInfoPF1Q2
    ElseIf unTabInd = 3 Then
        Set DonnerColPF = uneStruct.mesInfoPF2Q1
    ElseIf unTabInd = 4 Then
        Set DonnerColPF = uneStruct.mesInfoPF2Q2
    ElseIf unTabInd = 5 Then
        Set DonnerColPF = uneStruct.mesInfoPF2PlusQ1
    ElseIf unTabInd = 6 Then
        Set DonnerColPF = uneStruct.mesInfoPF2PlusQ2
    ElseIf unTabInd = 7 Then
        Set DonnerColPF = uneStruct.mesInfoPF3Q1
    ElseIf unTabInd = 8 Then
        Set DonnerColPF = uneStruct.mesInfoPF3Q2
    Else
        MsgBox MsgErreurProg + MsgErreurCollectionInconnue + MsgIn + "Utilitaires:DonnerColPF", vbCritical
    End If
End Function

Public Sub RemplirQualit�Gel(unMat As Object, uneQualit� As String, unAGel As Single, unBGel As Single)
    If (TypeOf unMat Is MatSimple) Or (TypeOf unMat Is Mat�riau) Then
        unMat.maQualit� = uneQualit�
        unMat.monAGel = unAGel
        unMat.monBGel = unBGel
    ElseIf (TypeOf unMat Is MatCompos�) Then
        unMat.monAGel = unAGel
        unMat.monBGel = unBGel
    Else
        MsgBox MsgErreurProg + MsgErreurMat�riauInconnu + MsgIn + "Utilitaires:RemplirQualit�Gel", vbCritical
    End If
End Sub

Public Sub LireString(unFichId As Integer, uneString As String)
    'Lit la taille et le contenu d'une String � longueur variable
    'dans un fichier binaire unFichId o� elle a �t� �crite par la
    'fonction EcrireString qui �crit la taille puis le contenu
    'd'une string
    Dim uneLongString As Long
    
    'Lecture de la longueur de la chaine �crite par EcrireString
    Get #unFichId, , uneLongString
    'Initialisation de la chaine � lire avec le bon nombre de caract�res
    uneString = String(uneLongString, " ")
    'Lecture du contenu de la chaine
    Get #unFichId, , uneString
End Sub


Public Sub ViderCollection(uneCol As Collection)
    'Proc�dure vidant une collection
    For i = 1 To uneCol.Count
        uneCol.Remove 1
    Next i
End Sub

Public Function TrouverIndexInCol(uneCol As Collection, unAbregMatBF As String)
    'Proc�dure retournant l'index d'un mat�riau base/fondation
    'dans une collection perso ou CERTU � partir de son abr�g�
    For i = 1 To uneCol.Count
        If uneCol(i).monAbr�g� = unAbregMatBF Then
            TrouverIndexInCol = i
            Exit For
        End If
    Next i
End Function


Public Function DonnerSepMillier() As String
    'Fonction retournant le caract�re qui est
    'le s�parateur de millier en cours de windows
    DonnerSepMillier = Mid(Format("1234", "#,###"), 2, 1)
End Function

Public Function CorrigerNomFichier(unFileName As String) As String
    'Fonction retournant un nom de fichier corrig�
    'de double / par un seul
    Dim unePos As Integer, uneStringRes As String
    
    unePos = 1
    uneStringRes = unFileName
    
    Do
        unePos = InStr(1, uneStringRes, "\\")
        If unePos > 0 Then
            uneStringRes = Mid(uneStringRes, 1, unePos) + Mid(uneStringRes, unePos + 2)
        End If
    Loop While unePos > 0
    
    CorrigerNomFichier = uneStringRes
End Function
