Attribute VB_Name = "ModuleSaisie"

Public Sub VerifSaisieEntierPositif(KeyAscii As Integer, unControl As Control, uneValeurDefaut As String)
    Dim uneChaineTmp
    
    If KeyAscii = 27 Or KeyAscii = 13 Then
        'Cas de la frappe des touches Echap ou Retour Chariot
        Exit Sub
    End If
    
    uneChaineTmp = " " + unControl.Text 'car la fonction Str rajoute un blanc pour les valeurs > 0
    If unControl.Text = "" Or Mid(unControl.Text, 1, 1) = "0" Then
        'Cas o� la zone de saisie est vide, on remet la valeur par d�faut
        unControl.Text = uneValeurDefaut
    ElseIf uneChaineTmp <> Str(Val(unControl.Text)) Or IsNumeric(unControl.Text) = False Then
        MsgBox MsgSaisieEntierPositif, vbCritical
        unControl.Text = uneValeurDefaut
    End If
End Sub

Public Function SaisieEntierPositifEntreMinMax(KeyCode As Integer, unControl As Control, uneValeurDefaut As String, unIntMin As Integer, unIntMax As Integer, uneString As String) As Boolean
    SaisieEntierPositifEntreMinMax = False
    Call VerifSaisieEntierPositif(KeyCode, unControl, uneValeurDefaut)
    If Val(unControl.Text) < unIntMin Or Val(unControl.Text) > unIntMax Then
        unMsg = uneString + " " + MsgSupA + Format(unIntMin)
        unMsg = unMsg + " " + MsgEtInfA + Format(unIntMax)
        MsgBox unMsg, vbCritical
        unControl.Text = uneValeurDefaut
   Else
        SaisieEntierPositifEntreMinMax = True
    End If
End Function

Public Sub VerifierSortieSaisieEntierPositif(unControl As Control, unIntMin As Integer, unIntMax As Integer)
    Dim uneValInt As Integer
    
    'Cas o� on n'a pas cliqu� sur le bouton Annuler
    'pour une boite de dialogue
    If unControl.Text = "" Or IsNumeric(unControl.Text) = False Then
        uneValInt = -1
    Else
        uneValInt = Val(unControl.Text)
    End If
    
    If uneValInt < unIntMin Or uneValInt > unIntMax Then
        MsgBox MsgDur�eService + MsgSupA + Format(unIntMin) + " " + MsgEtInfA + Format(unIntMax), vbCritical
        unControl.SetFocus
    End If
End Sub
    
Public Function VerifierMinMaxTraficIni(uneForm As Form, unTextTrafIni As String) As Boolean
    Dim uneValMin As Long, uneValMax As Long
    Dim uneValMinTol As Long, uneValMaxTol As Long
    Dim unNomTypeVoie As String, unTypeVoie As Integer, unMsg As String
    
    'V�rification qu'on est entre le min et le max permis
    If unTextTrafIni = "" Then
        'Aucune saisie, on met � vide aussi le trafic cumul�
        uneForm.TextTrafCUM.Text = ""
        VerifierMinMaxTraficIni = True
        Exit Function
    End If
    
    'R�cup du domaine de validit� suivant le type de voie
    DonnerMinMaxTraficIni uneForm, uneValMinTol, uneValMin, uneValMax, uneValMaxTol
    unMsg = ""
    unNomTypeVoie = DonnerNomTypeVoie(uneForm)
    
    If CLng(unTextTrafIni) < uneValMinTol Or CLng(unTextTrafIni) > uneValMaxTol Then
        'Cas d'erreur non admise dans le domaine de validit�
        unMsg = MsgTraficIni + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
        'Affichage du domaine de tol�rance plus grande que le domaine de validit�
        unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
        unTypeIcone = vbCritical
        'On remet dans le bon onglet et la bonne zone de saisie
        uneForm.TabData.Tab = OngletTrafic
        uneForm.TextTrafIni.SetFocus
        VerifierMinMaxTraficIni = False
    ElseIf (CLng(unTextTrafIni) >= uneValMinTol And CLng(unTextTrafIni) < uneValMin) Or (CLng(unTextTrafIni) > uneValMax And CLng(unTextTrafIni) <= uneValMaxTol) Then
        'Cas d'erreur tol�r�e dans le domaine de validit�
        unMsg = MsgTraficIni + UCase(unNomTypeVoie) + " " + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax)
        unMsg = unMsg + Chr(13) + Chr(13) + MsgValTol + Format(uneValMinTol) + " " + MsgAnd + Format(uneValMaxTol) + MsgIsTol
        unTypeVoie = DonnerTypeVoie(uneForm)
        If CLng(unTextTrafIni) > uneValMax Then
            'Cas o� on d�passe la valeur maxi
            If (unTypeVoie >= TypeVoieTraficLourd And unTypeVoie <= TypeVoieBus) Or (unTypeVoie = TypeGiratoireTraficLourd) Then
                'Cas des voie bus, voie principale avec PL
                'et Giratoire sur voie principale PL
                'Si la valeur > au max et < max tol�r�e, conseiller de faire
                'une �tude en laboratoire
                unMsg = unMsg + Chr(13) + Chr(13) + MsgValLabo + Format(uneValMax) + MsgIsLabo
            End If
        End If
        'R�sultat OK
        unTypeIcone = vbInformation
        VerifierMinMaxTraficIni = True
    Else
        'Cas o� on se trouve dans le domaine de validit�
        VerifierMinMaxTraficIni = True
    End If
    
    If unMsg <> "" Then MsgBox unMsg, unTypeIcone
End Function

Public Function VerifierMinMaxDur�eService(uneForm As Form) As Boolean
    Dim uneValMin As Integer, uneValMax As Integer
    
    uneValMin = 5
    uneValMax = 50
    
    If uneForm.TextDur�eS.Text = "" Then
        MsgBox MsgDur�eService + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax), vbCritical
        'On remet dans le bon onglet et la bonne zone de saisie
        uneForm.TabData.Tab = OngletTrafic
        uneForm.TextDur�eS.SetFocus
        VerifierMinMaxDur�eService = False
    ElseIf CInt(uneForm.TextDur�eS.Text) > uneValMax Or CInt(uneForm.TextDur�eS.Text) < uneValMin Then
        MsgBox MsgDur�eService + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax), vbCritical
        'On remet dans le bon onglet et la bonne zone de saisie
        uneForm.TabData.Tab = OngletTrafic
        uneForm.TextDur�eS.SetFocus
        VerifierMinMaxDur�eService = False
    Else
        VerifierMinMaxDur�eService = True
    End If
End Function

Public Function VerifierMinMaxCAM(uneForm As Form, unTextCAM As String) As Boolean
    Dim uneValMin As Single, uneValMax As Single, uneValPrec As Single
    Dim unMsg As String
    
    'V�rification qu'on est entre le min et le max permis
    
    'R�cup du domaine de validit� suivant le type de voie
    unMsg = DonnerPrecMinMaxCAM(uneForm, uneValPrec, uneValMin, uneValMax)
    
    If CSng(unTextCAM) > uneValMax Or CSng(unTextCAM) < uneValMin Then
        unMsg = "Pour une " + unMsg + " :" + Chr(13) + Chr(13)
        MsgBox unMsg + MsgCoefCAM + MsgSupA + Format(uneValMin) + " " + MsgEtInfA + Format(uneValMax), vbCritical
        'On remet dans le bon onglet et la bonne zone de saisie
        uneForm.TabData.Tab = OngletCAM
        uneForm.MaskCAM.SetFocus
        VerifierMinMaxCAM = False
    Else
        VerifierMinMaxCAM = True
    End If
End Function

Public Sub VerifierSaisieEntier(uneTextBox As TextBox)
    'Verification de saisie d'un entier positif ou nul dans une textbox
    'A utiliser de pr�f�rence dans un change event
    If uneTextBox.Text = "" Or Format(Val(uneTextBox.Text)) = uneTextBox.Text Then
        'Cas o� le text est un entier positif ou nul
        uneTextBox.ForeColor = QBColor(12)
    Else
        If Mid(uneTextBox.Text, 1, 1) <> "0" Then MsgBox MsgSaisieEntierPositif, vbExclamation
        'Affichage jusqu'au dernier caract�re valide = chiffre)
        uneTextBox.Text = Format(Val(uneTextBox.Text))
        'Mise du curseur en fin de texte
        uneTextBox.SelStart = Len(uneTextBox.Text)
    End If
End Sub

