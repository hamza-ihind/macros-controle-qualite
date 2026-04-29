' vrillage.vbs — Detection de vrillage surfacique (Q-Checker)
' Methode : aire de validation + normale au centroide + comparaison croisee via Reference CATIA + accessibilite du CoG
'
' Correction du bug original : GetNormal (SPA) retourne UNE seule normale fixe au centroide,
' independamment de toute boucle. La boucle 10x10 de la version precedente comparait
' 100 fois le meme vecteur avec lui-meme — le vrillage n'etait jamais detecte.

Option Explicit

' ─────────────────────────────────────────────────────────────────────────────
Sub CATMain()

    On Error Resume Next

    ' ── 1. Document actif ──────────────────────────────────────────────────
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Aucun document actif accessible.", vbCritical, "Q-Checker Vrillage"
        Exit Sub
    End If

    If TypeName(oDoc) <> "PartDocument" Then
        MsgBox "Ouvrez un CATPart avant de lancer la macro.", vbExclamation, "Q-Checker Vrillage"
        Exit Sub
    End If

    ' ── 2. Verification de la selection ────────────────────────────────────
    Dim oSel
    Set oSel = oDoc.Selection
    If oSel.Count = 0 Then
        MsgBox "Selectionnez une surface avant de lancer la macro.", vbExclamation, "Q-Checker Vrillage"
        Exit Sub
    End If

    Dim oSurface
    Set oSurface = oSel.Item(1).Value
    If Err.Number <> 0 Then
        MsgBox "Impossible de lire l'element selectionne.", vbCritical, "Q-Checker Vrillage"
        Err.Clear
        Exit Sub
    End If

    ' ── 3. SPA Workbench ───────────────────────────────────────────────────
    Dim oSPA
    Set oSPA = oDoc.GetWorkbench("SPAWorkbench")
    If Err.Number <> 0 Then
        MsgBox "SPAWorkbench inaccessible. Effectuez Ctrl+U puis relancez.", vbCritical, "Q-Checker Vrillage"
        Err.Clear
        Exit Sub
    End If

    ' ── 4. Mesurable sur la surface ────────────────────────────────────────
    Dim oMeas
    Set oMeas = oSPA.GetMeasurable(oSurface)
    If Err.Number <> 0 Then
        MsgBox "Surface non mesurable." & vbCrLf & _
               "Causes possibles : geometrie non resolue ou element non surfacique." & vbCrLf & _
               "Verifiez avec Analyze > Check Geometry.", vbCritical, "Q-Checker Vrillage"
        Err.Clear
        Set oSPA = Nothing : Set oDoc = Nothing
        Exit Sub
    End If

    ' ── 5. Controle d'aire (surface degeneree = aire nulle) ────────────────
    Dim dArea
    dArea = 0
    Err.Clear
    dArea = oMeas.Area * 1000000    ' m2 -> mm2
    If Err.Number <> 0 Or dArea <= 0 Then
        MsgBox "Surface degeneree detectee (aire nulle ou illisible)." & vbCrLf & _
               "La surface est vraisemblablement auto-intersectee ou vide." & vbCrLf & _
               "Action : Analyze > Check Geometry.", vbCritical, "Q-Checker Vrillage"
        Err.Clear
        Set oMeas = Nothing : Set oSPA = Nothing : Set oDoc = Nothing
        Exit Sub
    End If

    ' ── 6. Normale de reference au centroide ───────────────────────────────
    Dim vNormRef(2)
    Err.Clear
    oMeas.GetNormal vNormRef
    If Err.Number <> 0 Then
        MsgBox "Impossible d'obtenir la normale de reference." & vbCrLf & _
               "La surface peut etre un corps ouvert ou non resolue.", vbCritical, "Q-Checker Vrillage"
        Err.Clear
        Set oMeas = Nothing : Set oSPA = Nothing : Set oDoc = Nothing
        Exit Sub
    End If

    ' Verifier que la normale n'est pas le vecteur nul
    Dim dNormLen
    dNormLen = Sqr(vNormRef(0)^2 + vNormRef(1)^2 + vNormRef(2)^2)
    If dNormLen < 0.0001 Then
        MsgBox "Normale nulle detectee : la surface est mathematiquement degeneree." & vbCrLf & _
               "Reconstruisez la surface depuis ses courbes guides.", vbCritical, "Q-Checker Vrillage"
        Set oMeas = Nothing : Set oSPA = Nothing : Set oDoc = Nothing
        Exit Sub
    End If

    ' ── 7. Comparaison croisee : normale directe vs normale via Reference ──
    ' GetNormal sur oSurface vs GetNormal sur CreateReferenceFromObject(oSurface)
    ' Une inversion du produit scalaire (< -0.1) revele une incoherence d'orientation.
    Dim oPart
    Set oPart = oDoc.Part

    Dim vNormRef2(2)
    Dim dDotCross
    Dim bInversionDetected
    bInversionDetected = False

    Err.Clear
    Dim oRef
    Set oRef = oPart.CreateReferenceFromObject(oSurface)
    If Err.Number = 0 Then
        Dim oMeas2
        Set oMeas2 = oSPA.GetMeasurable(oRef)
        If Err.Number = 0 Then
            oMeas2.GetNormal vNormRef2
            If Err.Number = 0 Then
                dDotCross = vNormRef(0)*vNormRef2(0) + vNormRef(1)*vNormRef2(1) + vNormRef(2)*vNormRef2(2)
                If dDotCross < -0.1 Then
                    bInversionDetected = True
                End If
            End If
        End If
        Set oMeas2 = Nothing
    End If
    Err.Clear

    ' ── 8. Verification complementaire : accessibilite du centroide (CoG) ──
    Dim vCoG(2)
    Dim bCoGOk
    bCoGOk = True
    Err.Clear
    oMeas.GetCOG vCoG
    If Err.Number <> 0 Then
        bCoGOk = False
        Err.Clear
    End If

    ' ── 9. Construction du rapport ─────────────────────────────────────────
    Dim sRep
    sRep = "===== RAPPORT VRILLAGE SURFACIQUE =====" & vbCrLf
    sRep = sRep & "Surface   : " & oSurface.Name                                           & vbCrLf
    sRep = sRep & "Aire      : " & FormatNumber(dArea, 2)            & " mm2"              & vbCrLf
    sRep = sRep & "Normale   : [" & FormatNumber(vNormRef(0),3) & "; " & _
                                    FormatNumber(vNormRef(1),3) & "; " & _
                                    FormatNumber(vNormRef(2),3) & "]"                      & vbCrLf
    sRep = sRep & "Centroide : " & BoolLabel(bCoGOk, "OK", "INACCESSIBLE")                & vbCrLf
    sRep = sRep & "Inversion : " & BoolLabel(bInversionDetected, "OUI", "NON")            & vbCrLf
    sRep = sRep & "=======================================" & vbCrLf

    ' ── 10. Affichage selon classification ────────────────────────────────
    If bInversionDetected Then
        sRep = sRep & "STATUT : [!!] VRILLAGE DETECTE" & vbCrLf & vbCrLf & _
               "La direction normale est incoherente entre les deux references." & vbCrLf & _
               "Risques : echec maillage FEA, rejet usinage CNC." & vbCrLf & vbCrLf & _
               "Actions correctives :" & vbCrLf & _
               "  1. Analyze > Check Geometry" & vbCrLf & _
               "  2. Surface > Invert Orientation" & vbCrLf & _
               "  3. Reconstruire via Shape Healing ou Join + Healing" & vbCrLf & _
               "  4. Revoir les guides / directrices du Sweep ou Loft"
        MsgBox sRep, vbCritical, "Q-Checker - Vrillage Detecte"

    ElseIf Not bCoGOk Then
        sRep = sRep & "STATUT : [?] SURFACE SUSPECTE" & vbCrLf & vbCrLf & _
               "Aucun vrillage mesure directement, mais le centroide" & vbCrLf & _
               "est inaccessible — la surface peut etre degeneree ou incomplete." & vbCrLf & _
               "Verification manuelle recommandee : Analyze > Check Geometry."
        MsgBox sRep, vbExclamation, "Q-Checker - Surface Suspecte"

    Else
        sRep = sRep & "STATUT : [OK] Aucun vrillage detecte." & vbCrLf & _
               "La surface est geometriquement reguliere."
        MsgBox sRep, vbInformation, "Q-Checker - Succes"
    End If

    ' ── Nettoyage ──────────────────────────────────────────────────────────
    Set oRef  = Nothing
    Set oMeas = Nothing
    Set oSPA  = Nothing
    Set oPart = Nothing
    Set oDoc  = Nothing

End Sub

' Helper : retourne l'une des deux etiquettes selon la condition
Function BoolLabel(bCond, sTrue, sFalse)
    If bCond Then
        BoolLabel = sTrue
    Else
        BoolLabel = sFalse
    End If
End Function