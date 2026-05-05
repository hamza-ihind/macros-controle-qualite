' ============================================================
'  DetectSliverFaces.vbs
'  Detecte les Sliver Faces dans le Part actif
'  Definition : faces ultra-fines (aire < seuil) creees par
'  quasi-coincidence geometrique lors de modelisation ou import
'  Auteur  : PFE Macro — CATIA V5
'  Usage   : Tools > Macro > Macros > Run
' ============================================================

Sub CATMain()

    ' ---- Seuil de detection (en mm2) -----------------------
    ' Modifier cette valeur selon le contexte du projet :
    '   0.001 mm2  -> pieces aeronautiques (haute precision)
    '   0.01  mm2  -> pieces automobiles standard
    '   0.1   mm2  -> modeles pour impression 3D
    '   1.0   mm2  -> import IGES/STEP (seuil large)
    Dim SEUIL_MM2
    SEUIL_MM2 = 0.01

    ' ---- Declaration des variables -------------------------
    Dim oDoc
    Dim oPart
    Dim oSPA
    Dim oHybridBodies
    Dim oHybridBody
    Dim oHybridShapes
    Dim oShape
    Dim oRef
    Dim oMeasure
    Dim dAire_m2
    Dim dAire_mm2
    Dim iNbSliver
    Dim iNbTotal
    Dim sListeSliver
    Dim sMsg
    Dim iIcon
    Dim i
    Dim j

    ' ---- Gestion des erreurs globales ----------------------
    On Error GoTo ErrHandler

    ' ---- 1. Recuperer le document actif --------------------
    Set oDoc = CATIA.ActiveDocument

    ' ---- 2. Verifier que c'est un PartDocument -------------
    If TypeName(oDoc) <> "PartDocument" Then
        MsgBox "ERREUR : Le document actif n'est pas un fichier .CATPart." & Chr(13) & Chr(13) & _
               "Veuillez ouvrir un fichier .CATPart avant de lancer la macro.", _
               16, "Detection Sliver Faces — PFE"
        Exit Sub
    End If

    ' ---- 3. Acceder a l'objet Part -------------------------
    Set oPart = oDoc.Part

    ' ---- 4. Verifier qu'il y a des HybridBodies ------------
    Set oHybridBodies = oPart.HybridBodies

    If oHybridBodies.Count = 0 Then
        MsgBox "AVERTISSEMENT : Aucun corps surfacique (HybridBody) detecte." & Chr(13) & Chr(13) & _
               "Le Part ne contient pas de surfaces GSD a analyser.", _
               48, "Detection Sliver Faces — PFE"
        GoTo Nettoyage
    End If

    ' ---- 5. Initialisation des compteurs -------------------
    iNbSliver    = 0
    iNbTotal     = 0
    sListeSliver = ""

    ' ---- 6. Instancier le SPAWorkbench ---------------------
    Set oSPA = oDoc.GetWorkbench("SPAWorkbench")

    ' ---- 7. Boucle sur tous les HybridBodies ---------------
    For i = 1 To oHybridBodies.Count

        Set oHybridBody   = oHybridBodies.Item(i)
        Set oHybridShapes = oHybridBody.HybridShapes

        ' ---- 8. Boucle sur chaque surface du HybridBody ----
        For j = 1 To oHybridShapes.Count

            Set oShape = oHybridShapes.Item(j)
            iNbTotal   = iNbTotal + 1

            ' ---- 9. Tenter la mesure (On Error Resume Next)
            On Error Resume Next
            Set oRef     = oPart.CreateReferenceFromObject(oShape)
            Set oMeasure = oSPA.GetMeasurable(oRef)

            ' Recuperer l'aire brute en m2
            dAire_m2  = 0
            dAire_m2  = oMeasure.Area

            On Error GoTo ErrHandler

            ' ---- 10. Convertir m2 -> mm2 (x 10^6) ----------
            dAire_mm2 = dAire_m2 * 1000000

            ' ---- 11. Comparer au seuil ---------------------
            If dAire_mm2 > 0 And dAire_mm2 < SEUIL_MM2 Then

                iNbSliver = iNbSliver + 1

                ' Stocker le nom + aire + body parent
                sListeSliver = sListeSliver & _
                    "  [" & iNbSliver & "] " & oShape.Name & _
                    " (HybridBody: " & oHybridBody.Name & ")" & _
                    " — Aire = " & FormatNumber(dAire_mm2, 6) & " mm2" & Chr(13)

            End If

            ' ---- Nettoyage des objets de mesure ------------
            Set oMeasure = Nothing
            Set oRef     = Nothing
            Set oShape   = Nothing

        Next j

        Set oHybridShapes = Nothing
        Set oHybridBody   = Nothing

    Next i

    ' ---- 12. Construire le rapport -------------------------
    sMsg = "======= RAPPORT — SLIVER FACES =======" & Chr(13)
    sMsg = sMsg & "Part          : " & oPart.Name               & Chr(13)
    sMsg = sMsg & "Surfaces anal.: " & iNbTotal                 & Chr(13)
    sMsg = sMsg & "Seuil utilise : " & SEUIL_MM2 & " mm2"      & Chr(13)
    sMsg = sMsg & "Sliver faces  : " & iNbSliver                & Chr(13)
    sMsg = sMsg & "======================================" & Chr(13)

    ' ---- 13. Verdict final ---------------------------------
    If iNbSliver = 0 Then
        sMsg  = sMsg & "[OK]  Aucune Sliver Face detectee." & Chr(13) & _
                "      Toutes les surfaces sont au-dessus du seuil."
        iIcon = 64
    Else
        sMsg  = sMsg & "[!!]  " & iNbSliver & " Sliver Face(s) detectee(s) :" & Chr(13) & Chr(13) & _
                sListeSliver & Chr(13) & _
                "Action recommandee :" & Chr(13) & _
                "  - Insert > Operations > Healing" & Chr(13) & _
                "  - Ou reconstruire la zone avec Fill / Blend" & Chr(13) & _
                "  - Ou augmenter la tolerance du Join"
        iIcon = 48
    End If

    ' ---- 14. Afficher le rapport ---------------------------
    MsgBox sMsg, iIcon, "Detection Sliver Faces — PFE"

Nettoyage:
    ' ---- 15. Liberation memoire ----------------------------
    Set oSPA          = Nothing
    Set oHybridBodies = Nothing
    Set oPart         = Nothing
    Set oDoc          = Nothing
    Exit Sub

' ---- Gestionnaire d'erreurs --------------------------------
ErrHandler:
    MsgBox "Erreur inattendue (#" & Err.Number & ") :" & Chr(13) & Chr(13) & _
           Err.Description & Chr(13) & Chr(13) & _
           "Verifiez que le Part est bien mis a jour (Ctrl+U).", _
           16, "Detection Sliver Faces — PFE"
    Resume Nettoyage

End Sub