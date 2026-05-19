' =============================================================================
' Silver Faces — Contrôle qualité surfaces minuscules (< SEUIL_MM2)
' Stratégies : [A] Suppression directe  |  [B] Healing
' =============================================================================
Const SEUIL_MM2    = 0.01    ' Seuil de détection (mm²)
Const SEUIL_DELETE = 0.0001  ' En dessous : suppression directe
Const HEALING_DIST = 0.1     ' Distance de fusion Healing (mm)

' =============================================================================
' Point d'entrée : détecte les silver faces et propose la correction.
' =============================================================================
Sub CATMain()

    Dim oDoc, iSliver, iTotal, sList, sMsg, iRep, i
    Set oDoc  = CATIA.ActiveDocument
    iSliver = 0 : iTotal = 0 : sList = ""

    Select Case TypeName(oDoc)

        ' ── CATPart ──────────────────────────────────────────────────────────
        Case "PartDocument"

            ' Guard: no HybridBodies → nothing to scan (per flowchart)
            If oDoc.Part.HybridBodies.Count = 0 Then
                MsgBox "Aucun HybridBody détecté dans le Part actif." & Chr(13) & _
                       "Vérifiez que la géométrie surfacique est bien présente.", _
                       48, "Silver Faces — Avertissement"
                Exit Sub
            End If

            On Error Resume Next
            ScanPart oDoc.Part, iSliver, iTotal, sList
            If Err.Number <> 0 Then
                MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & Chr(13) & _
                       "Vérifiez que le Part est mis à jour (Ctrl+U).", 16, "Silver Faces — Erreur"
                On Error GoTo 0 : Exit Sub
            End If
            On Error GoTo 0

        ' ── CATProduct ───────────────────────────────────────────────────────
        Case "ProductDocument"

            Dim bFound : bFound = False
            For i = 1 To CATIA.Documents.Count
                If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                    Dim oCurPart
                    Set oCurPart = CATIA.Documents.Item(i).Part
                    If oCurPart.HybridBodies.Count > 0 Then
                        bFound = True
                        On Error Resume Next
                        ScanPart oCurPart, iSliver, iTotal, sList
                        If Err.Number <> 0 Then
                            MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & Chr(13) & _
                                   "Vérifiez que le Part est mis à jour (Ctrl+U).", 16, "Silver Faces — Erreur"
                            On Error GoTo 0 : Exit Sub
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next

            If Not bFound Then
                MsgBox "Aucun CATPart avec géométrie surfacique trouvé dans le produit.", _
                       48, "Silver Faces — Avertissement"
                Exit Sub
            End If

        Case Else
            MsgBox "Ouvrez un .CATPart ou un .CATProduct.", 16, "Silver Faces"
            Exit Sub

    End Select

    ' ── Rapport ──────────────────────────────────────────────────────────────
    sMsg = "=== SILVER FACES ===" & Chr(13) & _
           "Seuil    : " & SEUIL_MM2 & " mm²" & Chr(13) & _
           "Surfaces : " & iTotal    & Chr(13) & _
           "Anomalies: " & iSliver   & Chr(13) & _
           "===================="

    If iSliver = 0 Then
        MsgBox sMsg & Chr(13) & Chr(13) & "[OK] Aucune silver face détectée.", _
               64, "Silver Faces"
        Exit Sub
    End If

    ' ── Demande de correction ─────────────────────────────────────────────────
    iRep = MsgBox(sMsg & Chr(13) & Chr(13) & sList & Chr(13) & _
                  "Appliquer les corrections automatiques ?", _
                  36, "Silver Faces — Anomalies détectées")   ' 36 = Oui/Non

    If iRep = 6 Then   ' vbYes = 6
        Select Case TypeName(oDoc)
            Case "PartDocument"
                On Error Resume Next
                CorrigerPart oDoc.Part
                If Err.Number <> 0 Then
                    MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & Chr(13) & _
                           "Vérifiez que le Part est mis à jour (Ctrl+U).", 16, "Silver Faces — Erreur"
                    On Error GoTo 0 : Exit Sub
                End If
                On Error GoTo 0

            Case "ProductDocument"
                For i = 1 To CATIA.Documents.Count
                    If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                        If CATIA.Documents.Item(i).Part.HybridBodies.Count > 0 Then
                            On Error Resume Next
                            CorrigerPart CATIA.Documents.Item(i).Part
                            If Err.Number <> 0 Then
                                MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & Chr(13) & _
                                       "Vérifiez que le Part est mis à jour (Ctrl+U).", 16, "Silver Faces — Erreur"
                                On Error GoTo 0 : Exit Sub
                            End If
                            On Error GoTo 0
                        End If
                    End If
                Next
        End Select
    End If

    Exit Sub

End Sub

' =============================================================================
' Scanne tous les HybridShapes d'un Part et détecte les silver faces.
' Les compteurs et la liste sont passés ByRef pour mise à jour dans CATMain.
' =============================================================================
Sub ScanPart(ByRef oPart, ByRef iSliver, ByRef iTotal, ByRef sList)

    Dim oSPA, oHBs, oHB, oShape, oRef, oM
    Dim dAire, sStrat, i, j

    ' GetWorkbench appelé une seule fois par Part (pas dans la boucle)
    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHBs = oPart.HybridBodies

    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        For j = 1 To oHB.HybridShapes.Count
            Set oShape = oHB.HybridShapes.Item(j)
            iTotal = iTotal + 1
            dAire  = 0

            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000   ' m² → mm²
            On Error GoTo 0

            If dAire > 0 And dAire < SEUIL_MM2 Then
                iSliver = iSliver + 1

                If dAire < SEUIL_DELETE Then
                    sStrat = "[A] Suppression"
                Else
                    sStrat = "[B] Healing"
                End If

                sList = sList & iSliver & ". " & oShape.Name & _
                        "  (" & oPart.Name & " / " & oHB.Name & ")" & _
                        "  = " & FormatNumber(dAire, 6) & " mm²  " & sStrat & Chr(13)
            End If

            Set oM   = Nothing
            Set oRef = Nothing
            Set oShape = Nothing
        Next
        Set oHB = Nothing
    Next

    Set oHBs = Nothing
    Set oSPA = Nothing

End Sub

' =============================================================================
' Corrige les silver faces :
'   [A] Suppression directe (aire < SEUIL_DELETE)
'   [B] Healing sur le HybridBody (SEUIL_DELETE ≤ aire < SEUIL_MM2)
' =============================================================================
Sub CorrigerPart(ByRef oPart)

    Dim oSPA, oHSF, oHBs, oHB, oShape, oRef, oM, oHeal
    Dim dAire, i, j, k, iDel, iHeal, sLog, bNeedHeal

    ' GetWorkbench appelé une seule fois pour tout le Part
    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHSF = oPart.HybridShapeFactory
    Set oHBs = oPart.HybridBodies
    iDel = 0 : iHeal = 0 : sLog = ""

    For i = 1 To oHBs.Count
        Set oHB   = oHBs.Item(i)
        bNeedHeal = False

        ' Parcours en sens inverse : évite le décalage d'index lors des suppressions
        For j = oHB.HybridShapes.Count To 1 Step -1
            Set oShape = oHB.HybridShapes.Item(j)
            dAire = 0

            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000   ' m² → mm²
            On Error GoTo 0

            ' [A] Suppression directe : aire quasi-nulle
            If dAire > 0 And dAire < SEUIL_DELETE Then
                Dim oSel
                Set oSel = oPart.Parent.Selection
                On Error Resume Next
                oSel.Clear
                oSel.Add oShape
                oSel.Delete
                On Error GoTo 0
                Set oSel = Nothing
                iDel = iDel + 1
                sLog = sLog & "[A-SUPPRIMÉ] " & oShape.Name & _
                       "  (" & FormatNumber(dAire, 6) & " mm²)" & Chr(13)

            ' [B] Healing : surface petite mais non quasi-nulle
            ElseIf dAire > 0 And dAire < SEUIL_MM2 Then
                bNeedHeal = True
            End If

            Set oM     = Nothing
            Set oRef   = Nothing
            Set oShape = Nothing
        Next

        ' [B] Crée le Healing sur le corps entier si au moins une surface B détectée
        If bNeedHeal And oHB.HybridShapes.Count > 0 Then
            On Error Resume Next
            Set oHeal = oHSF.AddNewHeal()
            For k = 1 To oHB.HybridShapes.Count
                Set oRef = oPart.CreateReferenceFromObject(oHB.HybridShapes.Item(k))
                oHeal.AddElement oRef
                Set oRef = Nothing
            Next
            oHeal.MergingDistance = HEALING_DIST
            oHB.AppendHybridShape oHeal
            oPart.Update
            On Error GoTo 0
            Set oHeal = Nothing
            iHeal = iHeal + 1
            sLog = sLog & "[B-HEALING]  " & oHB.Name & _
                   "  (dist = " & HEALING_DIST & " mm)" & Chr(13)
        End If

        Set oHB = Nothing
    Next

    Set oHBs = Nothing
    Set oHSF = Nothing
    Set oSPA = Nothing

    MsgBox "Corrections appliquées sur : " & oPart.Name & Chr(13) & Chr(13) & _
           "[A] Suppressions : " & iDel  & Chr(13) & _
           "[B] Healings     : " & iHeal & Chr(13) & Chr(13) & sLog & Chr(13) & _
           "Cas nécessitant une correction manuelle :" & Chr(13) & _
           "  · Face résiduelle après Trim  →  Delete Face + Fill" & Chr(13) & _
           "  · Offset / Import dégradé     →  corriger dans le logiciel source", _
           64, "Silver Faces — Corrections"

End Sub