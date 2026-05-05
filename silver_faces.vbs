' ============================================================
'  silver_faces.vbs — v1.0
'  Detecte et corrige les silver faces (surfaces ultra-fines)
'  Fonctionne sur un .CATPart ou un .CATProduct
'  Usage : Tools > Macro > Macros > Run
'
'  Strategies de correction automatique (ref. Table 4.12) :
'    [A] aire < SEUIL_DELETE -> Suppression directe
'          Cas : "Sliver quasi-degeneree (aire ≈ 0)" — Table 4.12, cas 6
'          Methode : Selection.Delete sur la HybridShape
'
'    [B] aire < SEUIL_MM2    -> Healing du corps surfacique parent
'          Cas : import IGES/STEP, micro-ecart entre surfaces — Table 4.12, cas 1 & 2
'          Methode : HybridShapeFactory.AddNewHeal sur toutes les surfaces du corps
'
'  Strategies NON automatisees (intervention manuelle requise) :
'    - Face residuelle apres Trim  -> Delete Face + Fill (Table 4.12, cas 3)
'    - Offset avec forte courbure  -> reduire offset / segmenter (Table 4.12, cas 4)
'    - Import repete               -> nettoyer dans logiciel source (Table 4.12, cas 5)
' ============================================================

Const SEUIL_MM2    = 0.01    ' Seuil de detection (mm2)
Const SEUIL_DELETE = 0.0001  ' En dessous : suppression directe (quasi-degeneree)
Const HEALING_DIST = 0.1     ' Distance de fusion pour le Healing (mm)

' ============================================================
Sub CATMain()

    Dim oDoc, iSliver, iTotal, sList, sMsg, iRep, i
    Set oDoc  = CATIA.ActiveDocument
    iSliver = 0 : iTotal = 0 : sList = ""

    On Error GoTo ErrHandler

    Select Case TypeName(oDoc)

        Case "PartDocument"
            ScanPart oDoc.Part, iSliver, iTotal, sList

        Case "ProductDocument"
            ' Parcourir tous les Parts ouverts (v1.0 — sans filtrage produit)
            For i = 1 To CATIA.Documents.Count
                If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                    ScanPart CATIA.Documents.Item(i).Part, iSliver, iTotal, sList
                End If
            Next i

        Case Else
            MsgBox "Ouvrez un .CATPart ou .CATProduct.", 16, "Silver Faces"
            Exit Sub

    End Select

    ' --- Rapport ---
    sMsg = "=== SILVER FACES ===" & Chr(13) & _
           "Seuil    : " & SEUIL_MM2 & " mm2" & Chr(13) & _
           "Surfaces : " & iTotal    & Chr(13) & _
           "Anomalies: " & iSliver   & Chr(13) & _
           "===================="

    If iSliver = 0 Then
        MsgBox sMsg & Chr(13) & Chr(13) & "[OK] Aucune silver face detectee.", _
               64, "Silver Faces"
        Exit Sub
    End If

    ' --- Proposer la correction automatique ---
    iRep = MsgBox(sMsg & Chr(13) & Chr(13) & sList & Chr(13) & _
                  "Appliquer les corrections automatiques ?", _
                  36, "Silver Faces")  ' 36 = Oui/Non

    If iRep = 6 Then  ' vbYes = 6
        Select Case TypeName(oDoc)
            Case "PartDocument"
                CorrigerPart oDoc.Part
            Case "ProductDocument"
                For i = 1 To CATIA.Documents.Count
                    If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                        CorrigerPart CATIA.Documents.Item(i).Part
                    End If
                Next i
        End Select
    End If

    Exit Sub

ErrHandler:
    MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & _
           "Verifiez que le Part est mis a jour (Ctrl+U).", 16, "Silver Faces"

End Sub

' ============================================================
' Scanne toutes les HybridShapes d'un Part et signale les slivers.
' Indique pour chaque anomalie quelle strategie sera appliquee.
Sub ScanPart(oPart, iSliver, iTotal, sList)

    Dim oSPA, oHBs, oHB, oShape, oRef, oM, dAire, sStrat, i, j

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
            dAire    = oM.Area * 1000000  ' m2 -> mm2
            On Error GoTo 0

            If dAire > 0 And dAire < SEUIL_MM2 Then
                iSliver = iSliver + 1

                ' Indiquer la strategie qui sera appliquee lors de la correction
                If dAire < SEUIL_DELETE Then
                    sStrat = "[A] Suppression"   ' Table 4.12, cas 6 : quasi-degeneree
                Else
                    sStrat = "[B] Healing"        ' Table 4.12, cas 1/2 : sliver standard
                End If

                sList = sList & iSliver & ". " & oShape.Name & _
                        " (" & oPart.Name & "/" & oHB.Name & ")" & _
                        " = " & FormatNumber(dAire, 6) & " mm2  " & sStrat & Chr(13)
            End If
        Next j
    Next i

End Sub

' ============================================================
' Applique les corrections sur un Part :
'
'   Strategie [A] — Table 4.12, cas 6 (aire quasi-nulle)
'     -> Selection.Delete : supprime directement la HybridShape
'     -> Iteration en sens inverse pour eviter les decalages d'index
'
'   Strategie [B] — Table 4.12, cas 1 & 2 (sliver standard)
'     -> HybridShapeHealing : cree un feature Healing sur le corps parent
'        en ajoutant toutes ses surfaces, avec une distance de fusion HEALING_DIST
'     -> Equivalent a : Insert > Operations > Healing dans l'interface CATIA
Sub CorrigerPart(oPart)

    Dim oSPA, oHSF, oHBs, oHB, oShape, oRef, oM, oHeal
    Dim dAire, i, j, k, iDel, iHeal, sLog, bNeedHeal

    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHSF = oPart.HybridShapeFactory
    Set oHBs = oPart.HybridBodies
    iDel = 0 : iHeal = 0 : sLog = ""

    For i = 1 To oHBs.Count
        Set oHB   = oHBs.Item(i)
        bNeedHeal = False

        ' --- Strategie [A] : suppression des faces quasi-degenerees ---
        ' Sens inverse pour ne pas perturber les index apres chaque Delete
        For j = oHB.HybridShapes.Count To 1 Step -1
            Set oShape = oHB.HybridShapes.Item(j)
            dAire = 0
            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000
            On Error GoTo 0

            If dAire > 0 And dAire < SEUIL_DELETE Then
                ' Suppression via Selection (Table 4.12, cas 6)
                On Error Resume Next
                oPart.Parent.Selection.Clear
                oPart.Parent.Selection.Add oShape
                oPart.Parent.Selection.Delete
                On Error GoTo 0
                iDel = iDel + 1
                sLog = sLog & "[A-SUPPRIME] " & oShape.Name & _
                       " (" & FormatNumber(dAire, 6) & " mm2)" & Chr(13)

            ElseIf dAire > 0 And dAire < SEUIL_MM2 Then
                ' Sliver standard -> le corps sera traite par Healing
                bNeedHeal = True
            End If
        Next j

        ' --- Strategie [B] : Healing du corps surfacique parent ---
        ' Cree un feature "Heal" regroupant toutes les surfaces du corps.
        ' Le Healing comble les micro-ecarts entre surfaces adjacentes.
        If bNeedHeal Then
            On Error Resume Next
            Set oHeal = oHSF.AddNewHeal()
            For k = 1 To oHB.HybridShapes.Count
                Set oRef = oPart.CreateReferenceFromObject(oHB.HybridShapes.Item(k))
                oHeal.AddElement oRef
            Next k
            oHeal.MergingDistance = HEALING_DIST  ' distance de fusion en mm
            oHB.AppendHybridShape oHeal
            oPart.Update
            On Error GoTo 0
            iHeal = iHeal + 1
            sLog = sLog & "[B-HEALING]  " & oHB.Name & _
                   " (dist=" & HEALING_DIST & " mm)" & Chr(13)
        End If

    Next i

    MsgBox "Corrections sur : " & oPart.Name & Chr(13) & Chr(13) & _
           "[A] Suppressions : " & iDel  & Chr(13) & _
           "[B] Healings     : " & iHeal & Chr(13) & Chr(13) & sLog & Chr(13) & _
           "Strategies non automatisees (manuel) :" & Chr(13) & _
           "  - Face apres Trim  -> Delete Face + Fill" & Chr(13) & _
           "  - Offset / Import  -> voir logiciel source", _
           64, "Silver Faces — Corrections"

End Sub