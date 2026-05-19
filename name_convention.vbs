' ============================================================================
' Contrôle de la convention de nommage :
'   KVScode_CADTYPE_PDA_VERSION_DESCRIPTION
'   Ex. : 5FF821105H_PCA_TM_5_FENDER_PHEV
'
'   KVS code   : 3 alphanumériques + 6 chiffres + 1 alphanum  (ex. 5FF821105H)
'   CAD type   : DMU | PCA | GEO | VEE
'   PDA        : TM (officiel) | EN (proposition)
'   Version    : chiffre(s)
'   Description: lettres maj. + chiffres + underscores (ex. FENDER_PHEV)
'
' - Vérifie TOUS les éléments (pas seulement le premier)
' - Affiche l'emplacement complet dans l'arbre pour chaque erreur
' - Sélectionne / surligne tous les éléments non conformes dans CATIA
' ============================================================================

Sub CATMain()
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument

    ' Clear existing selection so highlighted items are only the bad ones
    Dim oSel
    Set oSel = oDoc.Selection
    oSel.Clear

    ' Single regex instance reused everywhere.
    ' Pattern: KVScode_CADTYPE_PDA_VERSION_DESCRIPTION
    '   KVS    : [A-Z0-9]{3}\d{6}[A-Z0-9]   e.g. 5FF821105H
    '   CADTYPE: DMU|PCA|GEO|VEE
    '   PDA    : TM|EN
    '   VERSION: \d+
    '   DESC   : [A-Z][A-Z0-9_]+             e.g. FENDER_PHEV
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^[A-Z0-9]{3}\d{6}[A-Z0-9]_(DMU|PCA|GEO|VEE)_(TM|EN)_\d+_[A-Z][A-Z0-9_]*$"
    regEx.IgnoreCase = False

    Dim report
    report = ""
    Dim errorCount
    errorCount = 0

    Dim docType
    docType = TypeName(oDoc)

    ' ── CATPart ──────────────────────────────────────────────────────────────
    If docType = "PartDocument" Then
        Dim oPart
        Set oPart = oDoc.Part

        ' Check the Part root name
        If Not regEx.Test(oPart.Name) Then
            Call AddError(report, errorCount, "[Part]     ", oPart.Name, oPart.Name)
        End If

        ' Check every Body (PartBody, Body.2, …)
        Dim i
        For i = 1 To oPart.Bodies.Count
            Dim oBody
            Set oBody = oPart.Bodies.Item(i)
            If Not regEx.Test(oBody.Name) Then
                Call AddError(report, errorCount, "[Corps]    ", oBody.Name, _
                              oPart.Name & " > " & oBody.Name)
                On Error Resume Next
                oSel.Add oBody
                On Error GoTo 0
            End If
        Next

        ' Check every Geometric Set (HybridBody), recursively
        Call ScanHybridBodies(oPart.HybridBodies, oPart.Name, regEx, _
                              report, errorCount, oSel)

    ' ── CATProduct ───────────────────────────────────────────────────────────
    ElseIf docType = "ProductDocument" Then
        Dim rootProd
        Set rootProd = oDoc.Product
        Call ScanProduct(rootProd, rootProd.PartNumber, regEx, _
                         report, errorCount, oSel)

    Else
        MsgBox "Document non supporté. Ouvrez un CATPart ou un CATProduct.", _
               vbExclamation, "Type Invalide"
        Exit Sub
    End If

    ' ── Final report ─────────────────────────────────────────────────────────
    If errorCount = 0 Then
        MsgBox "Succès : Tous les éléments respectent la convention de nommage!", _
               vbInformation, "Contrôle Qualité"
    Else
        MsgBox "Convention attendue : KVScode_CADTYPE_PDA_VERSION_DESCRIPTION" & vbCrLf & _
               "  KVS     : 3 alphanum + 6 chiffres + 1 alphanum  (ex. 5FF821105H)" & vbCrLf & _
               "  CAD type: DMU | PCA | GEO | VEE" & vbCrLf & _
               "  PDA     : TM (officiel) | EN (proposition)" & vbCrLf & _
               "  Version : chiffre(s)" & vbCrLf & _
               "  Desc.   : lettres maj + chiffres + _ (ex. FENDER_PHEV)" & vbCrLf & _
               "  Exemple : 5FF821105H_PCA_TM_5_FENDER_PHEV" & vbCrLf & _
               "────────────────────────────────────────────────────" & vbCrLf & _
               errorCount & " élément(s) non conforme(s) — surlignés dans l'arbre :" & vbCrLf & vbCrLf & _
               report, vbCritical, "Erreurs de Nommage Détectées"
    End If
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Appends one error block to the report and increments the counter.
' ─────────────────────────────────────────────────────────────────────────────
Sub AddError(ByRef report, ByRef errorCount, itemType, itemName, treePath)
    report = report & itemType & Chr(34) & itemName & Chr(34) & vbCrLf
    report = report & "           Emplacement : " & treePath & vbCrLf & vbCrLf
    errorCount = errorCount + 1
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Recursively walks a CATProduct tree and flags every non-conforming PartNumber.
' treePath accumulates the full path from the root down to the current node.
' ─────────────────────────────────────────────────────────────────────────────
Sub ScanProduct(prod, treePath, regEx, ByRef report, ByRef errorCount, oSel)
    If Not regEx.Test(prod.PartNumber) Then
        Call AddError(report, errorCount, "[Produit]  ", prod.PartNumber, treePath)
        On Error Resume Next
        oSel.Add prod
        On Error GoTo 0
    End If

    ' Safely read child count (leaf parts expose an empty collection, not an error)
    Dim n
    n = 0
    On Error Resume Next
    n = prod.Products.Count
    On Error GoTo 0

    Dim i
    For i = 1 To n
        Dim child
        Set child = prod.Products.Item(i)
        Call ScanProduct(child, treePath & " > " & child.PartNumber, _
                         regEx, report, errorCount, oSel)
    Next
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Recursively walks HybridBodies (Geometric Sets) inside a CATPart.
' ─────────────────────────────────────────────────────────────────────────────
Sub ScanHybridBodies(hbColl, parentPath, regEx, ByRef report, ByRef errorCount, oSel)
    Dim i
    For i = 1 To hbColl.Count
        Dim hb
        Set hb = hbColl.Item(i)
        Dim hbPath
        hbPath = parentPath & " > " & hb.Name
        If Not regEx.Test(hb.Name) Then
            Call AddError(report, errorCount, "[Géom.Set] ", hb.Name, hbPath)
            On Error Resume Next
            oSel.Add hb
            On Error GoTo 0
        End If
        ' Recurse into nested Geometric Sets
        Call ScanHybridBodies(hb.HybridBodies, hbPath, regEx, report, errorCount, oSel)
    Next
End Sub