' solide-plein.vbs — Verifie si tous les corps sont solides et pleins (CATPart ET CATProduct recursif)

Dim g_oSPA
Dim g_oPart
Dim g_sReport
Dim g_bAllOK

' Analyse un Body : volume, surface, diagnostic si KO
Sub AnalyseBody(oBody)
    Dim oRef
    Dim oMeasure
    Dim dVol_mm3
    Dim dArea_mm2
    Dim nShapes
    Dim sStatus

    On Error Resume Next

    Err.Clear
    nShapes = oBody.Shapes.Count
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & oBody.Name & " : lecture impossible — " & Err.Description & Chr(13)
        g_bAllOK = False
        Err.Clear
        Exit Sub
    End If

    If nShapes = 0 Then
        g_sReport = g_sReport & "  [--] " & oBody.Name & " : vide (0 feature)" & Chr(13)
        g_bAllOK = False
        Exit Sub
    End If

    ' Creer une Reference pour GetMeasurable
    Err.Clear
    Set oRef = g_oPart.CreateReferenceFromObject(oBody)
    If Err.Number <> 0 Then
        Err.Clear
        Set oMeasure = g_oSPA.GetMeasurable(oBody)   ' fallback
    Else
        Set oMeasure = g_oSPA.GetMeasurable(oRef)
    End If

    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & oBody.Name & " : mesure impossible — " & Err.Description & Chr(13) & _
                    "         -> Ctrl+U pour mettre le Part a jour." & Chr(13)
        g_bAllOK = False
        Err.Clear
        Exit Sub
    End If

    Err.Clear
    dVol_mm3 = oMeasure.Volume * 1000000000
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & oBody.Name & " : volume illisible — " & Err.Description & Chr(13)
        g_bAllOK = False
        Err.Clear
        Set oMeasure = Nothing
        Exit Sub
    End If

    Err.Clear
    dArea_mm2 = oMeasure.Area * 1000000
    If Err.Number <> 0 Then
        dArea_mm2 = 0
        Err.Clear
    End If
    Set oMeasure = Nothing

    If dVol_mm3 > 0 Then
        sStatus = "[OK]"
    Else
        sStatus  = "[!!]"
        g_bAllOK = False
    End If

    g_sReport = g_sReport & "  " & sStatus & " " & oBody.Name                              & Chr(13)
    g_sReport = g_sReport & "       Features : " & nShapes                                 & Chr(13)
    g_sReport = g_sReport & "       Volume   : " & FormatNumber(dVol_mm3,  3) & " mm3"     & Chr(13)
    g_sReport = g_sReport & "       Surface  : " & FormatNumber(dArea_mm2, 3) & " mm2"     & Chr(13)

    If dVol_mm3 <= 0 Then
        g_sReport = g_sReport & "       --- DIAGNOSTIC ---" & Chr(13)
        Err.Clear
        g_oPart.Update
        If Err.Number <> 0 Then
            g_sReport = g_sReport & "       [!] Mise a jour echouee : " & Err.Description & Chr(13) & _
                        "           -> Verifiez les features en rouge/jaune." & Chr(13)
            Err.Clear
        Else
            g_sReport = g_sReport & "       [!] Volume nul. Causes possibles :" & Chr(13) & _
                        "           - Pocket annulant le solide" & Chr(13) & _
                        "           - Sketch auto-intersecte" & Chr(13) & _
                        "           - Corps ouvert / face manquante" & Chr(13) & _
                        "           - Shaft a 0 degre" & Chr(13) & _
                        "           -> Analyze > Check Geometry" & Chr(13)
        End If
    End If
End Sub

' Analyse tous les Bodies d'un PartDocument donne
Sub AnalysePart(oPartDoc)
    Dim oBodies
    Dim i

    On Error Resume Next

    Set g_oPart = oPartDoc.Part
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & oPartDoc.Name & " : Part inaccessible — " & Err.Description & Chr(13)
        g_bAllOK = False
        Err.Clear
        Exit Sub
    End If

    Set g_oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & g_oPart.Name & " : SPAWorkbench inaccessible — " & Err.Description & Chr(13)
        g_bAllOK = False
        Err.Clear
        Exit Sub
    End If

    Set oBodies = g_oPart.Bodies
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  [??] " & g_oPart.Name & " : lecture Bodies impossible — " & Err.Description & Chr(13)
        g_bAllOK = False
        Err.Clear
        Exit Sub
    End If

    If oBodies.Count = 0 Then
        g_sReport = g_sReport & "  [--] " & g_oPart.Name & " : aucun corps trouve" & Chr(13)
        Set oBodies = Nothing
        Exit Sub
    End If

    g_sReport = g_sReport & "Part   : " & g_oPart.Name & Chr(13)
    g_sReport = g_sReport & "Corps  : " & oBodies.Count & Chr(13)
    g_sReport = g_sReport & "-------------------------------" & Chr(13)

    For i = 1 To oBodies.Count
        Call AnalyseBody(oBodies.Item(i))
    Next

    g_sReport = g_sReport & "-------------------------------" & Chr(13)

    Set oBodies = Nothing
End Sub

' Parcours recursif d'un CATProduct — appelle AnalysePart sur chaque feuille CATPart
Sub WalkProduct(oProd)
    Dim i
    Dim oSubDoc

    On Error Resume Next

    If oProd.Products.Count = 0 Then
        ' Feuille : recuperer le PartDocument associe
        Err.Clear
        Set oSubDoc = oProd.ReferenceProduct.Parent
        If Err.Number <> 0 Then
            g_sReport = g_sReport & "  [??] " & oProd.Name & " : document inaccessible — " & Err.Description & Chr(13)
            g_bAllOK = False
            Err.Clear
            Exit Sub
        End If
        If TypeName(oSubDoc) = "PartDocument" Then
            Call AnalysePart(oSubDoc)
        End If
    Else
        ' Sous-assemblage : recurser
        For i = 1 To oProd.Products.Count
            Call WalkProduct(oProd.Products.Item(i))
        Next
    End If
End Sub

' Point d'entree principal
Sub CATMain()
    Dim oDoc
    Dim iIcon
    Dim bError
    Dim sErrMsg

    On Error Resume Next
    bError  = False
    sErrMsg = ""

    Set oDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        sErrMsg = "Document actif inaccessible. Err#" & Err.Number & " : " & Err.Description
        bError = True
        Err.Clear
    End If

    If Not bError Then
        g_bAllOK  = True
        g_sReport = "===== VERIFICATION SOLIDE =====" & Chr(13)

        If TypeName(oDoc) = "PartDocument" Then
            Call AnalysePart(oDoc)

        ElseIf TypeName(oDoc) = "ProductDocument" Then
            g_sReport = g_sReport & "Produit : " & oDoc.Product.Name & Chr(13)
            g_sReport = g_sReport & "===============================" & Chr(13)
            Call WalkProduct(oDoc.Product)

        Else
            MsgBox "Le document actif n'est pas un .CATPart ou un .CATProduct." & Chr(13) & _
                   "Ouvrez un fichier valide avant de lancer la macro.", _
                   16, "Verification Solide"
            bError = True
        End If
    End If

    If Not bError Then
        g_sReport = g_sReport & "===============================" & Chr(13)

        If g_bAllOK Then
            g_sReport = g_sReport & "RESULTAT : Tous les corps sont SOLIDES et PLEINS."
            iIcon = 64
        Else
            g_sReport = g_sReport & "RESULTAT : Probleme detecte !" & Chr(13) & _
                        "Action   : Analyze > Check Geometry"
            iIcon = 48
        End If

        MsgBox g_sReport, iIcon, "Verification Solide"
    End If

    If bError And sErrMsg <> "" Then
        MsgBox sErrMsg & Chr(13) & "Verifiez que le Part est a jour (Ctrl+U).", _
               16, "Verification Solide"
    End If

    Set g_oSPA  = Nothing
    Set g_oPart = Nothing
    Set oDoc    = Nothing

End Sub