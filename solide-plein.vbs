' solide-plein.vbs — Verifie si tous les corps du Part sont solides et pleins

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

' Point d'entree principal
Sub CATMain()
    Dim oDoc
    Dim oBodies
    Dim iIcon
    Dim bError
    Dim sErrMsg
    Dim i

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
        If TypeName(oDoc) <> "PartDocument" Then
            MsgBox "Le document actif n'est pas un .CATPart." & Chr(13) & _
                   "Ouvrez un .CATPart avant de lancer la macro.", _
                   16, "Verification Solide"
            bError = True
        End If
    End If

    If Not bError Then
        Set g_oPart = oDoc.Part
        If Err.Number <> 0 Then
            sErrMsg = "Part inaccessible. Err#" & Err.Number & " : " & Err.Description
            bError = True
            Err.Clear
        End If
    End If

    If Not bError Then
        Set g_oSPA = oDoc.GetWorkbench("SPAWorkbench")
        If Err.Number <> 0 Then
            sErrMsg = "SPAWorkbench inaccessible. Err#" & Err.Number & " : " & Err.Description
            bError = True
            Err.Clear
        End If
    End If

    If Not bError Then
        Set oBodies = g_oPart.Bodies
        If Err.Number <> 0 Then
            sErrMsg = "Lecture Bodies impossible. Err#" & Err.Number & " : " & Err.Description
            bError = True
            Err.Clear
        End If
    End If

    If Not bError Then
        If oBodies.Count = 0 Then
            MsgBox "Aucun corps (Body) trouve dans ce Part." & Chr(13) & _
                   "Ajoutez au moins un Pad ou Shaft.", _
                   16, "Verification Solide"
            bError = True
        End If
    End If

    If Not bError Then
        g_bAllOK  = True
        g_sReport = "===== VERIFICATION SOLIDE =====" & Chr(13)
        g_sReport = g_sReport & "Part   : " & g_oPart.Name & Chr(13)
        g_sReport = g_sReport & "Corps  : " & oBodies.Count & Chr(13)
        g_sReport = g_sReport & "-------------------------------" & Chr(13)

        For i = 1 To oBodies.Count
            Call AnalyseBody(oBodies.Item(i))
        Next

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
    Set oBodies = Nothing
    Set oDoc    = Nothing

End Sub