' ============================================================
'  CheckSolidAndPlein.vbs
'  Verifie si le Part actif est solide et plein (Volume > 0)
'  Auteur  : PFE Macro — CATIA V5
'  Usage   : Tools > Macro > Macros > Run
' ============================================================

Sub CATMain()
    Dim oDoc
    Dim oPart
    Dim oBody
    Dim oSPA
    Dim oMeasure
    Dim dVolume
    Dim dArea
    Dim dVol_mm3
    Dim dArea_mm2
    Dim sMsg
    Dim iIcon

    ' ---- Gestion des erreurs globales
    On Error GoTo ErrHandler

    ' ---- 1. Recuperer le document actif
    Set oDoc = CATIA.ActiveDocument

    ' ---- 2. Verifier que c'est un PartDocument
    If TypeName(oDoc) <> "PartDocument" Then
        MsgBox "ERREUR : Le document actif n'est pas un fichier .CATPart." & Chr(13) & Chr(13) & _
               "Veuillez ouvrir un fichier .CATPart avant de lancer la macro.", _
               16, "Verification Solide — PFE"
        Exit Sub
    End If

    ' ---- 3. Acceder a l'objet Part
    Set oPart = oDoc.Part

    ' ---- 4. Acceder au PartBody principal
    Set oBody = oPart.MainBody

    ' ---- 5. Verifier que le PartBody n'est pas vide
    If oBody.Shapes.Count = 0 Then
        MsgBox "ERREUR : Le PartBody est vide." & Chr(13) & Chr(13) & _
               "Aucune feature solide detectee dans le corps principal." & Chr(13) & _
               "Creez un Pad, Shaft ou autre feature dans Part Design.", _
               16, "Verification Solide — PFE"
        GoTo Nettoyage
    End If

    ' ---- 6. Mesurer le volume via SPAWorkbench
    Set oSPA     = oDoc.GetWorkbench("SPAWorkbench")
    Set oMeasure = oSPA.GetMeasurable(oBody)

    ' CATIA renvoie en m3 -> conversion en mm3 (x 10^9)
    ' CATIA renvoie en m2 -> conversion en mm2 (x 10^6)
    dVolume   = oMeasure.Volume
    dArea     = oMeasure.Area
    dVol_mm3  = dVolume * 1000000000
    dArea_mm2 = dArea   * 1000000

    ' ---- 7. Construire le rapport
    sMsg = "======= RAPPORT — VERIFICATION SOLIDE =======" & Chr(13)
    sMsg = sMsg & "Part        : " & oPart.Name                            & Chr(13)
    sMsg = sMsg & "PartBody    : " & oBody.Name                            & Chr(13)
    sMsg = sMsg & "Nb features : " & oBody.Shapes.Count                   & Chr(13)
    sMsg = sMsg & "Volume      : " & FormatNumber(dVol_mm3,  3) & " mm3"  & Chr(13)
    sMsg = sMsg & "Surface     : " & FormatNumber(dArea_mm2, 3) & " mm2"  & Chr(13)
    sMsg = sMsg & "=============================================" & Chr(13)

    ' ---- 8. Verdict final
    If dVol_mm3 > 0 Then
        sMsg  = sMsg & "[OK]  Le Part est SOLIDE et PLEIN."
        iIcon = 64
    Else
        sMsg  = sMsg & "[!!]  Volume NUL detecte." & Chr(13) & Chr(13) & _
                "Causes possibles :" & Chr(13) & _
                "  - Un Pocket annule totalement un Pad" & Chr(13) & _
                "  - La geometrie contient une auto-intersection" & Chr(13) & _
                "  - Une face est manquante ou ouverte" & Chr(13) & Chr(13) & _
                "Action : Lancez Analyze > Check Geometry dans CATIA."
        iIcon = 48
    End If

    ' ---- 9. Afficher le resultat
    MsgBox sMsg, iIcon, "Verification Solide — PFE"

Nettoyage:
    ' ---- 10. Liberation memoire
    Set oMeasure = Nothing
    Set oSPA     = Nothing
    Set oBody    = Nothing
    Set oPart    = Nothing
    Set oDoc     = Nothing
    Exit Sub

' ---- Gestionnaire d'erreurs
ErrHandler:
    MsgBox "Erreur inattendue (#" & Err.Number & ") :" & Chr(13) & Chr(13) & _
           Err.Description & Chr(13) & Chr(13) & _
           "Verifiez que le Part est bien mis a jour (Ctrl+U).", _
           16, "Verification Solide — PFE"
    Resume Nettoyage

End Sub