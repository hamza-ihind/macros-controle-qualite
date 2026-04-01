Sub CATMain()

    ' 1. Vérification du type de document actif
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    If TypeName(oDoc) <> "PartDocument" Then
        MsgBox "Erreur : Le document actif n'est pas une pièce. Veuillez ouvrir un fichier Part.", vbCritical, "Erreur de contexte"
        Exit Sub
    End If

    ' 2. Accès à la pièce et au corps principal (PartBody)
    Dim oPart As Part
    Set oPart = oDoc.Part
    
    Dim oBody As Body
    Set oBody = oPart.MainBody

    ' 3. Accès à l'atelier de mesure (SPAWorkbench)
    Dim oSPA As SPAWorkbench
    Set oSPA = oDoc.GetWorkbench("SPAWorkbench")

    ' 4. Création d'une référence pour mesurer le corps
    Dim oRef As Reference
    Set oRef = oPart.CreateReferenceFromObject(oBody)
    
    Dim oMeasurable As Measurable
    Set oMeasurable = oSPA.GetMeasurable(oRef)

    ' 5. Récupération du volume
    Dim dVolume As Double
    ' CATIA renvoie le volume en mètres cubes par défaut dans certaines API, 
    ' mais l'objet Measurable renvoie souvent la valeur en mm cubes.
    dVolume = oMeasurable.Volume

    ' 6. Condition de validation
    ' On utilise une petite tolérance (0.001) pour éviter les erreurs de calcul flottant
    If dVolume > 0.001 Then
        ' Si valide -- message confirmant que tout va bien
        MsgBox "Validation Réussie :" & vbCrLf & "Le corps principal est solide et plein." & vbCrLf & "Volume mesuré : " & Round(dVolume, 2) & " mm³.", vbInformation, "Statut de la pièce"
    Else
        ' Si non valide -- message d'erreur expliquant le problème
        MsgBox "Erreur de Conception :" & vbCrLf & "La pièce n'est pas un volume plein." & vbCrLf & "Problème : Le corps principal est vide ou composé uniquement de surfaces ouvertes. Veuillez utiliser des fonctions volumiques (Extrusion, Remplissage de surface, etc.).", vbExclamation, "Échec de validation"
    End If

End Sub