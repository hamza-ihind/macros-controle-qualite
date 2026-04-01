Sub CATMain()
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' Vérification que le document actif est bien un assemblage (Product)
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "Veuillez ouvrir un assemblage (CATProduct) pour lancer le contrôle qualité."
        Exit Sub
    End If
    
    Dim rootProd As Product
    Set rootProd = oDoc.Product
    
    Dim errorReport As String
    errorReport = ""
    
    ' Lancement de la vérification récursive
    Call CheckNamingConvention(rootProd, errorReport)
    
    ' Restitution des résultats à l'utilisateur
    If errorReport = "" Then
        MsgBox "Succès : Toutes les pièces respectent la convention de nommage du projet!", vbInformation, "Contrôle Qualité"
    Else
        MsgBox "Les pièces suivantes ne respectent pas la convention de nommage :" & vbCrLf & vbCrLf & errorReport, vbCritical, "Erreur de Nommage Détectée"
    End If
End Sub

' Fonction récursive pour parcourir tout l'arbre
Sub CheckNamingConvention(currentProd As Product, ByRef errorReport As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Configuration de la règle de validation (Expression Régulière)
    ' Exemple ici : 3 Lettres majuscules, un tiret, puis 3 ou 4 chiffres (ex: PRJ-123)
    regEx.Pattern = "^[A-Z]{3}-\d{3,4}$"
    regEx.IgnoreCase = False
    regEx.Global = True
    
    ' Test du nom (PartNumber) de la pièce évaluée
    If Not regEx.Test(currentProd.PartNumber) Then
        ' Si le nom est incorrect, on l'ajoute au rapport d'erreurs
        errorReport = errorReport & "- " & currentProd.PartNumber & vbCrLf
    End If
    
    ' S'il y a des enfants (sous-produits), la macro relance cette même fonction pour chacun d'eux
    Dim i As Integer
    If currentProd.Products.Count > 0 Then
        For i = 1 To currentProd.Products.Count
            Call CheckNamingConvention(currentProd.Products.Item(i), errorReport)
        Next
    End If
End Sub