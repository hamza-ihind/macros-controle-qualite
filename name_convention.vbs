Sub CATMain()
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "Veuillez ouvrir un assemblage (CATProduct) pour lancer le contrôle qualité."
        Exit Sub
    End If
    
    Dim rootProd As Product
    Set rootProd = oDoc.Product
    
    Dim errorReport As String
    errorReport = ""
    
    Call CheckNamingConvention(rootProd, errorReport)
    
    If errorReport = "" Then
        MsgBox "Succès : Toutes les pièces respectent la convention de nommage du projet!", vbInformation, "Contrôle Qualité"
    Else
        MsgBox "Les pièces suivantes ne respectent pas la convention de nommage :" & vbCrLf & vbCrLf & errorReport, vbCritical, "Erreur de Nommage Détectée"
    End If
End Sub

Sub CheckNamingConvention(currentProd As Product, ByRef errorReport As String)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Pattern = "^[A-Z]{3}-\d{3,4}$"
    regEx.IgnoreCase = False
    regEx.Global = True
    
    
    If Not regEx.Test(currentProd.PartNumber) Then
        errorReport = errorReport & "- " & currentProd.PartNumber & vbCrLf
    End If
    
    
    Dim i As Integer
    If currentProd.Products.Count > 0 Then
        For i = 1 To currentProd.Products.Count
            Call CheckNamingConvention(currentProd.Products.Item(i), errorReport)
        Next
    End If
End Sub