Sub CATMain()
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument
    
    Dim errorReport
    errorReport = ""
    
    Dim docType
    docType = TypeName(oDoc)
    
    ' ── Handle CATPart ──────────────────────────────────────────
    If docType = "PartDocument" Then
        Dim oPart
        Set oPart = oDoc.Part
        
        Dim regExPart
        Set regExPart = CreateObject("VBScript.RegExp")
        regExPart.Pattern = "^[A-Z]{3}-\d{3,4}$"
        regExPart.IgnoreCase = False
        
        If Not regExPart.Test(oPart.Name) Then
            errorReport = "- " & oPart.Name
        End If
        
    ' ── Handle CATProduct ───────────────────────────────────────
    ElseIf docType = "ProductDocument" Then
        Dim rootProd
        Set rootProd = oDoc.Product
        Call CheckNamingConvention(rootProd, errorReport)
        
    Else
        MsgBox "Document non supporté. Ouvrez un CATPart ou un CATProduct.", vbExclamation, "Type Invalide"
        Exit Sub
    End If
    
    ' ── Report ──────────────────────────────────────────────────
    If errorReport = "" Then
        MsgBox "Succès : Tous les éléments respectent la convention de nommage!", vbInformation, "Contrôle Qualité"
    Else
        Dim totalLines
        totalLines = UBound(Split(Trim(errorReport), vbCrLf)) + 1
        
        MsgBox "Convention attendue : XXX-0000 (3 lettres majuscules, tiret, 3-4 chiffres)" & vbCrLf & _
               "────────────────────────────────" & vbCrLf & _
               totalLines & " élément(s) non conforme(s) :" & vbCrLf & vbCrLf & _
               errorReport, vbCritical, "Erreur de Nommage Détectée"
    End If
End Sub

Sub CheckNamingConvention(currentProd, ByRef errorReport)
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^[A-Z]{3}-\d{3,4}$"
    regEx.IgnoreCase = False
    
    If Not regEx.Test(currentProd.PartNumber) Then
        errorReport = errorReport & "- " & currentProd.PartNumber & vbCrLf
    End If
    
    Dim i
    If currentProd.Products.Count > 0 Then
        For i = 1 To currentProd.Products.Count
            Call CheckNamingConvention(currentProd.Products.Item(i), errorReport)
        Next
    End If
End Sub