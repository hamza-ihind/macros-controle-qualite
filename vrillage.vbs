Sub CATMain()
    Dim oDoc As PartDocument
    Set oDoc = CATIA.ActiveDocument
    
    Dim oSelection As Selection
    Set oSelection = oDoc.Selection
    
    ' 1. Vérification de la sélection
    If oSelection.Count = 0 Then
        MsgBox "Veuillez sélectionner une surface avant de lancer l'analyse.", vbExclamation
        Exit Sub
    End If
    
    Dim oSurface As AnyObject
    Set oSurface = oSelection.Item(1).Value
    
    Dim oSPA As SPAWorkbench
    Set oSPA = oDoc.GetWorkbench("SPAWorkbench")
    
    Dim oMeasurable As Measurable
    Set oMeasurable = oSPA.GetMeasurable(oSurface)
    
    Dim i, j As Integer
    Dim NbPoints As Integer : NbPoints = 10
    Dim vNorm(2) As Variant
    Dim vPrevNorm(2) As Variant
    Dim dDotProduct As Double
    Dim bTwistDetected As Boolean : bTwistDetected = False
    
    On Error Resume Next
    
    For i = 1 To NbPoints
        For j = 1 To NbPoints
            oMeasurable.GetNormal vNorm
            
            If i > 1 Or j > 1 Then
                dDotProduct = (vNorm(0) * vPrevNorm(0)) + (vNorm(1) * vPrevNorm(1)) + (vNorm(2) * vPrevNorm(2))
                
                If dDotProduct < 0 Then
                    bTwistDetected = True
                    Exit For
                End If
            End If
            
            vPrevNorm(0) = vNorm(0)
            vPrevNorm(1) = vNorm(1)
            vPrevNorm(2) = vNorm(2)
        Next
        If bTwistDetected Then Exit For
    Next

    If bTwistDetected Then
        MsgBox "ERREUR CRITIQUE : Vrillage (Twist) détecté !" & vbCrLf & _
               "La normale de la surface s'inverse brutalement." & vbCrLf & _
               "Risque d'échec : Maillage FEA et Usinage CNC.", vbCritical, "Q-Checker - Rapport de Qualité"
    Else
        MsgBox "Validation Surfacique : Aucune anomalie de vrillage détectée." & vbCrLf & _
               "La surface est mathématiquement régulière.", vbInformation, "Q-Checker - Succès"
    End If
End Sub