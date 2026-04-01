Sub CATMain()
    Dim oDoc As PartDocument
    Set oDoc = CATIA.ActiveDocument
    
    Dim oPart As Part
    Set oPart = oDoc.Part
    
    Dim oSelection As Selection
    Set oSelection = oDoc.Selection
    
    Dim oSurface As AnyObject
    Set oSurface = oSelection.Item(1).Value
    
    Dim oSPA As SPAWorkbench
    Set oSPA = oDoc.GetWorkbench("SPAWorkbench")
    
    Dim oMeasurable As Measurable
    Set oMeasurable = oSPA.GetMeasurable(oSurface)
    
    Dim dMinRadius As Double
    dMinRadius = oMeasurable.GetMinimumCurvatureRadius()
    
    Dim dRequiredThickness As Double
    dRequiredThickness = 2.5
    
    If dMinRadius < dRequiredThickness Then
        MsgBox "Le rayon de courbure (" & Round(dMinRadius, 3) & " mm) " & _
               "est inférieur à l'épaisseur requise (" & dRequiredThickness & " mm)." & vbCrLf & _
               "Risque d'auto-intersection et d'infaisabilité en production.", vbCritical, "Erreur Surfacique"
    Else
        MsgBox "La surface est apte à l'épaississement.", vbInformation
    End If
End Sub