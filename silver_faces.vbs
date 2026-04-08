' Macro : Detection_Sliver_Faces
' Objectif : Identifier les faces trop petites ou trop élancées

Sub CATMain()
    Dim oDoc As PartDocument : Set oDoc = CATIA.ActiveDocument
    Dim oPart As Part : Set oPart = oDoc.Part
    Dim oSelection As Selection : Set oSelection = oDoc.Selection
    
    ' On cherche toutes les faces du corps principal
    oSelection.Search "CATTopology.Face,all"
    
    Dim i As Integer
    Dim oSPA As SPAWorkbench : Set oSPA = oDoc.GetWorkbench("SPAWorkbench")
    Dim oMeasurable As Measurable
    Dim dArea As Double
    Dim iSliverCount As Integer : iSliverCount = 0
    
    ' Seuil de tolérance (à adapter selon le cahier des charges SEGULA)
    Dim dAreaLimit As Double : dAreaLimit = 0.01 ' en mm²
    
    For i = 1 To oSelection.Count
        Set oMeasurable = oSPA.GetMeasurable(oSelection.Item(i).Value)
        dArea = oMeasurable.Area
        
        ' Si l'aire est quasi nulle, c'est une sliver face potentielle
        If dArea < dAreaLimit And dArea > 0 Then
            iSliverCount = iSliverCount + 1
            ' On peut colorer la face en rouge pour la mise en évidence
            oSelection.Item(i).VisProperties.SetRealColor 255, 0, 0, 1
        End If
    Next
    
    If iSliverCount > 0 Then
        MsgBox iSliverCount & " Sliver Faces détectées (en rouge)." & vbCrLf & _
               "Ces faces doivent être supprimées ou fusionnées (Healing) " & _
               "pour garantir la validité du maillage FEA.", vbCritical
    Else
        MsgBox "Qualité Géométrique OK : Aucune face dégénérée détectée.", vbInformation
    End If
End Sub