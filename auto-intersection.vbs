Dim gTotal As Integer
Dim gFailed As Integer
Dim gReport As String

Sub CATMain()
    Dim oDoc As Document
    Dim sType As String

    gTotal  = 0
    gFailed = 0
    gReport = ""

    Set oDoc = CATIA.ActiveDocument
    sType = TypeName(oDoc)

    If sType = "PartDocument" Then
        ProcessPart oDoc.Part
    ElseIf sType = "ProductDocument" Then
        ScanProduct oDoc.Product
    Else
        MsgBox "Ouvrez un CATPart ou CATProduct.", vbExclamation
        Exit Sub
    End If

    If gTotal = 0 Then
        MsgBox "Aucune surface trouvee.", vbInformation
    ElseIf gFailed = 0 Then
        MsgBox "OK - " & gTotal & " surface(s) verifiee(s)." & vbCrLf & gReport, vbInformation
    Else
        MsgBox gFailed & " auto-intersection(s) sur " & gTotal & " surface(s)." & vbCrLf & gReport, vbCritical
    End If
End Sub

Sub ScanProduct(oProd As Product)
    Dim i      As Integer
    Dim oChild As Product
    Dim oDoc   As Document
    Dim sType  As String

    For i = 1 To oProd.Products.Count
        Set oChild = oProd.Products.Item(i)
        sType = ""
        On Error Resume Next
        Set oDoc = oChild.ReferenceProduct.Parent
        sType = TypeName(oDoc)
        On Error GoTo 0
        If sType = "PartDocument" Then
            ProcessPart oDoc.Part
        Else
            ScanProduct oChild
        End If
    Next
End Sub

Sub ProcessPart(oPart As Part)
    If oPart.HybridBodies.Count > 0 Then
        ScanHybridBodies oPart.HybridBodies, oPart
    End If
End Sub

Sub ScanHybridBodies(oHBs As HybridBodies, oPart As Part)
    Dim oHB As HybridBody
    Dim k   As Integer
    Dim oHS As HybridShape

    For Each oHB In oHBs
        For k = 1 To oHB.HybridShapes.Count
            Set oHS = oHB.HybridShapes.Item(k)
            CheckSurface oHS, oPart
        Next
        If oHB.HybridBodies.Count > 0 Then
            ScanHybridBodies oHB.HybridBodies, oPart
        End If
    Next
End Sub

Sub CheckSurface(oHS As HybridShape, oPart As Part)
    Dim oSPA    As SPAWorkbench
    Dim oRef    As Reference
    Dim oMeas   As Measurable
    Dim dArea   As Double
    Dim oHSF    As HybridShapeFactory
    Dim oTmpHB  As HybridBody
    Dim oJoin   As HybridShapeAssemble
    Dim oSel    As Selection
    Dim bFailed As Boolean

    dArea = 0
    On Error Resume Next
    Set oSPA  = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oRef  = oPart.CreateReferenceFromObject(oHS)
    Set oMeas = oSPA.GetMeasurable(oRef)
    dArea = oMeas.Area
    On Error GoTo 0

    If dArea <= 0 Then Exit Sub

    gTotal = gTotal + 1

    Set oHSF   = oPart.HybridShapeFactory
    Set oRef   = oPart.CreateReferenceFromObject(oHS)
    Set oTmpHB = oPart.HybridBodies.Add()
    oTmpHB.Name = "TMP_AI"
    Set oJoin  = oHSF.AddNewJoin(oRef, oRef)
    oTmpHB.AppendHybridShape oJoin

    bFailed = False
    On Error Resume Next
    oPart.UpdateObject oJoin
    If Err.Number <> 0 Then bFailed = True
    Err.Clear
    On Error GoTo 0

    If bFailed Then
        gFailed = gFailed + 1
        gReport = gReport & "[X] " & oHS.Name & vbCrLf
    Else
        gReport = gReport & "[OK] " & oHS.Name & vbCrLf
    End If

    On Error Resume Next
    Set oSel = oPart.Parent.Selection
    oSel.Clear
    oSel.Add oTmpHB
    oSel.Delete
    oPart.Update
    On Error GoTo 0
End Sub