Dim gTotal As Integer
Dim gFailed As Integer
Dim gReport As String

' Entry point: detects auto-intersecting surfaces in the active document.
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
        MsgBox "Type de document non pris en charge." & vbCrLf & _
               "Veuillez ouvrir un fichier CATPart ou CATProduct.", _
               vbExclamation, "Controle Auto-Intersection"
        Exit Sub
    End If

    Dim sHeader As String
    Dim sSep    As String
    sSep = String(52, "-") & vbCrLf

    If gTotal = 0 Then
        MsgBox "Aucune surface detectable n'a ete trouvee dans le document." & vbCrLf & _
               "Verifiez que le document contient des corps geometriques.", _
               vbInformation, "Controle Auto-Intersection"
    ElseIf gFailed = 0 Then
        sHeader = "RESULTAT : Toutes les surfaces sont valides." & vbCrLf & _
                  gTotal & " surface(s) analysee(s) -- 0 defaut detecte." & vbCrLf & _
                  sSep
        MsgBox sHeader & gReport, vbInformation, "Controle Auto-Intersection"
    Else
        sHeader = "RESULTAT : Auto-intersection(s) detectee(s) !" & vbCrLf & _
                  gFailed & " surface(s) en defaut sur " & gTotal & " analysee(s)." & vbCrLf & _
                  sSep
        MsgBox sHeader & gReport, vbCritical, "Controle Auto-Intersection"
    End If
End Sub

' Recursively scans a Product tree and processes each Part found.
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

' Processes a Part's HybridBodies for auto-intersection checks.
Sub ProcessPart(oPart As Part)
    If oPart.HybridBodies.Count > 0 Then
        ScanHybridBodies oPart.HybridBodies, oPart
    End If
End Sub

' Iterates all HybridShapes (including nested bodies) and checks each one.
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

' Checks a single surface for auto-intersection using a temporary Join feature.
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

    Dim sName   As String
    Dim sPad    As String
    Dim sStatus As String
    sName = oHS.Name
    If Len(sName) < 40 Then
        sPad = Space(40 - Len(sName))
    Else
        sPad = " "
    End If

    If bFailed Then
        gFailed = gFailed + 1
        sStatus = "Auto-Intersectante"
        gReport = gReport & "  " & sName & sPad & "-->  " & sStatus & vbCrLf
    Else
        sStatus = "OK"
        gReport = gReport & "  " & sName & sPad & "-->  " & sStatus & vbCrLf
    End If

    On Error Resume Next
    Set oSel = oPart.Parent.Selection
    oSel.Clear
    oSel.Add oTmpHB
    oSel.Delete
    oPart.Update
    On Error GoTo 0
End Sub