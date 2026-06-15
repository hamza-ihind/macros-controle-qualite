Dim g_oSPA
Dim g_oPart
Dim g_oMainDoc
Dim g_sReport
Dim g_bAllOK
Dim g_iTotal
Dim g_iBad

Sub HighlightBody(oBody)
    Dim oVP
    On Error Resume Next
    g_oMainDoc.Selection.Add oBody
    Set oVP = g_oMainDoc.Selection.VisProperties
    oVP.SetRealColor 255, 128, 0, 0
    Set oVP = Nothing
    Err.Clear
End Sub

Sub AnalyseBody(oBody)
    Dim oRef, oMeasure, dVol_mm3, nShapes, sName, sPad

    sName = oBody.Name
    If Len(sName) < 40 Then
        sPad = Space(40 - Len(sName))
    Else
        sPad = " "
    End If

    g_iTotal = g_iTotal + 1

    On Error Resume Next

    Err.Clear
    nShapes = oBody.Shapes.Count
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  " & sName & sPad & "-->  Lecture impossible" & vbCrLf
        g_bAllOK = False : g_iBad = g_iBad + 1
        HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    If nShapes = 0 Then
        g_sReport = g_sReport & "  " & sName & sPad & "-->  Corps vide" & vbCrLf
        g_bAllOK = False : g_iBad = g_iBad + 1
        HighlightBody oBody
        Exit Sub
    End If

    Err.Clear
    Set oRef = g_oPart.CreateReferenceFromObject(oBody)
    If Err.Number <> 0 Then
        Err.Clear
        Set oMeasure = g_oSPA.GetMeasurable(oBody)
    Else
        Set oMeasure = g_oSPA.GetMeasurable(oRef)
    End If

    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  " & sName & sPad & "-->  Mesure impossible" & vbCrLf
        g_bAllOK = False : g_iBad = g_iBad + 1
        HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    Err.Clear
    dVol_mm3 = oMeasure.Volume * 1000000000
    Set oMeasure = Nothing

    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  " & sName & sPad & "-->  Volume illisible" & vbCrLf
        g_bAllOK = False : g_iBad = g_iBad + 1
        HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    If dVol_mm3 <= 0 Then
        g_sReport = g_sReport & "  " & sName & sPad & "-->  Non solide" & vbCrLf
        g_bAllOK = False : g_iBad = g_iBad + 1
        HighlightBody oBody
    End If

End Sub

Sub AnalysePart(oPartDoc)
    Dim oBodies, i, iBadBefore, iLen, sNew

    On Error Resume Next

    Set g_oPart = oPartDoc.Part
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  " & oPartDoc.Name & Space(20) & "-->  Part inaccessible" & vbCrLf
        g_bAllOK = False
        Err.Clear : Exit Sub
    End If

    Set g_oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If Err.Number <> 0 Then
        g_sReport = g_sReport & "  " & g_oPart.Name & Space(20) & "-->  SPA inaccessible" & vbCrLf
        g_bAllOK = False
        Err.Clear : Exit Sub
    End If

    Set oBodies = g_oPart.Bodies
    If Err.Number <> 0 Or oBodies.Count = 0 Then
        Err.Clear
        Set oBodies = Nothing : Exit Sub
    End If

    iBadBefore = g_iBad
    iLen       = Len(g_sReport)

    For i = 1 To oBodies.Count
        AnalyseBody oBodies.Item(i)
    Next

    If g_iBad > iBadBefore Then
        sNew      = Mid(g_sReport, iLen + 1)
        g_sReport = Left(g_sReport, iLen) & g_oPart.Name & vbCrLf & String(52, "-") & vbCrLf & sNew
    End If

    Set oBodies = Nothing
End Sub

Sub WalkProduct(oProd)
    Dim i, oSubDoc

    On Error Resume Next

    If oProd.Products.Count = 0 Then
        Err.Clear
        Set oSubDoc = oProd.ReferenceProduct.Parent
        If Err.Number <> 0 Then
            g_sReport = g_sReport & "  " & oProd.Name & Space(20) & "-->  Document inaccessible" & vbCrLf
            g_bAllOK = False
            Err.Clear : Exit Sub
        End If
        If TypeName(oSubDoc) = "PartDocument" Then
            AnalysePart oSubDoc
        End If
    Else
        For i = 1 To oProd.Products.Count
            WalkProduct oProd.Products.Item(i)
        Next
    End If
End Sub

Sub CATMain()
    Dim oDoc, sSep, sHeader

    On Error Resume Next

    Set oDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Document actif inaccessible. Verifiez que CATIA est pret.", _
               vbCritical, "Verification Solide"
        Exit Sub
    End If

    Set g_oMainDoc = oDoc
    g_bAllOK  = True
    g_sReport = ""
    g_iTotal  = 0
    g_iBad    = 0
    sSep      = String(52, "-") & vbCrLf

    g_oMainDoc.Selection.Clear

    Select Case TypeName(oDoc)
        Case "PartDocument"
            AnalysePart oDoc
        Case "ProductDocument"
            WalkProduct oDoc.Product
        Case Else
            MsgBox "Type de document non pris en charge." & vbCrLf & _
                   "Veuillez ouvrir un fichier CATPart ou CATProduct.", _
                   vbExclamation, "Verification Solide"
            Exit Sub
    End Select

    If g_iTotal = 0 Then
        MsgBox "Aucun corps detecte dans le document.", vbInformation, "Verification Solide"
        Exit Sub
    End If

    If g_bAllOK Then
        sHeader = "RESULTAT : Tous les corps sont solides et pleins." & vbCrLf & _
                  g_iTotal & " corps analyse(s) -- 0 defaut detecte." & vbCrLf & sSep
        MsgBox sHeader, vbInformation, "Verification Solide"
    Else
        sHeader = "RESULTAT : Corps non conformes detectes !" & vbCrLf & _
                  g_iBad & " corps en defaut sur " & g_iTotal & " analyse(s)." & vbCrLf & sSep
        MsgBox sHeader & g_sReport, vbCritical, "Verification Solide"
    End If

    Set g_oSPA     = Nothing
    Set g_oPart    = Nothing
    Set g_oMainDoc = Nothing
    Set oDoc       = Nothing

End Sub