Const SEUIL_MM2    = 0.01
Const SEUIL_DELETE = 0.0001
Const HEALING_DIST = 0.1

Sub CATMain()

    Dim oDoc, iSliver, iTotal, sList, iRep, i, sSep, sHeader
    Set oDoc  = CATIA.ActiveDocument
    iSliver = 0 : iTotal = 0 : sList = ""
    sSep = String(52, "-") & vbCrLf

    Select Case TypeName(oDoc)

        Case "PartDocument"

            If oDoc.Part.HybridBodies.Count = 0 Then
                MsgBox "Aucun corps geometrique detecte dans le Part actif." & vbCrLf & _
                       "Verifiez que la geometrie surfacique est bien presente.", _
                       vbExclamation, "Silver Faces"
                Exit Sub
            End If

            On Error Resume Next
            ScanPart oDoc.Part, iSliver, iTotal, sList
            If Err.Number <> 0 Then
                MsgBox "Une erreur est survenue lors de l'analyse." & vbCrLf & _
                       "Verifiez que le Part est a jour (Ctrl+U).", vbCritical, "Silver Faces"
                On Error GoTo 0 : Exit Sub
            End If
            On Error GoTo 0

        Case "ProductDocument"

            Dim bFound : bFound = False
            For i = 1 To CATIA.Documents.Count
                If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                    Dim oCurPart
                    Set oCurPart = CATIA.Documents.Item(i).Part
                    If oCurPart.HybridBodies.Count > 0 Then
                        bFound = True
                        On Error Resume Next
                        ScanPart oCurPart, iSliver, iTotal, sList
                        If Err.Number <> 0 Then
                            MsgBox "Une erreur est survenue lors de l'analyse." & vbCrLf & _
                                   "Verifiez que le Part est a jour (Ctrl+U).", vbCritical, "Silver Faces"
                            On Error GoTo 0 : Exit Sub
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next

            If Not bFound Then
                MsgBox "Aucun CATPart avec geometrie surfacique trouve dans le produit.", _
                       vbExclamation, "Silver Faces"
                Exit Sub
            End If

        Case Else
            MsgBox "Type de document non pris en charge." & vbCrLf & _
                   "Veuillez ouvrir un fichier CATPart ou CATProduct.", _
                   vbExclamation, "Silver Faces"
            Exit Sub

    End Select

    If iTotal = 0 Then
        MsgBox "Aucune surface detectable trouvee dans le document.", _
               vbInformation, "Silver Faces"
        Exit Sub
    End If

    If iSliver = 0 Then
        sHeader = "RESULTAT : Toutes les faces sont valides." & vbCrLf & _
                  iTotal & " face(s) analysee(s) -- 0 silver face detectee." & vbCrLf & sSep
        MsgBox sHeader & sList, vbInformation, "Silver Faces"
        Exit Sub
    End If

    sHeader = "RESULTAT : Silver face(s) detectee(s) !" & vbCrLf & _
              iSliver & " face(s) en defaut sur " & iTotal & " analysee(s)." & vbCrLf & sSep

    iRep = MsgBox(sHeader & sList & vbCrLf & _
                  "Appliquer les corrections automatiques ?", _
                  vbYesNo + vbCritical, "Silver Faces")

    If iRep = vbYes Then
        Select Case TypeName(oDoc)
            Case "PartDocument"
                On Error Resume Next
                CorrigerPart oDoc.Part
                If Err.Number <> 0 Then
                    MsgBox "Une erreur est survenue lors de la correction." & vbCrLf & _
                           "Verifiez que le Part est a jour (Ctrl+U).", vbCritical, "Silver Faces"
                    On Error GoTo 0 : Exit Sub
                End If
                On Error GoTo 0

            Case "ProductDocument"
                For i = 1 To CATIA.Documents.Count
                    If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                        If CATIA.Documents.Item(i).Part.HybridBodies.Count > 0 Then
                            On Error Resume Next
                            CorrigerPart CATIA.Documents.Item(i).Part
                            If Err.Number <> 0 Then
                                MsgBox "Une erreur est survenue lors de la correction." & vbCrLf & _
                                       "Verifiez que le Part est a jour (Ctrl+U).", vbCritical, "Silver Faces"
                                On Error GoTo 0 : Exit Sub
                            End If
                            On Error GoTo 0
                        End If
                    End If
                Next
        End Select
    End If

End Sub

Sub ScanPart(ByRef oPart, ByRef iSliver, ByRef iTotal, ByRef sList)

    Dim oSPA, oHBs, oHB, oShape, oRef, oM
    Dim dAire, sName, sPad, sStatus, i, j

    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHBs = oPart.HybridBodies

    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        For j = 1 To oHB.HybridShapes.Count
            Set oShape = oHB.HybridShapes.Item(j)
            iTotal = iTotal + 1
            dAire  = 0

            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000
            On Error GoTo 0

            sName = oShape.Name
            If Len(sName) < 40 Then
                sPad = Space(40 - Len(sName))
            Else
                sPad = " "
            End If

            If dAire > 0 And dAire < SEUIL_MM2 Then
                iSliver  = iSliver + 1
                sStatus = "Silver Face"
            Else
                sStatus = "OK"
            End If

            sList = sList & "  " & sName & sPad & "-->  " & sStatus & vbCrLf

            Set oM     = Nothing
            Set oRef   = Nothing
            Set oShape = Nothing
        Next
        Set oHB = Nothing
    Next

    Set oHBs = Nothing
    Set oSPA = Nothing

End Sub

Sub CorrigerPart(ByRef oPart)

    Dim oSPA, oHSF, oHBs, oHB, oShape, oRef, oM, oHeal, oSel
    Dim dAire, i, j, k, iDel, iHeal, sLog, bNeedHeal, sSep

    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHSF = oPart.HybridShapeFactory
    Set oHBs = oPart.HybridBodies
    iDel = 0 : iHeal = 0 : sLog = ""
    sSep = String(52, "-") & vbCrLf

    For i = 1 To oHBs.Count
        Set oHB   = oHBs.Item(i)
        bNeedHeal = False

        For j = oHB.HybridShapes.Count To 1 Step -1
            Set oShape = oHB.HybridShapes.Item(j)
            dAire = 0

            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000
            On Error GoTo 0

            If dAire > 0 And dAire < SEUIL_DELETE Then
                Set oSel = oPart.Parent.Selection
                On Error Resume Next
                oSel.Clear
                oSel.Add oShape
                oSel.Delete
                On Error GoTo 0
                Set oSel = Nothing
                iDel = iDel + 1
                sLog = sLog & "  [Supprimee]  " & oShape.Name & _
                       "  (" & FormatNumber(dAire, 6) & " mm2)" & vbCrLf

            ElseIf dAire > 0 And dAire < SEUIL_MM2 Then
                bNeedHeal = True
            End If

            Set oM     = Nothing
            Set oRef   = Nothing
            Set oShape = Nothing
        Next

        If bNeedHeal And oHB.HybridShapes.Count > 0 Then
            On Error Resume Next
            Set oHeal = oHSF.AddNewHeal()
            For k = 1 To oHB.HybridShapes.Count
                Set oRef = oPart.CreateReferenceFromObject(oHB.HybridShapes.Item(k))
                oHeal.AddElement oRef
                Set oRef = Nothing
            Next
            oHeal.MergingDistance = HEALING_DIST
            oHB.AppendHybridShape oHeal
            oPart.Update
            On Error GoTo 0
            Set oHeal = Nothing
            iHeal = iHeal + 1
            sLog = sLog & "  [Healing]    " & oHB.Name & _
                   "  (dist = " & HEALING_DIST & " mm)" & vbCrLf
        End If

        Set oHB = Nothing
    Next

    Set oHBs = Nothing
    Set oHSF = Nothing
    Set oSPA = Nothing

    MsgBox "Corrections appliquees sur : " & oPart.Name & vbCrLf & sSep & _
           "  Suppressions : " & iDel  & vbCrLf & _
           "  Healings     : " & iHeal & vbCrLf & vbCrLf & sLog, _
           vbInformation, "Silver Faces"

End Sub