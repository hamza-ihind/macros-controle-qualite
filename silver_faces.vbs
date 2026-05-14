Const SEUIL_MM2    = 0.01
Const SEUIL_DELETE = 0.0001
Const HEALING_DIST = 0.1

' Entry point: scans the active document for silver faces and offers auto-correction.
Sub CATMain()

    Dim oDoc, iSliver, iTotal, sList, sMsg, iRep, i
    Set oDoc  = CATIA.ActiveDocument
    iSliver = 0 : iTotal = 0 : sList = ""

    On Error GoTo ErrHandler

    Select Case TypeName(oDoc)

        Case "PartDocument"
            ScanPart oDoc.Part, iSliver, iTotal, sList

        Case "ProductDocument"
            For i = 1 To CATIA.Documents.Count
                If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                    ScanPart CATIA.Documents.Item(i).Part, iSliver, iTotal, sList
                End If
            Next i

        Case Else
            MsgBox "Ouvrez un .CATPart ou .CATProduct.", 16, "Silver Faces"
            Exit Sub

    End Select

    ' --- Rapport ---
    sMsg = "=== SILVER FACES ===" & Chr(13) & _
           "Seuil    : " & SEUIL_MM2 & " mm2" & Chr(13) & _
           "Surfaces : " & iTotal    & Chr(13) & _
           "Anomalies: " & iSliver   & Chr(13) & _
           "===================="

    If iSliver = 0 Then
        MsgBox sMsg & Chr(13) & Chr(13) & "[OK] Aucune silver face detectee.", _
               64, "Silver Faces"
        Exit Sub
    End If

    ' --- Proposer la correction automatique ---
    iRep = MsgBox(sMsg & Chr(13) & Chr(13) & sList & Chr(13) & _
                  "Appliquer les corrections automatiques ?", _
                  36, "Silver Faces")  ' 36 = Oui/Non

    If iRep = 6 Then  ' vbYes = 6
        Select Case TypeName(oDoc)
            Case "PartDocument"
                CorrigerPart oDoc.Part
            Case "ProductDocument"
                For i = 1 To CATIA.Documents.Count
                    If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                        CorrigerPart CATIA.Documents.Item(i).Part
                    End If
                Next i
        End Select
    End If

    Exit Sub

ErrHandler:
    MsgBox "Erreur #" & Err.Number & " : " & Err.Description & Chr(13) & _
           "Verifiez que le Part est mis a jour (Ctrl+U).", 16, "Silver Faces"

End Sub

' Scans all HybridShapes in a Part and flags silver faces.
Sub ScanPart(oPart, iSliver, iTotal, sList)

    Dim oSPA, oHBs, oHB, oShape, oRef, oM, dAire, sStrat, i, j

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
            dAire    = oM.Area * 1000000  ' m2 -> mm2
            On Error GoTo 0

            If dAire > 0 And dAire < SEUIL_MM2 Then
                iSliver = iSliver + 1

                ' Pick correction strategy based on area size
                If dAire < SEUIL_DELETE Then
                    sStrat = "[A] Suppression"
                Else
                    sStrat = "[B] Healing"
                End If

                sList = sList & iSliver & ". " & oShape.Name & _
                        " (" & oPart.Name & "/" & oHB.Name & ")" & _
                        " = " & FormatNumber(dAire, 6) & " mm2  " & sStrat & Chr(13)
            End If
        Next j
    Next i

End Sub

' Corrects silver faces: deletes near-zero area shapes [A], applies Healing to the rest [B].
Sub CorrigerPart(oPart)

    Dim oSPA, oHSF, oHBs, oHB, oShape, oRef, oM, oHeal
    Dim dAire, i, j, k, iDel, iHeal, sLog, bNeedHeal

    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHSF = oPart.HybridShapeFactory
    Set oHBs = oPart.HybridBodies
    iDel = 0 : iHeal = 0 : sLog = ""

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

            ' [A] Delete quasi-null area shape
            If dAire > 0 And dAire < SEUIL_DELETE Then
                On Error Resume Next
                oPart.Parent.Selection.Clear
                oPart.Parent.Selection.Add oShape
                oPart.Parent.Selection.Delete
                On Error GoTo 0
                iDel = iDel + 1
                sLog = sLog & "[A-SUPPRIME] " & oShape.Name & _
                       " (" & FormatNumber(dAire, 6) & " mm2)" & Chr(13)

            ElseIf dAire > 0 And dAire < SEUIL_MM2 Then
                bNeedHeal = True
            End If
        Next j

        ' [B] Create a Healing feature on the body if any sliver was found
        If bNeedHeal Then
            On Error Resume Next
            Set oHeal = oHSF.AddNewHeal()
            For k = 1 To oHB.HybridShapes.Count
                Set oRef = oPart.CreateReferenceFromObject(oHB.HybridShapes.Item(k))
                oHeal.AddElement oRef
            Next k
            oHeal.MergingDistance = HEALING_DIST
            oHB.AppendHybridShape oHeal
            oPart.Update
            On Error GoTo 0
            iHeal = iHeal + 1
            sLog = sLog & "[B-HEALING]  " & oHB.Name & _
                   " (dist=" & HEALING_DIST & " mm)" & Chr(13)
        End If

    Next i

    MsgBox "Corrections sur : " & oPart.Name & Chr(13) & Chr(13) & _
           "[A] Suppressions : " & iDel  & Chr(13) & _
           "[B] Healings     : " & iHeal & Chr(13) & Chr(13) & sLog & Chr(13) & _
           "Strategies non automatisees (manuel) :" & Chr(13) & _
           "  - Face apres Trim  -> Delete Face + Fill" & Chr(13) & _
           "  - Offset / Import  -> voir logiciel source", _
           64, "Silver Faces — Corrections"
End Sub