Sub CATMain()
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument

    Dim oSel
    Set oSel = oDoc.Selection
    oSel.Clear

    Dim oDesignWork
    Set oDesignWork = Nothing

    Dim docType
    docType = TypeName(oDoc)

    If docType = "PartDocument" Then
        Call FindDesignWork(oDoc.Part.HybridBodies, oDesignWork)

    ElseIf docType = "ProductDocument" Then
        Dim i
        For i = 1 To CATIA.Documents.Count
            If Not (oDesignWork Is Nothing) Then Exit For
            Dim oCurDoc
            Set oCurDoc = CATIA.Documents.Item(i)
            If TypeName(oCurDoc) = "PartDocument" Then
                Call FindDesignWork(oCurDoc.Part.HybridBodies, oDesignWork)
            End If
        Next

    Else
        MsgBox "Unsupported document. Open a CATPart or CATProduct.", _
               vbExclamation, "Invalid Type"
        Exit Sub
    End If

    If oDesignWork Is Nothing Then
        MsgBox "Could not find a Geometric Set named 'Design_Work' in the tree." & vbCrLf & _
               "Make sure the set exists and is named exactly 'Design_Work'.", _
               vbExclamation, "Design_Work Not Found"
        Exit Sub
    End If

    Dim report
    report = ""
    Dim count
    count = 0

    Call ScanForDatums(oDesignWork, "Design_Work", report, count, oSel)

    Dim L
    L = vbCrLf

    If count = 0 Then
        MsgBox "No DATUM (Élément mort / Élément Isolé) found inside 'Design_Work'." & L & _
               "All features have an active construction history.", _
               vbInformation, "Datum Check — OK"
    Else
        MsgBox count & " DATUM (Élément mort / Élément Isolé) detected inside 'Design_Work'" & L & _
               "(all highlighted in the tree)" & L & _
               "────────────────────────────────────────────" & L & L & _
               report, _
               vbCritical, "Datum Check — " & count & " Issue(s) Found"
    End If
End Sub

Sub FindDesignWork(hbColl, ByRef result)
    If Not (result Is Nothing) Then Exit Sub

    Dim i
    For i = 1 To hbColl.Count
        If Not (result Is Nothing) Then Exit Sub

        Dim hb
        Set hb = hbColl.Item(i)

        If hb.Name = "Design_Work" Then
            Set result = hb
            Exit Sub
        End If

        Call FindDesignWork(hb.HybridBodies, result)
    Next
End Sub

Sub ScanForDatums(oHB, currentPath, ByRef report, ByRef count, oSel)
    Dim nShapes
    nShapes = 0
    On Error Resume Next
    nShapes = oHB.HybridShapes.Count
    On Error GoTo 0

    Dim i
    For i = 1 To nShapes
        Dim oShape
        Set oShape = Nothing
        On Error Resume Next
        Set oShape = oHB.HybridShapes.Item(i)
        On Error GoTo 0

        If Not (oShape Is Nothing) Then
            Dim tName
            tName = TypeName(oShape)
            If InStr(1, tName, "Datum", vbBinaryCompare) > 0 Then
                count = count + 1
                report = report & _
                    "  [" & count & "]  " & oShape.Name & "  (" & tName & ")" & vbCrLf & _
                    "        Path : " & currentPath & " > " & oShape.Name & vbCrLf & vbCrLf
                On Error Resume Next
                oSel.Add oShape
                On Error GoTo 0
            End If
        End If
    Next

    Dim nSubs
    nSubs = 0
    On Error Resume Next
    nSubs = oHB.HybridBodies.Count
    On Error GoTo 0

    Dim j
    For j = 1 To nSubs
        Dim oSubHB
        Set oSubHB = Nothing
        On Error Resume Next
        Set oSubHB = oHB.HybridBodies.Item(j)
        On Error GoTo 0

        If Not (oSubHB Is Nothing) Then
            Call ScanForDatums(oSubHB, currentPath & " > " & oSubHB.Name, _
                               report, count, oSel)
        End If
    Next
End Sub