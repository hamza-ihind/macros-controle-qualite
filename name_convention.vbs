Sub CATMain()
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument

    Dim docType
    docType = TypeName(oDoc)

    Dim rootName, rootComment
    rootComment = ""

    If docType = "PartDocument" Then
        rootName = oDoc.Part.PartNumber
        On Error Resume Next
        rootComment = oDoc.Product.Comment
        On Error GoTo 0

    ElseIf docType = "ProductDocument" Then
        rootName    = oDoc.Product.PartNumber
        rootComment = oDoc.Product.Comment

    Else
        MsgBox "Unsupported document. Open a CATPart or CATProduct.", _
               vbExclamation, "Invalid Type"
        Exit Sub
    End If

    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^[A-Z0-9]{3}\d{6}[A-Z0-9]_(DMU|PCA|GEO|VEE)_(TM|EN)_\d+_[A-Z][A-Z0-9_]*$"
    regEx.IgnoreCase = False

    Dim nameOK
    nameOK = regEx.Test(rootName)

    Dim commentOK
    commentOK = (InStr(1, rootComment, "SE316_PA2", vbBinaryCompare) > 0)

    Dim L
    L = vbCrLf

    Dim nameStatus, commentStatus
    If nameOK Then
        nameStatus = "[OK]  Compliant"
    Else
        nameStatus = "[!!] NOT Compliant"
    End If

    If commentOK Then
        commentStatus = "[OK]  Compliant  —  SE316_PA2 detected"
    Else
        commentStatus = "[!!] NOT Compliant  —  SE316_PA2 missing"
    End If

    Dim report
    report = "────────────────────────────────────────────" & L & _
             "  PART NUMBER" & L & _
             "────────────────────────────────────────────" & L & _
             "  Detected  :  " & rootName & L & _
             "  Status    :  " & nameStatus & L & _
             "  Should be :  KVScode_CADTYPE_PDA_VERSION_DESCRIPTION" & L & _
             "  e.g.         5FF821105H_PCA_TM_5_FENDER_PHEV" & L & L & _
             "────────────────────────────────────────────" & L & _
             "  COMMENT" & L & _
             "────────────────────────────────────────────" & L & _
             "  Detected  :  " & rootComment & L & _
             "  Status    :  " & commentStatus & L & _
             "  Should be :  SE316_PA2_[STATUS]_[DATE]" & L & _
             "  e.g.         SE316_PA2_WIP_250905"

    If nameOK And commentOK Then
        MsgBox report, vbInformation, "Naming Convention — OK"
    Else
        MsgBox report, vbCritical, "Naming Convention — Issues Found"
    End If
End Sub