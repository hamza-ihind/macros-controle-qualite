Sub CATMain()
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument

    Dim docType
    docType = TypeName(oDoc)

    Dim rootName, rootComment
    rootComment = ""

    If docType = "PartDocument" Then
        rootName = oDoc.Product.PartNumber
        On Error Resume Next
        rootComment = oDoc.Product.Comment
        On Error GoTo 0

    ElseIf docType = "ProductDocument" Then
        rootName    = oDoc.Product.PartNumber
        rootComment = oDoc.Product.Comment

    Else
        MsgBox "Open a CATPart or CATProduct.", vbExclamation, "Invalid Document"
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
        nameStatus = "[OK]"
    Else
        nameStatus = "[!!] NOT compliant"
    End If

    If commentOK Then
        commentStatus = "[OK]"
    Else
        commentStatus = "[!!] SE316_PA2 missing"
    End If

    Dim report
    report = "Part Number: " & rootName & "  " & nameStatus & L & _
             "Comment:     " & rootComment & "  " & commentStatus

    If nameOK And commentOK Then
        MsgBox report, vbInformation, "Naming Convention — OK"
    Else
        MsgBox report, vbCritical, "Naming Convention — Issues Found"
    End If
End Sub