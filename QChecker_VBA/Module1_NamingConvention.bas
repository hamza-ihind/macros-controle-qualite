Attribute VB_Name = "Module1_NamingConvention"
Option Explicit
' ============================================================
' Module 1 : Convention de nommage
' Verifie le PartNumber et le commentaire du document actif.
' Reutilise la logique de name_convention.vbs sans la modifier.
' ============================================================

' ------------------------------------------------------------
' Verifie la convention de nommage du document passe en parametre.
' Retourne True si conforme, False sinon.
' sReport : chaine alimentee avec les details du controle.
' ------------------------------------------------------------
Public Function RunCheck1(oDoc As Object, ByRef sReport As String) As Boolean

    Dim docType       As String
    Dim rootName      As String
    Dim rootComment   As String
    Dim regEx         As Object
    Dim nameOK        As Boolean
    Dim commentOK     As Boolean
    Dim nameStatus    As String
    Dim commentStatus As String

    rootName    = ""
    rootComment = ""
    docType     = TypeName(oDoc)

    ' Recuperation du PartNumber et du commentaire selon le type de document
    If docType = "PartDocument" Then
        On Error Resume Next
        rootName    = oDoc.Product.PartNumber
        rootComment = oDoc.Product.Comment
        On Error GoTo 0

    ElseIf docType = "ProductDocument" Then
        On Error Resume Next
        rootName    = oDoc.Product.PartNumber
        rootComment = oDoc.Product.Comment
        On Error GoTo 0

    Else
        sReport = sReport & "  Type de document non pris en charge." & vbCrLf
        RunCheck1 = False
        Exit Function
    End If

    ' Verification du pattern de nommage via expression reguliere
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern    = "^[A-Z0-9]{3}\d{6}[A-Z0-9]_(DMU|PCA|GEO|VEE)_(TM|EN)_\d+_[A-Z][A-Z0-9_]*$"
    regEx.IgnoreCase = False

    nameOK = regEx.Test(rootName)

    ' Verification de la presence de la reference projet SE316_PA2 dans le commentaire
    commentOK = (InStr(1, rootComment, "SE316_PA2", vbBinaryCompare) > 0)

    Set regEx = Nothing

    ' Formatage du statut de chaque critere
    If nameOK Then
        nameStatus = "[OK]"
    Else
        nameStatus = "[NON CONFORME]"
    End If

    If commentOK Then
        commentStatus = "[OK]"
    Else
        commentStatus = "[NON CONFORME] SE316_PA2 manquant"
    End If

    ' Alimentation du rapport
    sReport = sReport & "  Part Number : " & rootName    & "  " & nameStatus    & vbCrLf
    sReport = sReport & "  Commentaire : " & rootComment & "  " & commentStatus & vbCrLf

    RunCheck1 = (nameOK And commentOK)

End Function
