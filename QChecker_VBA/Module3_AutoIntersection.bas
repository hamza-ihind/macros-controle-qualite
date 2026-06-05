Attribute VB_Name = "Module3_AutoIntersection"
Option Explicit
' ============================================================
' Module 3 : Controle des auto-intersections
' Detecte les surfaces auto-intersectantes via un Join temporaire.
' Reutilise la logique de auto-intersection.vbs sans la modifier.
' ============================================================

' Variables de module (equivalentes aux globales gXxx de la macro originale)
Private m3_Total  As Integer
Private m3_Failed As Integer
Private m3_Report As String

' ------------------------------------------------------------
' Parcourt recursivement un Product pour traiter chaque Part
' ------------------------------------------------------------
Private Sub M3_ScanProduct(oProd As Object)
    Dim i      As Integer
    Dim oChild As Object
    Dim oDoc   As Object
    Dim sType  As String

    For i = 1 To oProd.Products.Count
        Set oChild = oProd.Products.Item(i)
        sType = ""
        On Error Resume Next
        Set oDoc = oChild.ReferenceProduct.Parent
        sType    = TypeName(oDoc)
        On Error GoTo 0
        If sType = "PartDocument" Then
            M3_ProcessPart oDoc.Part
        Else
            M3_ScanProduct oChild
        End If
    Next i
End Sub

' ------------------------------------------------------------
' Traite les HybridBodies d un Part
' ------------------------------------------------------------
Private Sub M3_ProcessPart(oPart As Object)
    If oPart.HybridBodies.Count > 0 Then
        M3_ScanHybridBodies oPart.HybridBodies, oPart
    End If
End Sub

' ------------------------------------------------------------
' Parcourt tous les HybridShapes, y compris les corps imbriques
' ------------------------------------------------------------
Private Sub M3_ScanHybridBodies(oHBs As Object, oPart As Object)
    Dim oHB As Object
    Dim k   As Integer
    Dim oHS As Object

    For Each oHB In oHBs
        For k = 1 To oHB.HybridShapes.Count
            Set oHS = oHB.HybridShapes.Item(k)
            M3_CheckSurface oHS, oPart
        Next k
        ' Recursion sur les corps geometriques imbriques
        If oHB.HybridBodies.Count > 0 Then
            M3_ScanHybridBodies oHB.HybridBodies, oPart
        End If
    Next
End Sub

' ------------------------------------------------------------
' Verifie une surface pour auto-intersection via un Join temporaire.
' Si le Join echoue a la mise a jour, la surface est auto-intersectante.
' ------------------------------------------------------------
Private Sub M3_CheckSurface(oHS As Object, oPart As Object)
    Dim oSPA    As Object
    Dim oRef    As Object
    Dim oMeas   As Object
    Dim dArea   As Double
    Dim oHSF    As Object
    Dim oTmpHB  As Object
    Dim oJoin   As Object
    Dim oSel    As Object
    Dim bFailed As Boolean
    Dim sName   As String
    Dim sPad    As String
    Dim sStatus As String

    ' Calcul de la surface pour verifier que l element est mesurable
    dArea = 0
    On Error Resume Next
    Set oSPA  = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oRef  = oPart.CreateReferenceFromObject(oHS)
    Set oMeas = oSPA.GetMeasurable(oRef)
    dArea     = oMeas.Area
    On Error GoTo 0

    ' Surface non mesurable : on ignore cet element
    If dArea <= 0 Then Exit Sub

    m3_Total = m3_Total + 1

    ' Creation d un Join temporaire pour tester l auto-intersection
    Set oHSF   = oPart.HybridShapeFactory
    Set oRef   = oPart.CreateReferenceFromObject(oHS)
    Set oTmpHB = oPart.HybridBodies.Add()
    oTmpHB.Name = "TMP_AI"
    Set oJoin  = oHSF.AddNewJoin(oRef, oRef)
    oTmpHB.AppendHybridShape oJoin

    ' La mise a jour echoue si la surface est auto-intersectante
    bFailed = False
    On Error Resume Next
    oPart.UpdateObject oJoin
    If Err.Number <> 0 Then bFailed = True
    Err.Clear
    On Error GoTo 0

    sName = oHS.Name
    If Len(sName) < 40 Then
        sPad = Space(40 - Len(sName))
    Else
        sPad = " "
    End If

    If bFailed Then
        m3_Failed = m3_Failed + 1
        sStatus   = "Auto-Intersectante"
        m3_Report = m3_Report & "  " & sName & sPad & "-->  " & sStatus & vbCrLf
    Else
        sStatus   = "OK"
        m3_Report = m3_Report & "  " & sName & sPad & "-->  " & sStatus & vbCrLf
    End If

    ' Suppression du corps geometrique temporaire cree pour le test
    On Error Resume Next
    Set oSel = oPart.Parent.Selection
    oSel.Clear
    oSel.Add oTmpHB
    oSel.Delete
    oPart.Update
    On Error GoTo 0
End Sub

' ============================================================
' Point d entree public - Module 3
' Retourne True si aucune auto-intersection n est detectee.
' sReport : chaine alimentee avec les details du controle.
' ============================================================
Public Function RunCheck3(oDoc As Object, ByRef sReport As String) As Boolean

    ' Initialisation des variables de module
    m3_Total  = 0
    m3_Failed = 0
    m3_Report = ""

    Select Case TypeName(oDoc)
        Case "PartDocument"
            M3_ProcessPart oDoc.Part
        Case "ProductDocument"
            M3_ScanProduct oDoc.Product
        Case Else
            sReport   = sReport & "  Type de document non pris en charge." & vbCrLf
            RunCheck3 = False
            Exit Function
    End Select

    ' Alimentation du rapport global
    If m3_Total = 0 Then
        sReport   = sReport & "  Aucune surface detectable trouvee dans le document." & vbCrLf
        RunCheck3 = True
    ElseIf m3_Failed = 0 Then
        sReport   = sReport & "  " & m3_Total & " surface(s) analysee(s) -- 0 auto-intersection detectee." & vbCrLf
        RunCheck3 = True
    Else
        sReport   = sReport & "  " & m3_Failed & " surface(s) auto-intersectante(s) sur " & m3_Total & " analysee(s)." & vbCrLf
        sReport   = sReport & m3_Report
        RunCheck3 = False
    End If

End Function
