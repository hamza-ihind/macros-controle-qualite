Attribute VB_Name = "Module2_SolidCheck"
Option Explicit
' ============================================================
' Module 2 : Verification solidite / volume
' Verifie que tous les corps ont un volume positif (solide plein).
' Reutilise la logique de solide-plein.vbs sans la modifier.
' ============================================================

' Variables de module (equivalentes aux g_xxx de la macro originale)
Private m2_oSPA     As Object
Private m2_oPart    As Object
Private m2_oMainDoc As Object
Private m2_sReport  As String
Private m2_bAllOK   As Boolean
Private m2_iTotal   As Integer
Private m2_iBad     As Integer

' ------------------------------------------------------------
' Met en evidence un corps defectueux en orange dans l'arbre CATIA
' ------------------------------------------------------------
Private Sub M2_HighlightBody(oBody As Object)
    Dim oVP As Object
    On Error Resume Next
    m2_oMainDoc.Selection.Add oBody
    Set oVP = m2_oMainDoc.Selection.VisProperties
    oVP.SetRealColor 255, 128, 0, 0
    Set oVP = Nothing
    Err.Clear
End Sub

' ------------------------------------------------------------
' Analyse un corps solide et mesure son volume
' ------------------------------------------------------------
Private Sub M2_AnalyseBody(oBody As Object)
    Dim oRef     As Object
    Dim oMeasure As Object
    Dim dVol_mm3 As Double
    Dim nShapes  As Integer
    Dim sName    As String
    Dim sPad     As String

    sName = oBody.Name
    If Len(sName) < 40 Then
        sPad = Space(40 - Len(sName))
    Else
        sPad = " "
    End If

    m2_iTotal = m2_iTotal + 1

    On Error Resume Next

    ' Verification de l accessibilite du corps
    Err.Clear
    nShapes = oBody.Shapes.Count
    If Err.Number <> 0 Then
        m2_sReport = m2_sReport & "  " & sName & sPad & "-->  Lecture impossible" & vbCrLf
        m2_bAllOK = False : m2_iBad = m2_iBad + 1
        M2_HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    If nShapes = 0 Then
        m2_sReport = m2_sReport & "  " & sName & sPad & "-->  Corps vide" & vbCrLf
        m2_bAllOK = False : m2_iBad = m2_iBad + 1
        M2_HighlightBody oBody
        Exit Sub
    End If

    ' Obtention de la mesure de volume
    Err.Clear
    Set oRef = m2_oPart.CreateReferenceFromObject(oBody)
    If Err.Number <> 0 Then
        Err.Clear
        Set oMeasure = m2_oSPA.GetMeasurable(oBody)
    Else
        Set oMeasure = m2_oSPA.GetMeasurable(oRef)
    End If

    If Err.Number <> 0 Then
        m2_sReport = m2_sReport & "  " & sName & sPad & "-->  Mesure impossible" & vbCrLf
        m2_bAllOK = False : m2_iBad = m2_iBad + 1
        M2_HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    Err.Clear
    dVol_mm3 = oMeasure.Volume * 1000000000  ' Conversion m3 en mm3
    Set oMeasure = Nothing

    If Err.Number <> 0 Then
        m2_sReport = m2_sReport & "  " & sName & sPad & "-->  Volume illisible" & vbCrLf
        m2_bAllOK = False : m2_iBad = m2_iBad + 1
        M2_HighlightBody oBody
        Err.Clear : Exit Sub
    End If

    ' Verification que le volume est positif (corps plein)
    If dVol_mm3 <= 0 Then
        m2_sReport = m2_sReport & "  " & sName & sPad & "-->  Non solide" & vbCrLf
        m2_bAllOK = False : m2_iBad = m2_iBad + 1
        M2_HighlightBody oBody
    End If

End Sub

' ------------------------------------------------------------
' Analyse tous les corps d un document Part
' ------------------------------------------------------------
Private Sub M2_AnalysePart(oPartDoc As Object)
    Dim oBodies    As Object
    Dim i          As Integer
    Dim iBadBefore As Integer
    Dim iLen       As Integer
    Dim sNew       As String

    On Error Resume Next

    Set m2_oPart = oPartDoc.Part
    If Err.Number <> 0 Then
        m2_sReport = m2_sReport & "  " & oPartDoc.Name & Space(20) & "-->  Part inaccessible" & vbCrLf
        m2_bAllOK = False
        Err.Clear : Exit Sub
    End If

    Set m2_oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    If Err.Number <> 0 Then
        m2_sReport = m2_sReport & "  " & m2_oPart.Name & Space(20) & "-->  SPA inaccessible" & vbCrLf
        m2_bAllOK = False
        Err.Clear : Exit Sub
    End If

    Set oBodies = m2_oPart.Bodies
    If Err.Number <> 0 Or oBodies.Count = 0 Then
        Err.Clear
        Set oBodies = Nothing : Exit Sub
    End If

    iBadBefore = m2_iBad
    iLen       = Len(m2_sReport)

    For i = 1 To oBodies.Count
        M2_AnalyseBody oBodies.Item(i)
    Next i

    ' Prefixe le nom du Part dans le rapport si des defauts ont ete trouves
    If m2_iBad > iBadBefore Then
        sNew       = Mid(m2_sReport, iLen + 1)
        m2_sReport = Left(m2_sReport, iLen) & m2_oPart.Name & vbCrLf & _
                     String(52, "-") & vbCrLf & sNew
    End If

    Set oBodies = Nothing
End Sub

' ------------------------------------------------------------
' Parcourt recursivement l arbre d un CATProduct
' ------------------------------------------------------------
Private Sub M2_WalkProduct(oProd As Object)
    Dim i       As Integer
    Dim oSubDoc As Object

    On Error Resume Next

    If oProd.Products.Count = 0 Then
        Err.Clear
        Set oSubDoc = oProd.ReferenceProduct.Parent
        If Err.Number <> 0 Then
            m2_sReport = m2_sReport & "  " & oProd.Name & Space(20) & "-->  Document inaccessible" & vbCrLf
            m2_bAllOK = False
            Err.Clear : Exit Sub
        End If
        If TypeName(oSubDoc) = "PartDocument" Then
            M2_AnalysePart oSubDoc
        End If
    Else
        For i = 1 To oProd.Products.Count
            M2_WalkProduct oProd.Products.Item(i)
        Next i
    End If
End Sub

' ============================================================
' Point d entree public - Module 2
' Retourne True si tous les corps sont solides et pleins.
' sReport : chaine alimentee avec les details du controle.
' ============================================================
Public Function RunCheck2(oDoc As Object, ByRef sReport As String) As Boolean

    ' Initialisation des variables de module
    m2_bAllOK  = True
    m2_sReport = ""
    m2_iTotal  = 0
    m2_iBad    = 0
    Set m2_oMainDoc = oDoc

    On Error Resume Next
    m2_oMainDoc.Selection.Clear
    On Error GoTo 0

    Select Case TypeName(oDoc)
        Case "PartDocument"
            M2_AnalysePart oDoc
        Case "ProductDocument"
            M2_WalkProduct oDoc.Product
        Case Else
            sReport = sReport & "  Type de document non pris en charge." & vbCrLf
            RunCheck2 = False
            Exit Function
    End Select

    ' Alimentation du rapport global
    If m2_iTotal = 0 Then
        sReport  = sReport & "  Aucun corps detectable dans le document." & vbCrLf
        RunCheck2 = True
    ElseIf m2_bAllOK Then
        sReport  = sReport & "  " & m2_iTotal & " corps analyse(s) -- 0 defaut detecte." & vbCrLf
        RunCheck2 = True
    Else
        sReport  = sReport & "  " & m2_iBad & " corps en defaut sur " & m2_iTotal & " analyse(s)." & vbCrLf
        sReport  = sReport & m2_sReport
        RunCheck2 = False
    End If

    ' Nettoyage des references objet
    Set m2_oSPA     = Nothing
    Set m2_oPart    = Nothing
    Set m2_oMainDoc = Nothing

End Function
