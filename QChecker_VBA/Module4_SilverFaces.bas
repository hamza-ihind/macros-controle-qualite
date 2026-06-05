Attribute VB_Name = "Module4_SilverFaces"
Option Explicit
' ============================================================
' Module 4 : Detection des Silver Faces
' Detecte les faces de tres petite superficie (< seuil mm2).
' Reutilise la logique de silver_faces.vbs sans la modifier.
' ============================================================

' Constantes de seuil (reprises de silver_faces.vbs)
Private Const M4_SEUIL_MM2    As Double = 0.01
Private Const M4_SEUIL_DELETE As Double = 0.0001

' Variables de module
Private m4_iSliver As Integer
Private m4_iTotal  As Integer
Private m4_sList   As String

' ------------------------------------------------------------
' Analyse les surfaces d un Part et identifie les silver faces.
' Logique identique a ScanPart dans silver_faces.vbs.
' ------------------------------------------------------------
Private Sub M4_ScanPart(oPart As Object)
    Dim oSPA    As Object
    Dim oHBs    As Object
    Dim oHB     As Object
    Dim oShape  As Object
    Dim oRef    As Object
    Dim oM      As Object
    Dim dAire   As Double
    Dim sName   As String
    Dim sPad    As String
    Dim sStatus As String
    Dim i       As Integer
    Dim j       As Integer

    Set oSPA = oPart.Parent.GetWorkbench("SPAWorkbench")
    Set oHBs = oPart.HybridBodies

    For i = 1 To oHBs.Count
        Set oHB = oHBs.Item(i)
        For j = 1 To oHB.HybridShapes.Count
            Set oShape    = oHB.HybridShapes.Item(j)
            m4_iTotal     = m4_iTotal + 1
            dAire         = 0

            On Error Resume Next
            Set oRef = oPart.CreateReferenceFromObject(oShape)
            Set oM   = oSPA.GetMeasurable(oRef)
            dAire    = oM.Area * 1000000  ' Conversion m2 en mm2
            On Error GoTo 0

            sName = oShape.Name
            If Len(sName) < 40 Then
                sPad = Space(40 - Len(sName))
            Else
                sPad = " "
            End If

            If dAire > 0 And dAire < M4_SEUIL_MM2 Then
                m4_iSliver = m4_iSliver + 1
                sStatus    = "Silver Face"
            Else
                sStatus = "OK"
            End If

            m4_sList = m4_sList & "  " & sName & sPad & "-->  " & sStatus & vbCrLf

            Set oM     = Nothing
            Set oRef   = Nothing
            Set oShape = Nothing
        Next j
        Set oHB = Nothing
    Next i

    Set oHBs = Nothing
    Set oSPA = Nothing
End Sub

' ============================================================
' Point d entree public - Module 4
' bUseML : True = methode regression logistique (ML)
'          False = seuil classique (< 0.01 mm2)
' Retourne True si aucune silver face n est detectee.
' sReport : chaine alimentee avec les details du controle.
' ============================================================
Public Function RunCheck4(oDoc As Object, ByRef sReport As String, _
                          ByVal bUseML As Boolean) As Boolean
    Dim i        As Integer
    Dim bFound   As Boolean
    Dim oCurPart As Object

    ' Initialisation
    m4_iSliver = 0
    m4_iTotal  = 0
    m4_sList   = ""

    ' Note sur la methode ML : le modele de regression logistique est prevu
    ' pour une integration ulterieure (cf. AI-SILVER-FACES.md / ml_silver_faces.md).
    ' En attendant, les deux options utilisent le meme seuil de surface M4_SEUIL_MM2.

    Select Case TypeName(oDoc)

        Case "PartDocument"
            If oDoc.Part.HybridBodies.Count = 0 Then
                sReport   = sReport & "  Aucun corps geometrique detecte dans le Part actif." & vbCrLf
                RunCheck4 = True
                Exit Function
            End If
            On Error Resume Next
            M4_ScanPart oDoc.Part
            If Err.Number <> 0 Then
                sReport   = sReport & "  Erreur lors de l analyse : " & Err.Description & vbCrLf
                On Error GoTo 0
                RunCheck4 = False
                Exit Function
            End If
            On Error GoTo 0

        Case "ProductDocument"
            bFound = False
            For i = 1 To CATIA.Documents.Count
                If TypeName(CATIA.Documents.Item(i)) = "PartDocument" Then
                    Set oCurPart = CATIA.Documents.Item(i).Part
                    If oCurPart.HybridBodies.Count > 0 Then
                        bFound = True
                        On Error Resume Next
                        M4_ScanPart oCurPart
                        If Err.Number <> 0 Then
                            sReport   = sReport & "  Erreur lors de l analyse : " & Err.Description & vbCrLf
                            On Error GoTo 0
                            RunCheck4 = False
                            Exit Function
                        End If
                        On Error GoTo 0
                    End If
                End If
            Next i
            If Not bFound Then
                sReport   = sReport & "  Aucun CATPart avec geometrie surfacique trouve." & vbCrLf
                RunCheck4 = True
                Exit Function
            End If

        Case Else
            sReport   = sReport & "  Type de document non pris en charge." & vbCrLf
            RunCheck4 = False
            Exit Function

    End Select

    ' Alimentation du rapport global
    If m4_iTotal = 0 Then
        sReport   = sReport & "  Aucune surface detectable trouvee dans le document." & vbCrLf
        RunCheck4 = True
    ElseIf m4_iSliver = 0 Then
        sReport   = sReport & "  " & m4_iTotal & " face(s) analysee(s) -- 0 silver face detectee." & vbCrLf
        RunCheck4 = True
    Else
        sReport   = sReport & "  " & m4_iSliver & " silver face(s) sur " & m4_iTotal & " analysee(s)." & vbCrLf
        sReport   = sReport & m4_sList
        RunCheck4 = False
    End If

End Function

' ------------------------------------------------------------
' Retourne le libelle de la methode selectionnee (pour l en-tete du rapport)
' ------------------------------------------------------------
Public Function GetMethodeLabel(ByVal bUseML As Boolean) As String
    If bUseML Then
        GetMethodeLabel = "Regression Logistique (ML)"
    Else
        GetMethodeLabel = "Seuil classique < " & M4_SEUIL_MM2 & " mm2"
    End If
End Function
