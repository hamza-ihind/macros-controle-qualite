Attribute VB_Name = "Module5_Report"
Option Explicit
' ============================================================
' Module 5 : Generation et export du rapport Q-Checker
' Fonctions utilitaires pour le formatage et l export du rapport.
' ============================================================

' Separateurs visuels reutilises dans le rapport
Public Const RPT_SEP  As String = "----------------------------------------"
Public Const RPT_SEP2 As String = "========================================"

' ------------------------------------------------------------
' Genere l en-tete du rapport avec nom, date, chemin et type du document.
' ------------------------------------------------------------
Public Function BuildReportHeader(oDoc As Object) As String
    Dim sHeader As String
    Dim sNom    As String
    Dim sChemin As String
    Dim sType   As String

    On Error Resume Next
    sNom    = oDoc.Name
    sChemin = oDoc.FullName
    sType   = TypeName(oDoc)
    On Error GoTo 0

    If sType = "PartDocument" Then
        sType = "CATPart"
    ElseIf sType = "ProductDocument" Then
        sType = "CATProduct"
    End If

    sHeader = RPT_SEP2 & vbCrLf
    sHeader = sHeader & "=== RAPPORT Q-CHECKER --- " & sNom & " ===" & vbCrLf
    sHeader = sHeader & "Date     : " & Format(Now, "DD/MM/YYYY HH:MM:SS") & vbCrLf
    sHeader = sHeader & "Fichier  : " & sChemin & vbCrLf
    sHeader = sHeader & "Type     : " & sType & vbCrLf
    sHeader = sHeader & RPT_SEP & vbCrLf

    BuildReportHeader = sHeader
End Function

' ------------------------------------------------------------
' Formate une section de module avec en-tete, statut et details.
' iModule   : numero du module (1 a 4)
' sModuleName : libelle du module (peut inclure la methode ex. Silver Faces (ML))
' bOK       : True si conforme, False sinon
' sDetails  : texte brut retourne par RunCheckX
' ------------------------------------------------------------
Public Function BuildModuleSection(ByVal iModule As Integer, _
                                   ByVal sModuleName As String, _
                                   ByVal bOK As Boolean, _
                                   ByVal sDetails As String) As String
    Dim sSection As String
    Dim sStatut  As String

    Select Case iModule
        Case 1, 2
            If bOK Then sStatut = "CONFORME" Else sStatut = "NON CONFORME"
        Case Else  ' Modules 3 et 4
            If bOK Then sStatut = "AUCUNE" Else sStatut = "DETECTEE(S)"
    End Select

    sSection = "[MODULE " & iModule & "] " & sModuleName & vbCrLf
    sSection = sSection & "  Statut : " & sStatut & vbCrLf
    sSection = sSection & sDetails
    sSection = sSection & RPT_SEP & vbCrLf

    BuildModuleSection = sSection
End Function

' ------------------------------------------------------------
' Genere le bilan final du rapport.
' iConformes : nombre de verifications conformes
' iTotal     : nombre total de verifications effectuees
' ------------------------------------------------------------
Public Function BuildReportBilan(ByVal iConformes As Integer, _
                                 ByVal iTotal As Integer) As String
    Dim sBilan As String
    sBilan = RPT_SEP2 & vbCrLf
    sBilan = sBilan & "BILAN : " & iConformes & "/" & iTotal & " verifications conformes" & vbCrLf
    sBilan = sBilan & RPT_SEP2 & vbCrLf
    BuildReportBilan = sBilan
End Function

' ------------------------------------------------------------
' Exporte le contenu du rapport dans un fichier .txt horodate.
' Sauvegarde dans le meme repertoire que le fichier CAD.
' Retourne True si l export a reussi, False sinon.
' ------------------------------------------------------------
Public Function ExportReport(ByVal sContent As String, _
                             ByVal sDocFullName As String) As Boolean
    Dim sDir      As String
    Dim sFileName As String
    Dim sFullPath As String
    Dim iFile     As Integer

    On Error Resume Next

    ' Determination du repertoire de sauvegarde (meme dossier que le CAD)
    If Len(sDocFullName) > 0 Then
        sDir = Left(sDocFullName, InStrRev(sDocFullName, "\"))
    End If
    ' Repli sur le profil utilisateur si le chemin est inaccessible
    If Len(sDir) = 0 Then sDir = Environ("USERPROFILE") & "\"

    ' Nom du fichier horodate pour eviter les ecrasements
    sFileName = "QChecker_Report_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    sFullPath = sDir & sFileName

    iFile = FreeFile
    Open sFullPath For Output As #iFile
    Print #iFile, sContent
    Close #iFile

    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l export du rapport :" & vbCrLf & Err.Description, _
               vbCritical, "Q-Checker --- Export"
        ExportReport = False
        Err.Clear
    Else
        MsgBox "Rapport exporte avec succes :" & vbCrLf & sFullPath, _
               vbInformation, "Q-Checker --- Export"
        ExportReport = True
    End If

    On Error GoTo 0
End Function
