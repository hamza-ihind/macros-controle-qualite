VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QCheckerForm 
   Caption         =   "CATIA Q-Checker -- Segula Technologies"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   794
      Caption         =   "CATIA Q-Checker -- Segula Technologies"
      TextAlign       =   2
   End
   Begin MSForms.Label lblDocLabel 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   477
      Caption         =   "Document actif :"
   End
   Begin MSForms.Label lblDocName 
      Height          =   270
      Left            =   2520
      TabIndex        =   2
      Top             =   660
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   477
      Caption         =   "(aucun document)"
   End
   Begin MSForms.Label lblTypeLabel 
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   990
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   477
      Caption         =   "Type :"
   End
   Begin MSForms.Label lblDocType 
      Height          =   270
      Left            =   2520
      TabIndex        =   4
      Top             =   990
      Width           =   2040
      _ExtentX        =   3599
      _ExtentY        =   477
      Caption         =   ""
   End
   Begin MSForms.Frame fraChecks 
      Caption         =   "Verifications"
      Height          =   1800
      Left            =   120
      TabIndex        =   5
      Top             =   1350
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   3175
      Begin MSForms.CheckBox chk1 
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   6720
         _ExtentX        =   11854
         _ExtentY        =   477
         Caption         =   "[1] Convention de nommage"
         Value           =   -1  'True
      End
      Begin MSForms.CheckBox chk2 
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   570
         Width           =   6720
         _ExtentX        =   11854
         _ExtentY        =   477
         Caption         =   "[2] Solidite / Volume"
         Value           =   -1  'True
      End
      Begin MSForms.CheckBox chk3 
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   900
         Width           =   6720
         _ExtentX        =   11854
         _ExtentY        =   477
         Caption         =   "[3] Auto-intersections"
         Value           =   -1  'True
      End
      Begin MSForms.CheckBox chk4 
         Height          =   270
         Left            =   240
         TabIndex        =   9
         Top             =   1230
         Width           =   6720
         _ExtentX        =   11854
         _ExtentY        =   477
         Caption         =   "[4] Silver Faces"
         Value           =   -1  'True
      End
   End
   Begin MSForms.Frame fraMethod 
      Caption         =   "Methode Silver Faces"
      Height          =   780
      Left            =   120
      TabIndex        =   10
      Top             =   3270
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   1376
      Begin MSForms.OptionButton opt1 
         Height          =   270
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   477
         Caption         =   "Seuil classique (< 0.01 mm2)"
         Value           =   -1  'True
      End
      Begin MSForms.OptionButton opt2 
         Height          =   270
         Left            =   3840
         TabIndex        =   12
         Top             =   270
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   477
         Caption         =   "Regression Logistique (ML)"
      End
   End
   Begin MSForms.CommandButton btnRun 
      Caption         =   "Tout verifier"
      Height          =   450
      Left            =   120
      TabIndex        =   13
      Top             =   4170
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   794
   End
   Begin MSForms.CommandButton btnClear 
      Caption         =   "Effacer resultats"
      Height          =   450
      Left            =   3240
      TabIndex        =   14
      Top             =   4170
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   794
   End
   Begin MSForms.TextBox txtReport 
      Height          =   3000
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2
      TabIndex        =   15
      Top             =   4770
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   5292
   End
   Begin MSForms.CommandButton btnExport 
      Caption         =   "Exporter rapport (.txt)"
      Height          =   450
      Left            =   120
      TabIndex        =   16
      Top             =   7920
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   794
   End
   Begin MSForms.CommandButton btnClose 
      Caption         =   "Fermer"
      Height          =   450
      Left            =   4200
      TabIndex        =   17
      Top             =   7920
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   794
   End
End
Attribute VB_Name = "QCheckerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' UserForm QCheckerForm -- Interface graphique Q-Checker
' CATIA V5 -- Segula Technologies
' ============================================================

' Chemin complet du document actif (utilise pour l export du rapport)
Private m_sDocFullName As String

' ============================================================
' Initialisation du formulaire :
'  - Detection du document actif
'  - Affichage du nom et du type
'  - Desactivation des controles si aucun document ouvert
' ============================================================
Private Sub UserForm_Initialize()
    Dim oDoc As Object

    m_sDocFullName = ""

    ' Mise en forme de l en-tete
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 11
    lblTitle.Font.Name = "Arial"

    ' Mise en forme du libelle du document (gras)
    lblDocName.Font.Bold = True

    ' TextBox : lecture seule, police a espacement fixe pour alignement
    txtReport.Locked    = True
    txtReport.BackColor = &H00F0F0F0
    txtReport.Font.Name = "Courier New"
    txtReport.Font.Size = 8

    ' Captions avec accents (impossibles dans l en-tete BEGIN...END)
    fraChecks.Caption    = "Verifications a effectuer"
    chk1.Caption         = "[1] Convention de nommage"
    chk2.Caption         = "[2] Solidite / Volume"
    chk3.Caption         = "[3] Auto-intersections"
    chk4.Caption         = "[4] Silver Faces"
    fraMethod.Caption    = "Methode Silver Faces"
    opt1.Caption         = "Seuil classique (< 0.01 mm2)"
    opt2.Caption         = "Regression Logistique (ML)"
    btnRun.Caption       = "Tout verifier"
    btnClear.Caption     = "Effacer resultats"
    btnExport.Caption    = "Exporter rapport (.txt)"
    btnClose.Caption     = "Fermer"
    lblDocLabel.Caption  = "Document actif :"
    lblTypeLabel.Caption = "Type :"

    ' Le bouton Export est desactive tant qu aucun rapport n est genere
    btnExport.Enabled = False

    ' Detection du document actif CATIA
    On Error Resume Next
    Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0

    If oDoc Is Nothing Then
        ' Aucun document ouvert
        lblDocName.Caption = "(aucun document ouvert)"
        lblDocType.Caption = "N/A"
        btnRun.Enabled     = False
        btnExport.Enabled  = False
        MsgBox "Aucun document CATIA n est ouvert." & vbCrLf & _
               "Veuillez ouvrir un fichier CATPart ou CATProduct.", _
               vbExclamation, "Q-Checker"
    Else
        ' Affichage du nom et du chemin complet du document
        On Error Resume Next
        lblDocName.Caption = oDoc.Name
        m_sDocFullName     = oDoc.FullName
        On Error GoTo 0

        ' Affichage du type de document
        Dim sType As String
        sType = TypeName(oDoc)
        If sType = "PartDocument" Then
            lblDocType.Caption = "CATPart"
        ElseIf sType = "ProductDocument" Then
            lblDocType.Caption = "CATProduct"
        Else
            lblDocType.Caption = sType & " (non pris en charge)"
            btnRun.Enabled = False
        End If
    End If

    Set oDoc = Nothing
End Sub

' ============================================================
' Bouton "Tout verifier" -- Execute les modules coches dans l ordre 1->4
' Affiche les resultats dans txtReport au fur et a mesure (DoEvents)
' ============================================================
Private Sub btnRun_Click()
    Dim oDoc       As Object
    Dim sReport    As String
    Dim sDetails   As String
    Dim bOK        As Boolean
    Dim iConformes As Integer
    Dim iTotal     As Integer
    Dim bUseML     As Boolean
    Dim sMethode   As String

    ' Recuperation du document actif
    On Error Resume Next
    Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0

    If oDoc Is Nothing Then
        MsgBox "Aucun document CATIA n est ouvert.", vbCritical, "Q-Checker"
        Exit Sub
    End If

    ' Desactivation des boutons pendant l execution pour eviter les doubles clics
    btnRun.Enabled    = False
    btnClear.Enabled  = False
    btnExport.Enabled = False
    txtReport.Locked  = False
    txtReport.Value   = "Analyse en cours, veuillez patienter..." & vbCrLf
    txtReport.Locked  = True
    DoEvents

    iConformes = 0
    iTotal     = 0
    bUseML     = (opt2.Value = True)

    ' Construction de l en-tete du rapport
    sReport = BuildReportHeader(oDoc)
    txtReport.Locked = False
    txtReport.Value  = sReport
    txtReport.Locked = True
    DoEvents

    ' ── Module 1 : Convention de nommage ──────────────────────────
    If chk1.Value Then
        iTotal   = iTotal + 1
        sDetails = ""
        On Error Resume Next
        bOK = RunCheck1(oDoc, sDetails)
        If Err.Number <> 0 Then
            sDetails = "  [ERREUR] " & Err.Description & vbCrLf
            bOK      = False
            Err.Clear
        End If
        On Error GoTo 0
        sReport = sReport & BuildModuleSection(1, "Convention de nommage", bOK, sDetails)
        If bOK Then iConformes = iConformes + 1
        txtReport.Locked = False
        txtReport.Value  = sReport
        txtReport.Locked = True
        DoEvents
    End If

    ' ── Module 2 : Solidite / Volume ──────────────────────────────
    If chk2.Value Then
        iTotal   = iTotal + 1
        sDetails = ""
        On Error Resume Next
        bOK = RunCheck2(oDoc, sDetails)
        If Err.Number <> 0 Then
            sDetails = "  [ERREUR] " & Err.Description & vbCrLf
            bOK      = False
            Err.Clear
        End If
        On Error GoTo 0
        sReport = sReport & BuildModuleSection(2, "Solidite / Volume", bOK, sDetails)
        If bOK Then iConformes = iConformes + 1
        txtReport.Locked = False
        txtReport.Value  = sReport
        txtReport.Locked = True
        DoEvents
    End If

    ' ── Module 3 : Auto-intersections ─────────────────────────────
    If chk3.Value Then
        iTotal   = iTotal + 1
        sDetails = ""
        On Error Resume Next
        bOK = RunCheck3(oDoc, sDetails)
        If Err.Number <> 0 Then
            sDetails = "  [ERREUR] " & Err.Description & vbCrLf
            bOK      = False
            Err.Clear
        End If
        On Error GoTo 0
        sReport = sReport & BuildModuleSection(3, "Auto-intersections", bOK, sDetails)
        If bOK Then iConformes = iConformes + 1
        txtReport.Locked = False
        txtReport.Value  = sReport
        txtReport.Locked = True
        DoEvents
    End If

    ' ── Module 4 : Silver Faces ───────────────────────────────────
    If chk4.Value Then
        iTotal    = iTotal + 1
        sDetails  = ""
        sMethode  = GetMethodeLabel(bUseML)
        On Error Resume Next
        bOK = RunCheck4(oDoc, sDetails, bUseML)
        If Err.Number <> 0 Then
            sDetails = "  [ERREUR] " & Err.Description & vbCrLf
            bOK      = False
            Err.Clear
        End If
        On Error GoTo 0
        sReport = sReport & BuildModuleSection(4, "Silver Faces (" & sMethode & ")", bOK, sDetails)
        If bOK Then iConformes = iConformes + 1
        txtReport.Locked = False
        txtReport.Value  = sReport
        txtReport.Locked = True
        DoEvents
    End If

    ' ── Bilan final ───────────────────────────────────────────────
    If iTotal = 0 Then
        sReport = sReport & "  Aucune verification selectionnee." & vbCrLf
    Else
        sReport = sReport & BuildReportBilan(iConformes, iTotal)
    End If
    txtReport.Locked = False
    txtReport.Value  = sReport
    txtReport.Locked = True

    ' Activation de l export maintenant que le rapport est disponible
    btnExport.Enabled = True

    ' Reactivation des boutons
    btnRun.Enabled   = True
    btnClear.Enabled = True

    Set oDoc = Nothing
End Sub

' ============================================================
' Bouton "Effacer resultats" -- Vide le TextBox du rapport
' ============================================================
Private Sub btnClear_Click()
    txtReport.Locked  = False
    txtReport.Value   = ""
    txtReport.Locked  = True
    btnExport.Enabled = False
End Sub

' ============================================================
' Bouton "Exporter rapport (.txt)" -- Sauvegarde dans un fichier horodate
' ============================================================
Private Sub btnExport_Click()
    If Len(txtReport.Value) = 0 Then
        MsgBox "Aucun rapport a exporter.", vbInformation, "Q-Checker"
        Exit Sub
    End If
    ExportReport txtReport.Value, m_sDocFullName
End Sub

' ============================================================
' Bouton "Fermer" -- Ferme et decharge le formulaire
' ============================================================
Private Sub btnClose_Click()
    Unload Me
End Sub

' ============================================================
' CheckBox Silver Faces -- Active ou desactive le choix de methode
' ============================================================
Private Sub chk4_Click()
    Dim bEnabled As Boolean
    bEnabled      = (chk4.Value = True)
    opt1.Enabled  = bEnabled
    opt2.Enabled  = bEnabled
    fraMethod.Enabled = bEnabled
End Sub

' ============================================================
' Nettoyage a la fermeture du formulaire
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    m_sDocFullName = ""
End Sub
