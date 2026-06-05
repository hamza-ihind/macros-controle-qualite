Attribute VB_Name = "Module_Main"
Option Explicit
' ============================================================
' Module principal -- Point d entree de la macro Q-Checker CATIA V5
' Segula Technologies
'
' Pour executer : Tools > Macro > Macros... > selectionner CATMain > Run
' Ou assignez la macro a un bouton de barre d outils.
' ============================================================

' ------------------------------------------------------------
' Point d entree CATIA V5 : affiche la UserForm QCheckerForm
' ------------------------------------------------------------
Sub CATMain()
    QCheckerForm.Show
End Sub
