Option Explicit
'norme de codage : 'g' au début d'une variable signifie qu'elle est globale
Global gIterMin As Integer

Global gIterMaxR As Integer
Global gIterMaxB As Integer
Global gIterMaxRF As Integer

Global gIterMoisCours As Integer 'pas sûr de l'utilité comme c'est en paramètre de la fonction?
Global gStep As Integer

Global gTotalR As Integer
Global gTotalB As Integer
Global gTotalRF As Integer
Global gLigneDebut As Integer 'PARAMETRER
Global gLigneFin As Integer
Global gLignesDates As Integer
Global excp(8) As Integer

'PENSER A AJOUTER VARIABLES DES LIGNES A EXCLURE

Public Function switchBehavior(ByVal Sheet As String)
    'cette fonction sert à changer le fonction de la macron en fonction de la feuille sur laquelle elle travaille
    Select Case Sheet
        Case "SUIVI PROJET"
            Worksheets("SUIVI PROJET").Activate
            gIterMin = 2 'colonne début iteration
            gLigneDebut = 3 'ligne démarrage
            gLigneFin = Cells(Rows.Count, 2).End(xlUp).Row
            gIterMaxR = 46
            gIterMaxB = 47
            gIterMaxRF = 48
            gTotalR = 51
            gTotalB = 52
            gTotalRF = 53
            gStep = 4
            gLignesDates = 1
            excp(0) = 7 'lignes à exclure de l'algo car pas à écrire ici
            excp(1) = 6
            excp(2) = 103
            
        Case "GESTION DES TEMPS"
            Worksheets("GESTION DES TEMPS").Activate
            gIterMin = 6 'colonne début iteration
            gLigneDebut = 9 'ligne démarrage
            gLigneFin = Cells(Rows.Count, 2).End(xlUp).Row
            gIterMaxR = 39
            gIterMaxB = 40
            gIterMaxRF = 41
            gTotalR = 43
            gTotalB = 44
            gTotalRF = 45
            gStep = 3
            gLignesDates = 6
            excp(0) = 0 'pas d'exception
            
        Case "PLAN TRESO PROJET"
            Worksheets("PLAN TRESO PROJET").Activate
            gIterMin = 2
            gLigneDebut = 3 'ligne démarrage
            gLigneFin = Cells(Rows.Count, 2).End(xlUp).Row
            gIterMaxR = 46
            gIterMaxB = 47
            gIterMaxRF = 48
            gTotalR = 51
            gTotalB = 52
            gTotalRF = 53
            gStep = 4
    End Select
    
End Function

Public Function getIterMois(ByVal moisCours As Date, ByVal selectedSheet As String) As Integer
    Dim i As Integer
    
    switchBehavior (selectedSheet)
    For i = gIterMin To gIterMaxRF Step gStep
        If Cells(gLignesDates, i).Value = moisCours Then
            getIterMois = i
        End If
    Next i
    
End Function
Public Function updateRRF(ByVal selectedSheet As String)
    Dim i As Integer
    Dim k As Integer
    Dim iterMoisPrec As Integer
    Dim rf_a_transf As Double
    Dim moisReel As Integer
    Dim moisDemande As Integer
    Dim moisActuel As Integer
    Dim last_row As Integer
    Dim moisCours As Date
    Worksheets("SUIVI PROJET").Activate
    
    switchBehavior (selectedSheet)
    
    moisCours = ActiveWorkbook.Sheets("REPORTING").Range("C2").Value
    i = 2
    k = 3
    last_row = Cells(Rows.Count, 2).End(xlUp).Row
    moisDemande = getIterMois(moisCours, selectedSheet) 'PENSER A GENERALISER CETTE FONCTION
    iterMoisPrec = moisDemande - 4
    For i = 2 To 46 Step 4
        If Cells(104, i).Value > 0 Then 'possibilité automatiser en remplçant 104 par last_row
            moisReel = i  ' détection du mois dans lequel le tableau est, par rapport à dans Reporting
        End If
    Next i
    iterMoisPrec = moisReel
    moisReel = moisReel + 4
    
    
    If moisDemande > moisReel Then ' cas où on demande une date future par rapport à l'état du tableau
        For i = moisReel To moisDemande - 4 Step 4
            For k = 3 To last_row - 1 Step 1
                    Cells(k, i) = Cells(k, i + 2).Value
            Next k
        Next i
    End If
    
    'NE PAS OUBLIER FAIRE LA BOUCLE DANS LA FONCTION IL FAUT ENCORE LA FINIR DE LA PROGRAMMER
    
    If moisDemande < moisReel Then ' cas où on demande une date antérieure par rapport à l'état du tableau
        For i = moisReel - 4 To moisDemande Step -4
            For k = 3 To last_row - 1 Step 1
                    Cells(k, i) = ""
            Next k
        Next i
         
    End If
    
    'Worksheets(SelectedSheet).Activate
    'PENSER A RAJOUTER LE CAS DE JANVIER !!
    'calcul du nouveau reel :
    
    


    'rf_a_transf = Cells(ligne_a_travailler, iterMoisPrec + 2).Value 'sauvegarde du RF à bouger
    'MsgBox iterMoisPrec
    'MsgBox rf_a_transf
    'Cells(ligne_a_travailler, iterMoisPrec) = rf_a_transf 'déplacement du RF en réel
    'updateRRF = iterMoisPrec 'sauvegarde de l'itération du mois précédant le mois en cours
    updateRRF = iterMoisPrec

End Function
Public Function calculTotalR(ByVal selectedSheet As String)
    
    Dim Total As Double
    Dim ligne As Integer
    Dim j As Integer
    Dim iterMois As Integer
    
    switchBehavior (selectedSheet)
    
    iterMois = getIterMois(ActiveWorkbook.Sheets("REPORTING").Range("C2").Value, selectedSheet)
 
    For ligne = gLigneDebut To gLigneFin
        Total = 0#
        If IsInArray(ligne, excp) = False Then
            For j = gIterMin To gIterMaxR Step gStep
                    Total = Total + Cells(ligne, j).Value
            Next j
            Cells(ligne, gTotalR).Value = Total
        End If
    Next ligne
    
    
End Function
Public Function calculTotalRF(ByVal selectedSheet As String)
    Dim Total As Double
    Dim i As Integer
    Dim j As Integer
    Dim ligne As Integer
    Dim ValAajouter As Double
    Dim iterMois As Integer
    Dim dateActu As Date
    
    dateActu = ActiveWorkbook.Sheets("REPORTING").Range("C2").Value
    
    iterMois = getIterMois(dateActu, selectedSheet)
        
    
    ValAajouter = 0#
    switchBehavior (selectedSheet)
    
    
    
    For ligne = gLigneDebut To gLigneFin
        Total = 0#
        If IsInArray(ligne, excp) = False Then
            For i = 2 To iterMois - 4 Step 4
                ValAajouter = Cells(ligne, i).Value
                If ligne = gLigneFin Then
                    MsgBox ValAajouter
                End If
                Total = Total + ValAajouter
            Next i
            
        
            For j = iterMois + 2 To 48 Step 4
                ValAajouter = Cells(ligne, j).Value
                If ligne = gLigneFin Then
                    MsgBox ValAajouter
                End If
                Total = Total + ValAajouter
            Next j
            
            Cells(ligne, gTotalRF).Value = Total
        End If
    Next ligne
    
    calculTotalRF = Total
    
    
End Function
Public Function calculAdd(val1 As Integer, val2 As Integer)
    calculAdd = val1 + val2
End Function
Public Function calculTotalB(ByVal selectedSheet As String) As Double
    Dim Total As Double
    Dim i As Integer
    Dim ligne As Integer
    
    switchBehavior (selectedSheet)
    
    
    For ligne = gLigneDebut To gLigneFin
        Total = 0#
        
        If IsInArray(ligne, excp) = False Then
            For i = gIterMin + 1 To gIterMaxB Step gStep
                    Total = Total + Cells(ligne, i).Value 'à finir à adapter
            Next i
            Cells(ligne, gTotalB).Value = Total
        End If
    
    Next ligne
    
    calculTotalB = Total
    
    
End Function
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function
Sub reforecast()
    Dim last_row As Long
    Dim moisEnCours As Date
    Dim i As Integer
    Dim Var As Double

    Dim selected_sheet As String
    
    selected_sheet = "SUIVI PROJET"
    
    Worksheets(selected_sheet).Activate
    last_row = Cells(Rows.Count, 2).End(xlUp).Row ' A CORRIGER
    

    Worksheets("REPORTING").Activate
    'Worksheets("GESTION DES TEMPS").Activate 'ATTENTION DS LE NOM DE BASE DE LA SHEET : ESPACE !!
    'celltotalrf = Cells(3, 53).Value
    'totaltxt = celltotalrf.Value
    
    Worksheets("SUIVI PROJET").Activate
    'Var = updateRRF(selected_sheet)
    switchBehavior ("SUIVI PROJET")
    

    'For i = 3 To last_row Step 1
    Var = calculTotalR("SUIVI PROJET") 'on envoie l'itération du mois en cours pour savoir où s'arrêter
    'Next i
    
    Var = calculTotalB("SUIVI PROJET")
    
    'For i = 3 To last_row Step 1
    Var = calculTotalRF("SUIVI PROJET")
    'Next i

    
        
End Sub
