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

Public Function getIterMois(ByVal moisCours As Date) As Integer
    Dim i As Integer

    For i = 6 To 54 Step 4
        If Cells(1, i).Value = moisCours Then
            getIterMois = i
        End If
    Next i
    
End Function
Public Function updateRRF(ByVal moisCours As Date)
    Dim i As Integer
    Dim k As Integer
    Dim iterMoisPrec As Integer
    Dim rf_a_transf As Double
    Dim moisReel As Integer
    Dim moisDemande As Integer
    Dim moisActuel As Integer
    Dim last_row As Integer
    Worksheets("SUIVI PROJET").Activate
    
    i = 2
    k = 3
    last_row = Cells(Rows.Count, 2).End(xlUp).Row
    moisDemande = getIterMois(moisCours)
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
    'ALORS LE SOUCIS C'EST QUE LA CA SEMBLE NE PAS PRENDRE EN COMPTE LA LIGNE 86 ? PQ JE NE SAIS PAS ENCORE
    
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
Public Function calculTotalR(ByVal ligne As Integer, ByVal iterMois As Integer, ByVal SelectedSheet As String)
    
    Dim Total As Double
    Dim j As Integer
    switchBehavior (SelectedSheet)
    
    Total = 0
    For j = gIterMin To gIterMaxR Step gStep
        Total = Total + Cells(ligne, j).Value
    Next j
    
    Cells(ligne, gTotalR).Value = Total
    
End Function
Public Function calculTotalRF(ByVal ligne As Integer, ByVal iterMois As Integer)
    Dim Total As Double
    Dim i As Integer
    Dim j As Integer
    
    
    Total = 0
    For i = 2 To iterMois - 4 Step 4
        Total = Total + Cells(ligne, i).Value
    Next i
    

    For j = iterMois + 2 To 48 Step 4
        Total = Total + Cells(ligne, j).Value
    Next j
    
    calculTotalRF = Total
    Cells(ligne, 53).Value = Total
    
End Function
Public Function calculAdd(val1 As Integer, val2 As Integer)
    calculAdd = val1 + val2
End Function
Public Function calculTotalB(ByVal sheetselect As String) As Double
    Dim Total As Double
    Dim i As Integer
    Dim ligne As Integer
    
    Total = 0#
    'Dim ligne As Integer
    For ligne = gLigneDebut To gLigneFin
        For i = gIterMin + 1 To gIterMaxB Step gStep
            Total = Total + Cells(ligne, i).Value 'à finir à adapter
        Next i
    Next ligne
    
    calculTotalB = Total
    Cells(ligne, gTotalB).Value = Total
    
End Function
Public Function calculREPORTING() As Double
    MsgBox "Test"
End Function
Sub reforecast()
    Dim last_row As Long
    Dim moisEnCours As Date
    Dim i As Integer
    Dim Var As Double
    Dim IterMoisCours As Integer
    Worksheets("SUIVI PROJET").Activate
    last_row = Cells(Rows.Count, 2).End(xlUp).Row ' A CORRIGER
    

    Worksheets("REPORTING").Activate
    moisEnCours = Range("C2").Value
    'Worksheets("GESTION DES TEMPS").Activate 'ATTENTION DS LE NOM DE BASE DE LA SHEET : ESPACE !!
    'celltotalrf = Cells(3, 53).Value
    'totaltxt = celltotalrf.Value
    
    Worksheets("SUIVI PROJET").Activate
    Var = updateRRF(moisEnCours)

    
    
    
    IterMoisCours = getIterMois(moisEnCours)
    For i = 9 To last_row Step 1
        Var = calculTotalR(i, IterMoisCours, "SUIVI PROJET") 'on envoie l'itération du mois en cours pour savoir où s'arrêter
    Next i
    
    Var = calculTotalB("SUIVI PROJET")
    
    For i = 3 To last_row Step 1
        Var = calculTotalRF(i, IterMoisCours)
    Next i
    
    'penser à adapter le reste des autres fonctions, pour que ça calcule à toutes les feuilles 
    
    
    'reste à faire : une fonction qui calcule le nouveau R puis RF total après cette modif
        
End Sub
