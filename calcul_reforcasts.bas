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
Global excp(100) As Integer 'tableau des lignes de suivi projet à exclure (totaux, lignes vides, checks...)
Global lignesRep(100) As Integer '
Global nblRep As Integer 'le nb de lignes dans reporting qu'il faudra chercher dans Suivi Projet
'Global excp_sp(100) As Integer 'tableau exceptions suivi projet

'PENSER A AJOUTER VARIABLES DES LIGNES A EXCLURE

Public Function switchBehavior(ByVal Sheet As String)
    'cette fonction sert à changer le fonction de la macron en fonction de la feuille sur laquelle elle travaille
    Select Case Sheet
        Case "SUIVI PROJET"
            Worksheets("SUIVI PROJET").Activate
            gIterMin = 2 'colonne début iteration
            gLigneDebut = 3 'ligne démarrage
            gLigneFin = Cells(Rows.Count, 3).End(xlUp).Row
            gIterMaxR = 46
            gIterMaxB = 47
            gIterMaxRF = 48
            gTotalR = 51
            gTotalB = 52
            gTotalRF = 53
            gStep = 4
            gLignesDates = 1
            
            
        Case "GESTION DES TEMPS"
            Worksheets("GESTION DES TEMPS").Activate
            gIterMin = 6 'colonne début iteration
            gLigneDebut = 9 'ligne démarrage
            gLigneFin = Cells(Rows.Count, 3).End(xlUp).Row
            gIterMaxR = 39
            gIterMaxB = 40
            gIterMaxRF = 41
            gTotalR = 43
            gTotalB = 44
            gTotalRF = 45
            gStep = 3
            gLignesDates = 6
            Erase excp
            excp(0) = 0 'pas d'exception
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
            For j = gIterMin To iterMois Step gStep
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
    
    
    dateActu = ActiveWorkbook.Sheets("REPORTING").Range("C2").Value 'on récupère la date à laquelle le reporting est réglé
    switchBehavior (selectedSheet)
    
    If Month(dateActu) = 1 Then
        For i = gLigneDebut To gLigneFin Step 1
            Cells(i, gTotalRF).Value = Cells(i, gTotalB).Value
        Next i
    Else
    
        iterMois = getIterMois(dateActu, selectedSheet) 'on récupère l'itération à laquelle s'arrêter
            
        
        ValAajouter = 0#
        
        
        
        
        For ligne = gLigneDebut To gLigneFin
            Total = 0#
            If IsInArray(ligne, excp) = False Then
                For i = gIterMin To iterMois - gStep Step gStep
                    ValAajouter = Cells(ligne, i).Value
                    Total = Total + ValAajouter
                Next i
                
            
                For j = iterMois + 2 To gIterMaxRF Step gStep
                    ValAajouter = Cells(ligne, j).Value
                    Total = Total + ValAajouter
                Next j
                
                Cells(ligne, gTotalRF).Value = Total
            End If
        Next ligne
        
        calculTotalRF = Total
    
    End If
    
    
    
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
Private Function reporting()
Dim i As Integer
Dim dateActu As Date
Dim iterMois As Integer
Dim k As Integer
Dim resultat As Double


dateActu = ActiveWorkbook.Sheets("REPORTING").Range("C2").Value
iterMois = getIterMois(dateActu, "SUIVI PROJET") 'on récupère l'itération à laquelle s'arrêter
switchBehavior ("SUIVI PROJET") ' on passe en mode suivi projet comme on aura besoin que de ça
resultat = 0


For i = 0 To nblRep - 1 Step 1 'calcul budget
    For k = gIterMin + 1 To iterMois + 4 Step gStep 'calcul des budgets initiaux
        resultat = resultat + Cells(lignesRep(i), k).Value 'on est dans la ligne est on avance col. par col.
    Next k
    Sheets("REPORTING").Cells(i + 5, 3).Value = resultat 'on écrit le résultat dans la case / on fait +5 parce que dans l'onglet reporting, ça commence à la ligne 5, et le i commence à 0, donc comme ça, il fait 0+5 = 5, et il va chercher la bonne ligne, et ainsi de suite
    resultat = 0 'on remet les compteurs à 0
Next i

For i = 0 To nblRep - 1 Step 1 'calcul reforecast
    For k = gIterMin To iterMois - 4 Step gStep 'calcul des budgets initiaux // -4 car sinon on va compter le réel du mois en cours or on voudra le RF
        resultat = resultat + Cells(lignesRep(i), k).Value 'on est dans la ligne est on avance col. par col.
    Next k
    resultat = resultat + Cells(lignesRep(i), iterMois + 2).Value ' ajout RF du mois en cours
    Sheets("REPORTING").Cells(i + 5, 4).Value = resultat 'on écrit le résultat dans la case
    resultat = 0 'on remet les compteurs à 0
Next i

For i = 0 To nblRep - 1 Step 1 'calcul réel
    For k = gIterMin To iterMois Step gStep 'calcul des budgets initiaux
        resultat = resultat + Cells(lignesRep(i), k).Value 'on est dans la ligne est on avance col. par col.
    Next k
    Sheets("REPORTING").Cells(i + 5, 5).Value = resultat 'on écrit le résultat dans la case
    resultat = 0 'on remet les compteurs à 0
Next i

' ---------------- CHECKS ---------------- :
resultat = 0 'check budget
For k = gIterMin + 1 To iterMois + 4 Step gStep 'calcul des budgets initiaux
        resultat = resultat + Cells(gLigneFin, k).Value 'on est dans la ligne est on avance col. par col.
Next k
Sheets("REPORTING").Cells(5 + nblRep + 3, 3).Value = Sheets("REPORTING").Range("C10").Value - resultat 'on écrit le résultat dans la case
'on va chercher la case 4 + nblRep + 3 , pour trouver la ligne de "checks", car elle se trouve toujours 3 lignes en dessous des lignes de reporting, qui elles mêmes démarrent toujours à la ligne 5 (mais on fait +4 pas +5 car on démarre déjà à la ligne 1 dans Excel)

resultat = 0 'check rf
For k = gIterMin To iterMois - 4 Step gStep 'calcul des budgets initiaux
        resultat = resultat + Cells(gLigneFin, k).Value 'on est dans la ligne est on avance col. par col.
Next k
resultat = resultat + Cells(gLigneFin, iterMois + 2).Value ' ajout RF du mois en cours
Sheets("REPORTING").Cells(5 + nblRep + 3, 4).Value = Sheets("REPORTING").Range("D10").Value - resultat 'on écrit le résultat dans la case


resultat = 0 'check reel
For k = gIterMin To iterMois Step gStep 'calcul des budgets initiaux
        resultat = resultat + Cells(gLigneFin, k).Value 'on est dans la ligne est on avance col. par col.
Next k
Sheets("REPORTING").Cells(5 + nblRep + 3, 5).Value = Sheets("REPORTING").Range("E10").Value - resultat 'on écrit le résultat dans la case



'Range("C13").Value = Application.Sum(Range(Cells(2, 1), Cells(3, 2)))
'Range("C13").Value = resultat

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
    Dim myValue As Integer
    
    i = 0
    myValue = 1
    nblRep = 0 ' compteur du nb de lignes Reporting à 0
    
    Worksheets("Suivi Projet").Activate

    While myValue > 0
        myValue = InputBox("Insérez, une par une, les lignes à exclure dans SUIVI PROJET (lignes vides, lignes de total, et checks). Quand vous n'avez plus de ligne à excure, entrez simplement le chiffre 0.")
        excp(i) = myValue
        i = i + 1
    Wend
    
    i = 0
    myValue = 1 'on réinitialise les valeurs
    While myValue <> 0
        myValue = InputBox("Insérez, une par une, les lignes correspondantes au Reporting dans Suivi Projet. Entrez 0 une fois que vous les avez toutes entrées.")
        lignesRep(i) = myValue
        i = i + 1
        nblRep = nblRep + 1
    Wend
    
    nblRep = nblRep - 1 ' on enlève 1 dans le compteur de lignes parce qu'il y a le "0" qu'on a ajouté à la fin
    
    
    
    switchBehavior ("SUIVI PROJET")

    Var = calculTotalR("SUIVI PROJET")
    Var = calculTotalB("SUIVI PROJET")
    Var = calculTotalRF("SUIVI PROJET")

    switchBehavior ("GESTION DES TEMPS")
    
    Var = calculTotalR("GESTION DES TEMPS")
    Var = calculTotalB("GESTION DES TEMPS")
    Var = calculTotalRF("GESTION DES TEMPS")
    Var = reporting()
    
        
End Sub
