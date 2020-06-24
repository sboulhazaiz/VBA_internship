Attribute VB_Name = "Module1"
Public Function getIterMois(ByVal moisCours As Date) As Integer
    For i = 6 To 54 Step 4
        If Cells(1, i).Value = moisCours Then
            getIterMois = i
        End If
    Next i
    
End Function
Public Function updateRRF(ByVal moisCours As Date, ByVal ligne_a_travailler As Integer)
    Worksheets("SUIVI PROJET").Activate
    'moisCours = Range("C2").Value
    'calcul du nouveau reel :
    For i = 6 To 54 Step 4
        If Cells(1, i).Value = moisCours Then
            iterMoisPrec = i - 4 'on enregistre à quelle itération se trouve le mois précédant
            'MsgBox i
            'MsgBox iterMoisPrec 'reconnait le mois à mettre à jour (rf), celui qui précède le mois en cours donc
        End If
    Next i
    
    rf_a_transf = Cells(ligne_a_travailler, iterMoisPrec + 2).Value 'sauvegarde du RF à bouger
    'MsgBox iterMoisPrec
    'MsgBox rf_a_transf
    Cells(ligne_a_travailler, iterMoisPrec) = rf_a_transf 'déplacement du RF en réel
    'updateRRF = iterMoisPrec 'sauvegarde de l'itération du mois précédant le mois en cours
    updateRRF = iterMoisPrec

End Function
Public Function calculTotalR(ByVal ligne As Integer, ByVal iterMois As Integer)
    Total = 0 'initilisation
    For j = 2 To iterMois Step 4
        Total = Total + Cells(ligne, j).Value
    Next j
    
    Cells(ligne, 51).Value = Total
    
End Function
Public Function calculTotalRF(ByVal ligne As Integer, ByVal iterMois As Integer)
    Total = 0
    For i = 2 To iterMois - 4 Step 4
        Total = Total + Cells(ligne, i).Value
        MsgBox Cells(ligne, i).Value
    Next i
    

    For j = iterMois To 48 Step 4
        Total = Total + Cells(ligne, j).Value
        MsgBox Cells(ligne, j).Value
    Next j
    
    Cells(ligne, 53).Value = Total
    
End Function

Sub reforecast()
    Dim last_row As Long
    last_row = Cells(Rows.Count, 2).End(xlUp).Row ' A CORRIGER
    

    Worksheets("REPORTING").Activate
    moisEnCours = Range("C2").Value
    Worksheets("SUIVI PROJET").Activate 'ATTENTION DS LE NOM DE BASE DE LA SHEET : ESPACE !!
    'celltotalrf = Cells(3, 53).Value
    'totaltxt = celltotalrf.Value
    'MsgBox celltotalrf
    'MsgBox moisEnCours
    'moisPrec = updateRRF(moisEnCours)  15 954 €
    For i = 3 To last_row - 1 Step 1
        moisPrec = updateRRF(moisEnCours, i)
    Next i
    IterMoisCours = getIterMois(moisEnCours)
    
    For j = 3 To last_row Step 1
        Var = calculTotalR(j, IterMoisCours) 'on envoie l'itération du mois en cours pour savoir où s'arrêter
    Next j
    
    Var = calculTotalRF(3, IterMoisCours)
    
    'alors je sais pas pq la fonction RF donne tjrs total de 0 ? soucsi dans la 2ème boucle à priori
    
    
    'reste à faire : une fonction qui calcule le nouveau R puis RF total après cette modif
    
    
    
    
    
    
End Sub
