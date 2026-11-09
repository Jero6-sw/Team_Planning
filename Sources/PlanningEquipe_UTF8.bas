Attribute VB_Name = "M_PlanningEquipe"

Option Explicit

'====================================================================================
'  PLANNING D'ÉQUIPE - Module VBA unique
'------------------------------------------------------------------------------------
'  MACROS A UTILISER (Alt+F8) :
'    - ConstruireClasseur : construit les 5 feuilles du classeur (à exécuter 1 fois)
'    - MajTout            : recalcule les tâches et met à jour Charge et Résumé
'    - AjouterFeriesAnnee : ajoute les jours fériés français d'une année donnée
'
'  FEUILLES GÉRÉES :
'    Saisie       : 1 ligne = 1 tâche (prénom, projet, description, durée...)
'    Charge       : grille de charge par personne et par semaine (vision couleur)
'    Résumé       : résumé numérique (personnes, projets, semaines, tâches)
'    Équipe       : membres, absences (congés / formations), paramètres
'    Jours Fériés : jours fériés France + jours exceptionnels (ponts...)
'
'  RÈGLES DE CALCUL :
'    - Jour disponible = lundi à vendredi, hors jours fériés, hors absences
'      de la personne (congés, formations, autre)
'    - Dans Saisie, un Début laissé vide = la tâche s'enchaîne après la fin
'      de la tâche précédente de la même personne (ordre du tableau)
'    - Retard = jours ouvrés entre la fin calculée et aujourd'hui (si non terminée)
'    - Charge d'une semaine = jours de tâches réellement disponibles dans la semaine,
'      ramenés à la capacité de la personne (colonne Capacité de la feuille Équipe)
'
'  COMPATIBILITÉ (syntaxe vérifiée) :
'    - Vérifié : Microsoft 365, Office 2024, Excel 2016 / 2019 / 2021
'      (Excel desktop Windows, 32 et 64 bits)
'    - Le module n'utilise que des instructions VBA et des constantes Excel
'      standard (xlValidateList, xlCellValue, xlGreater, xlBetween, etc.)
'    - Non concerné : Excel Online et Excel mobile (les macros n'y tournent pas)
'
'  ENCODAGE DU FICHIER :
'    - PlanningEquipe.bas : encodage ANSI (Windows-1252), fins de ligne CRLF,
'      à importer directement dans VBE (Fichier > Importer un fichier) ;
'      les accents sont préservés à l'import
'    - PlanningEquipe_UTF8.bas : même code en UTF-8 pour lecture et copier-coller
'      depuis github.com (ne pas copier la ligne "Attribute VB_Name" dans ce cas)
'====================================================================================

'--------------------------------------------------------------------------------
' Constantes et couleurs
'--------------------------------------------------------------------------------
Private Const SH_SAI As String = "Saisie"
Private Const SH_CHA As String = "Charge"
Private Const SH_RES As String = "Résumé"
Private Const SH_EQ As String = "Équipe"
Private Const SH_FER As String = "Jours Fériés"

Private Const ETAT_NC As String = "Non commencée"
Private Const ETAT_EC As String = "En cours"
Private Const ETAT_T As String = "Terminée"

' Couleurs de la grille de charge : modifiez les valeurs dans InitCouleurs
Private coulVertClair As Long
Private coulVert As Long
Private coulOrange As Long
Private coulRouge As Long
Private coulConge As Long
Private coulFormation As Long
Private coulIndispo As Long

Private Sub InitCouleurs()
    coulVertClair = RGB(221, 245, 221)   ' sous-charge
    coulVert = RGB(198, 239, 206)         ' charge correcte (jusqu'à 100%)
    coulOrange = RGB(255, 217, 150)       ' surcharge modérée
    coulRouge = RGB(255, 160, 160)       ' surcharge élevée
    coulConge = RGB(180, 210, 255)        ' congé
    coulFormation = RGB(215, 185, 240)    ' formation
    coulIndispo = RGB(217, 217, 217)     ' semaine sans aucun jour disponible
End Sub

'--------------------------------------------------------------------------------
' Accès aux feuilles et au tableau de saisie
'--------------------------------------------------------------------------------
Private Function Feuille(nom As String) As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
End Function

Private Function FeuilleReq(nom As String) As Worksheet
    Set FeuilleReq = Feuille(nom)
    If FeuilleReq Is Nothing Then
        Err.Raise vbObjectError + 513, , _
            "La feuille '" & nom & "' est introuvable. Exécutez d'abord la macro ConstruireClasseur."
    End If
End Function

Private Function TableSaisie() As ListObject
    Dim s As Worksheet
    Set s = Feuille(SH_SAI)
    If Not s Is Nothing Then
        On Error Resume Next
        Set TableSaisie = s.ListObjects("TblSaisie")
        On Error GoTo 0
    End If
End Function

'--------------------------------------------------------------------------------
' Petites fonctions utilitaires
'--------------------------------------------------------------------------------
Private Function NumOr(c As Range, def As Double) As Double
    If IsEmpty(c.Value) Then
        NumOr = def
    ElseIf IsNumeric(c.Value) Then
        NumOr = CDbl(c.Value)
    Else
        NumOr = def
    End If
End Function

Private Function MinD(a As Double, b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Public Function SemaineISO(d As Date) As Long
    ' Numéro de semaine ISO (semaines du lundi au dimanche)
    Dim jan4 As Date, lundi1 As Date
    jan4 = DateSerial(Year(d), 1, 4)
    lundi1 = jan4 - (Weekday(jan4, vbMonday) - 1)
    If d < lundi1 Then
        SemaineISO = SemaineISO(DateSerial(Year(d) - 1, 12, 28))
    Else
        SemaineISO = (d - lundi1) \ 7 + 1
    End If
End Function

'--------------------------------------------------------------------------------
' Calculs de jours : fériés, absences, jours disponibles
'--------------------------------------------------------------------------------
Public Function EstFerie(d As Date) As Boolean
    Dim s As Worksheet, i As Long, last As Long
    Set s = Feuille(SH_FER)
    If s Is Nothing Then Exit Function
    last = s.Cells(s.Rows.Count, 1).End(xlUp).Row
    For i = 3 To last
        If IsDate(s.Cells(i, 1).Value) Then
            If Int(CDate(s.Cells(i, 1).Value)) = d Then
                EstFerie = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function EstAbsent(prenom As String, d As Date) As Boolean
    Dim s As Worksheet, i As Long, last As Long
    Set s = Feuille(SH_EQ)
    If s Is Nothing Then Exit Function
    last = s.Cells(s.Rows.Count, 6).End(xlUp).Row
    For i = 3 To last
        If Len(Trim$(CStr(s.Cells(i, 6).Value))) > 0 Then
            If StrComp(Trim$(CStr(s.Cells(i, 6).Value)), prenom, vbTextCompare) = 0 Then
                If IsDate(s.Cells(i, 8).Value) And IsDate(s.Cells(i, 9).Value) Then
                    If d >= Int(CDate(s.Cells(i, 8).Value)) And d <= Int(CDate(s.Cells(i, 9).Value)) Then
                        EstAbsent = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Function JourDispo(prenom As String, d As Date) As Boolean
    If Weekday(d, vbMonday) > 5 Then Exit Function          ' samedi / dimanche
    If EstFerie(d) Then Exit Function                        ' jour férié
    If EstAbsent(prenom, d) Then Exit Function               ' congé / formation
    JourDispo = True
End Function

Public Function JoursDispoEntre(prenom As String, d1 As Date, d2 As Date) As Double
    Dim d As Date, n As Double
    If d2 < d1 Then Exit Function
    For d = d1 To d2
        If JourDispo(prenom, d) Then n = n + 1
    Next d
    JoursDispoEntre = n
End Function

Public Function JoursOuvresEntre(d1 As Date, d2 As Date) As Long
    ' Jours ouvrés hors week-ends et jours fériés (sans tenir compte des absences)
    Dim d As Date, n As Long
    If d2 < d1 Then Exit Function
    For d = d1 To d2
        If Weekday(d, vbMonday) <= 5 Then
            If Not EstFerie(d) Then n = n + 1
        End If
    Next d
    JoursOuvresEntre = n
End Function

Public Function ProchainJourDispo(prenom As String, d As Date) As Date
    Dim x As Date, garde As Long
    x = d
    Do While Not JourDispo(prenom, x)
        x = x + 1
        garde = garde + 1
        If garde > 365 Then Exit Do
    Loop
    ProchainJourDispo = x
End Function

Public Function FinTache(prenom As String, debut As Date, duree As Long) As Date
    ' Renvoie la date de fin : duree jours disponibles en comptant le jour de début
    Dim d As Date, n As Long
    d = ProchainJourDispo(prenom, debut)
    If duree <= 0 Then
        FinTache = d
        Exit Function
    End If
    Do While n < duree
        If JourDispo(prenom, d) Then n = n + 1
        If n < duree Then d = d + 1
    Loop
    FinTache = d
End Function

'--------------------------------------------------------------------------------
' Lecture de la configuration (feuille Équipe)
'--------------------------------------------------------------------------------
Private Function ListeMembres() As Variant
    Dim s As Worksheet, last As Long, i As Long, n As Long
    Dim res() As Variant
    Set s = Feuille(SH_EQ)
    If s Is Nothing Then Exit Function
    last = s.Cells(s.Rows.Count, 1).End(xlUp).Row
    If last < 3 Then Exit Function
    ReDim res(1 To last - 2)
    For i = 3 To last
        If Len(Trim$(CStr(s.Cells(i, 1).Value))) > 0 Then
            n = n + 1
            res(n) = Trim$(CStr(s.Cells(i, 1).Value))
        End If
    Next i
    If n = 0 Then Exit Function
    If n < last - 2 Then ReDim Preserve res(1 To n)
    ListeMembres = res
End Function

Private Function Capacite(prenom As String) As Double
    ' Capacité hebdomadaire de la personne (colonne B de la feuille Équipe)
    Dim s As Worksheet, last As Long, i As Long
    Dim cap As Double
    Set s = Feuille(SH_EQ)
    If s Is Nothing Then
        Capacite = 5
        Exit Function
    End If
    cap = NumOr(s.Range("M6"), 5)
    last = s.Cells(s.Rows.Count, 1).End(xlUp).Row
    For i = 3 To last
        If StrComp(Trim$(CStr(s.Cells(i, 1).Value)), prenom, vbTextCompare) = 0 Then
            cap = NumOr(s.Cells(i, 2), cap)
            Exit For
        End If
    Next i
    Capacite = cap
End Function

Private Function JoursAbsence(prenom As String, d1 As Date, d2 As Date, typ As String) As Double
    ' Nombre de jours d'absence du type demandé entre d1 et d2 (hors week-ends et fériés)
    Dim s As Worksheet, i As Long, last As Long
    Dim a As Date, b As Date, d As Date, n As Double
    Set s = Feuille(SH_EQ)
    If s Is Nothing Then Exit Function
    last = s.Cells(s.Rows.Count, 6).End(xlUp).Row
    For i = 3 To last
        If StrComp(Trim$(CStr(s.Cells(i, 6).Value)), prenom, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(s.Cells(i, 7).Value)), typ, vbTextCompare) = 0 Then
                If IsDate(s.Cells(i, 8).Value) And IsDate(s.Cells(i, 9).Value) Then
                    a = Int(CDate(s.Cells(i, 8).Value))
                    b = Int(CDate(s.Cells(i, 9).Value))
                    If a <= d2 And b >= d1 Then
                        For d = MaxD(a, d1) To MinD(b, d2)
                            If Weekday(d, vbMonday) <= 5 Then
                                If Not EstFerie(d) Then n = n + 1
                            End If
                        Next d
                    End If
                End If
            End If
        End If
    Next i
    JoursAbsence = n
End Function

Private Function AutoDebutGrille() As Date
    ' Début automatique de la grille : le plus ancien entre aujourd'hui
    ' et les débuts effectifs des tâches
    Dim d As Date, lo As ListObject, r As Long, v As Variant
    d = Date
    Set lo = TableSaisie()
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For r = 1 To lo.DataBodyRange.Rows.Count
                v = lo.DataBodyRange.Cells(r, 7).Value
                If IsDate(v) Then
                    If Int(CDate(v)) < d Then d = Int(CDate(v))
                End If
            Next r
        End If
    End If
    AutoDebutGrille = d
End Function

Private Sub Horizon(ByRef debLundi As Date, ByRef nbSem As Long, _
                    ByRef seuilBas As Double, ByRef seuilHaut As Double)
    ' Lit les paramètres de la feuille Équipe (cellules M2 à M6)
    Dim sEq As Worksheet, v As Variant, dt As Date
    Set sEq = FeuilleReq(SH_EQ)
    seuilBas = NumOr(sEq.Range("M4"), 80) / 100
    If seuilBas <= 0 Then seuilBas = 0.8
    seuilHaut = NumOr(sEq.Range("M5"), 120) / 100
    If seuilHaut <= 1 Then seuilHaut = 1.2
    nbSem = CLng(NumOr(sEq.Range("M3"), 26))
    If nbSem < 1 Then nbSem = 1
    If nbSem > 104 Then nbSem = 104
    v = sEq.Range("M2").Value
    If Not IsEmpty(v) And IsDate(v) Then
        dt = Int(CDate(v))
    Else
        dt = AutoDebutGrille()
    End If
    debLundi = dt - (Weekday(dt, vbMonday) - 1)
End Sub

Private Function ChargeSemaine(prenom As String, d1 As Date, d2 As Date) As Double
    ' Charge planifiée de la personne entre d1 et d2 (jours couverts par des tâches)
    Dim lo As ListObject, r As Long
    Dim deb As Date, fin As Date, a As Date, b As Date
    Set lo = TableSaisie()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value)), prenom, vbTextCompare) = 0 Then
            If IsDate(lo.DataBodyRange.Cells(r, 7).Value) And IsDate(lo.DataBodyRange.Cells(r, 8).Value) Then
                deb = Int(CDate(lo.DataBodyRange.Cells(r, 7).Value))
                fin = Int(CDate(lo.DataBodyRange.Cells(r, 8).Value))
                If deb <= d2 And fin >= d1 Then
                    a = MaxD(deb, d1)
                    b = MinD(fin, d2)
                    ChargeSemaine = ChargeSemaine + JoursDispoEntre(prenom, a, b)
                End If
            End If
        End If
    Next r
End Function

Private Function CoulCharge(charge As Double, dispo As Double, _
                            seuilBas As Double, seuilHaut As Double) As Long
    If dispo <= 0 Then
        CoulCharge = coulIndispo
    ElseIf charge <= 0 Then
        CoulCharge = RGB(255, 255, 255)
    ElseIf charge / dispo < seuilBas Then
        CoulCharge = coulVertClair
    ElseIf charge / dispo <= 1 Then
        CoulCharge = coulVert
    ElseIf charge / dispo <= seuilHaut Then
        CoulCharge = coulOrange
    Else
        CoulCharge = coulRouge
    End If
End Function

Private Sub CompterTaches(prenom As String, ByRef nNC As Long, ByRef nEC As Long, _
                         ByRef nT As Long, ByRef ret As Long)
    Dim lo As ListObject, r As Long
    Dim p As String, et As String, rv As Variant
    nNC = 0: nEC = 0: nT = 0: ret = 0
    Set lo = TableSaisie()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For r = 1 To lo.DataBodyRange.Rows.Count
        p = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value))
        If StrComp(p, prenom, vbTextCompare) = 0 Then
            et = Trim$(CStr(lo.DataBodyRange.Cells(r, 10).Value))
            If StrComp(et, ETAT_T, vbTextCompare) = 0 Then
                nT = nT + 1
            ElseIf StrComp(et, ETAT_EC, vbTextCompare) = 0 Then
                nEC = nEC + 1
            Else
                nNC = nNC + 1
            End If
            rv = lo.DataBodyRange.Cells(r, 11).Value
            If Not IsEmpty(rv) Then
                If IsNumeric(rv) Then ret = ret + CLng(rv)
            End If
        End If
    Next r
End Sub

Private Function ListeProjets() As Collection
    Dim c As New Collection, lo As ListObject, r As Long, pr As String
    Set lo = TableSaisie()
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For r = 1 To lo.DataBodyRange.Rows.Count
                pr = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value))
                If Len(pr) > 0 Then
                    On Error Resume Next
                    c.Add pr, pr
                    On Error GoTo 0
                End If
            Next r
        End If
    End If
    Set ListeProjets = c
End Function

'--------------------------------------------------------------------------------
' MISE À JOUR : recalcul des tâches (feuille Saisie)
'--------------------------------------------------------------------------------
Public Sub RecalculerSaisie()
    Dim lo As ListObject, r As Long, num As Long
    Dim prenom As String, etat As String
    Dim duree As Long, retard As Long
    Dim deb As Date, fin As Date, av As Double
    Dim v As Variant, finPrec As Variant
    Dim chaines As New Collection

    Set lo = TableSaisie()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    For r = 1 To lo.DataBodyRange.Rows.Count
        prenom = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value))
        If Len(prenom) > 0 Then

            ' Numéro automatique si la cellule est vide
            v = lo.DataBodyRange.Cells(r, 1).Value
            If IsEmpty(v) Or Not IsNumeric(v) Then
                num = num + 1
                lo.DataBodyRange.Cells(r, 1).Value = num
            End If

            ' Durée en jours ouvrés
            v = lo.DataBodyRange.Cells(r, 5).Value
            If Not IsEmpty(v) And IsNumeric(v) Then duree = CLng(v) Else duree = 0
            If duree < 0 Then duree = 0

            ' Avancement de 0 à 100
            v = lo.DataBodyRange.Cells(r, 9).Value
            If Not IsEmpty(v) And IsNumeric(v) Then av = CDbl(v) Else av = 0
            If av < 0 Then av = 0
            If av > 100 Then av = 100

            ' Date de début : saisie, ou enchaînement automatique
            v = lo.DataBodyRange.Cells(r, 6).Value
            If Not IsEmpty(v) And IsDate(v) Then
                deb = Int(CDate(v))
            Else
                finPrec = Empty
                On Error Resume Next
                finPrec = chaines(prenom)
                On Error GoTo 0
                If IsDate(finPrec) And Not IsEmpty(finPrec) Then
                    deb = ProchainJourDispo(prenom, CDate(finPrec) + 1)
                Else
                    deb = ProchainJourDispo(prenom, Date)
                End If
            End If

            fin = FinTache(prenom, deb, duree)
            lo.DataBodyRange.Cells(r, 7).Value = deb
            lo.DataBodyRange.Cells(r, 8).Value = fin

            On Error Resume Next
            chaines.Remove prenom
            On Error GoTo 0
            chaines.Add fin, prenom

            ' État automatique si la cellule est vide
            etat = Trim$(CStr(lo.DataBodyRange.Cells(r, 10).Value))
            If Len(etat) = 0 Then
                If av >= 100 Then
                    etat = ETAT_T
                ElseIf av > 0 Then
                    etat = ETAT_EC
                Else
                    etat = ETAT_NC
                End If
                lo.DataBodyRange.Cells(r, 10).Value = etat
            End If

            ' Retard : jours ouvrés entre la fin calculée et aujourd'hui
            If StrComp(etat, ETAT_T, vbTextCompare) <> 0 And fin < Date Then
                retard = JoursOuvresEntre(fin + 1, Date)
            Else
                retard = 0
            End If
            lo.DataBodyRange.Cells(r, 11).Value = retard
        End If
    Next r
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------
' MISE À JOUR : grille de charge (feuille Charge)
'--------------------------------------------------------------------------------
Public Sub MajCharge()
    Dim sEq As Worksheet, s As Worksheet
    Dim membres As Variant, nM As Long
    Dim debLundi As Date, nbSem As Long
    Dim seuilBas As Double, seuilHaut As Double
    Dim w As Long, col As Long, r As Long, i As Long
    Dim lundi As Date, vendredi As Date
    Dim p As String, txt As String
    Dim dispo As Double, charge As Double, cap As Double
    Dim nc As Double, nf As Double
    Dim totCharge() As Double, totDispo() As Double
    Dim rTot As Long, rAbs As Long, rr As Long, rLeg As Long

    Set sEq = FeuilleReq(SH_EQ)
    Set s = FeuilleReq(SH_CHA)
    InitCouleurs
    Call Horizon(debLundi, nbSem, seuilBas, seuilHaut)
    membres = ListeMembres()
    If IsArray(membres) Then nM = UBound(membres) Else nM = 0
    ReDim totCharge(0 To nbSem - 1)
    ReDim totDispo(0 To nbSem - 1)

    Application.ScreenUpdating = False
    s.Cells.Clear

    ' Titre
    s.Range("A1").Value = "CHARGE PAR PERSONNE ET PAR SEMAINE"
    s.Rows(1).RowHeight = 24
    FormatTitre s.Range("A1")

    ' En-têtes des semaines
    s.Range("A2").Value = "Pers. / Semaine"
    s.Range("A3").Value = "Lundi"
    For w = 0 To nbSem - 1
        col = 2 + w
        lundi = debLundi + 7 * w
        s.Cells(2, col).Value = "S" & SemaineISO(lundi)
        s.Cells(3, col).Value = lundi
        s.Cells(3, col).NumberFormat = "ddd dd/mm"
        s.Cells(2, col).HorizontalAlignment = xlCenter
        s.Cells(3, col).HorizontalAlignment = xlCenter
    Next w
    StyleEntete s.Range(s.Cells(2, 1), s.Cells(3, 1 + nbSem))

    ' Une ligne par personne
    r = 4
    For i = 1 To nM
        p = CStr(membres(i))
        s.Cells(r, 1).Value = p
        s.Cells(r, 1).Font.Bold = True
        cap = Capacite(p)
        For w = 0 To nbSem - 1
            lundi = debLundi + 7 * w
            vendredi = lundi + 4
            dispo = JoursDispoEntre(p, lundi, vendredi)
            If dispo > cap Then dispo = cap
            charge = ChargeSemaine(p, lundi, vendredi)
            totCharge(w) = totCharge(w) + charge
            totDispo(w) = totDispo(w) + dispo
            If charge > 0 Then
                s.Cells(r, 2 + w).Value = charge
                s.Cells(r, 2 + w).NumberFormat = "0.0"
            End If
            s.Cells(r, 2 + w).HorizontalAlignment = xlCenter
            s.Cells(r, 2 + w).Interior.Color = CoulCharge(charge, dispo, seuilBas, seuilHaut)
        Next w
        r = r + 1
    Next i

    ' Lignes Dispo équipe et TOTAL charge
    s.Cells(r, 1).Value = "Dispo équipe (j)"
    s.Cells(r, 1).Font.Italic = True
    s.Cells(r + 1, 1).Value = "TOTAL charge (j)"
    s.Cells(r + 1, 1).Font.Bold = True
    For w = 0 To nbSem - 1
        s.Cells(r, 2 + w).Value = totDispo(w)
        s.Cells(r, 2 + w).NumberFormat = "0.0"
        s.Cells(r, 2 + w).Font.Italic = True
        s.Cells(r, 2 + w).HorizontalAlignment = xlCenter
        If totCharge(w) > 0 Then
            s.Cells(r + 1, 2 + w).Value = totCharge(w)
        End If
        s.Cells(r + 1, 2 + w).NumberFormat = "0.0"
        s.Cells(r + 1, 2 + w).Font.Bold = True
        s.Cells(r + 1, 2 + w).HorizontalAlignment = xlCenter
        s.Cells(r + 1, 2 + w).Interior.Color = CoulCharge(totCharge(w), totDispo(w), seuilBas, seuilHaut)
    Next w
    rTot = r + 1
    If nM > 0 Then
        Bordures s.Range(s.Cells(4, 1), s.Cells(rTot, 1 + nbSem))
    End If

    ' Grille des absences
    rAbs = rTot + 3
    s.Cells(rAbs, 1).Value = "ABSENCES (C = congé, F = formation : nombre de jours dans la semaine)"
    FormatTitre s.Cells(rAbs, 1)
    s.Cells(rAbs, 1).Font.Size = 11
    s.Cells(rAbs + 1, 1).Value = "Pers. / Semaine"
    s.Cells(rAbs + 2, 1).Value = "Lundi"
    For w = 0 To nbSem - 1
        col = 2 + w
        lundi = debLundi + 7 * w
        s.Cells(rAbs + 1, col).Value = "S" & SemaineISO(lundi)
        s.Cells(rAbs + 2, col).Value = lundi
        s.Cells(rAbs + 2, col).NumberFormat = "ddd dd/mm"
        s.Cells(rAbs + 1, col).HorizontalAlignment = xlCenter
        s.Cells(rAbs + 2, col).HorizontalAlignment = xlCenter
    Next w
    StyleEntete s.Range(s.Cells(rAbs + 1, 1), s.Cells(rAbs + 2, 1 + nbSem))

    rr = rAbs + 3
    For i = 1 To nM
        p = CStr(membres(i))
        s.Cells(rr, 1).Value = p
        s.Cells(rr, 1).Font.Bold = True
        For w = 0 To nbSem - 1
            lundi = debLundi + 7 * w
            vendredi = lundi + 4
            nc = JoursAbsence(p, lundi, vendredi, "Congé")
            nf = JoursAbsence(p, lundi, vendredi, "Formation")
            txt = ""
            If nc > 0 Then txt = "C" & Format$(nc, "0")
            If nf > 0 Then
                If Len(txt) > 0 Then txt = txt & " "
                txt = txt & "F" & Format$(nf, "0")
            End If
            s.Cells(rr, 2 + w).HorizontalAlignment = xlCenter
            If Len(txt) > 0 Then
                s.Cells(rr, 2 + w).Value = txt
                If nc > 0 And nf = 0 Then
                    s.Cells(rr, 2 + w).Interior.Color = coulConge
                ElseIf nf > 0 Then
                    s.Cells(rr, 2 + w).Interior.Color = coulFormation
                End If
            End If
        Next w
        rr = rr + 1
    Next i
    If nM > 0 Then
        Bordures s.Range(s.Cells(rAbs + 1, 1), s.Cells(rr - 1, 1 + nbSem))
    End If

    ' Légende
    rLeg = rr + 2
    s.Cells(rLeg, 1).Value = "LÉGENDE"
    FormatTitre s.Cells(rLeg, 1)
    s.Cells(rLeg, 1).Font.Size = 11
    AjouterLegende s, rLeg + 1, coulVertClair, "Sous-charge : moins de " & Format$(seuilBas, "0%") & " de la capacité"
    AjouterLegende s, rLeg + 2, coulVert, "Charge correcte : entre " & Format$(seuilBas, "0%") & " et 100%"
    AjouterLegende s, rLeg + 3, coulOrange, "Surcharge modérée : jusqu'à " & Format$(seuilHaut, "0%")
    AjouterLegende s, rLeg + 4, coulRouge, "Surcharge élevée : plus de " & Format$(seuilHaut, "0%")
    AjouterLegende s, rLeg + 5, coulConge, "Congé (grille des absences)"
    AjouterLegende s, rLeg + 6, coulFormation, "Formation (grille des absences)"
    AjouterLegende s, rLeg + 7, RGB(255, 255, 255), "Blanc : aucune tâche planifiée"
    AjouterLegende s, rLeg + 8, coulIndispo, "Gris : aucun jour disponible (férié, absence)"

    ' Largeurs de colonnes
    s.Columns(1).ColumnWidth = 22
    s.Range(s.Columns(2), s.Columns(1 + nbSem)).ColumnWidth = 7.5

    Application.ScreenUpdating = True
End Sub

Private Sub AjouterLegende(s As Worksheet, r As Long, coul As Long, texte As String)
    s.Cells(r, 2).Interior.Color = coul
    s.Cells(r, 2).Borders.LineStyle = xlContinuous
    s.Cells(r, 3).Value = texte
End Sub

'--------------------------------------------------------------------------------
' MISE À JOUR : résumé numérique (feuille Résumé)
'--------------------------------------------------------------------------------
Public Sub MajResume()
    Dim s As Worksheet
    Dim membres As Variant, nM As Long, i As Long, w As Long, r As Long
    Dim debLundi As Date, nbSem As Long, finGrille As Date
    Dim seuilBas As Double, seuilHaut As Double
    Dim lo As ListObject
    Dim lundi As Date, p As String, prj As Variant
    Dim ch As Double, dp As Double, tch As Double, tdp As Double
    Dim nNC As Long, nEC As Long, nT As Long, ret As Long
    Dim v As Variant, deb As Date, fin As Date
    Dim a As Date, b As Date, dur As Double, av As Double
    Dim nTot As Long, nNCt As Long, nECt As Long, nTt As Long, nRet As Long
    Dim nP As Long, retP As Long, chP As Double
    Dim sumD As Double, sumDA As Double
    Dim rD As Long, rF As Long

    Set s = FeuilleReq(SH_RES)
    InitCouleurs
    Call Horizon(debLundi, nbSem, seuilBas, seuilHaut)
    finGrille = debLundi + 7 * nbSem - 1
    membres = ListeMembres()
    If IsArray(membres) Then nM = UBound(membres) Else nM = 0
    Set lo = TableSaisie()

    Application.ScreenUpdating = False
    s.Cells.Clear
    s.Range("A1").Value = "RÉSUMÉ NUMÉRIQUE DE LA CHARGE"
    s.Rows(1).RowHeight = 24
    FormatTitre s.Range("A1")

    ' ------------------------------------------------ PAR PERSONNE
    r = 3
    s.Cells(r, 1).Value = "PAR PERSONNE (sur l'horizon de la grille de charge)"
    FormatTitre s.Cells(r, 1)
    s.Cells(r, 1).Font.Size = 11
    r = r + 1
    EcrireLigne s, r, Array("Prénom", "Charge planifiée (j)", "Disponible (j)", "Taux de charge", _
                            "Non commencées", "En cours", "Terminées", "Retard cumulé (j)")
    StyleEntete s.Range(s.Cells(r, 1), s.Cells(r, 8))
    r = r + 1
    rD = r
    For i = 1 To nM
        p = CStr(membres(i))
        ch = 0
        dp = 0
        For w = 0 To nbSem - 1
            lundi = debLundi + 7 * w
            ch = ch + ChargeSemaine(p, lundi, lundi + 4)
            dp = dp + MinD(JoursDispoEntre(p, lundi, lundi + 4), Capacite(p))
        Next w
        Call CompterTaches(p, nNC, nEC, nT, ret)
        s.Cells(r, 1).Value = p
        s.Cells(r, 2).Value = ch
        s.Cells(r, 3).Value = dp
        If dp > 0 Then
            s.Cells(r, 4).Value = ch / dp
            s.Cells(r, 4).Interior.Color = CoulCharge(ch, dp, seuilBas, seuilHaut)
        Else
            s.Cells(r, 4).Value = 0
        End If
        s.Cells(r, 5).Value = nNC
        s.Cells(r, 6).Value = nEC
        s.Cells(r, 7).Value = nT
        s.Cells(r, 8).Value = ret
        r = r + 1
    Next i
    rF = r - 1
    If rF >= rD Then
        s.Range(s.Cells(rD, 2), s.Cells(rF, 3)).NumberFormat = "0.0"
        s.Range(s.Cells(rD, 4), s.Cells(rF, 4)).NumberFormat = "0%"
        s.Range(s.Cells(rD, 5), s.Cells(rF, 8)).NumberFormat = "0"
        s.Range(s.Cells(rD, 4), s.Cells(rF, 4)).HorizontalAlignment = xlCenter
        Bordures s.Range(s.Cells(rD - 1, 1), s.Cells(rF, 8))
    End If

    ' ------------------------------------------------ PAR PROJET
    r = r + 1
    s.Cells(r, 1).Value = "PAR PROJET (charge sur l'horizon, avancement pondéré par la durée)"
    FormatTitre s.Cells(r, 1)
    s.Cells(r, 1).Font.Size = 11
    r = r + 1
    EcrireLigne s, r, Array("Projet", "Charge planifiée (j)", "Avancement moyen", "Nb tâches", "Tâches en retard")
    StyleEntete s.Range(s.Cells(r, 1), s.Cells(r, 5))
    r = r + 1
    rD = r
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For Each prj In ListeProjets()
                chP = 0
                sumD = 0
                sumDA = 0
                nP = 0
                retP = 0
                For i = 1 To lo.DataBodyRange.Rows.Count
                    If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, 3).Value)), CStr(prj), vbTextCompare) = 0 Then
                        nP = nP + 1
                        p = Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value))
                        If IsDate(lo.DataBodyRange.Cells(i, 7).Value) And IsDate(lo.DataBodyRange.Cells(i, 8).Value) Then
                            deb = Int(CDate(lo.DataBodyRange.Cells(i, 7).Value))
                            fin = Int(CDate(lo.DataBodyRange.Cells(i, 8).Value))
                            a = MaxD(deb, debLundi)
                            b = MinD(fin, finGrille)
                            If a <= b Then
                                chP = chP + JoursDispoEntre(p, a, b)
                            End If
                        End If
                        v = lo.DataBodyRange.Cells(i, 5).Value
                        If Not IsEmpty(v) And IsNumeric(v) Then dur = CDbl(v) Else dur = 0
                        v = lo.DataBodyRange.Cells(i, 9).Value
                        If Not IsEmpty(v) And IsNumeric(v) Then av = CDbl(v) Else av = 0
                        If av < 0 Then av = 0
                        If av > 100 Then av = 100
                        sumD = sumD + dur
                        sumDA = sumDA + dur * av
                        v = lo.DataBodyRange.Cells(i, 11).Value
                        If Not IsEmpty(v) And IsNumeric(v) Then
                            If CLng(v) > 0 Then retP = retP + 1
                        End If
                    End If
                Next i
                s.Cells(r, 1).Value = CStr(prj)
                s.Cells(r, 2).Value = chP
                If sumD > 0 Then
                    s.Cells(r, 3).Value = sumDA / sumD / 100
                Else
                    s.Cells(r, 3).Value = 0
                End If
                s.Cells(r, 4).Value = nP
                s.Cells(r, 5).Value = retP
                r = r + 1
            Next prj
        End If
    End If
    rF = r - 1
    If rF >= rD Then
        s.Range(s.Cells(rD, 2), s.Cells(rF, 2)).NumberFormat = "0.0"
        s.Range(s.Cells(rD, 3), s.Cells(rF, 3)).NumberFormat = "0%"
        s.Range(s.Cells(rD, 4), s.Cells(rF, 5)).NumberFormat = "0"
        s.Range(s.Cells(rD, 3), s.Cells(rF, 3)).HorizontalAlignment = xlCenter
        Bordures s.Range(s.Cells(rD - 1, 1), s.Cells(rF, 5))
    End If

    ' ------------------------------------------------ PAR SEMAINE
    r = r + 1
    s.Cells(r, 1).Value = "PAR SEMAINE (équipe complète)"
    FormatTitre s.Cells(r, 1)
    s.Cells(r, 1).Font.Size = 11
    r = r + 1
    EcrireLigne s, r, Array("Semaine", "Lundi", "Charge équipe (j)", "Dispo équipe (j)", "Taux")
    StyleEntete s.Range(s.Cells(r, 1), s.Cells(r, 5))
    r = r + 1
    rD = r
    For w = 0 To nbSem - 1
        lundi = debLundi + 7 * w
        tch = 0
        tdp = 0
        For i = 1 To nM
            p = CStr(membres(i))
            tch = tch + ChargeSemaine(p, lundi, lundi + 4)
            tdp = tdp + MinD(JoursDispoEntre(p, lundi, lundi + 4), Capacite(p))
        Next i
        s.Cells(r, 1).Value = "S" & SemaineISO(lundi)
        s.Cells(r, 2).Value = lundi
        s.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
        s.Cells(r, 3).Value = tch
        s.Cells(r, 4).Value = tdp
        If tdp > 0 Then
            s.Cells(r, 5).Value = tch / tdp
            s.Cells(r, 5).Interior.Color = CoulCharge(tch, tdp, seuilBas, seuilHaut)
        Else
            s.Cells(r, 5).Value = 0
            s.Cells(r, 5).Interior.Color = coulIndispo
        End If
        s.Cells(r, 3).NumberFormat = "0.0"
        s.Cells(r, 4).NumberFormat = "0.0"
        s.Cells(r, 5).NumberFormat = "0%"
        s.Cells(r, 5).HorizontalAlignment = xlCenter
        r = r + 1
    Next w
    rF = r - 1
    If rF >= rD Then
        Bordures s.Range(s.Cells(rD - 1, 1), s.Cells(rF, 5))
    End If

    ' ------------------------------------------------ TÂCHES
    r = r + 1
    s.Cells(r, 1).Value = "TÂCHES"
    FormatTitre s.Cells(r, 1)
    s.Cells(r, 1).Font.Size = 11
    r = r + 1
    EcrireLigne s, r, Array("Indicateur", "Valeur")
    StyleEntete s.Range(s.Cells(r, 1), s.Cells(r, 2))
    r = r + 1
    nTot = 0
    nNCt = 0
    nECt = 0
    nTt = 0
    nRet = 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For i = 1 To lo.DataBodyRange.Rows.Count
                If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value))) > 0 Then
                    nTot = nTot + 1
                    p = Trim$(CStr(lo.DataBodyRange.Cells(i, 10).Value))
                    If StrComp(p, ETAT_T, vbTextCompare) = 0 Then
                        nTt = nTt + 1
                    ElseIf StrComp(p, ETAT_EC, vbTextCompare) = 0 Then
                        nECt = nECt + 1
                    Else
                        nNCt = nNCt + 1
                    End If
                    v = lo.DataBodyRange.Cells(i, 11).Value
                    If Not IsEmpty(v) And IsNumeric(v) Then
                        If CLng(v) > 0 Then nRet = nRet + 1
                    End If
                End If
            Next i
        End If
    End If
    EcrireLigne s, r, Array("Nombre total de tâches", nTot): r = r + 1
    EcrireLigne s, r, Array("Non commencées", nNCt): r = r + 1
    EcrireLigne s, r, Array("En cours", nECt): r = r + 1
    EcrireLigne s, r, Array("Terminées", nTt): r = r + 1
    EcrireLigne s, r, Array("Tâches en retard", nRet): r = r + 1
    s.Range(s.Cells(r - 5, 2), s.Cells(r - 1, 2)).NumberFormat = "0"
    s.Range(s.Cells(r - 6, 1), s.Cells(r - 1, 2)).Borders.LineStyle = xlContinuous
    If nRet > 0 Then
        s.Cells(r - 1, 1).Font.Bold = True
        s.Cells(r - 1, 2).Font.Bold = True
        s.Cells(r - 1, 2).Font.Color = RGB(192, 0, 0)
    End If

    ' Largeurs
    s.Columns(1).ColumnWidth = 26
    s.Range(s.Columns(2), s.Columns(8)).ColumnWidth = 15

    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------
' MISE À JOUR GLOBALE
'--------------------------------------------------------------------------------
Public Sub MajTout()
    Dim noms As Variant, i As Long
    noms = Array(SH_SAI, SH_CHA, SH_RES, SH_EQ, SH_FER)
    For i = LBound(noms) To UBound(noms)
        If Feuille(CStr(noms(i))) Is Nothing Then
            MsgBox "La feuille '" & noms(i) & "' est absente." & vbCrLf & _
                   "Exécutez d'abord la macro ConstruireClasseur.", vbExclamation
            Exit Sub
        End If
    Next i

    Application.ScreenUpdating = False
    On Error GoTo Erreur
    RecalculerSaisie
    MajCharge
    MajResume
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur pendant la mise à jour : " & Err.Description, vbExclamation
End Sub

'--------------------------------------------------------------------------------
' JOURS FÉRIÉS FRANÇAIS (calcul automatique de Pâques, Ascension, Pentecôte)
'--------------------------------------------------------------------------------
Private Function DimanchePaques(an As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim f As Long, g As Long, h As Long, i As Long, k As Long
    Dim l As Long, m As Long, mois As Long, jour As Long
    a = an Mod 19
    b = an \ 100
    c = an Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mois = (h + l - 7 * m + 114) \ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    DimanchePaques = DateSerial(an, mois, jour)
End Function

Private Sub AjouterFerie(s As Worksheet, ByRef ligne As Long, d As Date, libelle As String)
    s.Cells(ligne, 1).Value = d
    s.Cells(ligne, 2).Value = libelle
    ligne = ligne + 1
End Sub

Private Sub RemplirFeries(s As Worksheet, an As Long, ByRef ligne As Long)
    Dim paques As Date
    paques = DimanchePaques(an)
    AjouterFerie s, ligne, DateSerial(an, 1, 1), "Jour de l'an"
    AjouterFerie s, ligne, paques + 1, "Lundi de Pâques"
    AjouterFerie s, ligne, DateSerial(an, 5, 1), "Fête du travail"
    AjouterFerie s, ligne, DateSerial(an, 5, 8), "Victoire 1945"
    AjouterFerie s, ligne, paques + 39, "Ascension"
    AjouterFerie s, ligne, paques + 50, "Lundi de Pentecôte"
    AjouterFerie s, ligne, DateSerial(an, 7, 14), "Fête nationale"
    AjouterFerie s, ligne, DateSerial(an, 8, 15), "Assomption"
    AjouterFerie s, ligne, DateSerial(an, 11, 1), "Toussaint"
    AjouterFerie s, ligne, DateSerial(an, 11, 11), "Armistice"
    AjouterFerie s, ligne, DateSerial(an, 12, 25), "Noël"
End Sub

Public Sub AjouterFeriesAnnee()
    Dim rep As String, an As Long, s As Worksheet, ligne As Long
    rep = InputBox("Année des jours fériés à ajouter (par exemple 2028) :", "Jours fériés")
    If Len(Trim$(rep)) = 0 Then Exit Sub
    If Not IsNumeric(rep) Then
        MsgBox "Année invalide.", vbExclamation
        Exit Sub
    End If
    an = CLng(Val(rep))
    Set s = FeuilleReq(SH_FER)
    ligne = s.Cells(s.Rows.Count, 1).End(xlUp).Row + 1
    RemplirFeries s, an, ligne
    MsgBox "Jours fériés de " & an & " ajoutés à la feuille 'Jours Fériés'.", vbInformation
End Sub

'--------------------------------------------------------------------------------
' MISE EN FORME : petites procédures partagées
'--------------------------------------------------------------------------------
Private Sub FormatTitre(c As Range)
    c.Font.Bold = True
    c.Font.Size = 14
    c.Font.Color = RGB(31, 78, 120)
End Sub

Private Sub StyleEntete(rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(68, 114, 196)
    rng.Font.Color = RGB(255, 255, 255)
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub Bordures(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(166, 166, 166)
    End With
End Sub

Private Sub EcrireLigne(s As Worksheet, r As Long, valeurs As Variant, Optional col As Long = 1)
    Dim i As Long
    For i = LBound(valeurs) To UBound(valeurs)
        s.Cells(r, col + i - LBound(valeurs)).Value = valeurs(i)
    Next i
End Sub

'--------------------------------------------------------------------------------
' CONSTRUCTION DU CLASSEUR
'--------------------------------------------------------------------------------
Public Sub ConstruireClasseur()
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Cette macro construit (ou reconstruit) les 5 feuilles du planning :" & vbCrLf & _
                     "Saisie, Charge, Résumé, Équipe, Jours Fériés." & vbCrLf & vbCrLf & _
                     "Les données des anciennes versions de ces feuilles seront effacées." & vbCrLf & _
                     "Continuer ?", vbYesNo + vbQuestion, "Construction du classeur")
    If reponse <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Erreur

    Dim s As Worksheet, lo As ListObject, lr As ListRow, btn As Button
    Dim ligne As Long, i As Long

    ' ============================ FEUILLE ÉQUIPE ============================
    Set s = FeuilleNeuve(SH_EQ, RGB(128, 128, 128))
    s.Range("A1").Value = "MEMBRES DE L'ÉQUIPE"
    FormatTitre s.Range("A1")
    EcrireLigne s, 2, Array("Prénom", "Capacité (j/sem)", "Rôle / Notes")
    StyleEntete s.Range("A2:C2")
    s.Range("A3").Value = "Alice": s.Range("B3").Value = 5
    s.Range("A4").Value = "Bruno": s.Range("B4").Value = 5
    s.Range("A5").Value = "Chloé": s.Range("B5").Value = 4: s.Range("C5").Value = "Temps partiel (4 j/sem)"
    s.Range("A6").Value = "David": s.Range("B6").Value = 5
    Bordures s.Range("A2:C6")

    s.Range("F1").Value = "ABSENCES (congés, formations...)"
    FormatTitre s.Range("F1")
    EcrireLigne s, 2, Array("Prénom", "Type", "Début", "Fin", "Commentaire"), 6
    StyleEntete s.Range("F2:J2")
    s.Range("F3").Value = "Alice": s.Range("G3").Value = "Congé"
    s.Range("H3").Value = DateSerial(2026, 12, 21): s.Range("I3").Value = DateSerial(2027, 1, 1)
    s.Range("J3").Value = "Congés de fin d'année"
    s.Range("F4").Value = "Bruno": s.Range("G4").Value = "Formation"
    s.Range("H4").Value = DateSerial(2026, 11, 16): s.Range("I4").Value = DateSerial(2026, 11, 17)
    s.Range("J4").Value = "Formation sécurité"
    s.Range("H3:I500").NumberFormat = "dd/mm/yyyy"
    Bordures s.Range("F2:J4")

    s.Range("L1").Value = "PARAMÈTRES"
    FormatTitre s.Range("L1")
    s.Range("L2").Value = "Début de la grille (lundi, vide = auto)"
    s.Range("L3").Value = "Nombre de semaines"
    s.Range("L4").Value = "Seuil sous-charge (%)"
    s.Range("L5").Value = "Seuil surcharge (%)"
    s.Range("L6").Value = "Capacité par défaut (j/sem)"
    s.Range("M3").Value = 26
    s.Range("M4").Value = 80
    s.Range("M5").Value = 120
    s.Range("M6").Value = 5
    s.Range("L2:L6").Font.Bold = True
    Bordures s.Range("L1:M6")

    s.Columns(1).ColumnWidth = 14
    s.Columns(2).ColumnWidth = 15
    s.Columns(3).ColumnWidth = 24
    s.Columns(6).ColumnWidth = 12
    s.Columns(7).ColumnWidth = 12
    s.Columns(8).ColumnWidth = 12
    s.Columns(9).ColumnWidth = 12
    s.Columns(10).ColumnWidth = 24
    s.Columns(12).ColumnWidth = 36
    s.Columns(13).ColumnWidth = 10

    ' Validations de la feuille Équipe
    With s.Range("F3:F500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeEquipe"
    End With
    With s.Range("G3:G500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Congé,Formation,Autre"
    End With

    ' Nom dynamique pour la liste des membres
    ThisWorkbook.Names.Add Name:="ListeEquipe", RefersTo:="=" & SH_EQ & "!$A$3:$A$200"

    ' ============================ FEUILLE JOURS FÉRIÉS ============================
    Set s = FeuilleNeuve(SH_FER, RGB(128, 128, 128))
    s.Range("A1").Value = "JOURS FÉRIÉS (France)"
    FormatTitre s.Range("A1")
    EcrireLigne s, 2, Array("Date", "Libellé")
    StyleEntete s.Range("A2:B2")
    ligne = 3
    RemplirFeries s, 2026, ligne
    RemplirFeries s, 2027, ligne
    s.Range(s.Cells(3, 1), s.Cells(ligne - 1, 1)).NumberFormat = "dd/mm/yyyy"
    s.Range("D2").Value = "Ajoutez ici vos jours exceptionnels (ponts, fermeture d'entreprise...), puis relancez la mise à jour."
    s.Range("D2").Font.Italic = True
    s.Range("D2").Font.Color = RGB(120, 120, 120)
    s.Columns(1).ColumnWidth = 12
    s.Columns(2).ColumnWidth = 22
    Bordures s.Range(s.Cells(2, 1), s.Cells(ligne - 1, 2))

    ' ============================ FEUILLE SAISIE ============================
    Set s = FeuilleNeuve(SH_SAI, RGB(0, 153, 0))
    s.Range("A1").Value = "SAISIE DES TÂCHES"
    s.Rows(1).RowHeight = 24
    FormatTitre s.Range("A1")
    EcrireLigne s, 2, Array("N°", "Prénom", "Projet", "Description", "Durée (j ouvrés)", _
                            "Début (vide = auto)", "Début effectif", "Fin calculée", _
                            "Avancement (%)", "État", "Retard (j)", "Notes")
    Set lo = s.ListObjects.Add(xlSrcRange, s.Range("A2:L2"), , xlYes)
    lo.Name = "TblSaisie"
    On Error Resume Next
    lo.TableStyle = "TableStyleMedium9"
    On Error GoTo Erreur
    StyleEntete s.Range("A2:L2")
    s.Range("G2").Interior.Color = RGB(112, 134, 178)
    s.Range("H2").Interior.Color = RGB(112, 134, 178)
    s.Range("K2").Interior.Color = RGB(112, 134, 178)

    ' Exemples de tâches (remplaçables par vos données)
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(1, "Alice", "Projet Alpha", "Cadrage et spécifications", 5, _
                           DateSerial(2026, 11, 9), Empty, Empty, 40, ETAT_EC, Empty, "")
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(2, "Alice", "Projet Alpha", "Développement module A", 8, _
                           Empty, Empty, Empty, 0, Empty, Empty, "Début vide : s'enchaîne sur la tâche n°1")
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(3, "Bruno", "Projet Bêta", "Audit sécurité", 3, _
                           DateSerial(2026, 11, 11), Empty, Empty, 60, ETAT_EC, Empty, "Le 11/11 (férié) est sauté")
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(4, "Chloé", "Projet Alpha", "Recette fonctionnelle", 6, _
                           Empty, Empty, Empty, 0, Empty, Empty, "")
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(5, "David", "Projet Gamma", "Rédaction documentation", 2, _
                           DateSerial(2026, 10, 26), Empty, Empty, 100, ETAT_T, Empty, "")
    Set lr = lo.ListRows.Add
    lr.Range.Value = Array(6, "Chloé", "Projet Bêta", "Refonte de la page d'accueil", 4, _
                           Empty, Empty, Empty, 10, ETAT_EC, Empty, "")

    ' Formats de colonnes
    s.Range("F3:G1000").NumberFormat = "dd/mm/yyyy"
    s.Range("H3:H1000").NumberFormat = "dd/mm/yyyy"
    s.Range("I3:I1000").NumberFormat = "0"" %"""
    s.Range("K3:K1000").NumberFormat = "0"
    s.Range("D3:D1000").WrapText = True
    s.Columns(1).ColumnWidth = 5
    s.Columns(2).ColumnWidth = 12
    s.Columns(3).ColumnWidth = 16
    s.Columns(4).ColumnWidth = 42
    s.Columns(5).ColumnWidth = 9
    s.Columns(6).ColumnWidth = 12
    s.Columns(7).ColumnWidth = 12
    s.Columns(8).ColumnWidth = 12
    s.Columns(9).ColumnWidth = 12
    s.Columns(10).ColumnWidth = 15
    s.Columns(11).ColumnWidth = 9
    s.Columns(12).ColumnWidth = 36

    ' Validations de saisie
    With s.Range("B3:B1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeEquipe"
    End With
    With s.Range("E3:E1000").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, Formula1:="0", Formula2:="1000"
    End With
    With s.Range("I3:I1000").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="0", Formula2:="100"
    End With
    With s.Range("J3:J1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=ETAT_NC & "," & ETAT_EC & "," & ETAT_T
    End With

    ' Mises en forme conditionnelles (État et Retard)
    s.Range("J3:J1000").FormatConditions.Delete
    With s.Range("J3:J1000").FormatConditions.Add(xlCellValue, xlEqual, """" & ETAT_T & """")
        .Interior.Color = RGB(198, 239, 206)
    End With
    With s.Range("J3:J1000").FormatConditions.Add(xlCellValue, xlEqual, """" & ETAT_EC & """")
        .Interior.Color = RGB(189, 215, 238)
    End With
    With s.Range("J3:J1000").FormatConditions.Add(xlCellValue, xlEqual, """" & ETAT_NC & """")
        .Interior.Color = RGB(217, 217, 217)
    End With
    s.Range("K3:K1000").FormatConditions.Delete
    With s.Range("K3:K1000").FormatConditions.Add(xlCellValue, xlGreater, "0")
        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)
        .Interior.Color = RGB(255, 199, 206)
    End With

    ' Aide contextuelle sur les en-têtes
    s.Range("F2").ClearComments
    s.Range("F2").AddComment "Laissez vide pour enchaîner automatiquement après la fin de la tâche précédente de la même personne (ordre du tableau)."
    s.Range("G2").ClearComments
    s.Range("G2").AddComment "Calculé automatiquement (date retenue après jours fériés et absences)."
    s.Range("H2").ClearComments
    s.Range("H2").AddComment "Calculée automatiquement : durée en jours disponibles."
    s.Range("K2").ClearComments
    s.Range("K2").AddComment "Calculé : jours ouvrés entre la fin calculée et aujourd'hui si la tâche n'est pas terminée."

    ' Bouton de mise à jour
    Set btn = s.Buttons.Add(200, 2, 190, 20)
    btn.Caption = "Mettre à jour le planning"
    btn.OnAction = "MajTout"
    btn.Font.Size = 10
    btn.Font.Bold = True

    ' ============================ FEUILLE CHARGE ============================
    Set s = FeuilleNeuve(SH_CHA, RGB(0, 102, 204))
    s.Range("A1").Value = "CHARGE PAR PERSONNE ET PAR SEMAINE"
    s.Rows(1).RowHeight = 24
    FormatTitre s.Range("A1")
    Set btn = s.Buttons.Add(330, 2, 150, 20)
    btn.Caption = "Mettre à jour"
    btn.OnAction = "MajTout"
    btn.Font.Size = 10
    btn.Font.Bold = True

    ' ============================ FEUILLE RÉSUMÉ ============================
    Set s = FeuilleNeuve(SH_RES, RGB(237, 125, 49))
    s.Range("A1").Value = "RÉSUMÉ NUMÉRIQUE DE LA CHARGE"
    s.Rows(1).RowHeight = 24
    FormatTitre s.Range("A1")
    Set btn = s.Buttons.Add(330, 2, 150, 20)
    btn.Caption = "Mettre à jour"
    btn.OnAction = "MajTout"
    btn.Font.Size = 10
    btn.Font.Bold = True

    ' ============================ NETTOYAGE ET ORDRE ============================
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(i).Name <> SH_EQ And ThisWorkbook.Worksheets(i).Name <> SH_FER _
           And ThisWorkbook.Worksheets(i).Name <> SH_SAI _
           And ThisWorkbook.Worksheets(i).Name <> SH_CHA _
           And ThisWorkbook.Worksheets(i).Name <> SH_RES Then
            If Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets(i).Cells) = 0 Then
                ThisWorkbook.Worksheets(i).Delete
            End If
        End If
    Next i
    Application.DisplayAlerts = True

    ThisWorkbook.Worksheets(SH_SAI).Move Before:=ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets(SH_CHA).Move After:=ThisWorkbook.Worksheets(SH_SAI)
    ThisWorkbook.Worksheets(SH_RES).Move After:=ThisWorkbook.Worksheets(SH_CHA)
    ThisWorkbook.Worksheets(SH_EQ).Move After:=ThisWorkbook.Worksheets(SH_RES)
    ThisWorkbook.Worksheets(SH_FER).Move After:=ThisWorkbook.Worksheets(SH_EQ)

    ' Première mise à jour : dates, grille de charge et résumé
    MajTout

    ' Volets figés
    ThisWorkbook.Worksheets(SH_SAI).Activate
    ThisWorkbook.Worksheets(SH_SAI).Range("C3").Select
    ActiveWindow.FreezePanes = True
    ThisWorkbook.Worksheets(SH_CHA).Activate
    ThisWorkbook.Worksheets(SH_CHA).Range("B4").Select
    ActiveWindow.FreezePanes = True
    ThisWorkbook.Worksheets(SH_SAI).Activate
    ThisWorkbook.Worksheets(SH_SAI).Range("A3").Select

    Application.ScreenUpdating = True
    MsgBox "Le classeur est prêt !" & vbCrLf & vbCrLf & _
           "- Saisissez vos tâches dans la feuille Saisie" & vbCrLf & _
           "- Configurez l'équipe, les congés, formations et jours fériés" & vbCrLf & _
           "- Cliquez sur 'Mettre à jour le planning' après chaque modification" & vbCrLf & vbCrLf & _
           "Pensez à enregistrer le fichier au format .xlsm.", vbInformation, "Planning d'équipe"
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Erreur pendant la construction : " & Err.Description, vbCritical
End Sub

Private Function FeuilleNeuve(nom As String, coulOnglet As Long) As Worksheet
    Dim s As Worksheet
    Set s = Feuille(nom)
    If Not s Is Nothing Then
        Application.DisplayAlerts = False
        s.Delete
        Application.DisplayAlerts = True
    End If
    Set FeuilleNeuve = ThisWorkbook.Worksheets.Add
    FeuilleNeuve.Name = nom
    FeuilleNeuve.Tab.Color = coulOnglet
End Function
