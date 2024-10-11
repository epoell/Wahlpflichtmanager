'Konstanten
'----------
'Anzahl moeglicher Faecher, muss fest sein um Arrays initiieren zu koennen
Public Const Faecherzahl As Long = 5
'Tabellennamen
Public Const StrWahlen As String = "Wahlen"
Public Const StrWahlopt As String = "Wahlmoeglichkeiten"
Public Const StrZuteilung As String = "Zuteilung"
'Spaltensortierung Wahloptionen
'Kennziffer | Fachname | Kursgroesse
Public Const SpalteZiffer   As Long = 1
Public Const SpalteFach     As Long = 2
Public Const SpalteGroesse  As Long = 3
'Spaltensortierung Wahlen
'Vorname | Nachname | Klasse | 1. Wunsch | 2. Wunsch | 3. Wunsch | 4. Wunsch | 5. Wunsch
Public Const SpalteVorname  As Long = 1
Public Const SpalteNachname As Long = 2
Public Const SpalteKlasse   As Long = 3
Public Const Spalte1Wunsch  As Long = 4
Public Const Spalte2Wunsch  As Long = 5
Public Const Spalte3Wunsch  As Long = 6
Public Const Spalte4Wunsch  As Long = 7
Public Const Spalte5Wunsch  As Long = 8
'Spaltensortieung ZuteilungsSheet
'Wie bei Wahlen, zuteilungSpalte daneben mit Leerspalte fuer Optik
Public Const ZutSpalte      As Long = 10
Public Const ZutFachSpalte  As Long = 11
'Ausgabe wer mit wem getauscht
Public Const AusgabeBln As Boolean = True

Sub FachVerteilung()
    '------------------------------------------------------
    'Pruefen ob Arbeitsmappe korrekt aufgebaut
    'und Konstanten/Variablen richtig belegt
    '------------------------------------------------------
    Dim Response As Long
    Dim Msg, Style, Title
    
    'Wahloptionen liegen in "Wahlmoeglichkeiten"
    If WorksheetMissing(StrWahlopt) Then
        Title = "Tabelle kontrollieren"
        Msg = "Die Wahlmoeglichkeiten muessen im Tabellenblatt '" & StrWahlopt & "' abgelegt werden." & vbNewLine & "Bitte Tabellenblatt-Namen anpassen."
        Style = vbOKOnly Or vbCritical  'Nur die Schaltflaehen OK anzeigen und das Symbol Kritische Meldung anzeigen.
        Response = MsgBox(Msg, Style, Title)
        Exit Sub
    End If
    
    ThisWorkbook.Sheets(StrWahlopt).Activate
    
    'In Wahloptionnen hat Titel, es geht mit Kennziffer 1 los
    If ThisWorkbook.Sheets(StrWahlopt).Cells(2, SpalteZiffer).Value <> "1" Then
        Title = "Tabelle kontrollieren"
        Msg = "Im Tabellenblatt '" & StrWahlopt & "' muessen in Spalte A die Kennziffern stehen (aufsteigend sortiert) - und in Zelle A2 mit 1 beginnen." _
                & vbNewLine & "Bitte anpassen."
        Style = vbOKOnly Or vbCritical  'Nur die Schaltflaehen OK anzeigen und das Symbol Kritische Meldung anzeigen.
        Response = MsgBox(Msg, Style, Title)
        Exit Sub
    End If
    
    'Wahloptionen hat die richtigen Spalten
    Title = "Tabelle kontrollieren"
    Msg = "Im Tabellenblatt '" & StrWahlopt & "' stehen in Spalte A ist die Kennziffern (aufsteigend sortiert)." & vbNewLine _
            & "Spalte B gibt das Fach an und Spalte C die Kursgroesse." & vbNewLine _
            & "Korrekt?"
    Style = vbOKCancel Or vbInformation 'Nur die Schaltflaechen OK und Abbrechen anzeigen und das Symbol Informationsmeldung anzeigen.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbCancel Then
        'Abbruch des Subs bei falschem Format -> Abbrechen-Button geklickt
        Exit Sub
    End If
    
    'Faecherzahl ist die richtige Konstante
    If Faecherzahl <> Faecherzaehlen Then
        Title = "Konstante anpassen"
        Msg = "Die Konstante fuer die Anzahl der Wahlmoeglichkeiten ist nicht aktuell. Bitte anpassen!" & vbNewLine _
                & "Im Makro ganz oben bei den Konstanten: 'Faecherzahl' sollte " & Faecherzaehlen & " sein."
        Style = vbOKOnly Or vbCritical  'Nur die Schaltflaehen OK anzeigen und das Symbol Kritische Meldung anzeigen.
        Response = MsgBox(Msg, Style, Title)
        Exit Sub
    End If
    
    '---
    
    'Wahlen der Schueler:innen liegen in "Wahlen"
    If WorksheetMissing(StrWahlen) Then
        Title = "Tabelle kontrollieren"
        Msg = "Die Wahlen der Schueler:innen muessen im Tabellenblatt '" & StrWahlen & "' abgelegt werden." & vbNewLine _
                & "Bitte Tabellenblatt-Namen anpassen."
        Style = vbOKOnly Or vbCritical  'Nur die Schaltflaehen OK anzeigen und das Symbol Kritische Meldung anzeigen.
        Response = MsgBox(Msg, Style, Title)
        Exit Sub
    End If
    
    ThisWorkbook.Sheets(StrWahlen).Activate
    
    'Wahlen hat die richtigen Spalten
    'Vorname | Nachname | Klasse | 1. Wunsch | 2. Wunsch | 3. Wunsch | 4. Wunsch | 5. Wunsch
    Title = "Tabelle kontrollieren"
    Msg = "Im Tabellenblatt '" & StrWahlen & "' stehen die Namen in Spalte A und B," & vbNewLine & "die Klasse in Spalte C," & vbNewLine _
            & "und die Wuensche in Spalten D bis H. " & vbNewLine _
            & "Die Spalten fuer Klasse (C) und die 3. - 5. Wuensche (F, G, H) duerfen leer sein." _
            & vbNewLine & vbNewLine & "Ist die Tabelle korrekt formatiert?"
    Style = vbOKCancel Or vbInformation 'Nur die Schaltflaechen OK und Abbrechen anzeigen und das Symbol Informationsmeldung anzeigen.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbCancel Then
        'Abbruch des Subs bei falschem Format -> Abbrechen-Button geklickt
        Exit Sub
    End If
    
    '---
    
    Dim Sorted As Boolean
    Sorted = False 'Gibt an, ob Schuelerliste nach Priotitaet sortiert ist, oder (bei False) zufaellig sortiert werden soll
    
    'Schuelerliste nach Prio oder zufaellg sortieren?
    Title = "Sortierunng waehlen"
    Msg = "Ist die Liste der Schueler:innen nach absteigender Prioritaet sortiert?" & vbNewLine & vbNewLine & "Z.B. nach Eingang der Rueckmeldungen, sodass Schueler:innen weiter oben in der Liste eher ihren Erstwunsch bekommen."
    Style = vbYesNo Or vbInformation 'Die Schaltflaechen Ja und Nein anzeigen und das Symbol Informationsmeldung anzeigen.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then
        Sorted = True
    Else
        Sorted = False
    End If
    
    '------------------------------------------------------
   
   
    '------------------------------------------------------
    'SetUp
    'Anlegen des Tabellenblatts zur Verteilung und der Arrays mit den Wahlen und freien Plaetzen
    '------------------------------------------------------
    
    'Aktivitaten verbergen, bis alles fertig ist
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim count As Long
    Dim student As Long
    Dim fach As Long
    Dim ZuteilungsSheet As Worksheet
    Dim Sheetname As String
    
    '1. Zuteilung-Tabellenblatt anlegen und Schueler:innenliste mit Wuenschen kopieren
    With ThisWorkbook
        count = .Sheets.count
        Worksheets(StrWahlen).Copy After:=.Sheets(count)
        'Tabelle umbennnen
        'Um bei jeder euen Verteilung einen individ. Namen anzugeben, wird einfach gezaehlt, minus die zwei festen Tabellen Moeglichkeiten und Wahlen
        i = count - 2
        Sheetname = StrZuteilung & i
        'wenn Name noch nicht vergeben, dann benutzen (sonst Neuen erzeugen)
'Sprungmaker SetName
SetName:
        If WorksheetMissing(Sheetname) Then
            ActiveSheet.name = Sheetname
        Else
            i = i + 1
            Sheetname = StrZuteilung & i
            GoTo SetName
        End If
        
        'Das neu erzeugte Blatt merken
        Set ZuteilungsSheet = ActiveSheet
        
        ' Zellen leeren, die spaeter eh ueberschrieben werden
        ZuteilungsSheet.Range(Columns(Spalte5Wunsch + 1), Columns(ZutFachSpalte + 1)).Clear
    End With
    
    ZuteilungsSheet.Cells(1, ZutSpalte) = "Zuteilung"
    ZuteilungsSheet.Cells(1, ZutFachSpalte) = "Fachname"
    
    '2. Zaehlen, wie oft welches Fach gewaehlt wurde und in 3 Arrays ablegen, Index ist Fachkennziffer
    Dim Erstwunsch(Faecherzahl) As Long
    Dim Zweitwunsch(Faecherzahl) As Long
    Dim Drittwunsch(Faecherzahl) As Long
    Dim Viertwunsch(Faecherzahl) As Long
    Dim Fuenftwunsch(Faecherzahl) As Long
    Dim FirstStudent As Long
    Dim LastStudent As Long
    Dim BelowStudents As Long
    Dim BelowSubjects As Long
    FirstStudent = 2
    LastStudent = Schuelerzaehlen
    BelowStudents = LastStudent + 5
    BelowSubjects = BelowStudents + Faecherzahl + 2
    'Hinweistabelle vorbereiten:
    ZuteilungsSheet.Cells(BelowStudents, 1) = "Kennziffer"
    ZuteilungsSheet.Cells(BelowStudents, 2) = "Fach"
    ZuteilungsSheet.Cells(BelowStudents, 3) = "# Erstwunsch"
    ZuteilungsSheet.Cells(BelowStudents, 4) = "# Zweitwunsch"
    ZuteilungsSheet.Cells(BelowStudents, 5) = "# Drittwunsch"
    ZuteilungsSheet.Cells(BelowStudents, 6) = "# Viertwunsch"
    ZuteilungsSheet.Cells(BelowStudents, 7) = "# Fuenftwunsch"
    
    'Zaehlen
    For fach = 1 To Faecherzahl
        With ThisWorkbook.Sheets(StrWahlen)
            Erstwunsch(fach) = WorksheetFunction.CountIf(.Range(.Cells(FirstStudent, Spalte1Wunsch), .Cells(LastStudent, Spalte1Wunsch)), fach)   'C2:C LastStudent
            Zweitwunsch(fach) = WorksheetFunction.CountIf(.Range(.Cells(FirstStudent, Spalte2Wunsch), .Cells(LastStudent, Spalte2Wunsch)), fach)  'D2:D LastStudent
            Drittwunsch(fach) = WorksheetFunction.CountIf(.Range(.Cells(FirstStudent, Spalte3Wunsch), .Cells(LastStudent, Spalte3Wunsch)), fach)  'E2:E LastStudent
            Viertwunsch(fach) = WorksheetFunction.CountIf(.Range(.Cells(FirstStudent, Spalte4Wunsch), .Cells(LastStudent, Spalte4Wunsch)), fach)
            Fuenftwunsch(fach) = WorksheetFunction.CountIf(.Range(.Cells(FirstStudent, Spalte5Wunsch), .Cells(LastStudent, Spalte5Wunsch)), fach)
        End With
        'Als Hinweis ablegen:
        ZuteilungsSheet.Cells(BelowStudents + fach, 1) = fach
        ZuteilungsSheet.Cells(BelowStudents + fach, 2) = GetSubjectname(fach)
        ZuteilungsSheet.Cells(BelowStudents + fach, 3) = Erstwunsch(fach)
        ZuteilungsSheet.Cells(BelowStudents + fach, 4) = Zweitwunsch(fach)
        ZuteilungsSheet.Cells(BelowStudents + fach, 5) = Drittwunsch(fach)
        ZuteilungsSheet.Cells(BelowStudents + fach, 6) = Viertwunsch(fach)
        ZuteilungsSheet.Cells(BelowStudents + fach, 7) = Fuenftwunsch(fach)
        
    Next fach

        
    '3. Array anlegen, mit verfuegbaren Plaetzen je Fach, Index ist Fachkennziffer
    Dim FreiePlaetze(Faecherzahl) As Long
    
    For fach = 1 To Faecherzahl
        FreiePlaetze(fach) = ThisWorkbook.Sheets(StrWahlopt).Cells(fach + 1, SpalteGroesse).Value 'ConvertCellToInt(ThisWorkbook.Sheets(StrWahlopt).Cells(i, SpalteGroesse))
    Next fach
        
    'Wenn Liste ohne Priorisierung, dann zufaellig sortieren um Vergabe so fair wie moeglich zu machen
    If Sorted = False Then
        'Spalte fuer Zufallszahlen einfuegen
        ZuteilungsSheet.Range("A:A").Insert
        'Zufallszahlen
        Dim max As Long
        max = LastStudent * 10
        For student = FirstStudent To LastStudent
            Randomize
            ZuteilungsSheet.Cells(student, 1).Value = WorksheetFunction.RandBetween(1, max)
        Next student
        'nach Zufallszahlen sortieren
        Range(Columns(1), Columns(Spalte5Wunsch + 1)).Sort key1:=Range("A2"), order1:=xlAscending, Header:=xlYes
        'Spalte mit Zufallszahlen loeschen
        ZuteilungsSheet.Range("A:A").Delete
    End If
    '------------------------------------------------------
        
    
    '------------------------------------------------------
    'Verteilung
    'Annahme bei Sorted = True: Schuelerinnen die frueh abgeben, stehen oben in der Liste. Je weiter oben jemand steht, desto eher bekommen sie ihren Erstwunsch
    'Sonst: erstmal zufaellig sortieren
    '------------------------------------------------------
    Dim wunsch1 As Long
    Dim wunsch2 As Long
    Dim ohnePlatz As Long
    ohnePlatz = 0
    
    'Schueler:innen von oben nach unten durchgehen
    For student = FirstStudent To LastStudent
        'Wuensche aus Tabelle auslesen (Erstwunsch in Zpalte C, usw.)
        wunsch1 = GetWishy(ZuteilungsSheet, student, 1)
        wunsch2 = GetWishy(ZuteilungsSheet, student, 2)
        
        '4.Erstwunsch zuteilen, bis Kurs voll
        If FreiePlaetze(wunsch1) > 0 Then
            'Zuteilung aufschreiben
            ZuteilungsSheet.Cells(student, ZutSpalte) = wunsch1 'Kennziffer
            ZuteilungsSheet.Cells(student, ZutFachSpalte) = GetSubjectname(wunsch1) 'Fachname
            'Freie Plaetze verringern
            FreiePlaetze(wunsch1) = FreiePlaetze(wunsch1) - 1
            
        '5. Wenn Erstwunsch schon voll, Zweitwunsch vergeben
        ElseIf FreiePlaetze(wunsch2) > 0 Then
            'Zuteilung aufschreiben
            ZuteilungsSheet.Cells(student, ZutSpalte) = wunsch2 'Kennziffer
            ZuteilungsSheet.Cells(student, ZutSpalte + 1) = GetSubjectname(wunsch2) 'Fachname
            'Freie Plaetze verringern
            FreiePlaetze(wunsch2) = FreiePlaetze(wunsch2) - 1
        End If
    Next student
    
    '------------------------------------------------------

    Dim wunsch1x As Long
    Dim wunsch2x As Long
    Dim wunsch3x As Long
    Dim zuteilungy As Long
    Dim wunsch1y As Long
    Dim wunsch2y As Long
    
    If AusgabeBln Then
        i = 1
    End If
    
    '6-9. Bei allen verbleibenden gucken, ob man Tauschen kann
    With ZuteilungsSheet
        'von oben nach unten Schueler:in x waehlen, die noch kein Fach haben
        For x = FirstStudent To LastStudent
            If IsEmpty(.Cells(x, ZutSpalte).Value) Then 'Zuteilung wuerde in Cells(x, zutSpalte) stehen
                'Wuensche von x bestimmen
                wunsch1x = GetWishy(ZuteilungsSheet, x, 1)
                wunsch2x = GetWishy(ZuteilungsSheet, x, 2)
                wunsch3x = GetWishy(ZuteilungsSheet, x, 3)
                wunsch4x = GetWishy(ZuteilungsSheet, x, 4)
                wunsch5x = GetWishy(ZuteilungsSheet, x, 5)
                
                '---------------------------------------------------
                'Fuer 3.Wuensche
                '---------------------------------------------------
                '6.a von unten nach oben: Schueler:in y finden, (von unten nach oben relevant, wenn Liste nach Priortaet sortiert, sonst egal)
                For y = LastStudent To FirstStudent Step -1
                    'mit Erstwunsch im Fach, das x als Zweitwunsch hat, und mit Zweitwunsch von y noch nicht voll
                    zuteilungy = ConvertCellToInt(.Cells(y, ZutSpalte))
                    wunsch1y = GetWishy(ZuteilungsSheet, y, 1)
                    wunsch2y = GetWishy(ZuteilungsSheet, y, 2)
                    If zuteilungy = wunsch1y And wunsch2x = wunsch1y And FreiePlaetze(wunsch2y) > 0 Then
                        'tauschen
                        'x den Zweitwunsch geben (erstwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch2x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch2x) 'Fachname
                        'y den Zweitwunsch geben
                        .Cells(y, ZutSpalte) = wunsch2y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch2y) 'Fachname
                        FreiePlaetze(wunsch2y) = FreiePlaetze(wunsch2y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Zweitwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                    End If
                Next y
                
                '6.b wenn es keine:n y gibt, dann bekommt x Drittwunsch, wenn noch Platz da ist
                If FreiePlaetze(wunsch3x) > 0 Then
                    .Cells(x, ZutSpalte) = wunsch3x 'Kennziffer
                    .Cells(x, ZutSpalte + 1) = GetSubjectname(wunsch3x) 'Fachname
                    FreiePlaetze(wunsch3x) = FreiePlaetze(wunsch3x) - 1
                    
                    'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                    GoTo NextStudent
                End If
                
                '6.c wenn in Drittwunsch nichts mehr frei ist fuer x, dann ein neues y finden, um Erstwunsch zu tauschen
                'von unten nach oben: Schueler:in y finden,
                For y = LastStudent To FirstStudent Step -1
                    'mit Erstwunsch im Fach, das x als Erstwunsch hat, und mit Zweitwunsch von y noch nicht voll
                    zuteilungy = ConvertCellToInt(.Cells(y, ZutSpalte))
                    wunsch1y = GetWishy(ZuteilungsSheet, y, 1)
                    wunsch2y = GetWishy(ZuteilungsSheet, y, 2)
                    If zuteilungy = wunsch1y And wunsch1x = wunsch1y And FreiePlaetze(wunsch2y) > 0 Then
                        'tauschen
                        'x den Erstwunsch geben (erstwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch1x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch1x) 'Fachname
                        'y den Zweitwunsch geben
                        .Cells(y, ZutSpalte) = wunsch2y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch2y) 'Fachname
                        FreiePlaetze(wunsch2y) = FreiePlaetze(wunsch2y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Erstwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                    End If
                Next y
                
                
                '---------------------------------------------------
                'Fuer 4.Wuensche
                '---------------------------------------------------
                '7.a von unten nach oben: Schueler:in y finden, (von unten nach oben relevant, wenn Liste nach Priortaet sortiert, sonst egal)
                For y = LastStudent To FirstStudent Step -1
                    'mit Erstwunsch im Fach, das x als Drittwunsch hat, und mit Zweitwunsch von y noch nicht voll
                    zuteilungy = ConvertCellToInt(.Cells(y, ZutSpalte))
                    wunsch1y = GetWishy(ZuteilungsSheet, y, 1)
                    wunsch2y = GetWishy(ZuteilungsSheet, y, 2)
                    wunsch3y = GetWishy(ZuteilungsSheet, y, 3)
                    If zuteilungy = wunsch1y And wunsch3x = wunsch1y And FreiePlaetze(wunsch2y) > 0 Then
                        'tauschen
                        'x den Drittwunsch geben (erstwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch3x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch3x) 'Fachname
                        'y den Zweitwunsch geben
                        .Cells(y, ZutSpalte) = wunsch2y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch2y) 'Fachname
                        FreiePlaetze(wunsch2y) = FreiePlaetze(wunsch2y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Drittwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                        
                    'Sonst: mit Zweitwunsch im Fach, das x als Drittwunsch hat, und mit Drittwunsch von y noch nicht voll
                    ElseIf zuteilungy = wunsch2y And wunsch3x = wunsch2y And FreiePlaetze(wunsch3y) > 0 Then
                        'tauschen
                        'x den Drittwunsch geben (erstwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch3x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch3x) 'Fachname
                        'y den Drittwunsch geben
                        .Cells(y, ZutSpalte) = wunsch3y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch3y) 'Fachname
                        FreiePlaetze(wunsch3y) = FreiePlaetze(wunsch3y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Drittwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                    End If
                Next y
                
                '7.b wenn es keine:n y gibt, dann bekommt x Viertwunsch, wenn noch Platz da ist
                If FreiePlaetze(wunsch4x) > 0 Then
                    .Cells(x, ZutSpalte) = wunsch4x 'Kennziffer
                    .Cells(x, ZutSpalte + 1) = GetSubjectname(wunsch4x) 'Fachname
                    FreiePlaetze(wunsch4x) = FreiePlaetze(wunsch4x) - 1
                    
                    'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                    GoTo NextStudent
                End If
                
                '---------------------------------------------------
                'Fuer 5.Wuensche
                '---------------------------------------------------
                '8.a von unten nach oben: Schueler:in y finden, (von unten nach oben relevant, wenn Liste nach Priortaet sortiert, sonst egal)
                For y = LastStudent To FirstStudent Step -1
                   'mit Zweitwunsch im Fach, das x als Viertwunsch hat, und mit Drittwunsch von y noch nicht voll
                    If zuteilungy = wunsch2y And wunsch4x = wunsch2y And FreiePlaetze(wunsch3y) > 0 Then
                        'tauschen
                        'x den Viertwunsch geben (Zweitwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch4x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch4x) 'Fachname
                        'y den Drittwunsch geben
                        .Cells(y, ZutSpalte) = wunsch3y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch3y) 'Fachname
                        FreiePlaetze(wunsch3y) = FreiePlaetze(wunsch3y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Viertwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                        
                    'mit Drittwunsch im Fach, das x als Viertwunsch hat, und mit Viertwunsch von y noch nicht voll
                    ElseIf zuteilungy = wunsch3y And wunsch4x = wunsch3y And FreiePlaetze(wunsch4y) > 0 Then
                        'tauschen
                        'x den Viertwunsch geben (Drittwunsch von y)
                        .Cells(x, ZutSpalte) = wunsch4x 'Kennziffer
                        .Cells(x, ZutFachSpalte) = GetSubjectname(wunsch4x) 'Fachname
                        'y den Viertwunsch geben
                        .Cells(y, ZutSpalte) = wunsch4y 'Kennziffer
                        .Cells(y, ZutFachSpalte) = GetSubjectname(wunsch4y) 'Fachname
                        FreiePlaetze(wunsch4y) = FreiePlaetze(wunsch4y) - 1
                        
                        'Ggf. Tausch notieren
                        If AusgabeBln Then
                            PrintSwitchInfo ZuteilungsSheet, BelowSubjects + i, x, y, "Viertwunsch"
                            i = i + 1
                        End If
                        
                        'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                        GoTo NextStudent
                    End If
                Next y
                
                '8.b wenn es keine:n y gibt, dann bekommt x Fuenftwunsch, wenn noch Platz da ist
                If FreiePlaetze(wunsch5x) > 0 Then
                    .Cells(x, ZutSpalte) = wunsch4x 'Kennziffer
                    .Cells(x, ZutSpalte + 1) = GetSubjectname(wunsch5x) 'Fachname
                    FreiePlaetze(wunsch5x) = FreiePlaetze(wunsch5x) - 1
                    
                    'Bei Erfolg naechste:n Schueler:in x ohne Zuteilung finden
                    GoTo NextStudent
                End If
                
                
                '---------------------------------------------------
                'Fuer Verbleibende
                '---------------------------------------------------
                '9.wenn kein Platz fuer x ist: Verteilung auslassen und Zelle farbig hinterlegen
                .Cells(x, ZutSpalte).Interior.ColorIndex = 39 '39: violett, 3: knallrot, 6: gelb, 33: hellblau
                ohnePlatz = ohnePlatz + 1
            End If 'isEmpty
            
'Sprungmarke um nach Tausch Schleife zu durchbrechen
NextStudent:
        Next x
    End With
    
    '------------------------------------------------------

        
    '10. uebrig gebliebene Plaetze in Hinweistabelle vermerken
    ZuteilungsSheet.Cells(BelowStudents, ZutSpalte) = "Verbleibende Plaetze"
    For i = 1 To Faecherzahl
        ZuteilungsSheet.Cells(BelowStudents + i, ZutSpalte) = FreiePlaetze(i)
    Next i
    
    ZuteilungsSheet.Cells(BelowStudents, ZutSpalte + 2) = "Schueler:innen ohne Platz"
    ZuteilungsSheet.Cells(BelowStudents + 1, ZutSpalte + 2) = ohnePlatz
    
    '11. Wenn zufaellig sortiert, wieder nach Klasse und Name sortieren
    If Sorted = False Then
        Range(ZuteilungsSheet.Cells(1, 1), ZuteilungsSheet.Cells(LastStudent, ZutFachSpalte)).Sort _
            key1:=Columns(SpalteKlasse), order1:=xlAscending, _
            key2:=Columns(SpalteNachname), order2:=xlAscending, Header:=xlYes
    End If
        
    '12. Obere Zelle auswaehlen fuer Uebersicht
    ZuteilungsSheet.Cells(1, 1).Activate
    
    '13. Aenderungen wieder anzeigen
    Application.ScreenUpdating = True
    
End Sub
   
   
'------------------------------------------------------
'Hilfsfunktionen
'------------------------------------------------------

'Zaehlen wie vile Schueler:innen in der Liste stehen
'Annahme: es stehen nur Faecher in Spalte B in Wahlenmoeglichkeiten
Function Faecherzaehlen() As Long
    Faecherzaehlen = ThisWorkbook.Sheets(StrWahlopt).Cells(Rows.count, 1).End(xlUp).Row - 1 '-1 da Titel vorhanden
End Function

'Zaehlen wie vile Schueler:innen in der Liste stehen
'Annahme: es stehen nur Schuelerinnen in Spalte A in Wahlen
'Ist eig 1 zu viel, da Titel mitgezaehlt, ist aber gut, da es eh nur fuer Zeilennummer genutzt wird
Function Schuelerzaehlen() As Long
    Schuelerzaehlen = ThisWorkbook.Sheets(StrWahlen).Cells(Rows.count, 1).End(xlUp).Row
End Function

'Den Wunsch mit der Prioritaet Prio von dem:der gegegebenen Schueler:in bestimmen
Function GetWishy(ZutSheet As Worksheet, Stud As Variant, Prio As Long) As Long
    'Spalten: 1.Wunsch: C->3 = 2+Prio=1, 2.Wunsch: D->4 = 2+Prio=2, 3.Wunsch: E->5 = 2+Prio=3, ...
    GetWishy = ConvertCellToInt(ZutSheet.Cells(Stud, SpalteKlasse + Prio))
End Function

'Den Namen eines Fachs an Hand der Kennziffer bestimmen
Function GetSubjectname(Subjectnumber As Variant) As String
    GetSubjectname = ThisWorkbook.Sheets(StrWahlopt).Cells(Subjectnumber + 1, SpalteFach).Value
End Function

'Ausgabe wer mit wem getauscht wurde
Sub PrintSwitchInfo(ZutSheet As Worksheet, pos As Long, Stud1 As Variant, Stud2 As Variant, Wishno As String)
    Dim name1 As String
    Dim name2 As String
    
    With ZutSheet
        name1 = .Cells(Stud1, SpalteVorname).Value & " " & .Cells(Stud1, SpalteNachname).Value
        name2 = .Cells(Stud2, SpalteVorname).Value & " " & .Cells(Stud2, SpalteNachname).Value
        
        .Cells(pos, 1) = "Fuer " & Wishno & " von " & name1 & " Platz von " & name2 & " uebernommen"
    End With
End Sub

'Zellinhalt in Integer umwandeln
'Dank an stackoverflow
Function ConvertCellToInt(Cell As Range) As Integer
   On Error GoTo NOT_AN_INTEGER
   ConvertCellToInt = CInt(Cell.Value)
   Exit Function
NOT_AN_INTEGER:
   ConvertCellToInt = 0
End Function

'Testen, ob es ein Worksheet mit dem gegebenen Namen gibt
'Dank an Tim Williams, https://stackoverflow.com/a/6688482
Function WorksheetMissing(shtName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(shtName)
    On Error GoTo 0
    WorksheetMissing = sht Is Nothing 'sht Is Nothing ist True wenn kein Sheet mit dem Namen gefunden wurde
End Function
