# Wahlpflichtmanager
VBA Macro, dass so fair wie möglich Fächer an Schüler:innen verteilt
VBA Macro, that tries to distribute chosen subjects as fair as possible among students/pupils.

In einer Excelarbeitsmappe gibt es zwei Tabellenblätter:
"Wahlmoeglichkeiten" listet die Namen der moeglichen Faecher auf und die Kursgröße je Fach.
"Wahlen" listet alle Schüler:innen auf mit Nachname, Vorname und jeweils deren Erstwunsch, Zweitwunsch und Drittwunsch der möglichen Fächer.

Aus diesen Angaben wird dann ein neues Tabellenblatt mit einer Zuteilung erzeugt.
Die Zuteilung ist nicht maximal optimiert, sondern folgt pragmatisch dem folgenden Vorgehen:
1. Alle Erstwünsche werden von oben nach unten vergeben, bis alle Plätze voll sind
2. Alle Zweitwünsche werden von oben nach unten vergeben, bis alle Plätze voll sind
3. Wenn jemand noch keinen Platz hat wird versucht zu tauschen, mit jemand der den Erstwunsch bekommen hat, oder man bekommt den Drittwunsch.
4. Das Verfahren geht nicht unbedingt auf. Schüler:innen deren Wünsche nicht erfüllt werden konnten werden lila markiert, zur manuellen Nachberarbeitung.

Achtung! Schüler:innen, die weiter oben in der Liste stehen, werden deshalb aktiv bevorteilt.
(Annahme: je weiter oben jemand steht, desto früher wurde der Wahlzettel mit den Wünschen eingereicht. Das dient als ANreiz, dass Schüler:innen diesen schnell ausfüllen und abgeben.) Alternativ kann man auch die Liste mit den Schüler:innen zufällig sortieren lassen. Dann sollte aber auch im Macro gegen Ende die Zuteilung des Drittwunsches erst nach dem Tausch eines Erst- gegen einen Zweitwunsch geschehen. Die Code-Blöcke kann man einfach tauschen.
