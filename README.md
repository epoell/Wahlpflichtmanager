# Wahlpflichtmanager
VBA Macro für Excel, dass so fair wie möglich Fächer an Schüler:innen verteilt.

VBA Macro for Excel, that tries to distribute chosen subjects as fair as possible among students/pupils.


## Vorgehen
Gegeben ist eine Liste mit Schüler:innen und deren Fach-Wünschen und eine Liste mit möglichen Fächern und deren maximaler Kursgröße.

Daraus wird (in einem neuen Tabellenblatt) eine Zuteilung wie folgt errechnet, um möglichst vielen einen möglichst hohen Wunsch zu ermöglichen:
1. Alle Erstwünsche werden von oben nach unten vergeben, bis alle Plätze voll sind
2. Alle Zweitwünsche werden von oben nach unten vergeben, bis alle Plätze voll sind
3. Wenn jemand noch keinen Platz hat wird versucht zu tauschen, mit jemand der den Erstwunsch bekommen hat. (Bzw. dann für den Zweit-, Dritt- und Viertwunsch).
4. Sonst wird der höchstmögliche noch freie Wunsch zugeteilt.
5. Die Verteilung geht nicht unbedingt auf. Schüler:innen deren Wünsche nicht erfüllt werden konnten werden lila markiert, zur manuellen Nachberarbeitung.

In Testläufen konnten aber 390 aus 400 Schüler:innen zugetilt werden.

Es kann ausgwählt werden (durch einen Dialog beim Start des Macros), ob die Liste der Schüler:innen nach Priorität sorteiert ist, um die Schüler:innen weiter oben in der Liste bei der Wahl zu bevorteilen. Das kann z.B. der chronologischen Reihenfolge entsprechen, in der Schüler:innen die Rückmeldungen einreichen. Was als Anreiz gegenüber den Schüler:innen genutzt werden kann die Wünsche zeitnah einzureichen.
Andernfalls wird die Liste der Schüler:innen zu beginn zufällig sortiert, die Fächer zugeteilt und dann wieder nach Klasse und Nachname sortiert.


## Vorgesehende Formatierung
In einer Excelarbeitsmappe gibt es zwei Tabellenblätter:
"Wahlmoeglichkeiten" listet die Namen der moeglichen Faecher auf und die Kursgröße je Fach.
"Wahlen" listet alle Schüler:innen auf mit Namen, Klasse und jeweils deren Erst- bis Fünftwunsch der möglichen Fächer.

Die Zuteilung wird in einem neuen vollständig genereirten Tabellenblatt erzeugt, sodass die ursprünglichen Liste unverändert bleiben. Man kann meherere Zuteilungen nacheinander in der selben Mappe generieren lassen und die generierten Tabellenblätter können nach Belieben gelöscht werden. Konflikte in der Tabellenbennenung werden im Makro umgangen.

![Bildschirmfoto vom 2023-05-09 12-16-57](https://github.com/epoell/Wahlpflichtmanager/assets/47521842/fe7a1b14-3e8f-4c84-a6b1-4b30dcd81149)
Formatierung der Wahlmoeglichkeiten

![Bildschirmfoto vom 2023-05-22 12-47-20](https://github.com/epoell/Wahlpflichtmanager/assets/47521842/8c9b62a9-ebd8-4c3a-b74a-9a9d4c54db42)
Formatierung des Zuteilungsergebnisses (analog der Wahlen)
