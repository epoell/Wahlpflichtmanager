# Wahlpflichtmanager

## Was
VBA Makro für Excel, dass so fair wie möglich Fächer an Schüler:innen verteilt.

VBA Macro for Excel, that tries to distribute chosen subjects as fair as possible among students/pupils.

Ausgehend von einer Liste mit Schüler:innen und deren Fach-Wünschen und einer Liste mit möglichen Fächern und deren maximaler Kursgröße, wird eine möglichst Optimale Zuteilung von Kursen zu Schüler:innen erzeugt. Das Makro soll z.B. Stufenkoordinator:innen helfen bei der Verteilung von Wahlpflicht-Fächern/Wahlpflichtunterricht an der Schule, bei der Zuordnung zu Projektwochen, AGs und Kursen. Generell kann hiermit jede Zuteilung aus einer Mehrfachwahl erzeugt werden. Die Anzahl der Fächer, Beteiligten und Wahlmöglichkeiten ist dabei beliebig. Das Makro arbeitet einfach auf den verfügbaren Informationen.

Urpsrünglich wurde dieses Programm für einen Freund und ganz tollen Lehrer geschrieben, damit er mehr Zeit hat unser Rugby-Team zu trainieren. Wenn das Makro gefällt, freue ich mich sehr über einen Stern.


## Vorgehen
Aus den beidenen gegebenen Tabellenblättern wird (in einem neuen Tabellenblatt) eine Zuteilung wie folgt errechnet, um möglichst vielen einen möglichst hohen Wunsch zu ermöglichen:
1. Alle Erstwünsche werden von oben nach unten vergeben, bis alle Kursplätze voll sind
2. Alle Zweitwünsche werden von oben nach unten vergeben, bis alle Plätze voll sind
3. Wenn jemand noch keinen Platz hat wird versucht zu tauschen, mit jemand der den Erstwunsch bekommen hat. (Bzw.. dann für den Zweit-, Dritt- und Viertwunsch).
4. Sonst wird der höchstmögliche noch freie Wunsch zugeteilt.
5. Die Verteilung geht nicht unbedingt auf. Schüler:innen deren Wünsche nicht erfüllt werden konnten werden lila markiert, zur manuellen Nachbearbeitung.

In Testläufen konnten aber 390 aus 400 Schüler:innen zugeteilt werden. Die Laufzeit betrug wenige Sekunden.

Es kann ausgewählt werden (durch einen Dialog beim Start des Macros), ob die Liste der Schüler:innen nach Priorität sortiert ist, um die Schüler:innen weiter oben in der Liste bei der Wahl zu bevorteilen. Das kann z.B. der chronologischen Reihenfolge entsprechen, in der Schüler:innen die Rückmeldungen einreichen. Was als Anreiz gegenüber den Schüler:innen genutzt werden kann die Wünsche zeitnah einzureichen.
Andernfalls wird die Liste der Schüler:innen zu beginn zufällig sortiert, die Fächer zugeteilt und dann wieder nach Klasse und Nachname sortiert.


## Vorgesehene Formatierung
In einer Excelarbeitsmappe gibt es zwei Tabellenblätter:
"Wahlmöglichkeiten" listet die Namen der möglichen Fächer auf und die Kursgröße je Fach.
"Wahlen" listet alle Schüler:innen auf mit Namen, Klasse und jeweils deren Erst- bis Fünftwunsch der möglichen Fächer.

Die Zuteilung wird in einem neuen vollständig generierten Tabellenblatt erzeugt, sodass die ursprünglichen Liste unverändert bleiben. Man kann mehrere Zuteilungen nacheinander in der selben Mappe generieren lassen und die generierten Tabellenblätter können nach Belieben gelöscht werden. Konflikte in der Tabellenbenennung werden im Makro umgangen.


![Bildschirmfoto vom 2023-05-22 12-43-42](https://github.com/epoell/Wahlpflichtmanager/assets/47521842/9f4a94b3-40cc-4fa2-85a9-f28bea24e1e6)

Formatierung der Wahlmöglichkeiten

![Bildschirmfoto vom 2023-05-22 12-47-20](https://github.com/epoell/Wahlpflichtmanager/assets/47521842/8c9b62a9-ebd8-4c3a-b74a-9a9d4c54db42)

Formatierung des Zuleitungsergebnisses (analog der Wahlen). (Die Daten sind natürlich randomisiert)
