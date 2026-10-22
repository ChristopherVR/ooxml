---
title: Zusammenarbeit
description: Eine Live-Co-Editing-Sitzung starten oder ihr beitreten, die Anwesenheitsanzeigen lesen, verstehen wie gleichzeitige Bearbeitungen zusammengeführt werden, und eine Bildschirmpräsentation übertragen.
---

# Zusammenarbeit

pptx-viewer unterstützt Echtzeit-Co-Editing: Mehrere Personen öffnen dieselbe Präsentation und sehen die Änderungen der anderen, während sie passieren.

::: info Optional und vom Entwickler konfiguriert
Zusammenarbeit ist eine optionale Funktion. Sehen Sie keine Schaltfläche **Teilen**, wurde die verwendete App ohne sie gebaut. Entwickler: siehe [/react/collaboration](/react/collaboration).
:::

## Eine Sitzung starten

1. Klicken Sie auf die Schaltfläche **Teilen** am rechten Ende der Registerkartenzeile des Menübands (oder **Datei > Freigeben > Mit Personen teilen**). Der Dialog **Präsentation teilen** öffnet sich auf seiner Registerkarte **Sitzung erstellen**.
2. Füllen Sie aus:
   - **Sitzungsname** - die Raumkennung (Buchstaben, Zahlen, Bindestriche, Unterstriche). Mitarbeitende brauchen diesen Namen.
   - **Ihr Anzeigename** - wie Sie für andere erscheinen.
   - **Zusammenarbeitsserver** - optional. Lassen Sie ihn leer für eine Peer-to-Peer-Sitzung, oder geben Sie die Adresse des Zusammenarbeitsservers Ihrer Organisation ein (eine `wss://...`-URL).
3. Klicken Sie auf **Teilen starten**.

![Dialog zum Teilen und zur Zusammenarbeit](/user-guide/collaboration-share-dialog.jpg)

Sobald die Sitzung läuft, zeigt der Dialog einen **Freigabelink** mit einer Schaltfläche **URL kopieren**. Senden Sie diesen Link an Ihre Mitarbeitenden. Der Dialog listet außerdem die **Verbundenen Nutzer** auf und auf welcher Folie jede Person ist. Die Schaltfläche Teilen in der Symbolleiste wird grün und zeigt die Teilnehmerzahl, solange das Teilen aktiv ist.

Um die Sitzung zu beenden, öffnen Sie den Dialog erneut und klicken Sie auf **Teilen beenden**.

## Einer Sitzung beitreten

Öffnen Sie entweder einen Freigabelink, den Ihnen jemand geschickt hat, oder treten Sie manuell bei:

1. Klicken Sie auf **Teilen** und wechseln Sie zur Registerkarte **Sitzung beitreten**.
2. Fügen Sie den Einladungslink oder die Raum-ID in **Einladungslink oder Raum-ID** ein.
3. Geben Sie Ihren Anzeigenamen ein und, falls die Sitzung einen verwendet, die Adresse des Zusammenarbeitsservers.
4. Klicken Sie auf **Sitzung beitreten**.

## Anwesenheitsanzeigen

Solange eine Sitzung aktiv ist, sehen Sie:

- **Avatarleiste** - farbige Kreise in der Symbolleiste für jeden Teilnehmer (bis zu fünf, danach ein Überlauf-Chip "+N"). Ihr eigener Avatar hat einen weißen Rand.
- **Remote-Cursor** - der Zeiger jedes Mitarbeiters bewegt sich mit einer farbigen Namensbeschriftung auf der Zeichenfläche. Cursor erscheinen nur für Personen, die dieselbe Folie ansehen wie Sie.
- **Remote-Auswahlen** - markiert jemand ein Element, erhält es einen farbigen Umriss mit seinem Namen, sodass Sie wissen, woran er arbeitet.
- **Folgeleiste** - ein Streifen oben mit der Aufschrift **Folgen Sie einem Mitarbeiter**. Klicken Sie auf eine Person, um ihr zu folgen: Ihre Ansicht wechselt dann die Folie, wann immer sie es tut. Klicken Sie auf **Stoppen**, um wieder selbst zu navigieren.
- **Verbindungsstatus** - ein Punkt in der Statusleiste: grün **Verbunden** (mit der Nutzerzahl), gelb **Verbinden...**, grau **Getrennt** oder rot **Verbindungsfehler** mit einer Schaltfläche **Erneut versuchen**.

## Wie sich gleichzeitiges Bearbeiten verhält

Bearbeitungen werden automatisch mit konfliktfreien Datenstrukturen zusammengeführt:

- Zwei Personen, die verschiedene Folien oder verschiedene Elemente auf derselben Folie ändern, überschreiben sich nie gegenseitig.
- Zwei Personen, die im selben Textfeld tippen, werden auf Zeichenebene zusammengeführt, wie in einem gemeinsamen Dokumenteditor.
- Bricht Ihre Verbindung ab, können Sie weiter bearbeiten; Ihre Änderungen werden synchronisiert, sobald die Verbindung zurückkehrt. Beobachten Sie den Statuspunkt: Solange er **Getrennt** zeigt, sehen andere Ihre neuesten Änderungen noch nicht.

::: tip Große Eingriffe abstimmen
Das Zusammenführen schützt Sie vor verlorenen Änderungen, nicht vor Überraschungen. Das Löschen einer Folie, an der jemand anderes arbeitet, löscht sie dennoch für alle; einigen Sie sich bei größeren Umstrukturierungen daher darauf, wem welcher Abschnitt gehört.
:::

## Peer-to-Peer oder Relay-Server

Das Feld **Zusammenarbeitsserver** entscheidet, wie sich Teilnehmer verbinden:

- **Leer (Peer-to-Peer)** - Browser-Tabs auf demselben Gerät finden sich immer, und Personen auf anderen Geräten verbinden sich über öffentliche WebRTC-Signalisierungsserver direkt mit Ihnen. Das ist der schnellste Weg, die Zusammenarbeit auszuprobieren; die gehostete Demo funktioniert so.
- **Serveradresse gesetzt** - alle übertragen über einen dedizierten Zusammenarbeitsserver. Verwenden Sie das, wenn Ihre Organisation einen betreibt: Es funktioniert auch durch strenge Firewalls und NAT, wo direkte Peer-Verbindungen scheitern, hält den Raum verfügbar, solange ein Teilnehmer verbunden ist, und kann von der Host-Anwendung mit Authentifizierung und serverseitiger Speicherung kombiniert werden.

::: warning wss:// auf sicheren Seiten verwenden
Wird die App über `https` ausgeliefert, muss die Serveradresse `wss://` (einen sicheren WebSocket) verwenden, sonst blockiert der Browser die Verbindung.
:::

## Eine Bildschirmpräsentation übertragen

Die Übertragung ist einseitige Zusammenarbeit: Sie präsentieren, die Zuschauer folgen.

1. Klicken Sie auf **Bildschirmpräsentation > Übertragen** (oder **Online präsentieren** im Menü Präsentieren). Der Dialog **Präsentation übertragen** öffnet sich.
2. Prüfen Sie den vorausgefüllten Namen der **Übertragungssitzung** und Ihren **Präsentatorname**n, und setzen oder leeren Sie den **Zusammenarbeitsserver** wie oben.
3. Klicken Sie auf **Übertragung starten**. Die App wechselt in den Präsentationsmodus, und der Dialog zeigt einen **Zuschauer-Link** mit **URL kopieren**.
4. Teilen Sie den Zuschauer-Link. Personen, die ihn öffnen, sehen Ihre Folien in Echtzeit weiterschalten; der Dialog zählt Ihre Zuschauer.
5. Klicken Sie auf **Übertragung beenden**, wenn Sie fertig sind.

## Kommentare funktionieren neben dem Live-Bearbeiten

Kommentare (siehe [Bearbeiten > Kommentare](/de/user/editing#kommentare)) werden mit der Präsentation gespeichert und sind eine gute Möglichkeit, Feedback für Mitarbeitende zu hinterlassen, die nicht gleichzeitig online sind.

## Weiter

- [Folien bearbeiten](/de/user/editing)
- [Tastaturkurzbefehle](/de/user/shortcuts)
