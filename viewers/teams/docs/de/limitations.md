# Einschränkungen

OpenTeams befindet sich in der Beta-Phase. Diese Seite listet auf, was es nicht kann, damit Sie
vor dem Einsatz entscheiden können.

## Nicht Microsoft Teams

OpenTeams ist nicht Microsoft Teams, steht in keiner Verbindung zu Microsoft und erhebt keinen Anspruch auf Funktionsgleichheit.
Es implementiert:

- Kanäle, Beiträge mit Antworten und Thread-Bereichen, Reaktionen, Bearbeiten und Löschen nur durch den Autor, Suche und Zähler für Ungelesenes;
- 1:1- und Gruppenchats, jeweils in einem eigenen Raum, dem nur die Mitglieder beitreten können;
- einen Aktivitäts-Feed und Benachrichtigungen in der App (Hinweis, Browser-Benachrichtigung bei geöffnetem Tab, Ton);
- einen Kalender mit Besprechungen, Wiederholungen, Erinnerungen und `.ics`-Export;
- Präsenz: wer online ist, wer tippt, und einen Status, den Sie selbst setzen (verfügbar, beschäftigt, nicht stören, gleich zurück, abwesend, offline) mit optionaler Rücksetzzeit, eine Statusnachricht und automatisch abwesend bei Inaktivität;
- Besprechungen in einem Kanal: Audio, Video, Bildschirmfreigabe und Handheben, mit einem Bildschirm vor dem Beitritt;
- Dateien, die in der Unterhaltung als Links geteilt werden, wobei Word-, Excel-, PowerPoint- und Visio-Dateien
  am Namen erkannt werden.

Es implementiert **nicht** Aufzeichnung, Hintergrundunschärfe, Live-Untertitel, Push- oder E-Mail-Benachrichtigungen
(Benachrichtigungen funktionieren nur bei geöffneter App), Einladungen außerhalb des Arbeitsbereichs oder Zusagen (RSVP).

## Inhaltsbereich

Word zeigt eine schreibgeschützte Vorschau; Excel bietet lokales Bearbeiten und das Speichern neuer Kanalkopien,
wenn Speicher konfiguriert ist; Visio nutzt den vorhandenen Viewer.
Es gibt kein gemeinsames Bearbeiten, kein Überschreiben und kein versioniertes Zurückspeichern. PowerPoint benötigt eine
vom Host konfigurierte eingebettete Viewer-Seite. Fehlende Integrationen werden im Bereich gemeldet.
Textlesevorgänge sind auf 2 MiB und native Office-Lesevorgänge auf 32 MiB begrenzt. Die Oberfläche gibt es auf Englisch, Deutsch, Spanisch, Französisch und in vereinfachtem Chinesisch (Eigenschaft `locale`); von Personen geschriebener Text wird nie übersetzt.

Die Dateiansichten unterstützen direktes Hochladen, das Anlegen leerer Excel-Dateien und die Metadatensuche, wenn
Speicher konfiguriert ist. Uploads zeigen einen Beschäftigt-Zustand und erlauben Wiederholen, ohne Byte-
Fortschritt oder Abbruch. Ein fehlgeschlagener Stapel kann Dateien hinterlassen, die vor dem Fehler hochgeladen wurden und an keiner Nachricht
hängen; der Referenzserver entfernt seine unvollständigen Uploads und lässt die hochladende Person oder einen
Moderator den Rest löschen (`DELETE /files/<scope>/<name>`), tut das aber nicht von selbst. Der Referenzserver
beschränkt Dateien privater Chats auf deren Mitglieder und zählt den Speicher pro Workspace; Kanaldateien kann jede
angemeldete Person des Servers lesen. Ordner, Indizierung des Dokumentinhalts und Dateiversionsverlauf sind
nicht implementiert.

Markdown unterstützt Überschriften, Aufzählungen, Zitate, umzäunten Code, einfache Hervorhebung, Web-Links, flache Aufgabenlisten, Tabellen auf oberster Ebene und dateirelative Web-Links. Es implementiert nicht das vollständige
CommonMark/GFM, verschachtelte Strukturen oder Referenzlinks. Rohes HTML wird als Text angezeigt.

Websites und HTML-Anhänge nutzen eine Sandbox ohne Same-Origin-, Popup- oder Top-Navigations-
Berechtigung. Manche Websites blockieren das Einbetten oder verlangen Berechtigungen, die diese
Sandbox nicht gewährt. Das externe Öffnen bleibt verfügbar. Definitionen von Datei- und Website-Tabs können in Kanälen
geteilt werden; die Plattform für eingebettete Apps, die Authentifizierung und die Berechtigungen
von Microsoft Teams sind nicht implementiert. Prüfungen auf den Tab-Ersteller laufen
im Client und benötigen eine Serverautorisierung, um durchsetzbar zu werden.

## Sicherheit und Datenschutz

- **Keine Ende-zu-Ende-Verschlüsselung.** Der Chat läuft als Yjs-Updates über Ihren Server, der sie lesen kann
  und speichert. Hochgeladene Dateien werden unverändert gespeichert. Anruf-Medien sind DTLS-SRTP zwischen den Teilnehmern (ein
  TURN-Relay kann sie nicht entschlüsseln).
- **Autorisierung ist Aufgabe des Servers.** Autor- und Rollenprüfungen laufen im Client; ein feindlicher
  Client kann alles schreiben, was der Server akzeptiert.
- **Authentifizierung ist standardmäßig Pflicht.** Der Referenzserver startet nur mit Konten, einem `authenticate`-Hook oder `authorizeSync`; `--insecure-dev` wählt einen offenen Entwicklungsmodus mit Warnung. Mit Konten signiert er Sitzungen für konfigurierte Benutzer (oder Ihre `authenticate`-Funktion) und übernimmt die Benutzer-ID aus der Sitzung; `authorizeSync` kann den Zugriff pro Raum entscheiden. Ein gemeinsames `TEAMS_TOKEN` allein ist ein Geheimnis im Browser, keine Authentifizierung pro Benutzer, und Benutzer-IDs sind Behauptungen. Es gibt kein SSO, kein Zurücksetzen von Passwörtern und keine Zugriffssteuerung für Teams und Kanäle. Siehe [was der Referenzserver nicht kann](/de/server#was-der-referenzserver-nicht-kann).
- **Präsenz ist beratend und nicht authentifiziert**, wie jede Yjs-Awareness: Ein Client kann jeden
  Namen angeben.
- Alles, was ein Teilnehmer sendet, wird beim Zurücklesen erneut validiert und als Text dargestellt, nie als Markup.

## Skalierung

- **Mesh-Anrufe** verbinden jeden Teilnehmer mit jedem anderen. Das braucht keinen Medienserver und
  funktioniert für eine Handvoll Personen (die Obergrenze ist standardmäßig 12), nicht für ein Webinar.
- **Der Verlauf wächst.** Der Referenzserver hält einen Snapshot pro Raum und verdichtet ihn nie.
- **Ein Serverprozess.** Der Referenzserver hält Räume im Arbeitsspeicher und skaliert nicht
  horizontal.

## Lokaler Modus und die Live-Demos

`mode: 'local'` nutzt `BroadcastChannel`, sodass sich nur Tabs desselben Browsers auf derselben Website
sehen. Nichts wird an einen Server gesendet, Dateien werden nur per Name geteilt (es gibt keinen Ort zum Hochladen)
und das gemeinsame Dokument wird im `localStorage` dieses Browsers gehalten. Die
[Live-Demos](/demos) laufen so, weil GitHub Pages keinen Server hosten kann.

## Anrufe über Netzwerke hinweg

Öffentliches STUN funktioniert in offenen Netzen. Hinter symmetrischem NAT oder strikten Firewalls benötigen Sie einen eigenen TURN-
Server; ohne ihn kommen manche Anrufe nicht zustande. Kameras und Mikrofone erfordern `localhost` oder
HTTPS.
