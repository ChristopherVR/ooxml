# Einschränkungen

OpenTeams ist eine frühe Implementierung. Diese Seite listet auf, was es nicht kann, damit Sie
vor dem Einsatz entscheiden können.

## Nicht Microsoft Teams

OpenTeams ist nicht Microsoft Teams, steht in keiner Verbindung zu Microsoft und erhebt keinen Anspruch auf Funktionsgleichheit.
Es implementiert:

- Kanäle, Beiträge mit Antworten, Reaktionen, Bearbeiten und Löschen nur durch den Autor, Suche und Zähler für Ungelesenes;
- Präsenz: wer online ist, wer tippt, und eine Verfügbarkeit, die Sie selbst setzen (verfügbar, beschäftigt, abwesend);
- Besprechungen in einem Kanal: Audio, Video, Bildschirmfreigabe und Handheben, mit einem Bildschirm vor dem Beitritt;
- Dateien, die in der Unterhaltung als Links geteilt werden, wobei Word-, Excel-, PowerPoint- und Visio-Dateien
  am Namen erkannt werden.

Es implementiert **nicht** 1:1- oder Gruppenchats außerhalb von Kanälen, Threads als eigene Bereiche,
Benachrichtigungen, einen Kalender, Aufzeichnung, Hintergrundunschärfe oder Live-Untertitel.

## Inhaltsbereich

Word zeigt eine schreibgeschützte Vorschau; Excel bietet lokales Bearbeiten und das Speichern neuer Kanalkopien,
wenn Speicher konfiguriert ist; Visio nutzt den vorhandenen Viewer.
Es gibt kein gemeinsames Bearbeiten, kein Überschreiben und kein versioniertes Zurückspeichern. PowerPoint benötigt eine
vom Host konfigurierte eingebettete Viewer-Seite. Fehlende Integrationen werden im Bereich gemeldet.
Textlesevorgänge sind auf 2 MiB und native Office-Lesevorgänge auf 32 MiB begrenzt.

Die Dateiansichten unterstützen direktes Hochladen, das Anlegen leerer Excel-Dateien und die Metadatensuche, wenn
Speicher konfiguriert ist. Uploads zeigen einen Beschäftigt-Zustand und erlauben Wiederholen, ohne Byte-
Fortschritt oder Abbruch. Fehlgeschlagene Stapel können nicht referenzierte Speicherobjekte hinterlassen;
der Referenzserver hat keinen Bereinigungsvertrag für diese Objekte. Ordner, Indizierung des Dokument-
inhalts, Dateiversionsverlauf und authentifizierte Dateiberechtigungen sind
nicht implementiert.

Markdown unterstützt Überschriften, Aufzählungen, Zitate, umzäunten Code, einfache Hervorhebungen und Web-
Links. Es implementiert noch nicht das vollständige CommonMark/GFM, Tabellen, Aufgabenlisten, verschachtelte
Strukturen oder relative Links. Rohes HTML wird als Text angezeigt.

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
- **Der Referenzserver hat keine Benutzerkonten.** Sein Token ist ein gemeinsames Geheimnis, im
  Browser gespeichert, keine Authentifizierung pro Benutzer. Siehe [was der Referenzserver nicht kann](/de/server#was-der-referenzserver-nicht-kann).
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
