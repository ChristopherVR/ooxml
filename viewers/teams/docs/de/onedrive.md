# OneDrive-Links und Office-Dokumente

Für eine Schritt-für-Schritt-Einrichtung (Registrierung in Microsoft Entra, Anmeldung und ein vollständiges Beispiel)
beginnen Sie mit dem [Schnellstart für OneDrive](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

OpenTeams kann einen OneDrive-Link zu einer DOCX-, XLSX-, PPTX- oder VSDX-Datei enthalten. Die integrierte Office Suite
öffnet diese Datei im passenden Dokumenteditor und über die gemeinsame Verbindung `ooxml-core/embedding/onedrive`.
OpenTeams selbst ist ein Arbeitsbereich, kein fünfter Office-Dateieditor und nicht Microsoft Teams.

Dieselbe Unterscheidung gilt für `openteams-react-viewer`, `openteams-vue-viewer`, `openteams-angular-viewer`,
`openteams-svelte-viewer`, `openteams-solid-viewer`, `openteams-vanilla-viewer` und das Element `ooxml-ui/teams`:

| Gemeinsames Element             | Wo Änderungen gespeichert werden                                                                         |
| ------------------------------- | -------------------------------------------------------------------------------------------------------- |
| OneDrive-Link                   | Nur im ursprünglichen OneDrive-Element, wenn der berechtigte Host über seine OneDrive-Sitzung speichert. |
| Link zu lokaler Arbeitsdatei    | Im lokalen Office-Arbeitsbereich; es ist kein geräteübergreifender Cloud-Link.                           |
| Anhang auf dem OpenTeams-Server | Im Speicher dieses Servers; eine hochgeladene Kopie ist vom OneDrive-Original unabhängig.                |

Der Teilen-Ablauf für verbundene Dateien in der Suite fordert eine ausdrückliche Rolle für Anzeigen oder Bearbeiten
sowie einen Geltungsbereich für den Link an. Einen bestehenden Element-URL zu posten oder zu kopieren, gewährt einer
anderen Person selbst noch keinen Zugriff. Erzeugen Sie niemals einen anonymen Link als Ersatz, wenn ein Mandant einen
eingeschränkten Link ablehnt. `openteams-server` erwirbt keine Microsoft-Anmeldedaten, spiegelt keine Ordner aus dem
Laufwerk und führt keine Office-Dateiänderungen zusammen.

Für Authentifizierung, Paketverweise und berechtigungsbewusstes Teilen nutzen Sie den
[Leitfaden zur Office-OneDrive-Integration](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams).
Sein [Testleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) unterscheidet
testlokale Transporte von echten Kontoabnahmen. Eine Simulationsoberfläche oder einen Dienst dafür gibt es nicht.

Die Chat- und Präsenz-Kanäle von Yjs sowie die Räume für die gemeinsame Dokumentbearbeitung sind getrennt von der
gemeinsamen Bearbeitung in Microsoft Office. Wenn Sie denselben OneDrive-Link an zwei Personen schicken, werden ihre
Editoren nicht automatisch synchronisiert.
