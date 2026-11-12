---
title: OneDrive-Integration testen
description: Prüfen Sie die gemeinsame Office-OneDrive-Verbindung mit testlokalen Transporten und der Abnahme mit einem autorisierten Microsoft-Konto.
---

# OneDrive-Integration testen

Die OneDrive-Verbindung wird von Word, Excel, PowerPoint und Visio gemeinsam genutzt. Ihre automatisierten Transporte gehören ausschließlich zu den Testfixtures. Es gibt keine ausgelieferte Simulationsoberfläche, keinen öffentlichen Mock-Dienst und keine OneDrive-Simulationsdemo, die Sie konfigurieren oder bereitstellen müssten.

Befehle, Prüfungen je Format und die Matrix für die Abnahme mit echten Konten finden Sie im [Office-Testleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md). Der [Integrationsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) beschreibt die Einbindung von Paketen und Frameworks, die Authentifizierung, die Freigabe und die Wiederherstellung bei Konflikten.

Automatisierte Regressionstests können die Abfolge der Graph-Anfragen und die Speicherverdrahtung des Editors ohne Microsoft-Zugangsdaten prüfen. Sie decken Snapshots und Dateiidentität ab, Berechtigungseinschränkungen, Revisionsbedingungen, erhaltene lokale Änderungen, getrennte Cloud-Kopien und die Fehlerbehandlung. Sie können jedoch nicht belegen, dass die Microsoft-Anmeldung, die Zustimmung, der bedingte Zugriff, CORS auf der bereitgestellten Domain oder die Treue zu nativem PowerPoint funktionieren.

Microsoft Graph Explorer erlaubt manuelle API-Prüfungen mit der bestehenden App-Registrierung von Microsoft. Er prüft jedoch weder die eigene App-Registrierung Ihrer Anwendung noch deren Anmeldeablauf. Verwenden Sie eine Wegwerf-Präsentation und folgen Sie dem Testleitfaden. Geben Sie die Client-ID des Graph Explorer oder fremde Zugangsdaten nicht als Authentifizierungslösung Ihrer Anwendung aus.

Für das tatsächliche Bearbeiten über die Suite konfigurieren Sie den Microsoft-Token-Provider des Hosts, bevor dessen Einstieg lädt. Verwenden Sie danach **Open from OneDrive** (Aus OneDrive öffnen). Echte persönliche und geschäftliche Konten sowie die Richtlinien von SharePoint-Mandanten bleiben als Abnahmefälle erforderlich. Kein Test ohne Registrierung belegt eine Live-Zusammenarbeit mit Microsoft oder garantiert gleichzeitiges Bearbeiten ohne Konflikte.
