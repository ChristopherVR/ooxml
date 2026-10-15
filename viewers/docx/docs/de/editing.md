# Text bearbeiten

Jedes Framework bindet denselben Web-Component-Editor ein. Befehle, Tastenkürzel, Suche, Sprachmetadaten und Dokumentkonvertierung haben genau eine Implementierung.

## Suchen und Ersetzen

Verwenden Sie **Suchen und Ersetzen** oder **Strg+F** (**Befehl+F** unter macOS), solange der Editor den Fokus hat. Geben Sie eine wörtliche Suchphrase ein und springen Sie mit **Weitersuchen** oder **Vorheriges suchen** durch die Treffer. **Groß-/Kleinschreibung beachten** steuert die Unterscheidung von Groß- und Kleinbuchstaben. Mit Escape schließen Sie das Panel und der Fokus kehrt zum Dokument zurück.

Die Suche erstreckt sich über benachbarte formatierte Runs innerhalb eines Absatzes. Sie findet keine Treffer über Absätze, Tabellenzellen oder harte Zeilenumbrüche hinweg. Treffer, die ein Unicode-Graphem trennen würden, etwa einen Teil einer kombinierenden Zeichenfolge oder eines zusammengesetzten Emojis, werden vermieden. Die Suche ist wörtlich; reguläre Ausdrücke und sprachspezifischer linguistischer Abgleich werden nicht unterstützt.

**Ersetzen** ändert den aktuellen Treffer; **Alle ersetzen** ändert alle Treffer des Dokuments in einem einzigen rückgängig machbaren Vorgang. Der Ersetzungstext übernimmt die Formatierung des ersten Zeichens des Treffers. Text außerhalb des Treffers behält seine Formatierung. Im schreibgeschützten Modus funktioniert die Suche; das Ersetzen ist deaktiviert. Die bestehenden Prüfungen zur DOCX-Erhaltung gelten weiterhin, wenn Änderungen rund um nicht unterstützte Quellinhalte exportiert werden.

## Mehrsprachige Dokumente

Absatzrichtung und Textsprache sind getrennte Eigenschaften. Verwenden Sie **Absatzrichtung** für ein Layout von links nach rechts oder von rechts nach links und **Run-Richtung** für eine ausdrückliche Richtungsvorgabe innerhalb eines Absatzes. Die Sprachsteuerelemente speichern die Sprachkennungen von Word für lateinische, ostasiatische und komplexe Schriften. Sprachmetadaten übersetzen keinen Text und installieren keine Rechtschreibprüfung.

Der Editor erhält Unicode-Text, darunter Arabisch, Hebräisch, CJK, kombinierende Zeichen und Emojis. Browserschriften, Textformung (Shaping) und bidirektionales Layout bestimmen die Darstellung auf dem Bildschirm. Diese Steuerelemente stellen weder eine mit Word gleichwertige Typografie für komplexe Schriften noch Seitenumbrüche oder das IME-Verhalten des Betriebssystems her.

Die Wortzählung verwendet die Unicode-Wortsegmentierung des Browsers, wo unterstützt auch die wörterbuchbasierte Segmentierung für Schriften ohne Leerzeichen. Die Zählungen können sich zwischen ICU-Versionen und von Microsoft Word unterscheiden. Ältere Browser verwenden eine weniger genaue Ersatzlösung.

## Integration der Koautorenschaft

Der [Leitfaden zur Zusammenarbeit](/de/collaboration) beschreibt das versionierte Bearbeitungsprotokoll und wie eine Anwendung den gemeinsamen Editor mit ihrer Autorität verbindet. Dies ist eine Integrations-API, kein gehosteter Dienst für die Zusammenarbeit. Zugriffskontrolle, Netzwerk, Speicherung und Wiederverbindungsrichtlinie liegen bei der Anwendung.

## Oberflächensprache

Setzen Sie die Eigenschaft `locale` des gemeinsamen Editors (oder die Prop `locale` des Bindings) auf `en`, `fr`, `de`, `es` oder `zh-CN`. Regionale und Schriftvarianten werden auf die nächste unterstützte Sprache abgebildet (`fr-CA`, `de-DE`, `es-MX`, `zh-Hans`, `zh-SG`); nicht unterstützte Locales, einschließlich Traditionell Chinesisch (`zh-TW`, `zh-Hant`), fallen auf Englisch zurück. `EDITOR_LOCALES` listet die Codes auf und `normalizeEditorLocale()` bildet jedes Tag auf einen davon ab. Jede Locale definiert jeden Zeichenfolgenschlüssel (zur Kompilierzeit und durch Tests geprüft). Die deutschen, spanischen und vereinfacht chinesischen Zeichenfolgen wurden nicht von Muttersprachlern geprüft. Ein Wechsel der Oberflächensprache erhält das aktuelle Dokument, die Auswahl, den Verlauf zum Rückgängigmachen und die Sprachmetadaten des Dokuments.

## Importierte Absatzformatvorlagen

Der Formatvorlagenkatalog im Register „Start“ listet die Absatzformatvorlagen des geladenen DOCX auf. Der Editor löst unterstützte Ausrichtung, Richtung, Abstände und Einzüge über die Dokumentstandards und die Vererbungskette der gewählten Formatvorlage auf. Direkte Absatzformatierung hat weiterhin Vorrang. Die Wahl einer Formatvorlage ändert deren Referenz, ohne vererbte Eigenschaften in jeden Absatz zu übernehmen.

Der ursprüngliche Formatvorlagenteil bleibt beim Speichern des geladenen Dokuments erhalten. Das Erstellen oder Bearbeiten von Formatvorlagendefinitionen, Zeichenformatvorlagen, vererbte Run-Formatierung und die Auflösung von Designs werden noch nicht unterstützt. Der Export eines neuen Pakets mit Formatvorlagenkatalog wird abgelehnt; verwenden Sie „Speichern“ im geladenen Dokument, um dessen ursprüngliche Formatvorlagendefinitionen zu erhalten.
