---
title: Stapelweise Elementaktualisierungen
description: Aktualisieren Sie Elemente über mehrere Folien hinweg als einen rückgängig machbaren Vorgang über die öffentliche Viewer-API.
---

# Stapelweise Elementaktualisierungen

Verwenden Sie `updateElements`, wenn eine Aktion des Hosts Elemente auf mehreren Folien aktualisieren muss. Ein Beispiel: Eine Schaltfläche „Adjust report titles“ (Berichtstitel anpassen) kann Titel auf zwei Seiten neu positionieren, ohne dass der Benutzer die Seite verlassen muss, die er gerade bearbeitet.

Die Methode steht am öffentlichen Viewer-Handle in React, Vue, Angular, Svelte und Vanilla JS zur Verfügung. `ElementUpdate` und `ElementUpdateOptions` werden aus dem Komponentenpaket jedes Bindings exportiert.

```ts
updateElements(
	updates: readonly ElementUpdate[],
	options?: ElementUpdateOptions,
): Promise<void>;

interface ElementUpdate {
	slideId: string;
	elementId: string;
	patch: Partial<PptxElement>;
}

interface ElementUpdateOptions {
	label?: string;
}
```

Hier ist `viewer` das aktuelle öffentliche Komponenten-Handle, und beide Folien enthalten bereits ein Titelelement:

```ts
const [first, second] = viewer.getSlides();

await viewer.updateElements(
	[
		{
			slideId: first.id,
			elementId: first.elements[0].id,
			patch: { x: 84 },
		},
		{
			slideId: second.id,
			elementId: second.elements[0].id,
			patch: { x: 90 },
		},
	],
	{ label: 'Adjust report titles' },
);

viewer.undo(); // Restore both titles in one step.
viewer.redo(); // Reapply both updates in one step.
```

## Commit- und Verlaufsverhalten {#commit-and-history-behavior}

- Jeder wirksame Stapel erzeugt genau einen Rückgängig-Schritt, unabhängig von vorherigen und nachfolgenden Bearbeitungen. Das gilt auch für einen weiteren Stapel, der im selben JavaScript-Durchlauf abgeschickt wird.
- Das Anwenden eines Stapels behält die aktive Folie und die Elementauswahl bei. Es springt nicht zu den Zielfolien. Rückgängig und Wiederholen verwenden das vorhandene Verhalten des Viewers beim Wiederherstellen von Snapshots.
- Warten Sie auf den Abschluss, bevor Sie das Ergebnis lesen oder einen abhängigen Vorgang starten. In React holen Sie `ref.current` erneut, wenn Sie über das öffentliche Handle lesen.
- Ausstehender Text aus einer Texteingabe wird vor einem wirksamen Stapel separat übernommen. Senden Sie Stapel zwischen abgeschlossenen Zeigerinteraktionen, nicht während eines Ziehens oder Größenänderns. Alle fünf Bindings lehnen einen Stapel ab, solange eine Zeigerinteraktion aktiv ist. Ein abgelehnter Stapel lässt ausstehenden Text, den Dokumentzustand und den Rückgängig-Verlauf unverändert.
- Leere Stapel und Stapel ohne Nettoänderung lassen den Verlauf und den Änderungsstatus unverändert. Abgelehnte Stapel bewahren das Dokument sowie die Rückgängig- und Wiederholen-Stapel.
- Die vorhandene Methode `updateElement` behält ihre Semantik bei: Sie wirkt auf die aktuelle Folie und auf einzelne Bearbeitungen.

## Ziele, Patches und Fehler {#targets-patches-and-errors}

Stapel setzen ein geladenes, bearbeitbares Dokument im normalen Folien-Bearbeitungsmodus voraus. Kontexte zum reinen Lesen, zur Vorschau, zur Präsentation sowie zur Bearbeitung von Mastern und Vorlagen lehnen das zurückgegebene Promise ab. Ziel-IDs müssen eindeutig auflösbar sein: `slideId` identifiziert eine Folie, und `elementId` identifiziert ein Element der obersten Ebene auf dieser Folie. Verschachtelte Gruppenkinder sowie Master- und Layoutelemente liegen außerhalb des Geltungsbereichs dieser Methode.

Jedes Ziel und jeder Patch wird geprüft, bevor eine Aktualisierung übernommen wird. Fehlt ein Ziel oder ist es mehrdeutig, wird der gesamte Stapel abgelehnt. Patches müssen Objekte sein und dürfen die `id` oder den `type` eines Elements nicht ersetzen. Felder für Position, Größe und Drehung müssen endliche Zahlen sein. Breite und Höhe dürfen nicht negativ sein. Andere Felder behalten den bestehenden Vertrag `Partial<PptxElement>`. Dies ist kein vollständiger Schema-Validator für Elemente.

Patches verwenden wie `updateElement` eine flache Ersetzung: Geben Sie bei einer verschachtelten Eigenschaft immer einen vollständigen Wert an. Mehrere Einträge für dasselbe Element werden in der Eingabereihenfolge angewendet. Eingaben werden defensiv kopiert, daher verändert eine Änderung am Patch-Objekt nach dem Aufruf der Methode den abgeschickten Stapel nicht.
