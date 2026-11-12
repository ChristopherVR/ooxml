---
title: Konto und Anmeldung
description: Was File > Account standardmäßig anzeigt und wie Sie einen Anmeldevorgang an den standardmäßig deaktivierten Haltepunkt anbinden.
---

# Konto und Anmeldung

Das Öffnen und Speichern von Microsoft-gehosteten Präsentationen ist eine optionale Host-Integration. Siehe
[OneDrive-Integration](/de/guide/onedrive) für den Speicheradapter und die Aufgaben der Authentifizierung sowie
[Testfixtures im Vergleich zur Abnahme mit einem echten Konto](/de/guide/onedrive-testing).

**File > Account** (Datei > Konto) enthält einen lokalen Profileditor, ein Speicher- und Datenschutzpanel mit aktuellen Werten aus dem lokalen Speicher, einen Info-Bereich und einen optionalen Anmeldebereich. Dieser bleibt ausgeblendet, bis ein Host ihn ausdrücklich anbindet.

## Was standardmäßig enthalten ist {#what-ships-by-default}

**Profil.** Ein Feld für den Anzeigenamen und eine Reihe von Farbfeldern für Avatare. Nichts davon wird irgendwohin gesendet; die Angaben dienen rein der Darstellung und werden im `localStorage` (Schlüssel `pptx-viewer-prefs`) zusammen mit dem Fallback für Design und Sprache aus [Theming](/de/guide/theming) gespeichert.

**Speicher und Datenschutz.** Eine Live-Anzahl der Präsentationen mit lokaler Autosave- bzw. Wiederherstellungskopie sowie deren Gesamtgröße. Die Werte stammen aus demselben IndexedDB-Speicher, den der Ablauf zur Wiederherstellung nach einem Neuladen verwendet. Die Schaltfläche „Clear local data“ (Lokale Daten löschen) entfernt nach einer Bestätigung jede lokale Kopie und jede gespeicherte Einstellung.

**Info.** Der Paketname und die Version des Bindings.

**Anmeldung.** Standardmäßig nicht vorhanden. Der Bereich wird nur angezeigt, wenn Sie `accountAuth` mit `enabled: true` übergeben (siehe unten).

Für keinen dieser Bereiche ist eine Einrichtung nötig. Er funktioniert in einem nicht konfigurierten Viewer genauso wie in einer vollständig eingebundenen Host-Anwendung.

## Einen echten Anmeldevorgang anbinden {#wiring-a-real-sign-in-flow}

Der Viewer legt sich nicht auf eine Authentifizierung fest: Er enthält keinen OAuth-Client, keine Sitzungscookies und keinen Backend-Aufruf. `accountAuth` ist ein schlanker Haltepunkt. Ist er aktiviert, zeigt die Konto-Seite eine Anmeldeaufforderung und ruft Ihre Anwendung zurück. Was „angemeldet“ bedeutet, entscheiden Sie; das Ergebnis geben Sie anschließend wieder an den Viewer weiter.

```ts
interface AccountAuthConfig {
	enabled: boolean;
	onSignIn: () => void;
	signedInUser?: {
		name: string;
		email?: string;
		avatarUrl?: string;
	};
}
```

- `enabled: false` (oder das vollständige Weglassen von `accountAuth`) ist der Standard: Der Anmeldebereich wird gar nicht gerendert.
- `onSignIn` wird ausgelöst, wenn der Benutzer auf die Anmeldeschaltfläche klickt. Starten Sie hier den eigentlichen Ablauf Ihrer Anwendung (Weiterleitung, OAuth-Popup, ein eigenes Modal usw.). Der Viewer wartet darauf nicht und verwaltet keinen Ladezustand.
- Sobald Ihre Anwendung weiß, wer angemeldet ist, übergeben Sie `signedInUser` über dieselbe `accountAuth`-Prop beim nächsten Render bzw. Update zurück. Die Konto-Seite zeigt dann „Signed in as {name}“ (Angemeldet als {name}) anstelle der Anmeldeschaltfläche.

```tsx
// React
const [user, setUser] = useState<{ name: string; email?: string } | undefined>();

<PowerPointViewer
	content={bytes}
	accountAuth={{
		enabled: true,
		onSignIn: () => startOAuthFlow().then(setUser),
		signedInUser: user,
	}}
/>;
```

```vue
<!-- Vue -->
<PowerPointViewer
	:content="bytes"
	:accountAuth="{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }"
/>
```

```html
<!-- Angular -->
<pptx-viewer
	[content]="bytes"
	[accountAuth]="{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }"
/>
```

```ts
// Vanilla
const viewer = createPptxViewer(host, {
	source: bytes,
	accountAuth: { enabled: true, onSignIn: startOAuthFlow, signedInUser: user },
});
```

```svelte
<!-- Svelte -->
<PowerPointViewer source={content} accountAuth={{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }} />
```

Die Form ist in allen fünf Bindings identisch.

## Nächste Schritte {#next-steps}

- [Theming](/de/guide/theming): derselbe `localStorage`-Fallback speist sowohl den Profileditor als auch die Auswahl für Erscheinungsbild und Sprache.
- [Lokalisierung (i18n)](/de/guide/localization): Jeder Text der Konto-Seite läuft über die Standard-Übersetzungsschlüssel `pptx.account.*`.
