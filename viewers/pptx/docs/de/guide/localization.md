---
title: Lokalisierung (i18n)
description: Wie die UI-Strings von pptx-viewer mit React, Vue 3, Angular, Vanilla und Svelte funktionieren, wie man eine Sprache in seiner App hinzufügt und wie man eine Übersetzung beiträgt.
---

# Lokalisierung (i18n)

**Der Viewer sucht jedes UI-Label über einen `pptx.*`-Übersetzungsschlüssel; Ihre App stellt das Wörterbuch bereit.** Jedes Binding liefert Englisch mit. Vollständige Referenzwörterbücher für Französisch, Spanisch, Deutsch und vereinfachtes Chinesisch werden außerdem über die Unterpfade `i18n/<locale>` jedes Viewer-Pakets ausgeliefert. Ihre Quelle wird einmal im privaten Workspace `pptx-viewer-locales` gepflegt, der in die Viewer-Pakete eingebunden wird und nicht separat installiert werden muss. React, Vue und Angular delegieren an die i18n-Bibliothek ihres Host-Frameworks; Vanilla und Svelte liefern einen kleinen eingebauten Übersetzer:

| Binding | Übersetzungsaufruf, den der Viewer macht                   | Bibliothek, die Sie bereitstellen                                                 |
| ------- | ---------------------------------------------------------- | --------------------------------------------------------------------------------- |
| React   | `t(key, opts)` von `react-i18next`                         | [i18next](https://www.i18next.com/) / [react-i18next](https://react.i18next.com/) |
| Vue 3   | `useI18n().t(key, opts)` von `vue-i18n`                    | [vue-i18n](https://vue-i18n.intlify.dev/)                                         |
| Angular | `translate()`-Signal / `TranslatePipe` von `ngx-translate` | [@ngx-translate/core](https://github.com/ngx-translate/core)                      |
| Vanilla | Das eigene eingebaute `t(key, params)` des Pakets          | keine: ein `messages`-Wörterbuch direkt übergeben, keine i18n-Bibliothek nötig    |
| Svelte  | Das eigene eingebaute `t(key, params)` des Pakets          | keine: `registerTranslations(locale, dict)` aufrufen, keine i18n-Bibliothek nötig |

So bleibt der Viewer für jedes Ökosystem idiomatisch (Sie schließen React/Vue/Angular an das i18n-Setup an, das Ihre App bereits hat; Vanilla/Svelte brauchen nichts Zusätzliches), statt Bindings, die noch keine haben, eine Übersetzungs-Laufzeitabhängigkeit aufzuzwingen.

## Das englische Wörterbuch

Jedes Paket re-exportiert ein fertiges englisches Ressourcenpaket, damit Sie keines von Grund auf schreiben müssen:

```ts
// React and Vue: subpath export
import { translationsEn, keyToLabel } from 'pptx-react-viewer/i18n'; // or 'pptx-vue-viewer/i18n'
```

```ts
// Angular: exported from the package root, not a subpath
import { translationsEn, keyToLabel } from 'pptx-angular-viewer';
```

```ts
// Vanilla: exported from the package root, like Angular
import { translationsEn, keyToLabel } from 'pptx-vanilla-viewer';
```

```ts
// Svelte: subpath export, like React/Vue
import { translationsEn, keyToLabel } from 'pptx-svelte-viewer/i18n';
```

`translationsEn` ist ein flaches `Record<string, string>` aller `pptx.*`-Schlüssel, z. B. `'pptx.statusBar.allSaved': 'All saved'`. Werte mit dynamischem Inhalt verwenden `{{token}}`-Platzhalter, z. B. `'pptx.statusBar.slideOf': 'Slide {{current}} of {{total}}'` - Ihre i18n-Bibliothek ersetzt diese aus den an `t()`/`translate()` übergebenen `opts`.

`keyToLabel(key)` leitet eine lesbare Bezeichnung aus dem letzten Segment eines Schlüssels ab, wenn kein Wörterbucheintrag passt (`"pptx.slideSorter.zoomIn"` → `"Zoom In"`). Binden Sie es als Handler für fehlende Schlüssel Ihrer Bibliothek ein, damit jeder noch nicht übersetzte Schlüssel etwas Vernünftiges anzeigt statt des rohen Schlüsselstrings.

## Einbinden

### React

```ts
// i18n.ts
import { createInstance } from 'i18next';
import { translationsEn, keyToLabel } from 'pptx-react-viewer/i18n';
import { initReactI18next } from 'react-i18next';

const i18nInstance = createInstance();
i18nInstance.use(initReactI18next).init({
	resources: { en: { translation: translationsEn } },
	lng: 'en',
	fallbackLng: 'en',
	interpolation: { escapeValue: false }, // React already escapes
	parseMissingKeyHandler: (key: string) => keyToLabel(key),
	missingKeyHandler: false,
});

export default i18nInstance;
```

```tsx
import { I18nextProvider } from 'react-i18next';

// app entry
import i18nInstance from './i18n';

<I18nextProvider i18n={i18nInstance}>
	<App />
</I18nextProvider>;
```

Sprachen wechseln Sie mit `i18nInstance.changeLanguage('fr')`.

### Vue 3

```ts
// i18n.ts
import { translationsEn, keyToLabel } from 'pptx-vue-viewer/i18n';
import { createI18n } from 'vue-i18n';

const i18n = createI18n({
	legacy: false,
	locale: 'en',
	fallbackLocale: 'en',
	messages: { en: translationsEn },
	missing: (_locale, key) => keyToLabel(key),
	missingWarn: false,
	fallbackWarn: false,
});

export default i18n;
```

```ts
// main.ts
import { createApp } from 'vue';

import App from './App.vue';
import i18n from './i18n';

createApp(App).use(i18n).mount('#app');
```

Sprachen wechseln Sie, indem Sie `i18n.global.locale.value = 'fr'` setzen.

### Angular

```ts
// i18n.ts
import { Injectable } from '@angular/core';
import type { MissingTranslationHandlerParams } from '@ngx-translate/core';
import { MissingTranslationHandler, provideTranslateService } from '@ngx-translate/core';
import { keyToLabel } from 'pptx-angular-viewer';

@Injectable()
class LabelFallbackMissingTranslationHandler implements MissingTranslationHandler {
	handle(params: MissingTranslationHandlerParams): string {
		return keyToLabel(params.key);
	}
}

export const i18nProviders = provideTranslateService({
	lang: 'en',
	fallbackLang: 'en',
	missingTranslationHandler: {
		provide: MissingTranslationHandler,
		useClass: LabelFallbackMissingTranslationHandler,
	},
});
```

```ts
// main.ts
bootstrapApplication(AppComponent, { providers: [i18nProviders] });
```

Registrieren Sie das englische Wörterbuch einmal (typischerweise in Ihrer Wurzelkomponente) und wechseln Sie Sprachen über `TranslateService`:

```ts
import { inject } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { translationsEn } from 'pptx-angular-viewer';

const translate = inject(TranslateService);
translate.setTranslation('en', translationsEn);
// later, to switch:
translate.use('fr');
```

::: tip Übergeben Sie ein einfaches `Provider`-Objekt, keine bloße Klasse
Wenn Sie einen eigenen `MissingTranslationHandler` schreiben, übergeben Sie ihn als `{ provide: MissingTranslationHandler, useClass: YourHandler }` statt als bloße Klassenreferenz. Unter manchen Build-Toolchains (Vite/Rolldown mit Standard-Dekoratoren) kann die Heuristik `isClass()` von ngx-translate in Produktions-Builds danebenliegen und den Handler ohne `new` aufrufen, was beim Bootstrap einen Fehler auslöst.
:::

### Vanilla

Das Vanilla-Binding hat kein Framework, an das eine i18n-Bibliothek angeschlossen werden könnte, und liefert deshalb einen winzigen eingebauten Übersetzer: Übergeben Sie ein nach Locale geordnetes `messages`-Wörterbuch direkt an `createPptxViewer`, ohne separaten Einbindungsschritt.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

import { translationsFr } from './translations/fr';

const viewer = createPptxViewer(document.querySelector('#host')!, {
	source: '/deck.pptx',
	locale: 'fr',
	messages: { fr: translationsFr },
});
```

`messages` ordnet einer Locale ein flaches `Record<string, string>` mit `pptx.*`-Überschreibungen zu; Sie müssen nicht jeden Schlüssel angeben oder `translationsEn` selbst einmischen, da Englisch für alles Fehlende immer der Rückfall ist. Sprachen wechseln Sie zur Laufzeit mit `viewer.setLocale('fr')` (das baut die statischen Beschriftungen der Oberfläche unter der neuen Locale neu auf und rendert neu).

### Svelte

Auch Svelte hat keine ausgezeichnete i18n-Laufzeit, daher stellt das Binding `registerTranslations` über seinen Unterpfad `/i18n` bereit: Registrieren Sie das Wörterbuch jeder Locale einmal beim Start und übergeben Sie dann die aktive `locale` als Prop.

```ts
// i18n.ts
import { registerTranslations } from 'pptx-svelte-viewer/i18n';

import { translationsFr } from './translations/fr';

registerTranslations('fr', translationsFr);
```

```svelte
<!-- App.svelte -->
<script lang="ts">
	import { PowerPointViewer } from 'pptx-svelte-viewer';
	import './i18n';

	let locale = $state('fr');
	let source = $state<Uint8Array>();
	// ...fetch/read the .pptx bytes into `source` however your app loads files
</script>

<PowerPointViewer {source} {locale} />
```

Registrierungen werden über das bereits für diese Locale Registrierte gemischt, sodass Sie `registerTranslations` mehrfach aufrufen können (z. B. einmal pro Funktionsbereich), ohne frühere Überschreibungen zu verdrängen. Wie bei Vanilla ist Englisch für noch nicht übersetzte Schlüssel immer der Rückfall.

## Datei > Optionen > Sprache

Der Einstellungsdialog jedes Bindings hat außerdem ein Register **Sprache**: eine Liste von Locales, aus denen der Benutzer zur Laufzeit wählen kann, direkt neben dem Register [Darstellung](/guide/theming). Es bündelt oder lädt selbst nie Übersetzungsinhalte - es bietet nur eine Oberfläche zum Wechseln zwischen Locales, die _Sie bereits_ in dem oben eingerichteten i18n-Setup registriert haben. Konkret: Wenn Sie `availableLocales` nicht selbst angeben, untersucht jedes Binding, was bereits geladen ist:

| Binding | Wie es verfügbare Locales ermittelt                                                   |
| ------- | ------------------------------------------------------------------------------------- |
| React   | `i18n.options.resources` / `i18n.languages` der `react-i18next`-Instanz               |
| Vue 3   | `useI18n().availableLocales` (eine eingebaute Eigenschaft des `vue-i18n`-Composables) |
| Angular | `TranslateService.getLangs()`                                                         |
| Vanilla | das an `createPptxViewer` übergebene `messages`-Wörterbuch plus `'en'`                |
| Svelte  | jede an `registerTranslations` übergebene Locale                                      |

Haben Sie also nur `en` und `fr` registriert, bietet das Register Sprache genau diese zwei an - nie eine Locale ohne Wörterbuch dahinter. Codes werden über den gemeinsamen `LOCALE_CATALOG` beschriftet (Anzeigenamen für Englisch/Französisch/Spanisch/Deutsch/vereinfachtes Chinesisch), wenn sie erkannt werden, andernfalls als roher Code angezeigt.

Die Rangfolge entspricht dem Register Darstellung: Eine gewählte Locale wird sofort angewendet, indem direkt Ihre i18n-Instanz aufgerufen wird (`i18n.changeLanguage`, `locale.value =`, `TranslateService.use` usw.), und in `localStorage` (`pptx-viewer-prefs`) gespeichert - es sei denn, Sie übergeben `onLocaleChange`; dann berührt der Viewer Ihre i18n-Instanz nie selbst und ruft nur diesen Callback auf, sodass Speicherung und Anwendung ganz Ihnen überlassen bleiben. Die vollständige Referenz der Props `defaultLocale`/`availableLocales`/`onLocaleChange` (gleiche Form wie die Theme-Props, dort einmal statt zweimal dokumentiert) steht unter [Theming](/guide/theming).

## Vereinfachtes Chinesisch

Die Demos bieten **简体中文** (`zh-CN`) unter **Datei > Optionen > Sprache** an. Das
Referenzwörterbuch deckt dieselben Schlüssel und `{{placeholders}}` ab wie Englisch.
Es übersetzt die Editor-Oberfläche, nicht den Text in Ihren Folien.

Importieren Sie das chinesische Wörterbuch direkt aus dem Viewer-Paket, das Sie bereits nutzen:

```ts
import { translationsZhCN } from 'pptx-react-viewer/i18n/zh-CN';
// Other bindings expose the same subpath:
// 'pptx-vue-viewer/i18n/zh-CN'
// 'pptx-angular-viewer/i18n/zh-CN'
// 'pptx-svelte-viewer/i18n/zh-CN'
// 'pptx-vanilla-viewer/i18n/zh-CN'
```

Jeder Unterpfad enthält ein eigenständiges Wörterbuch mit ESM, CommonJS und
TypeScript-Deklarationen. Der Import lädt weder den Viewer noch andere
Sprachen. Die vorhandenen englischen Exporte bleiben an ihren üblichen Einstiegspunkten
verfügbar. Registrieren Sie das Wörterbuch bei Ihrem Binding:

| Binding | Importiertes Wörterbuch registrieren                                       | Sprache wechseln                                   |
| ------- | -------------------------------------------------------------------------- | -------------------------------------------------- |
| React   | `i18n.addResourceBundle('zh-CN', 'translation', translationsZhCN)`         | `i18n.changeLanguage('zh-CN')`                     |
| Vue     | `i18n.global.setLocaleMessage('zh-CN', toVueI18nSyntax(translationsZhCN))` | `i18n.global.locale.value = 'zh-CN'`               |
| Angular | `translate.setTranslation('zh-CN', translationsZhCN)`                      | `translate.use('zh-CN')`                           |
| Vanilla | `messages: { 'zh-CN': translationsZhCN }` in den Viewer-Optionen           | `viewer.setLocale('zh-CN')`                        |
| Svelte  | `registerTranslations('zh-CN', translationsZhCN)`                          | Die Prop `locale` des Viewers auf `'zh-CN'` setzen |

Importieren Sie für Vue `toVueI18nSyntax` aus `pptx-vue-viewer/i18n`; es wandelt die
`{{name}}`-Platzhalter des Referenzwörterbuchs in die Vue-Syntax `{name}` um.
Die Sprachauswahl erkennt `zh-CN` und zeigt **简体中文** an, sobald die
Anwendung das Wörterbuch registriert.

## Weitere Referenzsprachen

Alle fünf Viewer-Pakete bieten dieselben optionalen Locale-Unterpfade:

| Locale                   | Unterpfad    | Export             |
| ------------------------ | ------------ | ------------------ |
| Französisch              | `i18n/fr`    | `translationsFr`   |
| Spanisch                 | `i18n/es`    | `translationsEs`   |
| Deutsch                  | `i18n/de`    | `translationsDe`   |
| Vereinfachtes Chinesisch | `i18n/zh-CN` | `translationsZhCN` |

Zum Beispiel `import { translationsFr } from 'pptx-vue-viewer/i18n/fr'`.
Vue-Anwendungen sollten diese Wörterbücher vor der Registrierung durch `toVueI18nSyntax`
leiten, genau wie im chinesischen Beispiel oben. Die Wörterbücher
behalten die kanonischen `{{name}}`-Platzhalter, die die anderen Bindings verwenden.

Diese Übersetzungen werden mit der Viewer-Version ausgeliefert, sodass ein Update der Komponente
auch die Referenzwörterbücher aktualisiert. Sie sind optional und werden weder
aus dem Paketstamm noch aus dem vorhandenen englischen i18n-Einstiegspunkt re-exportiert.

## Eine Sprache in Ihrer App hinzufügen

Den Viewer in eine weitere Sprache zu übersetzen, erfordert keinerlei Änderung an diesem Repository - es ist genau dieselbe Art Arbeit wie das Übersetzen jedes anderen Teils Ihrer App. Erstellen Sie ein Wörterbuch mit denselben Schlüsseln wie `translationsEn` und registrieren Sie es als zweite Ressource/Locale/Übersetzung in der Bibliothek Ihres Frameworks (`resources.fr` bei i18next, `messages.fr` bei vue-i18n, `translate.setTranslation('fr', ...)` bei ngx-translate). Sie brauchen nicht schon am ersten Tag jeden Schlüssel - alles, was Sie noch nicht übersetzt haben, fällt automatisch über `keyToLabel` zurück, sodass eine teilweise Abdeckung sanft degradiert, statt leere Beschriftungen oder rohe Schlüssel zu zeigen.

Um fehlende oder falsch geschriebene Schlüssel schon zur Kompilierzeit zu finden, statt sich auf den Laufzeit-Rückfall zu verlassen, typisieren Sie Ihr Wörterbuch gegen die exportierte Union `TranslationKey` statt gegen ein bloßes `Record<string, string>`:

```ts
import type { TranslationKey } from 'pptx-react-viewer/i18n'; // 'pptx-vue-viewer/i18n', or from 'pptx-angular-viewer' root for Angular

export const translationsFr: Record<TranslationKey, string> = {
	'pptx.statusBar.allSaved': 'Tout enregistré',
	// TypeScript errors if you miss a key, misspell one, or add one that doesn't exist
	...
};
```

Dieses Muster verwenden alle fünf Demo-Apps. Funktionierende Beispiele von Anfang bis Ende finden Sie unter [In den Demos ausprobieren](#in-den-demos-ausprobieren) weiter unten.

## Eine Übersetzung upstream beitragen

Der private Workspace `packages/locales` enthält die Quelle der Referenzwörterbücher, die alle fünf Viewer-Pakete ausliefern. Bauen Sie ihn, bevor Sie ein Binding bauen; jedes Binding kopiert seine eigenständigen Artefakte in sein eigenes Verzeichnis `dist/i18n`. Führen Sie nach dem Bauen aller fünf Bindings `bun run test:locale-subpaths` aus, um die tatsächlichen npm-Tarballs, Importe und Deklarationen zu prüfen.
Jede Sprache ist in benannte Produktbereichsdateien wie `charts.ts`,
`presenting-and-slide-show.ts` und `text-and-equations.ts` gegliedert.

Muttersprachler und fließend Sprechende können helfen, ohne ein ganzes Wörterbuch durchzusehen:

1. Wählen Sie eine Produktbereichsdatei unter `packages/locales/src/fr`, `src/es`, `src/de` oder `src/zh-CN`.
2. Vergleichen Sie jeden Wert mit dem passenden Schlüssel in `src/ui/src/pptx/i18n/translations-en.ts`.
3. Verbessern Sie übersetzte Werte und lassen Sie die Schlüssel mit Punkten und jeden `{{token}}`-Platzhalter unverändert.
4. Bevorzugen Sie die Terminologie der lokalisierten Microsoft-PowerPoint-Oberfläche, besonders bei SmartArt, Morph-Übergängen, Diagrammen und Masteransichten.
5. Führen Sie `bun run --filter 'pptx-viewer-locales' test`, `typecheck` und `build` aus und nennen Sie im Pull Request die geprüfte Sprache und die Produktbereiche.

Wenn Englisch neue Schlüssel erhält, liest `bun run locales:generate` die vorhandenen
semantischen Dateien und bewahrt jede gültige Übersetzung, auch geprüfte Werte, die
absichtlich mit Englisch übereinstimmen. Es füllt nur fehlende Einträge oder Einträge mit
ungültigen Platzhaltern. Prüfen Sie maschinell unterstützte Ergänzungen, bevor Sie sie committen.
Der Generator schlägt fehl, wenn ein neues Schlüsselpräfix keinem benannten Abschnitt zugeordnet wurde.

Um eine weitere Sprache hinzuzufügen, ergänzen Sie ihren Locale-Einstiegspunkt, die Generator-Konfiguration,
den Root-Export und einen Testfall. Exakte Schlüsselgleichheit und Interpolationsplatzhalter werden
durch `packages/locales/src/locales.test.ts` erzwungen.

## In den Demos ausprobieren

Die Demos für [React](https://christophervr.github.io/ooxml/pptx/demo/), [Vue](https://christophervr.github.io/ooxml/pptx/demo-vue/), [Angular](https://christophervr.github.io/ooxml/pptx/demo-angular/), Vanilla und Svelte enthalten jeweils eine Sprachauswahl, die zwischen Englisch, Französisch, Spanisch, Deutsch und vereinfachtem Chinesisch wechselt. Alle vier nicht-englischen Wörterbücher decken jeden kanonischen Viewer-Schlüssel über den privaten Locale-Workspace des Repositorys ab.

Die Anbindung der Auswahl ist Demo-Code; die Wörterbücher werden über den Bun-Workspace aus `packages/locales` referenziert:

- **React**: `demos/pptx/demo-react/i18n.ts`, `LanguagePicker.tsx`
- **Vue**: `demos/pptx/demo-vue/src/i18n.ts`, `LanguagePicker.vue`
- **Angular**: `demos/pptx/demo-angular/src/i18n.ts`, `language-picker.component.ts`
- **Vanilla**: `demos/pptx/demo-vanilla/src/demo-i18n.ts`, `language-picker.ts`
- **Svelte**: `demos/pptx/demo-svelte/src/demo-i18n.svelte.ts`, `LanguagePicker.svelte`

## Nächste Schritte

- [Installation](/de/guide/installation) - Peer-Abhängigkeiten, einschließlich `i18next`/`react-i18next`.
- [React-Hooks](/react/hooks) - `t` wird durch die Hook-Schicht gereicht, falls Sie es außerhalb der eigenen Komponenten des Viewers aufrufen müssen.
