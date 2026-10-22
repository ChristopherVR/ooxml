---
title: Localisation (i18n)
description: Comment les chaînes de l'interface de pptx-viewer fonctionnent avec React, Vue 3, Angular, Vanilla et Svelte, comment ajouter une langue dans votre application et comment contribuer une traduction.
---

# Localisation (i18n)

**La visionneuse recherche chaque libellé d'interface via une clé de traduction `pptx.*` ; votre application fournit le dictionnaire.** Chaque binding livre l'anglais. Des dictionnaires de référence complets en français, espagnol, allemand et chinois simplifié sont aussi livrés via les sous-chemins `i18n/<locale>` de chaque package de visionneuse. Leur source est maintenue une seule fois dans le workspace privé `pptx-viewer-locales`, qui est intégré aux packages de visionneuse et n'a pas besoin d'être installé séparément. React, Vue et Angular délèguent à la bibliothèque i18n propre à leur framework hôte ; Vanilla et Svelte livrent un petit traducteur intégré :

| Binding | Appel de traduction que la visionneuse effectue           | Bibliothèque que vous fournissez                                                  |
| ------- | --------------------------------------------------------- | --------------------------------------------------------------------------------- |
| React   | `t(key, opts)` de `react-i18next`                         | [i18next](https://www.i18next.com/) / [react-i18next](https://react.i18next.com/) |
| Vue 3   | `useI18n().t(key, opts)` de `vue-i18n`                    | [vue-i18n](https://vue-i18n.intlify.dev/)                                         |
| Angular | signal `translate()` / `TranslatePipe` de `ngx-translate` | [@ngx-translate/core](https://github.com/ngx-translate/core)                      |
| Vanilla | le `t(key, params)` intégré au package                    | aucune : passez directement un dictionnaire `messages`, sans bibliothèque i18n    |
| Svelte  | le `t(key, params)` intégré au package                    | aucune : appelez `registerTranslations(locale, dict)`, sans bibliothèque i18n     |

Cela garde la visionneuse idiomatique pour chaque écosystème (vous branchez React/Vue/Angular sur la configuration i18n que votre application a déjà ; Vanilla/Svelte n'ont besoin de rien de plus) plutôt que d'imposer une dépendance d'exécution de traduction à des bindings qui n'en ont pas déjà une.

## Le dictionnaire anglais

Chaque package réexporte un bundle de ressources anglaises prêt à l'emploi pour que vous n'ayez pas à en écrire un de zéro :

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

`translationsEn` est un `Record<string, string>` plat de toutes les clés `pptx.*`, par exemple `'pptx.statusBar.allSaved': 'All saved'`. Les valeurs à contenu dynamique utilisent des espaces réservés d'interpolation `{{token}}`, par exemple `'pptx.statusBar.slideOf': 'Slide {{current}} of {{total}}'` : votre bibliothèque i18n les remplace à partir des `opts` passés à `t()`/`translate()`.

`keyToLabel(key)` dérive un libellé lisible du dernier segment d'une clé lorsqu'aucune entrée du dictionnaire ne correspond (`"pptx.slideSorter.zoomIn"` → `"Zoom In"`). Branchez-la comme gestionnaire de clé manquante de votre bibliothèque pour que toute clé non encore traduite affiche quelque chose de raisonnable plutôt que la chaîne brute de la clé.

## Mise en place

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

Changez de langue avec `i18nInstance.changeLanguage('fr')`.

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

Changez de langue en définissant `i18n.global.locale.value = 'fr'`.

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

Enregistrez le dictionnaire anglais une fois (typiquement dans votre composant racine) et changez de langue via `TranslateService` :

```ts
import { inject } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { translationsEn } from 'pptx-angular-viewer';

const translate = inject(TranslateService);
translate.setTranslation('en', translationsEn);
// later, to switch:
translate.use('fr');
```

::: tip Passez un objet `Provider` simple, pas une classe nue
Si vous écrivez votre propre `MissingTranslationHandler`, passez-le sous la forme `{ provide: MissingTranslationHandler, useClass: YourHandler }` plutôt que comme référence de classe nue. Avec certaines chaînes de build (Vite/Rolldown avec décorateurs standard), l'heuristique `isClass()` de ngx-translate peut se tromper dans les builds de production et appeler le gestionnaire sans `new`, ce qui provoque une erreur au démarrage.
:::

### Vanilla

Le binding Vanilla n'a pas de framework où brancher une bibliothèque i18n ; il livre donc un minuscule traducteur intégré : passez directement à `createPptxViewer` un dictionnaire `messages` indexé par locale, sans étape de configuration séparée.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

import { translationsFr } from './translations/fr';

const viewer = createPptxViewer(document.querySelector('#host')!, {
	source: '/deck.pptx',
	locale: 'fr',
	messages: { fr: translationsFr },
});
```

`messages` associe une locale à un `Record<string, string>` plat de surcharges `pptx.*` ; vous n'avez pas à inclure toutes les clés ni à répandre vous-même `translationsEn`, puisque l'anglais sert toujours de repli pour tout ce qui manque. Changez de langue à l'exécution avec `viewer.setLocale('fr')` (cela reconstruit les libellés statiques de l'interface sous la nouvelle locale et relance le rendu).

### Svelte

Svelte n'a pas non plus d'environnement i18n de référence ; le binding expose donc `registerTranslations` depuis son sous-chemin `/i18n` : enregistrez le dictionnaire de chaque locale une fois au démarrage, puis passez la `locale` active en prop.

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

Les enregistrements fusionnent par-dessus ce qui est déjà enregistré pour cette locale ; vous pouvez donc appeler `registerTranslations` plusieurs fois (par exemple une fois par domaine fonctionnel) sans écraser les surcharges précédentes. Comme pour Vanilla, l'anglais sert toujours de repli pour les clés que vous n'avez pas encore traduites.

## Fichier > Options > Langue

La boîte de dialogue Paramètres de chaque binding comporte aussi un onglet **Langue** : une liste de locales que l'utilisateur peut choisir à l'exécution, juste à côté de l'onglet [Apparence](/guide/theming). Il n'intègre ni ne récupère jamais de contenu de traduction lui-même : il propose seulement une interface pour basculer entre les locales _que vous avez déjà enregistrées_ dans la configuration i18n mise en place ci-dessus. Concrètement, lorsque vous ne fournissez pas vous-même `availableLocales`, chaque binding inspecte ce qui est déjà chargé :

| Binding | Comment il découvre les locales disponibles                                  |
| ------- | ---------------------------------------------------------------------------- |
| React   | `i18n.options.resources` / `i18n.languages` de l'instance `react-i18next`    |
| Vue 3   | `useI18n().availableLocales` (propriété du composable intégré de `vue-i18n`) |
| Angular | `TranslateService.getLangs()`                                                |
| Vanilla | le dictionnaire `messages` passé à `createPptxViewer`, plus `'en'`           |
| Svelte  | chaque locale passée à `registerTranslations`                                |

Ainsi, si vous n'avez enregistré que `en` et `fr`, l'onglet Langue propose exactement ces deux-là, jamais une locale sans dictionnaire derrière. Les codes sont libellés via le `LOCALE_CATALOG` partagé (noms d'affichage anglais/français/espagnol/allemand/chinois simplifié) lorsqu'ils sont reconnus, ou affichés tels quels sinon.

La priorité reflète celle de l'onglet Apparence : une locale choisie s'applique immédiatement en appelant directement votre instance i18n (`i18n.changeLanguage`, `locale.value =`, `TranslateService.use`, etc.) et est conservée dans `localStorage` (`pptx-viewer-prefs`), sauf si vous passez `onLocaleChange` : dans ce cas la visionneuse ne touche jamais elle-même à votre instance i18n et se contente d'appeler ce rappel, laissant l'application et la persistance entièrement à votre charge. Voir [Thèmes](/guide/theming) pour la référence complète des props `defaultLocale`/`availableLocales`/`onLocaleChange` (de même forme que les props de thème, documentées une seule fois là-bas).

## Chinois simplifié

Les démos proposent **简体中文** (`zh-CN`) dans **Fichier > Options > Langue**. Le
dictionnaire de référence couvre les mêmes clés et les mêmes `{{placeholders}}` que l'anglais.
Il traduit l'interface de l'éditeur ; il ne traduit pas le texte de vos diapositives.

Importez le dictionnaire chinois directement depuis le package de visionneuse que vous utilisez déjà :

```ts
import { translationsZhCN } from 'pptx-react-viewer/i18n/zh-CN';
// Other bindings expose the same subpath:
// 'pptx-vue-viewer/i18n/zh-CN'
// 'pptx-angular-viewer/i18n/zh-CN'
// 'pptx-svelte-viewer/i18n/zh-CN'
// 'pptx-vanilla-viewer/i18n/zh-CN'
```

Chaque sous-chemin contient un dictionnaire autonome avec ESM, CommonJS et
déclarations TypeScript. L'importer ne charge ni la visionneuse ni les autres
langues. Les exports anglais existants restent disponibles à leurs points d'entrée
habituels. Enregistrez le dictionnaire auprès de votre binding :

| Binding | Enregistrer le dictionnaire importé                                          | Changer de langue                                      |
| ------- | ---------------------------------------------------------------------------- | ------------------------------------------------------ |
| React   | `i18n.addResourceBundle('zh-CN', 'translation', translationsZhCN)`           | `i18n.changeLanguage('zh-CN')`                         |
| Vue     | `i18n.global.setLocaleMessage('zh-CN', toVueI18nSyntax(translationsZhCN))`   | `i18n.global.locale.value = 'zh-CN'`                   |
| Angular | `translate.setTranslation('zh-CN', translationsZhCN)`                        | `translate.use('zh-CN')`                               |
| Vanilla | `messages: { 'zh-CN': translationsZhCN }` dans les options de la visionneuse | `viewer.setLocale('zh-CN')`                            |
| Svelte  | `registerTranslations('zh-CN', translationsZhCN)`                            | Définir la prop `locale` de la visionneuse à `'zh-CN'` |

Pour Vue, importez `toVueI18nSyntax` depuis `pptx-vue-viewer/i18n` ; elle convertit
les espaces réservés `{{name}}` du dictionnaire de référence en syntaxe `{name}` de Vue.
Le sélecteur de langue reconnaît `zh-CN` et affiche **简体中文** dès que
l'application a enregistré le dictionnaire.

## Autres langues de référence

Les cinq packages de visionneuse fournissent les mêmes sous-chemins de locale optionnels :

| Locale            | Sous-chemin  | Export             |
| ----------------- | ------------ | ------------------ |
| Français          | `i18n/fr`    | `translationsFr`   |
| Espagnol          | `i18n/es`    | `translationsEs`   |
| Allemand          | `i18n/de`    | `translationsDe`   |
| Chinois simplifié | `i18n/zh-CN` | `translationsZhCN` |

Par exemple, `import { translationsFr } from 'pptx-vue-viewer/i18n/fr'`.
Les applications Vue doivent faire passer ces dictionnaires par `toVueI18nSyntax`
avant de les enregistrer, comme dans l'exemple chinois ci-dessus. Les dictionnaires
conservent les espaces réservés canoniques `{{name}}` utilisés par les autres bindings.

Ces traductions sont livrées avec la version de la visionneuse ; mettre à jour le composant
met donc aussi à jour les dictionnaires de référence. Elles sont optionnelles et ne sont pas
réexportées depuis la racine du package ni depuis le point d'entrée i18n anglais existant.

## Ajouter une langue dans votre application

Traduire la visionneuse dans une autre langue n'exige pas de toucher à ce dépôt : c'est exactement le même genre de travail que traduire n'importe quelle autre partie de votre application. Construisez un dictionnaire avec les mêmes clés que `translationsEn` et enregistrez-le comme deuxième ressource/locale/traduction dans la bibliothèque de votre framework (`resources.fr` pour i18next, `messages.fr` pour vue-i18n, `translate.setTranslation('fr', ...)` pour ngx-translate). Vous n'avez pas besoin de toutes les clés dès le premier jour : tout ce que vous n'avez pas encore traduit retombe automatiquement sur `keyToLabel`, de sorte qu'une couverture partielle se dégrade en douceur au lieu d'afficher des libellés vides ou des clés brutes.

Pour détecter les clés manquantes ou mal orthographiées à la compilation plutôt que de compter sur le repli à l'exécution, typez votre dictionnaire d'après l'union `TranslationKey` exportée plutôt que par un simple `Record<string, string>` :

```ts
import type { TranslationKey } from 'pptx-react-viewer/i18n'; // 'pptx-vue-viewer/i18n', or from 'pptx-angular-viewer' root for Angular

export const translationsFr: Record<TranslationKey, string> = {
	'pptx.statusBar.allSaved': 'Tout enregistré',
	// TypeScript errors if you miss a key, misspell one, or add one that doesn't exist
	...
};
```

C'est le schéma utilisé par les cinq applications de démonstration. Voir [Essayer dans les démos](#essayer-dans-les-demos) ci-dessous pour des exemples complets et fonctionnels.

## Contribuer une traduction en amont

Le workspace privé `packages/locales` contient la source des dictionnaires de référence livrés par les cinq packages de visionneuse. Construisez-le avant de construire un binding ; chaque binding copie ses artefacts autonomes dans son propre répertoire `dist/i18n`. Exécutez `bun run test:locale-subpaths` après avoir construit les cinq bindings pour vérifier les vrais tarballs npm, les imports et les déclarations.
Chaque langue est organisée en fichiers nommés par domaine fonctionnel, comme `charts.ts`,
`presenting-and-slide-show.ts` et `text-and-equations.ts`.

Les locuteurs natifs et courants peuvent aider sans relire un dictionnaire entier :

1. Choisissez un fichier de domaine fonctionnel sous `packages/locales/src/fr`, `src/es`, `src/de` ou `src/zh-CN`.
2. Comparez chaque valeur à la clé correspondante dans `src/ui/src/pptx/i18n/translations-en.ts`.
3. Améliorez les valeurs traduites en laissant inchangés les clés à points et chaque espace réservé `{{token}}`.
4. Privilégiez la terminologie de l'interface localisée de Microsoft PowerPoint, surtout pour SmartArt, les transitions morphe, les graphiques et les affichages masque.
5. Exécutez `bun run --filter 'pptx-viewer-locales' test`, `typecheck` et `build`, puis indiquez dans la pull request la langue relue et les domaines fonctionnels concernés.

Lorsque l'anglais gagne de nouvelles clés, `bun run locales:generate` lit les fichiers
sémantiques existants et préserve chaque traduction valide, y compris les valeurs relues
qui coïncident volontairement avec l'anglais. Il ne remplit que les entrées manquantes ou
celles dont les espaces réservés sont invalides. Relisez les ajouts assistés par machine
avant de les valider. Le générateur échoue si un nouveau préfixe de clé n'a pas été
affecté à une section nommée.

Pour ajouter une autre langue, ajoutez son point d'entrée de locale, sa configuration du
générateur, son export à la racine et un cas de test. La parité exacte des clés et des
espaces réservés d'interpolation est vérifiée par `packages/locales/src/locales.test.ts`.

## Essayer dans les démos

Les démos [React](https://christophervr.github.io/ooxml/pptx/demo/), [Vue](https://christophervr.github.io/ooxml/pptx/demo-vue/), [Angular](https://christophervr.github.io/ooxml/pptx/demo-angular/), Vanilla et Svelte comportent chacune un sélecteur de langue qui bascule entre anglais, français, espagnol, allemand et chinois simplifié. Les quatre dictionnaires non anglais couvrent chaque clé canonique de la visionneuse via le workspace de locales privé du dépôt.

Le câblage du sélecteur est du code de démo ; les dictionnaires sont référencés depuis `packages/locales` via le workspace Bun :

- **React** : `demos/pptx/demo-react/i18n.ts`, `LanguagePicker.tsx`
- **Vue** : `demos/pptx/demo-vue/src/i18n.ts`, `LanguagePicker.vue`
- **Angular** : `demos/pptx/demo-angular/src/i18n.ts`, `language-picker.component.ts`
- **Vanilla** : `demos/pptx/demo-vanilla/src/demo-i18n.ts`, `language-picker.ts`
- **Svelte** : `demos/pptx/demo-svelte/src/demo-i18n.svelte.ts`, `LanguagePicker.svelte`

## Étapes suivantes

- [Installation](/fr/guide/installation) - dépendances peer, dont `i18next`/`react-i18next`.
- [Hooks React](/react/hooks) - `t` est transmis à travers la couche de hooks si vous devez l'appeler en dehors des composants de la visionneuse.
