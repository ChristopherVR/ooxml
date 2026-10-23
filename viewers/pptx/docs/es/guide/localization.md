---
title: Localización (i18n)
description: Cómo funcionan las cadenas de interfaz de pptx-viewer con React, Vue 3, Angular, Vanilla y Svelte, cómo añadir un idioma en su aplicación y cómo contribuir una traducción.
---

# Localización (i18n)

**El visualizador busca cada etiqueta de interfaz mediante una clave de traducción `pptx.*`; su aplicación proporciona el diccionario.** Cada enlace incluye el inglés. Los diccionarios de referencia completos en francés, español, alemán y chino simplificado también se distribuyen a través de las subrutas `i18n/<locale>` de cada paquete del visualizador. Su código fuente se mantiene una sola vez en el workspace privado `pptx-viewer-locales`, que se empaqueta dentro de los paquetes del visualizador y no necesita instalarse por separado. React, Vue y Angular delegan en la biblioteca i18n de su framework anfitrión (React también acepta su propia función `translate` mediante `PptxI18nProvider`); Vanilla y Svelte incluyen un pequeño traductor integrado:

| Enlace  | Llamada de traducción que hace el visualizador                                         | Biblioteca que usted proporciona                                                                                                      |
| ------- | -------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| React   | `t(key, opts)` de `react-i18next`, o su propia `translate` mediante `PptxI18nProvider` | [i18next](https://www.i18next.com/) / [react-i18next](https://react.i18next.com/), o cualquier biblioteca mediante `PptxI18nProvider` |
| Vue 3   | `useI18n().t(key, opts)` de `vue-i18n`                                                 | [vue-i18n](https://vue-i18n.intlify.dev/)                                                                                             |
| Angular | señal `translate()` / `TranslatePipe` de `ngx-translate`                               | [@ngx-translate/core](https://github.com/ngx-translate/core)                                                                          |
| Vanilla | El `t(key, params)` integrado del propio paquete                                       | ninguna: pase un diccionario `messages` directamente, sin biblioteca i18n                                                             |
| Svelte  | El `t(key, params)` integrado del propio paquete                                       | ninguna: llame a `registerTranslations(locale, dict)`, sin biblioteca i18n                                                            |

Así el visualizador sigue siendo idiomático en cada ecosistema (usted conecta React/Vue/Angular a la configuración i18n que su aplicación ya tenga; Vanilla/Svelte no necesitan nada adicional) en lugar de imponer una dependencia de ejecución de traducción a enlaces que aún no la tienen.

## El diccionario en inglés

Cada paquete reexporta un bundle de recursos en inglés listo para usar, para que no tenga que escribir uno desde cero:

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

`translationsEn` es un `Record<string, string>` plano con todas las claves `pptx.*`, por ejemplo `'pptx.statusBar.allSaved': 'All saved'`. Los valores con contenido dinámico usan marcadores de interpolación `{{token}}`, por ejemplo `'pptx.statusBar.slideOf': 'Slide {{current}} of {{total}}'`; su biblioteca i18n los sustituye a partir de las `opts` pasadas a `t()`/`translate()`.

`keyToLabel(key)` deriva una etiqueta legible del último segmento de una clave cuando ninguna entrada del diccionario coincide (`"pptx.slideSorter.zoomIn"` → `"Zoom In"`). Conéctela como manejador de claves faltantes de su biblioteca para que cualquier clave aún sin traducir muestre algo razonable en lugar de la cadena de la clave sin procesar.

## Conexión

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

Cambie de idioma con `i18nInstance.changeLanguage('fr')`.

#### Usar su propia biblioteca de i18n

Para traducir el visualizador con Lingui, FormatJS o cualquier otra biblioteca en lugar de react-i18next, envuélvalo en `PptxI18nProvider` y pase una función `translate(key, params)` junto con la `locale` activa. Abarca `PowerPointViewer` y los bloques independientes (`Toolbar`, `SlideCanvas`, `SlideTemplateGalleryDialog`, ...):

```tsx
import { i18n } from '@lingui/core';
import { useLingui } from '@lingui/react';
import { LOCALE_CATALOG, PowerPointViewer, PptxI18nProvider } from 'pptx-react-viewer';
import type { PptxTranslate } from 'pptx-react-viewer';

import { dynamicActivate } from './i18n'; // loads a catalog, then calls i18n.activate(code)

// Module level, so it never changes: the viewer re-translates when `locale` changes.
const translate: PptxTranslate = (key, params) => i18n._(key, params);
const LOCALES = LOCALE_CATALOG.filter((entry) => ['en', 'fr'].includes(entry.code));

function TranslatedViewer({ content }: { content: Uint8Array }) {
	const { i18n: active } = useLingui(); // re-renders when the language changes
	return (
		<PptxI18nProvider locale={active.locale} translate={translate}>
			<PowerPointViewer
				content={content}
				availableLocales={LOCALES}
				onLocaleChange={(code) => void dynamicActivate(code)}
			/>
		</PptxI18nProvider>
	);
}
```

`translate` recibe los parámetros del visualizador sin cambios: `defaultValue`, `count` y los valores de los marcadores `{{name}}`. Use los mismos nombres en su catálogo (`{name}` en sintaxis ICU); `count` es un valor simple, así que la pluralización corresponde a su catálogo. Los marcadores `{{name}}` que queden en el texto que devuelve también se rellenan, de modo que los diccionarios incluidos funcionan tal cual: `const translate: PptxTranslate = (key) => translationsFr[key];`. Cuando `translate` devuelve la propia clave, una cadena vacía o nada, o lanza un error, el visualizador muestra su texto en inglés integrado, luego el texto predeterminado de la llamada y luego una etiqueta derivada de la clave, de modo que un catálogo parcial nunca muestra una clave `pptx.*` sin traducir. Mantenga `translate` estable (una función a nivel de módulo o `useCallback`): el visualizador vuelve a traducir cada vez que cambian `translate` o `locale`.

Dentro de un proveedor, el idioma pertenece al anfitrión: el visualizador nunca llama a `changeLanguage`, ignora `defaultLocale` y nunca restaura una elección de idioma guardada. Archivo > Opciones > Idioma solo aparece cuando pasa `onLocaleChange` y dos o más `availableLocales`; una elección llama a `onLocaleChange` y la selección sigue a `locale` (una `locale` regional como `fr-CA` selecciona un `fr` ofrecido). `locale` no establece la dirección del texto. Un `PowerPointViewer` lee el proveedor más cercano por encima de él; un proveedor dentro del visualizador se ignora, de modo que todo el visualizador comparte un idioma. `i18next` y `react-i18next` siguen siendo dependencias peer obligatorias en 4.x: instálelas, aunque dentro de un proveedor el visualizador no las usa. `PptxI18nProvider` es solo para React; los demás enlaces traducen como se describe a continuación.

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

Cambie de idioma asignando `i18n.global.locale.value = 'fr'`.

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

Registre el diccionario en inglés una sola vez (normalmente en su componente raíz) y cambie de idioma mediante `TranslateService`:

```ts
import { inject } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { translationsEn } from 'pptx-angular-viewer';

const translate = inject(TranslateService);
translate.setTranslation('en', translationsEn);
// later, to switch:
translate.use('fr');
```

::: tip Pase un objeto `Provider` simple, no una clase directa
Si escribe su propio `MissingTranslationHandler`, páselo como `{ provide: MissingTranslationHandler, useClass: YourHandler }` en lugar de una referencia directa a la clase. En algunas cadenas de compilación (Vite/Rolldown con decoradores estándar), la heurística `isClass()` de ngx-translate puede fallar en compilaciones de producción e invocar el manejador sin `new`, lo que provoca un error en el arranque.
:::

### Vanilla

El enlace Vanilla no tiene un framework donde conectar una biblioteca i18n, por lo que incluye un pequeño traductor integrado: pase un diccionario `messages` indexado por locale directamente a `createPptxViewer`, sin ningún paso de conexión aparte.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

import { translationsFr } from './translations/fr';

const viewer = createPptxViewer(document.querySelector('#host')!, {
	source: '/deck.pptx',
	locale: 'fr',
	messages: { fr: translationsFr },
});
```

`messages` asocia un locale con un `Record<string, string>` plano de reemplazos `pptx.*`; no necesita incluir todas las claves ni extender `translationsEn` usted mismo, ya que el inglés es siempre el respaldo para lo que falte. Cambie de idioma en tiempo de ejecución con `viewer.setLocale('fr')` (esto reconstruye las etiquetas estáticas del chrome con el nuevo locale y vuelve a renderizar).

### Svelte

Svelte tampoco tiene un runtime i18n oficial, así que el enlace expone `registerTranslations` desde su subruta `/i18n`: registre el diccionario de cada locale una vez al inicio y luego pase el `locale` activo como prop.

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

Los registros se combinan con lo que ya esté registrado para ese locale, por lo que puede llamar a `registerTranslations` más de una vez (por ejemplo, una por área funcional) sin pisar reemplazos anteriores. Como en Vanilla, el inglés es siempre el respaldo para las claves que aún no ha traducido.

## Archivo > Opciones > Idioma

El diálogo de Configuración de cada enlace también tiene una pestaña **Idioma**: una lista de locales entre los que el usuario puede elegir en tiempo de ejecución, junto a la pestaña [Apariencia](/guide/theming). Nunca empaqueta ni descarga contenido de traducción por sí misma: solo ofrece una interfaz para alternar entre los locales _que usted ya ha registrado_ con la configuración i18n que conectó más arriba. En concreto, cuando no proporciona `availableLocales` usted mismo, cada enlace inspecciona lo que ya está cargado:

| Enlace  | Cómo descubre los locales disponibles                                                                                                    |
| ------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| React   | `i18n.options.resources` / `i18n.languages` de la instancia de `react-i18next`; dentro de un `PptxI18nProvider`: solo `availableLocales` |
| Vue 3   | `useI18n().availableLocales` (propiedad integrada del composable de `vue-i18n`)                                                          |
| Angular | `TranslateService.getLangs()`                                                                                                            |
| Vanilla | el diccionario `messages` pasado a `createPptxViewer`, más `'en'`                                                                        |
| Svelte  | cada locale pasado a `registerTranslations`                                                                                              |

Así, si solo ha registrado `en` y `fr`, la pestaña Idioma ofrece exactamente esos dos, nunca un locale sin diccionario detrás. Los códigos se etiquetan mediante el `LOCALE_CATALOG` compartido (nombres de inglés, francés, español, alemán y chino simplificado) cuando se reconocen, o se muestra el código sin procesar en caso contrario.

La precedencia refleja la de la pestaña Apariencia: un locale elegido se aplica de inmediato llamando directamente a su instancia i18n (`i18n.changeLanguage`, `locale.value =`, `TranslateService.use`, etc.) y se guarda en `localStorage` (`pptx-viewer-prefs`), salvo que pase `onLocaleChange`; en ese caso el visualizador nunca toca su instancia i18n y solo invoca esa devolución de llamada, dejando la persistencia y la aplicación enteramente a su cargo. Dentro de un `PptxI18nProvider` el visualizador nunca aplica ni guarda una elección: llama a `onLocaleChange` y, sin esa devolución de llamada, la pestaña Idioma no aparece. Consulte [Temas](/guide/theming) para la referencia completa de las props `defaultLocale`/`availableLocales`/`onLocaleChange` (con la misma forma que las props de tema, documentadas allí una sola vez).

## Chino simplificado

Las demos ofrecen **简体中文** (`zh-CN`) en **Archivo > Opciones > Idioma**. El
diccionario de referencia cubre las mismas claves y `{{placeholders}}` que el inglés.
Traduce la interfaz del editor; no traduce el texto de sus diapositivas.

Importe el diccionario chino directamente desde el paquete del visualizador que ya usa:

```ts
import { translationsZhCN } from 'pptx-react-viewer/i18n/zh-CN';
// Other bindings expose the same subpath:
// 'pptx-vue-viewer/i18n/zh-CN'
// 'pptx-angular-viewer/i18n/zh-CN'
// 'pptx-svelte-viewer/i18n/zh-CN'
// 'pptx-vanilla-viewer/i18n/zh-CN'
```

Cada subruta contiene un diccionario independiente con ESM, CommonJS y
declaraciones de TypeScript. Importarlo no carga el visualizador ni otros
idiomas. Las exportaciones existentes en inglés siguen disponibles en sus
puntos de entrada habituales. Registre el diccionario con su enlace:

| Enlace  | Registre el diccionario importado                                          | Cambiar de idioma                                    |
| ------- | -------------------------------------------------------------------------- | ---------------------------------------------------- |
| React   | `i18n.addResourceBundle('zh-CN', 'translation', translationsZhCN)`         | `i18n.changeLanguage('zh-CN')`                       |
| Vue     | `i18n.global.setLocaleMessage('zh-CN', toVueI18nSyntax(translationsZhCN))` | `i18n.global.locale.value = 'zh-CN'`                 |
| Angular | `translate.setTranslation('zh-CN', translationsZhCN)`                      | `translate.use('zh-CN')`                             |
| Vanilla | `messages: { 'zh-CN': translationsZhCN }` en las opciones del visualizador | `viewer.setLocale('zh-CN')`                          |
| Svelte  | `registerTranslations('zh-CN', translationsZhCN)`                          | Asigne la prop `locale` del visualizador a `'zh-CN'` |

Para Vue, importe `toVueI18nSyntax` desde `pptx-vue-viewer/i18n`; convierte los
marcadores `{{name}}` del diccionario de referencia a la sintaxis `{name}` de Vue.
El selector de idioma reconoce `zh-CN` y muestra **简体中文** una vez que la
aplicación registra el diccionario.

## Otros idiomas de referencia

Los cinco paquetes del visualizador ofrecen las mismas subrutas de locale opcionales:

| Locale             | Subruta      | Exportación        |
| ------------------ | ------------ | ------------------ |
| Francés            | `i18n/fr`    | `translationsFr`   |
| Español            | `i18n/es`    | `translationsEs`   |
| Alemán             | `i18n/de`    | `translationsDe`   |
| Chino simplificado | `i18n/zh-CN` | `translationsZhCN` |

Por ejemplo, `import { translationsFr } from 'pptx-vue-viewer/i18n/fr'`.
Las aplicaciones Vue deben pasar estos diccionarios por `toVueI18nSyntax`
antes de registrarlos, igual que en el ejemplo chino anterior. Los diccionarios
conservan los marcadores canónicos `{{name}}` que usan los demás enlaces.

Estas traducciones se distribuyen con la versión del visualizador, por lo que
actualizar el componente también actualiza los diccionarios de referencia. Son
opcionales y no se reexportan desde la raíz del paquete ni desde el punto de
entrada i18n en inglés existente.

## Añadir un idioma en su aplicación

Traducir el visualizador a otro idioma no requiere tocar este repositorio en absoluto: es exactamente el mismo tipo de trabajo que traducir cualquier otra parte de su aplicación. Construya un diccionario con las mismas claves que `translationsEn` y regístrelo como un segundo recurso/locale/traducción en la biblioteca que use su framework (`resources.fr` para i18next, `messages.fr` para vue-i18n, `translate.setTranslation('fr', ...)` para ngx-translate). No necesita todas las claves desde el primer día: lo que aún no haya traducido recurre automáticamente a `keyToLabel`, de modo que una cobertura parcial se degrada con elegancia en lugar de mostrar etiquetas en blanco o claves sin procesar.

Para detectar claves faltantes o mal escritas en tiempo de compilación en lugar de depender del respaldo en ejecución, tipifique su diccionario con la unión `TranslationKey` exportada en lugar de un simple `Record<string, string>`:

```ts
import type { TranslationKey } from 'pptx-react-viewer/i18n'; // 'pptx-vue-viewer/i18n', or from 'pptx-angular-viewer' root for Angular

export const translationsFr: Record<TranslationKey, string> = {
	'pptx.statusBar.allSaved': 'Tout enregistré',
	// TypeScript errors if you miss a key, misspell one, or add one that doesn't exist
	...
};
```

Este es el patrón que usan las cinco aplicaciones de demostración. Consulte [Pruébelo en las demos](#pruebelo-en-las-demos) más abajo para ver ejemplos completos y funcionales.

## Contribuir una traducción al proyecto

El workspace privado `packages/locales` contiene el código fuente de los diccionarios de referencia que distribuyen los cinco paquetes del visualizador. Compílelo antes de compilar un enlace; cada enlace copia sus artefactos independientes en su propio directorio `dist/i18n`. Ejecute `bun run test:locale-subpaths` tras compilar los cinco enlaces para comprobar los tarballs de npm, las importaciones y las declaraciones reales.
Cada idioma se organiza en archivos con nombre por área de producto, como `charts.ts`,
`presenting-and-slide-show.ts` y `text-and-equations.ts`.

Los hablantes nativos y fluidos pueden ayudar sin revisar un diccionario completo:

1. Elija un archivo de área de producto en `packages/locales/src/fr`, `src/es`, `src/de` o `src/zh-CN`.
2. Compare cada valor con la clave correspondiente en `src/ui/src/pptx/i18n/translations-en.ts`.
3. Mejore los valores traducidos dejando sin cambios las claves con puntos y todos los marcadores `{{token}}`.
4. Prefiera la terminología de la interfaz localizada de Microsoft PowerPoint, especialmente para SmartArt, transiciones de transformación, gráficos y vistas patrón.
5. Ejecute `bun run --filter 'pptx-viewer-locales' test`, `typecheck` y `build`, e indique en la pull request el idioma y las áreas de producto revisadas.

Cuando el inglés gana claves nuevas, `bun run locales:generate` lee los archivos
semánticos existentes y conserva toda traducción válida, incluidos los valores
revisados que coinciden intencionadamente con el inglés. Solo rellena las entradas
faltantes o con marcadores no válidos. Revise las incorporaciones asistidas por
máquina antes de confirmarlas. El generador falla si un prefijo de clave nuevo no
se ha asignado a una sección con nombre.

Para añadir otro idioma, agregue su punto de entrada de locale, la configuración
del generador, la exportación raíz y el caso de prueba. La paridad exacta de claves
y los marcadores de interpolación se verifican en `packages/locales/src/locales.test.ts`.

## Pruébelo en las demos

Las demos de [React](https://christophervr.github.io/ooxml/pptx/demo/), [Vue](https://christophervr.github.io/ooxml/pptx/demo-vue/), [Angular](https://christophervr.github.io/ooxml/pptx/demo-angular/), Vanilla y Svelte incluyen cada una un selector de idioma que alterna entre inglés, francés, español, alemán y chino simplificado. Los cuatro diccionarios que no son inglés cubren todas las claves canónicas del visualizador mediante el workspace de locales privado del repositorio.

La conexión del selector es código de demostración; los diccionarios se referencian desde `packages/locales` mediante el workspace de Bun:

- **React**: `demos/pptx/demo-react/i18n.ts`, `LanguagePicker.tsx`
- **Vue**: `demos/pptx/demo-vue/src/i18n.ts`, `LanguagePicker.vue`
- **Angular**: `demos/pptx/demo-angular/src/i18n.ts`, `language-picker.component.ts`
- **Vanilla**: `demos/pptx/demo-vanilla/src/demo-i18n.ts`, `language-picker.ts`
- **Svelte**: `demos/pptx/demo-svelte/src/demo-i18n.svelte.ts`, `LanguagePicker.svelte`

## Siguientes pasos

- [Instalación](/es/guide/installation) - dependencias entre pares, incluidas `i18next`/`react-i18next`.
- [Hooks de React](/react/hooks) - `t` se propaga por la capa de hooks si necesita llamarla fuera de los componentes del propio visualizador.
