# pptx-viewer-locales

**A private workspace package with the French, Spanish, German, Simplified Chinese, Japanese and Korean dictionaries of the PowerPoint viewer UI. It is never published to npm.**

The five viewer packages bundle these dictionaries: after each framework build, `scripts/copy-subpaths.mjs` copies them into the viewer's `dist/i18n`, so applications import them from their viewer and never install this package:

```ts
import { translationsFr } from 'pptx-react-viewer/i18n/fr';
import { translationsZhCN } from 'pptx-vue-viewer/i18n/zh-CN';
```

Every viewer (`pptx-react-viewer`, `pptx-vue-viewer`, `pptx-angular-viewer`, `pptx-svelte-viewer`, `pptx-vanilla-viewer`) has the `i18n/fr`, `/es`, `/de`, `/zh-CN`, `/ja` and `/ko` subpaths. Inside the workspace, the demos import them by package name (`pptx-viewer-locales/de`, `translationsDe`, and so on). The English source is `src/ui/src/pptx/i18n/translations-en.ts` at the repository root; the [localisation guide](../../docs/guide/localization.md) explains how each binding registers a dictionary.

Each dictionary has every English key; the tests check exact key parity and the `{{placeholder}}` tokens. The translations began as machine-assisted drafts, so native-speaker review is welcome.

## Reviewing a translation

Files are split by product area under `src/<locale>/` (for example `charts.ts` or `presenting-and-slide-show.ts`). Find the key in the English source to see its context, change only the translated value, keep the `pptx.*` key and every `{{placeholder}}` exactly, and prefer the terms the localised PowerPoint UI uses. Then run:

```bash
bun run --filter 'pptx-viewer-locales' test
bun run --filter 'pptx-viewer-locales' typecheck
bun run --filter 'pptx-viewer-locales' build
```

## Adding keys

After English strings are added, run `bun run locales:generate` in `viewers/pptx`. It keeps every valid existing translation, machine-translates only missing entries or entries with broken placeholders, and fails when a new key prefix has no named section. Review what it adds before committing.

Build this package before the viewers. After building them, `bun run test:locale-subpaths` (in `viewers/pptx`) packs the real viewer packages and imports every language through ESM, CommonJS, a bundler and TypeScript.
