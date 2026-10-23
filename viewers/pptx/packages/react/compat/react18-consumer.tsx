/**
 * React 18 type-declaration guard (issue #105).
 *
 * The published `.d.ts` is generated against `@types/react` 19, so this file
 * type-checks a realistic consumer against the BUILT `dist/` declarations while
 * `react`/`react-dom` types are mapped to `@types/react` 18 (see
 * `tsconfig.react18.json`, run by `bun run typecheck:react18`). It compiles
 * nothing at runtime: it exists purely so `tsc` proves the declarations resolve
 * and narrow correctly under the older type definitions.
 *
 * React 18 and 19 differ here in ways that bite libraries: `RefObject<T>.current`
 * is READONLY in 18, `useRef()` still has a zero-argument overload, `ReactNode`
 * excludes promises, and function components do not accept `ref` as a plain
 * prop. Everything below sticks to the React 18 spelling on purpose.
 */
import type {
	PowerPointViewerHandle,
	PowerPointViewerProps,
	PptxI18nProviderProps,
	PptxTranslate,
	PptxTranslateParams,
	ViewerTheme,
} from 'pptx-react-viewer';
import {
	LOCALE_CATALOG,
	PowerPointViewer,
	PptxI18nProvider,
	vermilionDarkTheme,
} from 'pptx-react-viewer';
import type { ReactElement } from 'react';
import { useCallback, useRef, useState } from 'react';

/** A host component wiring the viewer the way the README documents it. */
export function React18Consumer({ content }: { content: Uint8Array }): ReactElement {
	// React 18 types: `useRef<T>(null)` yields `RefObject<T>` with a readonly
	// `current`, which must still satisfy the component's forwarded ref prop.
	const viewerRef = useRef<PowerPointViewerHandle>(null);
	const [theme, setTheme] = useState<ViewerTheme>(vermilionDarkTheme);
	// Compiles ONLY against @types/react 18: React 19 dropped the zero-argument
	// `useRef()` overload. Doubles as proof the tsconfig paths mapping took
	// effect, so a mis-wired path can't turn this check into a no-op.
	const react18Marker = useRef<number>();
	react18Marker.current = content.byteLength;

	const [dirty, setDirty] = useState(false);
	const [selected, setSelected] = useState<string[]>([]);

	const save = useCallback(async (): Promise<Uint8Array | undefined> => {
		const bytes = await viewerRef.current?.getContent();
		return typeof bytes === 'string' ? undefined : bytes;
	}, []);

	const props: PowerPointViewerProps = {
		content,
		canEdit: true,
		theme,
		onDirtyChange: (isDirty: boolean) => setDirty(isDirty),
		onSelectionChange: (elementIds: string[]) => setSelected(elementIds),
		onThemeChange: () => setTheme(vermilionDarkTheme),
	};

	return (
		<div data-dirty={dirty} data-selected={selected.length} onClick={() => void save()}>
			<PowerPointViewer ref={viewerRef} {...props} />
		</div>
	);
}

/** The shape of a host library's message lookup, e.g. Lingui's `i18n._(id, values)`. */
declare const hostI18n: { _(id: string, values?: Record<string, unknown>): string };

/** The documented module-level translator: `params` passes straight to the host library. */
export const moduleTranslate: PptxTranslate = (key, params) => hostI18n._(key, params);

/** The shipped dictionaries as they are: a missing entry is `undefined`, a miss. */
export const dictionaryTranslate =
	(dictionary: Readonly<Record<string, string>>): PptxTranslate =>
	(key) =>
		dictionary[key];

/** A host translating the viewer with its own i18n library through `PptxI18nProvider`. */
export function React18TranslatedConsumer({ content }: { content: Uint8Array }): ReactElement {
	const [locale, setLocale] = useState('en');
	// Stable per language, as the provider documents: a new function re-translates the viewer.
	const translate = useCallback<PptxTranslate>(
		(key: string, params?: PptxTranslateParams) =>
			locale === 'en' ? undefined : `${key}:${String(params?.count ?? '')}`,
		[locale],
	);
	const props: PptxI18nProviderProps = { translate, locale };
	return (
		<PptxI18nProvider {...props}>
			<PowerPointViewer
				content={content}
				availableLocales={LOCALE_CATALOG.slice(0, 2)}
				onLocaleChange={setLocale}
			/>
		</PptxI18nProvider>
	);
}
