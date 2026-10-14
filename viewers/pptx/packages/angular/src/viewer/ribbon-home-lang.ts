import type { Signal } from '@angular/core';
/**
 * ribbon-home-lang.ts: runtime-locale plumbing for the shared Home elements.
 *
 * The shared elements call `translate` while they render, outside Angular's
 * dependency tracking. Reading {@link homeLanguage} inside a state method
 * makes an OnPush view re-derive its state (and so re-translate) whenever the
 * language or a dictionary changes; the snapshot translator resolves the
 * families' labels at that moment.
 */
import { toSignal } from '@angular/core/rxjs-interop';
import type { TranslateService } from '@ngx-translate/core';
import { homeSnapshotTranslator } from 'ooxml-ui/pptx';
import type { RibbonHomeFamily, RibbonHomeViewState } from 'ooxml-ui/pptx';
import { merge } from 'rxjs';
import { map, startWith } from 'rxjs/operators';

/** The shared elements' translator: `params` fill a key's `{{name}}` placeholders. */
export type Translate = NonNullable<RibbonHomeViewState['translate']>;

/** Changes on every language or dictionary update; must run in an injection context. */
export function homeLanguage(translation: TranslateService | null): Signal<number> {
	const changes = translation
		? merge(translation.onLangChange, translation.onTranslationChange).pipe(
				map(() => Date.now()),
				startWith(0),
			)
		: undefined;
	return changes ? toSignal(changes, { initialValue: 0 }) : ((() => 0) as Signal<number>);
}

/** A translator for `families` resolved now, so a locale switch re-translates their labels. */
export function homeTranslator(
	translation: TranslateService | null,
	language: Signal<number>,
	families: readonly RibbonHomeFamily[],
): Translate {
	language();
	return homeSnapshotTranslator(families, ngxTranslate(translation));
}

/** `TranslateService.instant` as a shared `Translate` (the key itself when there is no service). */
export function ngxTranslate(translation: TranslateService | null): Translate {
	return (key, params) => {
		const text: unknown = translation?.instant(key, params ? { ...params } : undefined);
		return typeof text === 'string' ? text : key;
	};
}
