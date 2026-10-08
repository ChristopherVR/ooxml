/**
 * The pptx document-property shapes (`PptxCoreProperties`, `PptxAppProperties`,
 * `PptxCustomProperty`) mapped onto the shared `opc/properties` model, which parses and writes
 * the parts for every format.
 *
 * @module document-properties-model
 */

import {
	customPropertyFromText,
	type AppProperties,
	type CoreProperties,
	type CustomProperty,
	type CustomPropertyText,
} from '../../../opc/properties/index';
import type { PptxAppProperties, PptxCoreProperties, PptxCustomProperty } from '../types';

const CORE_FIELDS = [
	'title',
	'subject',
	'creator',
	'keywords',
	'description',
	'lastModifiedBy',
	'revision',
	'created',
	'modified',
	'category',
	'contentStatus',
] as const satisfies readonly (keyof PptxCoreProperties & keyof CoreProperties)[];

/** App fields a caller may override as text; an empty value removes the element. */
const APP_TEXT_OVERRIDES = [
	'application',
	'appVersion',
	'presentationFormat',
	'company',
	'manager',
	'template',
] as const;

/** App fields a caller may override as numbers (truncated; a non-number is ignored). */
const APP_NUMBER_OVERRIDES = [
	'slides',
	'hiddenSlides',
	'notes',
	'totalTime',
	'words',
	'paragraphs',
] as const;

const trimmed = (value: string | undefined): string | undefined => value?.trim() || undefined;

const anyDefined = (record: object): boolean =>
	Object.values(record).some((value) => value !== undefined);

/** Core properties as the pptx model exposes them: trimmed, `undefined` when all are unset. */
export function toPptxCoreProperties(core: CoreProperties): PptxCoreProperties | undefined {
	const result: PptxCoreProperties = {};
	for (const field of CORE_FIELDS) result[field] = trimmed(core[field]);
	return anyDefined(result) ? result : undefined;
}

/** App properties as the pptx model exposes them, `undefined` when all are unset. */
export function toPptxAppProperties(app: AppProperties): PptxAppProperties | undefined {
	const result: PptxAppProperties = {
		application: trimmed(app.application),
		appVersion: trimmed(app.appVersion),
		presentationFormat: trimmed(app.presentationFormat),
		slides: app.slides,
		hiddenSlides: app.hiddenSlides,
		notes: app.notes,
		totalTime: app.totalTime,
		words: app.words,
		paragraphs: app.paragraphs,
		company: trimmed(app.company),
		manager: trimmed(app.manager),
		template: trimmed(app.template),
		hyperlinkBase: trimmed(app.hyperlinkBase),
		docSecurity: app.docSecurity,
		mmClips: app.mmClips,
		scaleCrop: app.scaleCrop,
		linksUpToDate: app.linksUpToDate,
		sharedDoc: app.sharedDoc,
		hyperlinksChanged: app.hyperlinksChanged,
	};
	return anyDefined(result) ? result : undefined;
}

/**
 * The variant types the pptx model reads and writes by name; another type reads as `unknown`
 * and is written as `lpwstr`.
 */
const CUSTOM_TYPES = new Set(['lpwstr', 'i4', 'bool', 'filetime', 'r8', 'i2', 'ui4', 'lpstr']);

/** Custom properties as the pptx model exposes them: every value as text. */
export function toPptxCustomProperties(texts: readonly CustomPropertyText[]): PptxCustomProperty[] {
	return texts.flatMap((entry) => {
		const name = entry.name.trim();
		if (!name) return [];
		const known = CUSTOM_TYPES.has(entry.type);
		return [{ name, value: known ? entry.text : '', type: known ? entry.type : 'unknown' }];
	});
}

/** `core` with the caller's overrides applied: values are trimmed and an empty one is unset. */
export function applyCoreOverrides(
	core: CoreProperties,
	overrides: PptxCoreProperties | undefined,
): CoreProperties {
	const next: CoreProperties = { ...core };
	for (const field of CORE_FIELDS) {
		const value = overrides?.[field];
		if (value === undefined) continue;
		const text = String(value).trim();
		if (text) next[field] = text;
		else delete next[field];
	}
	return next;
}

/** `app` with the caller's text and number overrides applied. */
export function applyAppOverrides(
	app: AppProperties,
	overrides: PptxAppProperties | undefined,
): AppProperties {
	const next: AppProperties = { ...app };
	for (const field of APP_TEXT_OVERRIDES) {
		const value = overrides?.[field];
		if (value === undefined) continue;
		const text = String(value).trim();
		if (text) next[field] = text;
		else delete next[field];
	}
	for (const field of APP_NUMBER_OVERRIDES) {
		const value = Number(overrides?.[field]);
		if (overrides?.[field] !== undefined && Number.isFinite(value)) next[field] = Math.trunc(value);
	}
	return next;
}

/** A pptx custom property as a shared one, keeping its value text exactly. */
export function toCustomProperty(entry: PptxCustomProperty): CustomProperty {
	const type = String(entry.type || 'lpwstr')
		.trim()
		.toLowerCase();
	return customPropertyFromText(
		entry.name,
		CUSTOM_TYPES.has(type) ? type : 'lpwstr',
		String(entry.value ?? ''),
	);
}
