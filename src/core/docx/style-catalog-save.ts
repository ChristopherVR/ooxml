import type JSZip from 'jszip';
import { DEFAULT_STYLES_XML, defaultStyleCatalogs } from './default-styles';
import { orderChildren } from './element-order';
import { writeLatentStyles } from './latent-styles';
import type { DocumentModel } from './model';
import { documentPartPath, parseRelationships } from './package-parts';
import {
	patchStyleParagraph,
	patchStyleRun,
	setStyleAttribute,
	setStyleValue,
	styleChild,
} from './style-formatting-package';
import type { StyleMetadata } from './style-metadata';
import { patchTableStyle } from './style-table-write';
import {
	buildXml,
	children,
	elements,
	first,
	getW,
	makeW,
	parseXml,
	WORD_NS,
	type XmlElement,
} from './xml';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';

const STYLE_ORDER = [
	'name',
	'aliases',
	'basedOn',
	'next',
	'link',
	'autoRedefine',
	'hidden',
	'uiPriority',
	'semiHidden',
	'unhideWhenUsed',
	'qFormat',
	'locked',
	'personal',
	'personalCompose',
	'personalReply',
	'rsid',
	'pPr',
	'rPr',
	'tblPr',
	'trPr',
	'tcPr',
	'tblStylePr',
	'extLst',
];
const STYLES_PATH = 'word/styles.xml';

type Catalogs = Pick<
	DocumentModel,
	'paragraphStyles' | 'characterStyles' | 'tableStyles' | 'latentStyles'
>;
type StyleKind = 'paragraph' | 'character' | 'table';

function catalogStyles(catalogs: Catalogs, kind: StyleKind) {
	const entries = new Map<string, StyleMetadata>();
	if (kind === 'table') {
		for (const style of Object.values(catalogs.tableStyles?.styles ?? {}))
			entries.set(style.id, style);
		return entries;
	}
	for (const style of Object.values(catalogs.characterStyles?.styles ?? {}))
		if (style.type === kind) entries.set(style.id, style);
	if (kind === 'paragraph')
		for (const style of Object.values(catalogs.paragraphStyles?.styles ?? {}))
			entries.set(style.id, style);
	return entries;
}

function patchMetadata(
	node: XmlElement,
	next: StyleMetadata,
	prior: StyleMetadata | undefined,
	nextRun: StyleMetadata | undefined,
	priorRun: StyleMetadata | undefined,
): void {
	const value = <K extends keyof StyleMetadata>(key: K) =>
		next[key] !== prior?.[key]
			? next[key]
			: nextRun?.[key] !== priorRun?.[key]
				? nextRun?.[key]
				: next[key];
	for (const [key, local] of [
		['isDefault', 'default'],
		['customStyle', 'customStyle'],
	] as const) {
		const after = value(key);
		if (after !== prior?.[key]) setStyleAttribute(node, local, after);
	}
	for (const [key, local] of [
		['name', 'name'],
		['basedOn', 'basedOn'],
		['uiPriority', 'uiPriority'],
		['quickFormat', 'qFormat'],
	] as const) {
		const after = value(key);
		if (after !== prior?.[key]) setStyleValue(node, local, after);
	}
}

function patchCatalog(root: XmlElement, model: Catalogs, base: Catalogs, kind: StyleKind): void {
	const next = catalogStyles(model, kind);
	const prior = catalogStyles(base, kind);
	const existing = new Map(
		children(root, 'style')
			.filter((node) => getW(node, 'type') === kind)
			.map((node) => [getW(node, 'styleId'), node]),
	);
	for (const id of prior.keys()) {
		const node = existing.get(id);
		if (!next.has(id) && node) root.removeChild(node);
	}
	for (const [id, after] of next) {
		const before = prior.get(id);
		const nextRun = model.characterStyles?.styles[id];
		const priorRun = base.characterStyles?.styles[id];
		if (
			JSON.stringify(after) === JSON.stringify(before) &&
			(kind !== 'paragraph' || JSON.stringify(nextRun) === JSON.stringify(priorRun))
		)
			continue;
		const original = existing.get(id);
		const node = original ?? makeW(root.ownerDocument, 'style');
		if (!original) {
			setStyleAttribute(node, 'styleId', id);
			setStyleAttribute(node, 'type', kind);
			root.insertBefore(node, elements(root).find((child) => child.localName === 'extLst') ?? null);
		}
		patchMetadata(
			node,
			after,
			original ? before : undefined,
			kind === 'paragraph' ? nextRun : undefined,
			original && kind === 'paragraph' ? priorRun : undefined,
		);
		if (kind === 'table') {
			const table = model.tableStyles?.styles[id];
			if (table) patchTableStyle(node, table, original ? base.tableStyles?.styles[id] : undefined);
		} else {
			const p = kind === 'paragraph' ? model.paragraphStyles?.styles[id] : undefined;
			const oldP = original && kind === 'paragraph' ? base.paragraphStyles?.styles[id] : undefined;
			const oldRun = original ? priorRun : undefined;
			if (p?.next !== oldP?.next) setStyleValue(node, 'next', p?.next);
			const link = p?.linkedStyle !== oldP?.linkedStyle ? p?.linkedStyle : nextRun?.linkedStyle;
			if (link !== (oldP?.linkedStyle ?? oldRun?.linkedStyle)) setStyleValue(node, 'link', link);
			patchStyleParagraph(node, p?.formatting, oldP?.formatting, p?.numbering, oldP?.numbering);
			patchStyleRun(
				node,
				nextRun?.formatting,
				oldRun?.formatting,
				nextRun?.eastAsiaLanguage,
				oldRun?.eastAsiaLanguage,
			);
		}
		orderChildren(node, STYLE_ORDER);
	}
}

function patchDefaults(root: XmlElement, model: Catalogs, base: Catalogs): void {
	const p = model.paragraphStyles?.docDefaults;
	const oldP = base.paragraphStyles?.docDefaults;
	const r = model.characterStyles;
	const oldR = base.characterStyles;
	const paragraphChanged = JSON.stringify(p ?? {}) !== JSON.stringify(oldP ?? {});
	const runChanged =
		JSON.stringify(r?.docDefaults ?? {}) !== JSON.stringify(oldR?.docDefaults ?? {}) ||
		r?.docDefaultsEastAsiaLanguage !== oldR?.docDefaultsEastAsiaLanguage;
	if (!paragraphChanged && !runChanged) return;
	const defaults = first(root, 'docDefaults') ?? makeW(root.ownerDocument, 'docDefaults');
	if (!defaults.parentNode) root.insertBefore(defaults, root.firstChild);
	if (runChanged)
		patchStyleRun(
			styleChild(defaults, 'rPrDefault'),
			r?.docDefaults,
			oldR?.docDefaults,
			r?.docDefaultsEastAsiaLanguage,
			oldR?.docDefaultsEastAsiaLanguage,
		);
	if (paragraphChanged) patchStyleParagraph(styleChild(defaults, 'pPrDefault'), p, oldP);
	orderChildren(defaults, ['rPrDefault', 'pPrDefault']);
}

function patchStyles(xml: string, model: Catalogs, base: Catalogs): string {
	const doc = parseXml(xml);
	const root = doc.documentElement;
	patchDefaults(root, model, base);
	writeLatentStyles(root, model.latentStyles, base.latentStyles);
	for (const kind of ['paragraph', 'character', 'table'] as const)
		patchCatalog(root, model, base, kind);
	return buildXml(doc);
}

export async function applyStyleCatalogs(
	zip: JSZip,
	model: DocumentModel,
	base?: DocumentModel,
): Promise<void> {
	const keys = ['paragraphStyles', 'characterStyles', 'tableStyles', 'latentStyles'] as const;
	if (!keys.some((key) => JSON.stringify(model[key]) !== JSON.stringify(base?.[key]))) return;
	const rels = parseRelationships(await zip.file('word/_rels/document.xml.rels')?.async('string'));
	const relatedPath = documentPartPath(rels, 'styles', '');
	const path = relatedPath || STYLES_PATH;
	const file = zip.file(path);
	const defaults = !base && model.paragraphStyles ? defaultStyleCatalogs() : undefined;
	const next = defaults ? { ...defaults, ...model } : model;
	const source =
		(await file?.async('string')) ??
		(defaults ? DEFAULT_STYLES_XML : `<w:styles xmlns:w="${WORD_NS}"/>`);
	zip.file(path, patchStyles(source, next, base ?? defaults ?? {}));
	if (!file) {
		await ensureContentTypeOverride(
			zip,
			path,
			'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
		);
		if (!relatedPath)
			await ensureDocumentRelationship(
				zip,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
				'styles.xml',
			);
	}
	// Word can load this parallel catalog when it is present. Apply the same edits to its source.
	const effects = zip.file(
		documentPartPath(rels, 'stylesWithEffects', 'word/stylesWithEffects.xml'),
	);
	if (effects) zip.file(effects.name, patchStyles(await effects.async('string'), next, base ?? {}));
}
