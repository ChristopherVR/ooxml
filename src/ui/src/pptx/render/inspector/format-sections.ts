/**
 * The selection-aware Format pane: which collapsible sections the inspector's
 * Properties tab shows for the current selection, in which order, and which of
 * them start open. Written once here; every binding renders the returned list
 * and puts its existing panels inside the matching section.
 *
 * Section ids are stable (hosts hide and reorder them through
 * `inspector.hiddenSections` / `inspector.sectionOrder`, and bindings tag each
 * section with {@link FORMAT_SECTION_ATTRIBUTE}). Which panels go in which
 * section is recorded in {@link FORMAT_PANEL_SECTIONS} so the five bindings
 * place them the same way.
 *
 * @module render/inspector/format-sections
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getShapeType, hasShapeProperties, hasTextProperties } from 'ooxml-core/pptx';

import { shouldShowAccessibilitySection } from '../element-non-visual-description';
import { SHAPE_EFFECT_SECTIONS } from './shape-effect-sections';

/** What the Properties tab is formatting. */
export type FormatSelectionCategory =
	| 'slide'
	| 'text'
	| 'shape'
	| 'line'
	| 'picture'
	| 'table'
	| 'chart'
	| 'smartArt'
	| 'media'
	| 'group'
	| 'object'
	| 'mixed';

/** Every Format pane section, in the canonical order a mixed selection uses. */
export const FORMAT_SECTION_IDS = [
	'slide',
	'font',
	'paragraph',
	'picture',
	'table',
	'chart',
	'smartArt',
	'media',
	'object',
	'fill',
	'outline',
	'effects',
	'arrange',
	'sizePosition',
	'altText',
	'action',
] as const;

export type FormatSectionId = (typeof FORMAT_SECTION_IDS)[number];

/** The attribute a binding puts on each section element (`data-pptx-format-section="fill"`). */
export const FORMAT_SECTION_ATTRIBUTE = 'data-pptx-format-section';

/** What a binding draws for one section. */
export interface FormatSectionView {
	id: FormatSectionId;
	/** Translation key of the heading. */
	labelKey: string;
	/** Whether the section starts open (before the user toggles it). */
	defaultOpen: boolean;
	/** Line selections only: the Outline section also shows arrowheads. */
	arrowheads?: boolean;
}

/** Host ordering and hiding (from the resolved customization). */
export interface FormatSectionOptions {
	hidden?: Iterable<FormatSectionId>;
	/** Ids to move to the front, in this order; the rest keep their place. */
	order?: readonly FormatSectionId[];
}

const LABEL_KEYS: Record<FormatSectionId, string> = {
	slide: 'pptx.formatPane.slide',
	font: 'pptx.formatPane.font',
	paragraph: 'pptx.formatPane.paragraph',
	picture: 'pptx.formatPane.picture',
	table: 'pptx.inspector.table',
	chart: 'pptx.inspector.chart',
	smartArt: 'pptx.inspector.smartArt',
	media: 'pptx.inspector.media',
	object: 'pptx.formatPane.object',
	fill: 'pptx.inspector.fill',
	outline: 'pptx.formatPane.outline',
	effects: 'pptx.inspector.effects',
	arrange: 'pptx.inspector.arrange',
	sizePosition: 'pptx.formatPane.sizePosition',
	altText: 'pptx.formatPane.altText',
	action: 'pptx.formatPane.action',
};

type Entry = readonly [FormatSectionId, boolean];

const TAIL: readonly Entry[] = [
	['arrange', false],
	['sizePosition', false],
	['altText', false],
	['action', false],
];

/** Each category's sections and whether each starts open. */
const CATEGORY_SECTIONS: Record<Exclude<FormatSelectionCategory, 'mixed'>, readonly Entry[]> = {
	slide: [['slide', true]],
	text: [
		['font', true],
		['paragraph', true],
		['fill', false],
		['outline', false],
		['effects', false],
		...TAIL,
	],
	shape: [
		['fill', true],
		['outline', true],
		['effects', false],
		['font', false],
		['paragraph', false],
		...TAIL,
	],
	line: [['outline', true], ['effects', false], ['font', false], ['paragraph', false], ...TAIL],
	// A picture keeps its Fill (the area behind a transparent picture) reachable, closed.
	picture: [['picture', true], ['fill', false], ['outline', false], ['effects', false], ...TAIL],
	// Cell text formatting lives in the Table section's cell formatting; no table-wide text panel yet.
	table: [['table', true], ...TAIL],
	chart: [['chart', true], ...TAIL],
	smartArt: [['smartArt', true], ...TAIL],
	media: [['media', true], ...TAIL],
	group: [
		['arrange', true],
		['sizePosition', true],
		['action', false],
	],
	object: [['object', true], ...TAIL],
};

/** Whether `element` is drawn as a line (a connector or a `line` preset). */
export function isLineFormatElement(element: PptxElement): boolean {
	if (element.type === 'connector') {
		return true;
	}
	return (
		hasShapeProperties(element) && !!element.shapeType && getShapeType(element.shapeType) === 'line'
	);
}

/** The category of one element. */
export function formatElementCategory(element: PptxElement): FormatSelectionCategory {
	switch (element.type) {
		case 'text':
			return 'text';
		case 'image':
		case 'picture':
			return 'picture';
		case 'table':
			return 'table';
		case 'chart':
			return 'chart';
		case 'smartArt':
			return 'smartArt';
		case 'media':
			return 'media';
		case 'group':
			return 'group';
		case 'connector':
			return 'line';
		case 'shape':
			return isLineFormatElement(element) ? 'line' : 'shape';
		default:
			return 'object';
	}
}

/** The category of a selection: `slide` when empty, `mixed` when the kinds differ. */
export function resolveFormatCategory(
	elements: readonly (PptxElement | null | undefined)[],
): FormatSelectionCategory {
	const present = elements.filter((el): el is PptxElement => !!el);
	if (present.length === 0) {
		return 'slide';
	}
	const first = formatElementCategory(present[0]!);
	if (present.length === 1) {
		return first;
	}
	return present.every((el) => formatElementCategory(el) === first) ? first : 'mixed';
}

/** Whether a shape already carries a Format Shape effect (Effects then starts open). */
export function hasFormatShapeEffect(element: PptxElement): boolean {
	if (!hasShapeProperties(element)) {
		return false;
	}
	const style = element.shapeStyle;
	return SHAPE_EFFECT_SECTIONS.some((section) => section.toggle?.isOn(style) ?? false);
}

function entriesFor(element: PptxElement): readonly Entry[] {
	const category = formatElementCategory(element);
	const entries = CATEGORY_SECTIONS[category as Exclude<FormatSelectionCategory, 'mixed'>];
	// Text sections only for elements that carry text; shape sections only with a shape style.
	return entries.filter(([id]) => {
		if (id === 'font' || id === 'paragraph') {
			return hasTextProperties(element);
		}
		if (id === 'fill' || id === 'outline' || id === 'effects') {
			return hasShapeProperties(element);
		}
		if (id === 'altText') {
			// A picture's alt text lives in its Picture section.
			return shouldShowAccessibilitySection(element);
		}
		return true;
	});
}

function view(id: FormatSectionId, defaultOpen: boolean, line: boolean): FormatSectionView {
	const labelKey = id === 'outline' && line ? 'pptx.inspector.line' : LABEL_KEYS[id];
	return line && id === 'outline'
		? { id, labelKey, defaultOpen, arrowheads: true }
		: { id, labelKey, defaultOpen };
}

function applyOptions(
	views: FormatSectionView[],
	options: FormatSectionOptions | undefined,
): FormatSectionView[] {
	const hidden = new Set(options?.hidden ?? []);
	const visible = views.filter((v) => !hidden.has(v.id));
	const order = options?.order ?? [];
	if (order.length === 0) {
		return visible;
	}
	const rank = (id: FormatSectionId): number => {
		const at = order.indexOf(id);
		return at === -1 ? order.length : at;
	};
	return visible
		.map((v, index) => ({ v, index }))
		.sort((a, b) => rank(a.v.id) - rank(b.v.id) || a.index - b.index)
		.map(({ v }) => v);
}

/**
 * The sections to show for `elements` (empty or nulls: the slide). A
 * multi-selection shows only the sections every element supports, in the
 * category order of a uniform selection (canonical order when mixed), and
 * never Size and position.
 */
export function resolveFormatSections(
	elements: readonly (PptxElement | null | undefined)[],
	options?: FormatSectionOptions,
): FormatSectionView[] {
	const present = elements.filter((el): el is PptxElement => !!el);
	if (present.length === 0) {
		return applyOptions(
			CATEGORY_SECTIONS.slide.map(([id, open]) => view(id, open, false)),
			options,
		);
	}
	const line = present.every(isLineFormatElement);
	if (present.length === 1) {
		const el = present[0]!;
		const effect = hasFormatShapeEffect(el);
		return applyOptions(
			entriesFor(el).map(([id, open]) => view(id, open || (id === 'effects' && effect), line)),
			options,
		);
	}
	const category = resolveFormatCategory(present);
	const sets = present.map((el) => new Map(entriesFor(el)));
	const base =
		category === 'mixed'
			? FORMAT_SECTION_IDS.map((id): Entry => [id, sets[0]!.get(id) ?? false])
			: entriesFor(present[0]!);
	const effect = present.every(hasFormatShapeEffect);
	return applyOptions(
		base
			.filter(([id]) => id !== 'sizePosition' && sets.every((set) => set.has(id)))
			.map(([id, open]) => view(id, open || (id === 'effects' && effect), line)),
		options,
	);
}
