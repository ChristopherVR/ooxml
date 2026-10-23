/**
 * The ribbon galleries hand every binding a catalogue name as a dictionary
 * key plus parameters. This suite builds each gallery (the SmartArt Layouts
 * gallery after its lazy library has loaded), resolves every tile name and
 * heading through each shipped locale exactly as the shared gallery element
 * does, and fails when an English catalogue name comes back unchanged.
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	translationsDe,
	translationsEs,
	translationsFr,
	translationsZhCN,
} from '../../../../../../viewers/pptx/packages/locales/src';
import { translationsEn } from '../../i18n';
import { BUILTIN_SMARTART_LAYOUTS } from '../../../../../core/diagram/builtin/catalog';
import { smartArtLayoutNameKey } from './smartart-layouts-gallery';
import { galleryItemLabel } from './gallery-view';
import { buildRibbonGallery, RIBBON_GALLERY_IDS } from './gallery-registry';
import type { RibbonGalleryDescriptor, RibbonGalleryId } from './gallery-types';

const box = { x: 0, y: 0, width: 200, height: 100 };
const textShape = {
	...box,
	id: 's1',
	type: 'shape',
	shapeType: 'rect',
	shapeStyle: {},
	text: 'Text',
	textSegments: [{ text: 'Text', style: {} }],
} as unknown as PptxElement;
const picture = { ...box, id: 'p1', type: 'picture', shapeType: 'rect', shapeStyle: {} };
const chart = {
	...box,
	id: 'c1',
	type: 'chart',
	chartData: {
		chartType: 'bar',
		barDirection: 'col',
		categories: ['a'],
		series: [{ name: 'S', values: [1] }],
	},
};
const smartArt = {
	...box,
	id: 'sa1',
	type: 'smartArt',
	smartArtData: { nodes: [{ id: 'n1', text: 'One' }] },
};
const table = {
	...box,
	id: 't1',
	type: 'table',
	tableData: { columnWidths: [1], rows: [{ cells: [{ text: 'A' }] }] },
};

function elementFor(id: RibbonGalleryId): PptxElement | null {
	if (id.startsWith('picture')) return picture as unknown as PptxElement;
	if (id.startsWith('chart')) return chart as unknown as PptxElement;
	if (id.startsWith('smartArt')) return smartArt as unknown as PptxElement;
	if (id.startsWith('table')) return table as unknown as PptxElement;
	return id.startsWith('theme') ? null : textShape;
}

const locales: Record<string, Record<string, string>> = {
	de: translationsDe,
	es: translationsEs,
	fr: translationsFr,
	'zh-CN': translationsZhCN,
};

/** Fill `{{name}}` placeholders the way every binding's translator does. */
function translatorFor(dictionary: Record<string, string>) {
	return (key: string, params?: Readonly<Record<string, string | number>>): string => {
		const value = dictionary[key];
		if (value === undefined) return key;
		return value.replace(/\{\{(\w+)\}\}/gu, (_, name: string) => String(params?.[name] ?? ''));
	};
}

/** Names that are spelled the same in English and the locale (shared loanwords and brand names). */
const SAME_SPELLING =
	/^(Office|Venn|Matrix|Chevron|Monochrome|Pastel|Perspective|Corrections|Color|Sepia|Horizontal|Parallel|Oblique|Angle|Layouts?|Orange|Median|Marquee|Slipstream|Violet|Violet II|Aspect|Style|Triangle)(?![A-Za-z])/u;

async function build(id: RibbonGalleryId): Promise<RibbonGalleryDescriptor> {
	const context = {
		element: elementFor(id),
		themeColorMap: { dk1: '#000000', lt1: '#FFFFFF', accent1: '#156082', accent2: '#E97132' },
	};
	return buildRibbonGallery(id, context);
}

describe('ribbon gallery catalogue names in every locale', () => {
	for (const [locale, dictionary] of Object.entries(locales)) {
		const translate = translatorFor(dictionary);
		it(`${locale} shows no English catalogue name`, async () => {
			const english = translatorFor(translationsEn);
			const untranslated: string[] = [];
			for (const id of RIBBON_GALLERY_IDS) {
				// Font names are product names, spelled the same in every language.
				if (id === 'themeFonts') continue;
				const descriptor = await build(id);
				for (const section of descriptor.sections) {
					for (const item of section.items) {
						const shown = galleryItemLabel(item, translate);
						const base = galleryItemLabel(item, english);
						if (shown === base && /[A-Za-z]{3}/u.test(base) && !SAME_SPELLING.test(base)) {
							untranslated.push(`${id}: ${base}`);
						}
					}
				}
			}
			expect(untranslated).toStrictEqual([]);
		}, 60_000);
	}

	it.each(Object.entries(locales))(
		'%s names every built-in SmartArt layout',
		(_locale, dictionary) => {
			// The layout library itself loads lazily; its catalogue is the list of names to cover.
			const missing: string[] = [];
			const english: string[] = [];
			for (const { title } of BUILTIN_SMARTART_LAYOUTS) {
				const key = smartArtLayoutNameKey(title);
				if (!translationsEn[key]) missing.push(key);
				else if (!dictionary[key]) missing.push(key);
				else if (dictionary[key] === translationsEn[key] && !SAME_SPELLING.test(title.trim())) {
					english.push(title);
				}
			}
			expect(missing).toStrictEqual([]);
			expect(english).toStrictEqual([]);
		},
	);
});
