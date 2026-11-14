import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	formatMixedValue,
	formatSelectionElements,
	isFormatSectionMixed,
} from './format-section-mixed';
import { FORMAT_PANEL_SECTIONS } from './format-panel-sections';
import {
	FORMAT_SECTION_IDS,
	resolveFormatCategory,
	resolveFormatSections,
} from './format-sections';

const el = (over: Record<string, unknown>): PptxElement =>
	({ id: 'e', x: 0, y: 0, width: 10, height: 10, ...over }) as unknown as PptxElement;

const ids = (elements: PptxElement[]) => resolveFormatSections(elements).map((s) => s.id);
const open = (elements: PptxElement[]) =>
	resolveFormatSections(elements)
		.filter((s) => s.defaultOpen)
		.map((s) => s.id);

describe('resolveFormatCategory', () => {
	it('names each selection kind', () => {
		expect(resolveFormatCategory([])).toBe('slide');
		expect(resolveFormatCategory([null])).toBe('slide');
		expect(resolveFormatCategory([el({ type: 'text' })])).toBe('text');
		expect(resolveFormatCategory([el({ type: 'shape', shapeType: 'rect' })])).toBe('shape');
		expect(resolveFormatCategory([el({ type: 'shape', shapeType: 'Line' })])).toBe('line');
		expect(resolveFormatCategory([el({ type: 'connector' })])).toBe('line');
		expect(resolveFormatCategory([el({ type: 'picture' })])).toBe('picture');
		expect(resolveFormatCategory([el({ type: 'table' })])).toBe('table');
		expect(resolveFormatCategory([el({ type: 'chart' })])).toBe('chart');
		expect(resolveFormatCategory([el({ type: 'smartArt' })])).toBe('smartArt');
		expect(resolveFormatCategory([el({ type: 'media' })])).toBe('media');
		expect(resolveFormatCategory([el({ type: 'group' })])).toBe('group');
		expect(resolveFormatCategory([el({ type: 'ole' })])).toBe('object');
		expect(resolveFormatCategory([el({ type: 'text' }), el({ type: 'text' })])).toBe('text');
		expect(resolveFormatCategory([el({ type: 'text' }), el({ type: 'picture' })])).toBe('mixed');
	});
});

describe('resolveFormatSections', () => {
	it('shows the slide section with nothing selected', () => {
		expect(resolveFormatSections([])).toEqual([
			{ id: 'slide', labelKey: 'pptx.formatPane.slide', defaultOpen: true },
		]);
	});

	it('opens Font and Paragraph for a text box', () => {
		const text = [el({ type: 'text' })];
		expect(ids(text).slice(0, 7)).toEqual([
			'font',
			'paragraph',
			'fill',
			'outline',
			'effects',
			'arrange',
			'sizePosition',
		]);
		expect(open(text)).toEqual(['font', 'paragraph']);
	});

	it('opens Fill and Outline for a shape, and Effects once it has one', () => {
		const plain = el({ type: 'shape', shapeType: 'rect', shapeStyle: {} });
		expect(ids([plain]).slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(open([plain])).toEqual(['fill', 'outline']);
		const glowing = el({
			type: 'shape',
			shapeType: 'rect',
			shapeStyle: { glowColor: '#ff0000', glowRadius: 4 },
		});
		expect(open([glowing])).toEqual(['fill', 'outline', 'effects']);
	});

	it('labels Outline as Line with arrowheads for lines only', () => {
		const [line] = resolveFormatSections([el({ type: 'connector' })]);
		expect(line).toEqual({
			id: 'outline',
			labelKey: 'pptx.inspector.line',
			defaultOpen: true,
			arrowheads: true,
		});
		const shapeOutline = resolveFormatSections([el({ type: 'shape', shapeType: 'rect' })]).find(
			(s) => s.id === 'outline',
		);
		expect(shapeOutline).toEqual({
			id: 'outline',
			labelKey: 'pptx.formatPane.outline',
			defaultOpen: true,
		});
		expect(ids([el({ type: 'connector' })])).not.toContain('fill');
	});

	it('gives pictures, tables, charts and groups their own lists', () => {
		const picture = [el({ type: 'picture' })];
		expect(ids(picture)).toEqual([
			'picture',
			'fill',
			'outline',
			'effects',
			'arrange',
			'sizePosition',
			'action',
		]);
		expect(open(picture)).toEqual(['picture']);
		const table = [el({ type: 'table' })];
		expect(ids(table).slice(0, 3)).toEqual(['table', 'arrange', 'sizePosition']);
		expect(open(table)).toEqual(['table']);
		expect(ids([el({ type: 'chart' })]).slice(0, 3)).toEqual(['chart', 'arrange', 'sizePosition']);
		expect(open([el({ type: 'group' })])).toEqual(['arrange', 'sizePosition']);
		expect(ids([el({ type: 'ole' })])[0]).toBe('object');
		expect(ids([el({ type: 'ink' })])).not.toContain('object');
	});

	it('keeps only shared sections for a multi-selection, without Size and position', () => {
		const shapes = [
			el({ type: 'shape', shapeType: 'rect' }),
			el({ type: 'shape', shapeType: 'ellipse' }),
		];
		expect(ids(shapes)).not.toContain('sizePosition');
		expect(ids(shapes).slice(0, 2)).toEqual(['fill', 'outline']);
		const mixed = [el({ type: 'shape', shapeType: 'rect' }), el({ type: 'chart' })];
		expect(ids(mixed)).toEqual(['arrange', 'altText', 'action']);
	});

	it('applies host hiding and ordering', () => {
		const text = [el({ type: 'text' })];
		const views = resolveFormatSections(text, {
			hidden: ['fill'],
			order: ['sizePosition', 'arrange'],
		});
		expect(views.map((v) => v.id).slice(0, 4)).toEqual([
			'sizePosition',
			'arrange',
			'font',
			'paragraph',
		]);
		expect(views.map((v) => v.id)).not.toContain('fill');
	});

	it('maps every panel to a known section', () => {
		for (const section of Object.values(FORMAT_PANEL_SECTIONS)) {
			expect(FORMAT_SECTION_IDS).toContain(section);
		}
	});
});

describe('mixed values', () => {
	it('reports a field whose values differ', () => {
		const a = el({ type: 'shape', shapeStyle: { fillColor: '#ff0000' } });
		const b = el({ type: 'shape', shapeStyle: { fillColor: '#00ff00' } });
		const c = el({ type: 'shape', shapeStyle: { fillColor: '#ff0000' } });
		const fill = (e: PptxElement) =>
			(e as { shapeStyle?: { fillColor?: string } }).shapeStyle?.fillColor;
		expect(formatMixedValue([a, b], fill)).toEqual({ mixed: true });
		expect(formatMixedValue([a, c], fill)).toEqual({ mixed: false, value: '#ff0000' });
		expect(isFormatSectionMixed([a, b], 'fill')).toBe(true);
		expect(isFormatSectionMixed([a, c], 'fill')).toBe(false);
		expect(isFormatSectionMixed([a], 'fill')).toBe(false);
		expect(isFormatSectionMixed([a, b], 'outline')).toBe(false);
	});
});

describe('formatSelectionElements', () => {
	it('puts the primary first and resolves the other ids', () => {
		const a = el({ id: 'a', type: 'text' });
		const b = el({ id: 'b', type: 'text' });
		expect(formatSelectionElements(null, ['a'], [a])).toEqual([]);
		expect(formatSelectionElements(b, ['a', 'b', 'missing'], [a, b])).toEqual([b, a]);
	});
});
