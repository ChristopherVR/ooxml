import { keyToLabel, translationsEn } from 'ooxml-ui/pptx/i18n';
/**
 * Tests for the Angular SmartArt inspector logic.
 *
 * The Angular package's vitest setup has no Angular compiler,
 * so component/TestBed rendering is not exercised here. Instead these tests
 * target the pure helper layer in `smart-art-properties-helpers.ts` plus the
 * `editor-insert.ts` re-exports of the core editing ops. The text pane is the
 * shared `pptx-ui-smartart-text-pane` element, covered in `ooxml-ui` and by
 * `smart-art-properties.component.test.ts`.
 */
import type { PptxSmartArtData, PptxSmartArtNode } from 'pptx-viewer-core';
import { resetSmartArtEditCounter } from 'pptx-viewer-core';
import { beforeEach, describe, expect, it } from 'vitest';

import {
	addSmartArtNode,
	addSmartArtNodeAsChild,
	demoteSmartArtNode,
	promoteSmartArtNode,
	removeSmartArtNode,
	reorderSmartArtNode,
	switchSmartArtLayout,
	SWITCHABLE_LAYOUT_TYPES,
	updateSmartArtNodeText,
} from './editor-insert';
import {
	smartArtColorSchemeLabelKey,
	smartArtLayoutLabelKey,
	smartArtStyleLabelKey,
} from './schema-token-labels';
import {
	currentColorScheme,
	currentLayout,
	currentStyle,
	DEFAULT_COLOR_SCHEME,
	DEFAULT_LAYOUT,
	DEFAULT_STYLE,
	setColorScheme,
	setLayout,
	setStyle,
	SMART_ART_COLOR_SCHEMES,
	SMART_ART_STYLE_OPTIONS,
} from './smart-art-properties-helpers';

// ── Fixtures ─────────────────────────────────────────────────────────────────

function node(id: string, text: string, parentId?: string): PptxSmartArtNode {
	return { id, text, parentId };
}

/** A flat 3-node list diagram. */
function flatData(): PptxSmartArtData {
	return {
		resolvedLayoutType: 'list',
		nodes: [node('a', 'Alpha'), node('b', 'Bravo'), node('c', 'Charlie')],
	};
}

/** A diagram with one parent and two children. */
function nestedData(): PptxSmartArtData {
	return {
		resolvedLayoutType: 'hierarchy',
		nodes: [
			node('root', 'Root'),
			node('child1', 'Child 1', 'root'),
			node('child2', 'Child 2', 'root'),
		],
	};
}

beforeEach(() => {
	resetSmartArtEditCounter();
});

// ── Re-export surface ─────────────────────────────────────────────────────────

describe('editor-insert SmartArt re-exports', () => {
	it('re-exports every editing op from pptx-viewer-core', () => {
		expect(addSmartArtNode).toBeTypeOf('function');
		expect(addSmartArtNodeAsChild).toBeTypeOf('function');
		expect(removeSmartArtNode).toBeTypeOf('function');
		expect(updateSmartArtNodeText).toBeTypeOf('function');
		expect(reorderSmartArtNode).toBeTypeOf('function');
		expect(promoteSmartArtNode).toBeTypeOf('function');
		expect(demoteSmartArtNode).toBeTypeOf('function');
		expect(switchSmartArtLayout).toBeTypeOf('function');
	});

	it('re-exports the switchable layout list', () => {
		expect(SWITCHABLE_LAYOUT_TYPES.length).toBeGreaterThan(0);
		expect(SWITCHABLE_LAYOUT_TYPES).toContain('process');
		expect(SWITCHABLE_LAYOUT_TYPES).toContain('list');
	});
});

// ── Option constants ──────────────────────────────────────────────────────────

describe('option constants', () => {
	it('exposes five colour schemes including colorful1', () => {
		expect(SMART_ART_COLOR_SCHEMES).toContain('colorful1');
		expect(SMART_ART_COLOR_SCHEMES).toHaveLength(5);
	});

	it('exposes the three style intensities', () => {
		expect([...SMART_ART_STYLE_OPTIONS]).toStrictEqual(['flat', 'moderate', 'intense']);
	});
});

// ── Read helpers ──────────────────────────────────────────────────────────────

describe('read helpers', () => {
	it('falls back to defaults when scheme/style/layout are unset', () => {
		const bare: PptxSmartArtData = { nodes: [] };
		expect(currentColorScheme(bare)).toBe(DEFAULT_COLOR_SCHEME);
		expect(currentStyle(bare)).toBe(DEFAULT_STYLE);
		expect(currentLayout(bare)).toBe(DEFAULT_LAYOUT);
	});

	it('reads explicit scheme/style/layout when present', () => {
		const data: PptxSmartArtData = {
			nodes: [],
			colorScheme: 'colorful3',
			style: 'intense',
			resolvedLayoutType: 'cycle',
		};
		expect(currentColorScheme(data)).toBe('colorful3');
		expect(currentStyle(data)).toBe('intense');
		expect(currentLayout(data)).toBe('cycle');
	});
});

// ── Colour scheme / style ─────────────────────────────────────────────────────

describe('setColorScheme / setStyle', () => {
	it('setColorScheme applies the chosen scheme', () => {
		expect(setColorScheme(flatData(), 'monochromatic1').colorScheme).toBe('monochromatic1');
	});

	it('setStyle applies the chosen intensity', () => {
		expect(setStyle(flatData(), 'intense').style).toBe('intense');
	});

	it('setColorScheme and setStyle clear drawingShapes', () => {
		const data: PptxSmartArtData = { ...flatData(), drawingShapes: [] };
		expect(setColorScheme(data, 'colorful2').drawingShapes).toBeUndefined();
		expect(setStyle(data, 'moderate').drawingShapes).toBeUndefined();
	});
});

// ── Layout switch ─────────────────────────────────────────────────────────────

describe('setLayout', () => {
	it('switches to a new layout while preserving nodes', () => {
		const next = setLayout(flatData(), 'cycle');
		expect(next.resolvedLayoutType).toBe('cycle');
		expect(next.nodes.map((n) => n.id)).toStrictEqual(['a', 'b', 'c']);
	});

	it('returns the same object when the layout is unchanged', () => {
		const data = flatData(); // resolvedLayoutType: 'list'
		expect(setLayout(data, 'list')).toBe(data);
	});

	it('can switch through every switchable layout type', () => {
		for (const layout of SWITCHABLE_LAYOUT_TYPES) {
			const next = setLayout(flatData(), layout);
			expect(next.resolvedLayoutType).toBe(layout);
			expect(next.nodes).toHaveLength(3);
		}
	});
});

// ── Wire-token spelling ──────────────────────────────────────────────────────

/**
 * The three option groups in this panel (layout, colour variation, style) used
 * to print the `dgm:` token itself, so the picker offered "colorful1" and
 * "bending" as if they were English words. The values are a file-format
 * contract, so relabelling must leave them alone; these tests hold both ends.
 */
describe('option spelling', () => {
	/** The text a template renders for `key`, resolved the way the host does. */
	function renderedLabel(key: string): string {
		return (translationsEn as Record<string, string | undefined>)[key] ?? keyToLabel(key);
	}

	it('still offers the same five colour variations', () => {
		expect([...SMART_ART_COLOR_SCHEMES]).toStrictEqual([
			'colorful1',
			'colorful2',
			'colorful3',
			'monochromatic1',
			'monochromatic2',
		]);
	});

	it('spells each colour variation', () => {
		expect(
			SMART_ART_COLOR_SCHEMES.map((s) => renderedLabel(smartArtColorSchemeLabelKey(s))),
		).toStrictEqual([
			'Colourful 1',
			'Colourful 2',
			'Colourful 3',
			'Monochromatic 1',
			'Monochromatic 2',
		]);
	});

	it('still offers the same three style intensities, spelled as words', () => {
		expect([...SMART_ART_STYLE_OPTIONS]).toStrictEqual(['flat', 'moderate', 'intense']);
		expect(
			SMART_ART_STYLE_OPTIONS.map((s) => renderedLabel(smartArtStyleLabelKey(s))),
		).toStrictEqual(['Flat', 'Moderate', 'Intense']);
	});

	it('spells every switchable layout, with none falling back to its token', () => {
		const labels = SWITCHABLE_LAYOUT_TYPES.map((layout) =>
			renderedLabel(smartArtLayoutLabelKey(layout)),
		);

		expect(labels).toHaveLength(SWITCHABLE_LAYOUT_TYPES.length);
		for (const [index, layout] of SWITCHABLE_LAYOUT_TYPES.entries()) {
			expect(labels[index]).not.toBe(layout);
		}
		expect(labels).toContain('List');
		expect(labels).toContain('Hierarchy');
	});
});
