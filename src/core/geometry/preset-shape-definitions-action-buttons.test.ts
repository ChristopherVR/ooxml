import { describe, expect, it } from 'vitest';

import { ACTION_BUTTON_PRESET_DEFINITIONS } from './preset-shape-definitions-action-buttons';
import { evaluatePresetShape } from './preset-shape-evaluator';

// The 12 ECMA-376 actionButton* presets: a flat face plus the glyph as shaded
// sub-paths, measured on PowerPoint 16 (darken / darkenLess / lighten fills,
// outlines stroked, no bevel).

const REQUIRED_SHAPES = [
	'actionButtonBlank',
	'actionButtonHome',
	'actionButtonHelp',
	'actionButtonInformation',
	'actionButtonForwardNext',
	'actionButtonBackPrevious',
	'actionButtonEnd',
	'actionButtonBeginning',
	'actionButtonReturn',
	'actionButtonDocument',
	'actionButtonSound',
	'actionButtonMovie',
] as const;

describe('preset shape geometry: action buttons', () => {
	it('contains every required actionButton preset (12 shapes)', () => {
		for (const name of REQUIRED_SHAPES) {
			expect(ACTION_BUTTON_PRESET_DEFINITIONS[name], `missing preset ${name}`).toBeDefined();
		}
		expect(Object.keys(ACTION_BUTTON_PRESET_DEFINITIONS)).toHaveLength(12);
	});

	it('shape names match their keys and gd args mirror the formula tokens', () => {
		for (const [key, def] of Object.entries(ACTION_BUTTON_PRESET_DEFINITIONS)) {
			expect(def.name).toBe(key);
			for (const g of def.gdLst ?? []) {
				expect(g.args, `${key}/${g.name}`).toHaveLength(g.formula.trim().split(/\s+/).length - 1);
			}
		}
	});

	it('has NO adjustable guides (COM-verified: Shape.Adjustments.Count is 0)', () => {
		for (const name of REQUIRED_SHAPES) {
			expect(ACTION_BUTTON_PRESET_DEFINITIONS[name]?.avLst).toBeUndefined();
		}
	});

	it('draws a flat face and no bevel: Blank is the face alone', () => {
		expect(ACTION_BUTTON_PRESET_DEFINITIONS.actionButtonBlank!.pathLst).toHaveLength(1);
	});

	it('draws each glyph as shaded sub-paths on the face', () => {
		const modes = (n: string) =>
			evaluatePresetShape(n, 150, 90)!.paths.map((p) => p.fill ?? 'norm');
		expect(modes('actionButtonForwardNext')).toEqual(['norm', 'darken']);
		expect(modes('actionButtonHome')).toContain('darkenLess');
		expect(modes('actionButtonInformation')).toContain('lighten');
		expect(modes('actionButtonDocument')).toEqual(['norm', 'darkenLess', 'darken']);
	});

	it('puts the forward triangle at +-3/8 ss around the centre', () => {
		const d = evaluatePresetShape('actionButtonForwardNext', 150, 90)!.paths[1]!.d;
		expect(d.startsWith('M 41.25 11.25 L 108.75 45 L 41.25 78.75')).toBe(true);
	});
});
