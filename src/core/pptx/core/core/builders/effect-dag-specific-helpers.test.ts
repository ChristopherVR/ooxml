import { describe, it, expect } from 'vitest';

import type { ShapeStyle, XmlObject } from '../../types';
import { extractDagDuotone } from './effect-dag-specific-helpers';
import type { DagSpecificContext } from './effect-dag-specific-helpers';

// Helper to build a minimal DagSpecificContext.
function makeContext(overrides: Partial<DagSpecificContext> = {}): DagSpecificContext {
	return {
		parseColor: (colorNode: XmlObject | undefined) => {
			if (!colorNode) {
				return undefined;
			}
			const val = colorNode['@_val'];
			if (val) {
				return `#${val}`;
			}
			return undefined;
		},
		ensureArray: (value: unknown): XmlObject[] => {
			if (Array.isArray(value)) {
				return value as XmlObject[];
			}
			if (value === undefined || value === null) {
				return [];
			}
			return [value] as XmlObject[];
		},
		...overrides,
	};
}

describe('extractDagDuotone', () => {
	it('extracts two sRGB colors from duotone', () => {
		const dag: XmlObject = {
			'a:duotone': {
				'a:srgbClr': [{ '@_val': '000000' }, { '@_val': 'FFFFFF' }],
			},
		};
		const style: Partial<ShapeStyle> = {};
		extractDagDuotone(dag, style, makeContext());
		expect(style.dagDuotone).toStrictEqual({
			color1: '#000000',
			color2: '#FFFFFF',
		});
	});

	it('does nothing when fewer than 2 colors', () => {
		const dag: XmlObject = {
			'a:duotone': {
				'a:srgbClr': { '@_val': 'FF0000' },
			},
		};
		const style: Partial<ShapeStyle> = {};
		extractDagDuotone(dag, style, makeContext());
		expect(style.dagDuotone).toBeUndefined();
	});

	it('falls back to defaults when parseColor returns undefined', () => {
		const dag: XmlObject = {
			'a:duotone': {
				'a:schemeClr': [{ '@_val': 'accent1' }, { '@_val': 'accent2' }],
			},
		};
		const style: Partial<ShapeStyle> = {};
		const ctx = makeContext({
			parseColor: () => undefined,
		});
		extractDagDuotone(dag, style, ctx);
		expect(style.dagDuotone).toStrictEqual({
			color1: '#000000',
			color2: '#ffffff',
		});
	});

	it('does nothing when a:duotone is absent', () => {
		const dag: XmlObject = {};
		const style: Partial<ShapeStyle> = {};
		extractDagDuotone(dag, style, makeContext());
		expect(style.dagDuotone).toBeUndefined();
	});
});

// ---------------------------------------------------------------------------
// extractDagFillOverlay
// ---------------------------------------------------------------------------
