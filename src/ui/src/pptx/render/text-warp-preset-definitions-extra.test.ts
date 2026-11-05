import { describe, expect, it } from 'vitest';

import { GLYPH_ENVELOPE_PRESETS, hasGlyphEnvelope } from './text-warp-envelope-curves';
import { sampleWarpPresetCurve, sampleWarpPresetXExtent } from './text-warp-preset-sampler';

/**
 * Boundary of a block of full-block glyphs exported by PowerPoint 16 at 1280x720
 * (a 220 x 80 pt shape): `[top, bottom]` as a fraction of the box height at
 * x = 0.25, 0.5 and 0.75. Each row is `[preset, adj, adj2, samples]`.
 */
const MEASURED: ReadonlyArray<
	readonly [string, number | undefined, number | undefined, number[][]]
> = [
	[
		'textArchUpPour',
		undefined,
		undefined,
		[
			[0.066, 0.502],
			[0.002, 0.248],
			[0.073, 0.495],
		],
	],
	[
		'textArchDownPour',
		undefined,
		undefined,
		[
			[0.502, 0.938],
			[0.628, 1.001],
			[0.516, 0.93],
		],
	],
	[
		'textButtonPour',
		undefined,
		undefined,
		[
			[0.071, 0.5],
			[0.001, 0.247],
			[0.071, 0.5],
		],
	],
	[
		'textArchUpPour',
		10800000,
		30000,
		[
			[0.064, 0.5],
			[0.001, 0.352],
			[0.071, 0.5],
		],
	],
	[
		'textStop',
		undefined,
		undefined,
		[
			[0.059, 0.945],
			[0.002, 1.001],
			[0.059, 0.938],
		],
	],
	[
		'textStop',
		10000,
		undefined,
		[
			[0.038, 0.966],
			[0.002, 1.001],
			[0.038, 0.966],
		],
	],
	[
		'textTriangle',
		undefined,
		undefined,
		[
			[0.248, 1.001],
			[0.002, 1.001],
			[0.248, 1.001],
		],
	],
	[
		'textTriangle',
		20000,
		undefined,
		[
			[0.101, 1.001],
			[0.002, 1.001],
			[0.101, 1.001],
		],
	],
	[
		'textTriangleInverted',
		undefined,
		undefined,
		[
			[0.001, 0.753],
			[0.001, 0.992],
			[0.001, 0.753],
		],
	],
	[
		'textTriangleInverted',
		20000,
		undefined,
		[
			[0.001, 0.612],
			[0.001, 0.992],
			[0.001, 0.605],
		],
	],
	[
		'textChevron',
		undefined,
		undefined,
		[
			[0.127, 0.873],
			[0.001, 0.753],
			[0.127, 0.88],
		],
	],
	[
		'textChevron',
		10000,
		undefined,
		[
			[0.055, 0.948],
			[-0.001, 0.899],
			[0.055, 0.948],
		],
	],
	[
		'textChevronInverted',
		undefined,
		undefined,
		[
			[0.129, 0.874],
			[0.248, 0.994],
			[0.122, 0.874],
		],
	],
	[
		'textChevronInverted',
		10000,
		undefined,
		[
			[0.245, 0.752],
			[0.498, 0.998],
			[0.245, 0.752],
		],
	],
	[
		'textCurveUp',
		undefined,
		undefined,
		[
			[0.486, 0.957],
			[0.402, 0.999],
			[0.233, 0.957],
		],
	],
	[
		'textCurveUp',
		20000,
		undefined,
		[
			[0.216, 0.982],
			[0.18, 0.996],
			[0.103, 0.982],
		],
	],
	[
		'textCurveDown',
		undefined,
		undefined,
		[
			[0.233, 0.957],
			[0.402, 0.999],
			[0.486, 0.957],
		],
	],
	[
		'textCurveDown',
		20000,
		undefined,
		[
			[0.101, 0.98],
			[0.178, 1.001],
			[0.213, 0.98],
		],
	],
	[
		'textFadeRight',
		undefined,
		undefined,
		[
			[0.084, 0.92],
			[0.168, 0.836],
			[0.252, 0.752],
		],
	],
	[
		'textFadeRight',
		15000,
		undefined,
		[
			[0.038, 0.959],
			[0.08, 0.923],
			[0.115, 0.888],
		],
	],
	[
		'textFadeLeft',
		undefined,
		undefined,
		[
			[0.251, 0.75],
			[0.166, 0.834],
			[0.082, 0.919],
		],
	],
	[
		'textFadeLeft',
		15000,
		undefined,
		[
			[0.113, 0.887],
			[0.078, 0.922],
			[0.043, 0.964],
		],
	],
	[
		'textFadeUp',
		undefined,
		undefined,
		[
			[0.23, 0.996],
			[0.005, 0.996],
			[0.237, 0.996],
		],
	],
	[
		'textFadeUp',
		15000,
		undefined,
		[
			[-0.001, 0.998],
			[-0.001, 0.998],
			[-0.001, 0.998],
		],
	],
	[
		'textFadeDown',
		undefined,
		undefined,
		[
			[0.005, 0.764],
			[0.005, 0.996],
			[0.005, 0.757],
		],
	],
	[
		'textFadeDown',
		15000,
		undefined,
		[
			[-0.001, 0.998],
			[-0.001, 0.998],
			[-0.001, 0.998],
		],
	],
	[
		'textSlantUp',
		undefined,
		undefined,
		[
			[0.417, 0.86],
			[0.277, 0.727],
			[0.136, 0.586],
		],
	],
	[
		'textSlantUp',
		20000,
		undefined,
		[
			[0.152, 0.947],
			[0.103, 0.898],
			[0.054, 0.848],
		],
	],
	[
		'textSlantDown',
		undefined,
		undefined,
		[
			[0.143, 0.586],
			[0.277, 0.727],
			[0.417, 0.86],
		],
	],
	[
		'textSlantDown',
		20000,
		undefined,
		[
			[0.178, 0.466],
			[0.354, 0.649],
			[0.537, 0.825],
		],
	],
	[
		'textCascadeUp',
		undefined,
		undefined,
		[
			[0.108, 0.86],
			[0.073, 0.727],
			[0.038, 0.586],
		],
	],
	[
		'textCascadeUp',
		20000,
		undefined,
		[
			[0.136, 0.825],
			[0.094, 0.642],
			[0.045, 0.466],
		],
	],
	[
		'textCascadeDown',
		undefined,
		undefined,
		[
			[0.036, 0.584],
			[0.071, 0.725],
			[0.106, 0.866],
		],
	],
	[
		'textCascadeDown',
		20000,
		undefined,
		[
			[0.05, 0.472],
			[0.092, 0.648],
			[0.134, 0.823],
		],
	],
	[
		'textWave1',
		undefined,
		undefined,
		[
			[0.013, 0.759],
			[0.126, 0.878],
			[0.245, 0.991],
		],
	],
	[
		'textWave1',
		12500,
		-8000,
		[
			[0.03, 0.769],
			[0.066, 0.938],
			[0.234, 0.973],
		],
	],
	[
		'textWave1',
		12500,
		-8000,
		[
			[0.03, 0.769],
			[0.066, 0.938],
			[0.234, 0.973],
		],
	],
	[
		'textWave2',
		undefined,
		undefined,
		[
			[0.244, 0.989],
			[0.124, 0.877],
			[0.012, 0.757],
		],
	],
	[
		'textWave2',
		12500,
		-8000,
		[
			[0.219, 0.978],
			[0.184, 0.816],
			[0.022, 0.788],
		],
	],
	[
		'textWave2',
		12500,
		-8000,
		[
			[0.219, 0.978],
			[0.184, 0.816],
			[0.022, 0.788],
		],
	],
	[
		'textWave4',
		undefined,
		undefined,
		[
			[0.061, 0.94],
			[0.061, 0.947],
			[0.061, 0.94],
		],
	],
	[
		'textWave4',
		6250,
		-5000,
		[
			[0.11, 0.926],
			[0.019, 0.989],
			[0.075, 0.891],
		],
	],
	[
		'textWave4',
		6250,
		-5000,
		[
			[0.11, 0.926],
			[0.019, 0.989],
			[0.075, 0.891],
		],
	],
	[
		'textDoubleWave1',
		undefined,
		undefined,
		[
			[0.061, 0.94],
			[0.061, 0.94],
			[0.061, 0.94],
		],
	],
	[
		'textDoubleWave1',
		6250,
		-5000,
		[
			[0.016, 0.959],
			[0.108, 0.895],
			[0.045, 0.987],
		],
	],
	[
		'textDoubleWave1',
		6250,
		-5000,
		[
			[0.016, 0.959],
			[0.108, 0.895],
			[0.045, 0.987],
		],
	],
];

const X = [0.25, 0.5, 0.75];
const TOLERANCE = 0.015;

describe('measured envelope presets (path families)', () => {
	it.each(MEASURED)(
		'%s adj=%s adj2=%s follows the PowerPoint boundary',
		(preset, adj, adj2, rows) => {
			const extent = sampleWarpPresetXExtent(preset, adj, adj2);
			expect(extent).toBeDefined();
			let checked = 0;
			X.forEach((x, i) => {
				const [top, bottom] = rows[i] as [number, number];
				const ut = (x - extent!.top[0]) / (extent!.top[1] - extent!.top[0]);
				const ub = (x - extent!.bottom[0]) / (extent!.bottom[1] - extent!.bottom[0]);
				if (ut >= 0 && ut <= 1) {
					const curve = sampleWarpPresetCurve(preset, ut, adj, adj2);
					expect(Math.abs(curve!.top - top)).toBeLessThan(TOLERANCE);
					checked++;
				}
				if (ub >= 0 && ub <= 1) {
					const curve = sampleWarpPresetCurve(preset, ub, adj, adj2);
					expect(Math.abs(curve!.bottom - bottom)).toBeLessThan(TOLERANCE);
					checked++;
				}
			});
			expect(checked).toBeGreaterThan(0);
		},
	);

	it('routes every measured preset through the glyph envelope', () => {
		for (const [preset] of MEASURED) {
			expect(hasGlyphEnvelope(preset)).toBeTruthy();
			expect(GLYPH_ENVELOPE_PRESETS.has(preset)).toBeTruthy();
		}
	});

	it('shortens the fade presets to their plateau', () => {
		expect(sampleWarpPresetXExtent('textFadeUp', undefined, undefined)!.top[0]).toBeCloseTo(
			0.33333,
			4,
		);
		expect(sampleWarpPresetXExtent('textFadeDown', undefined, undefined)!.bottom[1]).toBeCloseTo(
			0.66667,
			4,
		);
	});

	it('limits the inner pour ellipse to the middle of the width', () => {
		const extent = sampleWarpPresetXExtent('textArchUpPour', undefined, undefined)!;
		expect(extent.top).toStrictEqual([0, 1]);
		expect(extent.bottom[0]).toBeCloseTo(0.25, 4);
		expect(extent.bottom[1]).toBeCloseTo(0.75, 4);
	});

	it('slides the wave paths against each other with adj2', () => {
		const shifted = sampleWarpPresetXExtent('textWave1', 12500, 8000)!;
		expect(shifted.top[1]).toBeCloseTo(0.84, 4);
		expect(shifted.bottom[0]).toBeCloseTo(0.16, 4);
	});
});
