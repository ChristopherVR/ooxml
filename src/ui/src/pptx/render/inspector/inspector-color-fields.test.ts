import type { ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createInspectorColorFields } from './inspector-color-fields';
import { SHAPE_EFFECT_SECTIONS } from './shape-effect-sections';

/** Two shapes' styles and a selection; patches go to the shape they name. */
function renderer() {
	const styles: Record<string, ShapeStyle> = {
		a: { glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 },
		b: { glowColor: '#ff0000', glowRadius: 3, glowOpacity: 0.9 },
	};
	let selected = 'a';
	const commits: Array<Partial<ShapeStyle>> = [];
	const recent: string[] = [];
	const colorField = createInspectorColorFields<ShapeStyle | undefined, Partial<ShapeStyle>>({
		sections: () => SHAPE_EFFECT_SECTIONS,
		target: () => styles[selected],
		targetKey: () => selected,
		preview(patch, target) {
			styles[target] = { ...styles[target], ...patch };
		},
		commit(patch) {
			commits.push(patch);
			styles[selected] = { ...styles[selected], ...patch };
		},
		remember: (hex) => recent.push(hex),
	});
	return {
		styles,
		commits,
		recent,
		glow: () => colorField('glow', 'color'),
		select(id: string) {
			selected = id;
		},
	};
}

describe('createInspectorColorFields', () => {
	it('keeps one controller per field, so a drag is one commit and one recent colour', () => {
		const r = renderer();
		r.glow().input('#00ff00');
		r.glow().input('#0000ff');
		r.glow().change('#0000ff');

		expect(r.commits).toStrictEqual([{ glowColor: '#0000ff', glowRadius: 8, glowOpacity: 0.5 }]);
		expect(r.recent).toStrictEqual(['#0000ff']);
	});

	it('puts a drag back on the shape it started on when the selection moves, and stops there', () => {
		const r = renderer();
		r.glow().input('#00ff00');
		r.select('b');
		r.glow().blur('#00ff00');

		expect(r.styles['a']).toStrictEqual({ glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 });
		expect(r.styles['b']).toStrictEqual({ glowColor: '#ff0000', glowRadius: 3, glowOpacity: 0.9 });
		expect(r.commits).toStrictEqual([]);
	});
});
