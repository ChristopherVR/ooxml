import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import type { VisioPage, VisioShape } from './model';
import { visioLineJumpPaths } from './line-jumps';

const native = JSON.parse(
	readFileSync(new URL('./__fixtures__/native-line-jump-policy.json', import.meta.url), 'utf8'),
) as {
	cases: [number, number, number, boolean, boolean][];
};
describe('native Visio connector jump policy matrix', () => {
	it.each(native.cases)(
		'page %s, horizontal %s, vertical %s',
		(pageCode, horizontal, vertical, hJump, vJump) => {
			const line = (id: string, path: string, lineJumpCode: number): VisioShape =>
				({
					id,
					name: id,
					kind: 'connector',
					width: 4,
					height: 4,
					transform: [1, 0, 0, 1, 0, 0],
					geometry: [{ path, fill: false, stroke: true }],
					connectorRoute: 'straight',
					children: [],
					hidden: false,
					lineJumpCode,
				}) as unknown as VisioShape;
			const h = line('1', 'M 1 2 L 5 2', horizontal);
			const v = line('2', 'M 3 0 L 3 4', vertical);
			const page: VisioPage = {
				id: '0',
				name: 'Page',
				width: 8,
				height: 6,
				isBackground: false,
				shapes: [h, v],
				connectors: [],
				layout: { lineJumpCode: pageCode, lineJumpStyle: 1 },
			};
			const paths = visioLineJumpPaths(page);
			expect(paths.has(h)).toBe(hJump);
			expect(paths.has(v)).toBe(vJump);
		},
	);
});
