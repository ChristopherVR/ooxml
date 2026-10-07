import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg';
import evidence from '../../../core/visio/__fixtures__/short-arrows-native.json';

describe('native short filled-arrow rendering', () => {
	it.each(evidence.cases)('anchors code $code size $size weight $lineWidth ratio $ratio', (row) => {
		const model = structuredClone(demoDocument),
			page = model.pages[0]!,
			shape = page.shapes[1]!;
		page.shapes = [shape];
		const path = `M 0 0 L ${row.length} 0`;
		shape.geometry = [{ path, stroke: true, fill: false }];
		shape.style.startArrow = 0;
		shape.style.endArrow = row.code;
		shape.style.endArrowSize = row.size;
		shape.style.lineWidth = row.lineWidth;
		const result = renderPage(model, page),
			marker = result.svg.querySelector('marker')!;
		expect(Number(marker.getAttribute('refX'))).toBeCloseTo(-row.setback, 12);
		expect(marker.querySelector('path')!.getAttribute('stroke')).toBe('none');
		expect(shape.geometry[0]!.path).toBe(path);
		expect(result.warnings.some((w) => /arrowhead/i.test(w))).toBe(false);
		result.dispose();
	});
});
const native = process.env.VISIO_NATIVE_SHORT_ARROWS_DIR;
describe.skipIf(!native)('native-authored short arrow matrix', () => {
	it('renders all saved cases through parsing and shared UI', async () => {
		const model = await parseVsdx(await readFile(resolve(native!, 'short-arrows.vsdx')));
		const result = renderPage(model, model.pages[0]!);
		expect(result.svg.querySelectorAll('marker')).toHaveLength(168);
		expect(result.warnings.some((w) => /arrowhead/i.test(w))).toBe(false);
		result.dispose();
	}, 30_000);
});
