import { describe, expect, it } from 'vitest';
import { editVsdx, parseVsdx } from '../index';
import { createSampleVsdx } from './sample-drawing';

describe('createSampleVsdx', () => {
	it('builds a two-page source-backed drawing with glued connectors', async () => {
		const model = await parseVsdx(await createSampleVsdx());
		expect(model.pages.map((page) => page.name)).toEqual(['Release workflow', 'Architecture']);
		const [workflow, architecture] = model.pages;
		expect(workflow!.shapes.map((shape) => shape.text?.plainText ?? '')).toContain('Ready?');
		expect(workflow!.connectors.length).toBe(6);
		expect(architecture!.connectors.length).toBe(4);
		expect(model.diagnostics.filter((d) => d.severity === 'warning')).toEqual([]);
	});

	it('accepts ordinary edits, so the sample is not read-only', async () => {
		const bytes = await createSampleVsdx();
		const moved = await editVsdx(bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: '2', x: 3, y: 4.65 },
			{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Edited' },
		]);
		const page = (await parseVsdx(moved.bytes)).pages[0]!;
		expect(page.shapes.find((shape) => shape.id === '1')!.text!.plainText).toBe('Edited');
		// The glued connectors follow the moved shape.
		expect(page.connectors.filter((connect) => connect.toShapeId === '2')).toHaveLength(2);
	});
});
