import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { cell, fixture, shape } from './test-fixtures';

describe('saved connector jump settings', () => {
	it.each([0, 1, 2, 3, 9])(
		'round-trips native page style %s without shifting its enum',
		async (lineJumpStyle) => {
			const bytes = await fixture();
			const saved = await editVsdx(bytes, [
				{ type: 'set-page-layout', pageId: '0', lineJumpStyle },
			]);
			expect((await parseVsdx(saved.bytes)).pages[0]!.layout?.lineJumpStyle ?? 0).toBe(
				lineJumpStyle,
			);
		},
	);
	it('reads master caches and local overrides without treating zero as absent', async () => {
		const bytes = await fixture({
			masters: [
				{
					id: '2',
					shapes: shape(
						'1',
						cell('OneD', 1) + cell('ConLineJumpStyle', '3.0') + cell('ConLineJumpDirX', '2e0'),
					),
				},
			],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('ConLineJumpDirX', 0) + cell('ConLineJumpDirY', 1), 'Master="2"')}</Shapes>`,
				},
			],
		});
		const result = await parseVsdx(bytes);
		expect(result.pages[0]!.shapes[0]).toMatchObject({
			lineJumpStyle: 3,
			lineJumpDirectionX: 0,
			lineJumpDirectionY: 1,
		});
	});
	it('reports unsupported caches instead of inventing a connector setting', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('OneD', 1) + cell('ConLineJumpStyle', 10) + cell('ConLineJumpDirX', 'NaN'))}</Shapes>`,
				},
			],
		});
		const result = await parseVsdx(bytes);
		expect(result.pages[0]!.shapes[0]!.lineJumpStyle).toBeUndefined();
		expect(result.diagnostics.filter((d) => d.code === 'unsupported-connector-jump')).toHaveLength(
			2,
		);
	});
	it('saves and reopens page jump directions while preserving connector overrides', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('OneD', 1) + cell('ConLineJumpStyle', 2))}</Shapes>`,
				},
			],
		});
		const edit = {
			type: 'set-page-layout' as const,
			pageId: '0',
			lineJumpDirectionX: 2,
			lineJumpDirectionY: 1,
		};
		const saved = await editVsdx(bytes, [edit]);
		const result = await parseVsdx(saved.bytes);
		expect(result.pages[0]!.layout).toMatchObject({ lineJumpDirectionX: 2, lineJumpDirectionY: 1 });
		expect(result.pages[0]!.shapes[0]!.lineJumpStyle).toBe(2);
		expect((await editVsdx(saved.bytes, [edit])).changedParts).toEqual([]);
		await expect(editVsdx(bytes, [{ ...edit, lineJumpDirectionX: 3 }])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});
});
