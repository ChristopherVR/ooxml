import { describe, expect, it } from 'vitest';
import { diagnosticCollector } from './diagnostics.js';
import { parseVsdx } from './index.js';
import { cell, fixture, row, section, shape } from './test-fixtures.js';

describe('bounded Visio metadata and diagnostics', () => {
	it.each([
		shape('1', '', `Name="${'x'.repeat(4097)}"`),
		shape('x'.repeat(257)),
		shape('1', section('Geometry', row(1, 'x'.repeat(129), ''))),
		shape('1', section('Character', row(0, '', cell('Font', 'x'.repeat(1025))))),
		shape('1', section('Paragraph', row(0, '', cell('BulletStr', 'x'.repeat(4097))))),
	])('rejects excessive metadata before inherited scene expansion', async (contents) => {
		await expect(
			parseVsdx(
				await fixture({
					masters: [{ id: '1', shapes: contents }],
					pages: [{ id: '0', contents: `<Shapes>${shape('2', '', 'Master="1"')}</Shapes>` }],
				}),
			),
		).rejects.toMatchObject({ code: 'METADATA_LIMIT' });
	});
	it('does not apply metadata caps to ordinary text or numeric geometry formulas', async () => {
		const value = 't'.repeat(5000);
		const document = await parseVsdx(
			await fixture({
				pages: [
					{ id: '0', contents: `<Shapes>${shape('1', '<Text>' + value + '</Text>')}</Shapes>` },
				],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.text.plainText).toBe(value);
	});
	it('bounds every diagnostic field before deduplication and keeps a truncation marker', () => {
		const collector = diagnosticCollector(20),
			large = 'x'.repeat(100_000);
		collector.report(large, large, { part: large, pageId: large, shapeId: large });
		collector.report(large, large, { part: large, pageId: large, shapeId: large });
		const diagnostics = collector.finish();
		expect(diagnostics).toHaveLength(2);
		expect(diagnostics[0]!.code.length).toBeLessThanOrEqual(64);
		expect(diagnostics[0]!.message.length).toBeLessThanOrEqual(2048);
		expect(diagnostics[0]!.part!.length).toBeLessThanOrEqual(1024);
		expect(diagnostics[0]!.pageId!.length).toBeLessThanOrEqual(256);
		expect(diagnostics[0]!.shapeId!.length).toBeLessThanOrEqual(1024);
		expect(diagnostics.at(-1)!.code).toBe('diagnostics-truncated');
	});
	it('caps aggregate diagnostic characters independently of the count limit', () => {
		const collector = diagnosticCollector(100_000);
		for (let i = 0; i < 2000; i++)
			collector.report('warning', 'x'.repeat(2000), { shapeId: String(i) });
		const diagnostics = collector.finish();
		const size = diagnostics
			.flatMap((item) => Object.values(item))
			.reduce((sum, value) => sum + value.length, 0);
		expect(size).toBeLessThanOrEqual(1_000_000);
		expect(diagnostics.length).toBeLessThan(500);
		expect(diagnostics.length).toBeGreaterThan(1);
		expect(diagnostics.at(-1)!.code).toBe('diagnostics-truncated');
	});
	it('retains bounded count semantics without sharing mutable marker objects', () => {
		const first = diagnosticCollector(1);
		first.report('first', 'first');
		first.report('second', 'second');
		first.finish()[0]!.message = 'Changed externally';
		const second = diagnosticCollector(1);
		second.report('first', 'first');
		second.report('second', 'second');
		expect(second.finish()[0]!.message).not.toBe('Changed externally');
	});
});
