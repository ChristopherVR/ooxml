import { describe, expect, it } from 'vitest';
import { parseVsdx } from './index.js';
import { cell, fixture, row, section, shape } from './test-fixtures.js';
const parseText = async (data: string, text: string, document = '', attrs = '') =>
	(
		await parseVsdx(
			await fixture({
				document,
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', data + `<Text>${text}</Text>`, attrs)}</Shapes>`,
					},
				],
			}),
		)
	).pages[0]!.shapes[0]!.text;
describe('cached Visio paragraph and text styles', () => {
	it('preserves mixed paragraph markers, offsets, alignment, RTL and exact line spacing', async () => {
		const paragraphs = section(
			'Paragraph',
			row(0, '', cell('HorzAlign', 0) + cell('SpLine', -1.5)) +
				row(1, '', cell('HorzAlign', 2) + cell('SpLine', 0.25) + cell('Flags', 1)),
		);
		const text = await parseText(paragraphs, '<pp IX="0"/>First\n<pp IX="1"/>Second');
		expect(text.paragraphs).toMatchObject([
			{
				start: 0,
				end: 5,
				horizontalAlign: 'left',
				direction: 'ltr',
				lineSpacing: { kind: 'multiple', value: 1.5 },
			},
			{
				start: 6,
				end: 12,
				horizontalAlign: 'right',
				direction: 'rtl',
				lineSpacing: { kind: 'exact', value: 0.25 },
			},
		]);
	});
	it('keeps indentation and before/after paragraph spacing in inches', async () => {
		const text = await parseText(
			section(
				'Paragraph',
				row(
					0,
					'',
					cell('IndLeft', 0.2) +
						cell('IndRight', 0.3) +
						cell('IndFirst', -0.1) +
						cell('SpBefore', 0.1) +
						cell('SpAfter', 0.2),
				),
			),
			'Text',
		);
		expect(text.paragraphs![0]).toMatchObject({
			indentLeft: 0.2,
			indentRight: 0.3,
			indentFirst: -0.1,
			spaceBefore: 0.1,
			spaceAfter: 0.2,
		});
	});
	it('preserves built-in/custom bullets and character font fallback', async () => {
		const data =
			section('Character', row(1, '', cell('Style', 1))) +
			section(
				'Paragraph',
				row(0, '', cell('Bullet', 1)) +
					row(
						1,
						'',
						cell('Bullet', 1) +
							cell('BulletStr', '→') +
							cell('BulletFontSize', 0.2) +
							cell('TextPosAfterBullet', 0.3),
					),
			);
		const style = `<StyleSheets><StyleSheet ID="0">${section('Character', row(0, '', cell('Font', 'Courier New') + cell('Size', 0.5) + cell('Color', '#123456')))}</StyleSheet></StyleSheets>`;
		const text = await parseText(data, '<cp IX="1"/>First\n<pp IX="1"/>Second', style);
		expect(text.runs[0]).toMatchObject({
			fontFamily: 'Courier New',
			fontSize: 0.5,
			color: '#123456',
			bold: true,
		});
		expect(text.paragraphs![0]!.bullet).toMatchObject({
			text: '•',
			fontFamily: 'Courier New',
			fontSize: 0.5,
		});
		expect(text.paragraphs![1]!.bullet).toMatchObject({ text: '→', fontSize: 0.2, offset: 0.3 });
	});
	it('inherits default DocumentSheet styles and implicit style zero', async () => {
		const styles = `<StyleSheets><StyleSheet ID="0">${section('Character', row(0, '', cell('Font', 'Georgia') + cell('Size', 0.3)))}</StyleSheet><StyleSheet ID="2">${section('Character', row(0, '', cell('Font', 'Verdana')))}</StyleSheet></StyleSheets>`;
		expect((await parseText('', 'Text', styles)).fontFamily).toBe('Georgia');
		expect(
			(await parseText('', 'Text', styles + '<DocumentSheet TextStyle="2"/>')).fontFamily,
		).toBe('Verdana');
	});
	it('does not inherit explicitly disabled style categories', async () => {
		const styles = `<StyleSheets><StyleSheet ID="1">${cell('EnableTextProps', 0)}${cell('EnableFillProps', 0)}${cell('EnableLineProps', 0)}${cell('FillForegnd', '#ff0000')}${cell('LineColor', '#ff0000')}${section('Character', row(0, '', cell('Font', 'Courier New') + cell('Size', 0.5)))}</StyleSheet></StyleSheets>`;
		const document = await parseVsdx(
			await fixture({
				document: styles,
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', '<Text>Text</Text>', 'TextStyle="1" FillStyle="1" LineStyle="1"')}</Shapes>`,
					},
				],
			}),
		);
		const item = document.pages[0]!.shapes[0]!;
		expect(item.text.fontFamily).toBe('Arial');
		expect(item.style.fill).toBe('#ffffff');
		expect(item.style.lineColor).toBe('#000000');
	});
	it('keeps UTF-16 offsets and excludes terminating paragraph newline', async () => {
		const text = await parseText('', '😀\n\nLast\n');
		expect(text.paragraphs!.map(({ start, end }) => [start, end])).toEqual([
			[0, 2],
			[3, 3],
			[4, 8],
		]);
	});
	it('inherits row defaults without leaking local sibling formatting', async () => {
		const styles = `<StyleSheets><StyleSheet ID="1">${section('Character', row(0, '', cell('Font', 'Courier New') + cell('Size', 0.5) + cell('Style', 0)))}</StyleSheet></StyleSheets>`;
		const local = section(
			'Character',
			row(0, '', cell('Style', 1)) + row(1, '', cell('Size', 0.75)),
		);
		const text = await parseText(local, 'Bold<cp IX="1"/>Normal', styles, 'TextStyle="1"');
		expect(text.runs).toMatchObject([
			{ bold: true, fontFamily: 'Courier New', fontSize: 0.5 },
			{ bold: false, fontFamily: 'Courier New', fontSize: 0.75 },
		]);
	});
	it('bounds normalized paragraphs and runs separately from text character counts', async () => {
		const paragraphs = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>a\nb\nc</Text>')}</Shapes>` }],
		});
		await expect(parseVsdx(paragraphs, { maxParagraphs: 2 })).rejects.toMatchObject({
			code: 'PARAGRAPH_LIMIT',
		});
		const runs = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', '<Text>a<cp IX="0"/>b<cp IX="0"/>c</Text>')}</Shapes>`,
				},
			],
		});
		await expect(parseVsdx(runs, { maxTextRuns: 2 })).rejects.toMatchObject({
			code: 'TEXT_RUN_LIMIT',
		});
	});
});
