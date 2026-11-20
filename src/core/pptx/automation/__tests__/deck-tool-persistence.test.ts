import { readFileSync } from 'node:fs';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { diffOpcPackages } from '../../../opc/diff-packages';
import { getThemePreset, PptxHandler } from '../../index';
import { pptxTools } from '../catalogue';
import { saveDeckBytes } from '../execution';
import { createTestPptxBytes } from './helpers/create-test-pptx';

const corpus = new Uint8Array(
	readFileSync(
		new URL(
			'../../__tests__/fixtures/corpus/master-layout-inheritance-fills.pptx',
			import.meta.url,
		),
	),
);

async function runEdit(name: string, bytes: Uint8Array, params: unknown): Promise<Uint8Array> {
	const tool = pptxTools.find((candidate) => candidate.name === name);
	if (tool?.kind !== 'edit') throw new Error(`${name} is not an edit tool`);
	const result = await tool.run(bytes, params);
	if (!result.bytes) throw new Error(`${name} reported no change`);
	return result.bytes;
}

const reload = (bytes: Uint8Array) => new PptxHandler().load(new Uint8Array(bytes).buffer);

async function partsText(bytes: Uint8Array, pattern: RegExp): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	const names = Object.keys(zip.files).filter((name) => pattern.test(name));
	const texts = await Promise.all(names.map((name) => zip.file(name)?.async('string') ?? ''));
	return texts.join('\n');
}

describe.each([
	['a corpus deck', async () => corpus],
	['a generated deck', () => createTestPptxBytes(2)],
])('deck tools persist their edits on %s', (_label, source) => {
	it('update_metadata writes core, app and custom properties', async () => {
		const bytes = await runEdit('update_metadata', await source(), {
			title: 'Saved title',
			company: 'Saved company',
			customProperties: [{ name: 'reviewed', value: 'yes' }],
		});

		const deck = await reload(bytes);
		expect(deck.coreProperties?.title).toBe('Saved title');
		expect(deck.appProperties?.company).toBe('Saved company');
		expect(deck.customProperties).toContainEqual(
			expect.objectContaining({ name: 'reviewed', value: 'yes' }),
		);
	});

	it('update_theme_colors changes the saved theme and only theme-related parts', async () => {
		const original = await source();
		const bytes = await runEdit('update_theme_colors', original, { accent1: '#123456' });

		expect((await reload(bytes)).theme?.colorScheme?.accent1?.toUpperCase()).toBe('#123456');
		const plain = await saveDeckBytes({ pptxData: await reload(original) }, original);
		const { added, removed, changed } = await diffOpcPackages(plain, bytes);
		expect(added).toEqual([]);
		expect(removed).toEqual([]);
		const unexpected = changed.filter((part) => !/^(ppt\/theme\/|docProps\/core)/u.test(part));
		expect(unexpected).toEqual([]);
		expect(changed.some((part) => part.startsWith('ppt/theme/'))).toBe(true);
	});

	it('update_theme_fonts changes the saved heading font', async () => {
		const bytes = await runEdit('update_theme_fonts', await source(), { majorFont: 'Saved Sans' });

		expect((await reload(bytes)).theme?.fontScheme?.majorFont?.latin).toBe('Saved Sans');
		expect(await partsText(bytes, /^ppt\/theme\//u)).toContain('Saved Sans');
	});

	it('apply_theme_preset writes the preset colours and name', async () => {
		const original = await source();
		const bytes = await runEdit('apply_theme_preset', original, { presetName: 'DARK' });

		const deck = await reload(bytes);
		const before = await reload(original);
		expect(deck.theme?.colorScheme).not.toEqual(before.theme?.colorScheme);
		const name = getThemePreset('DARK')?.name ?? '';
		expect(await partsText(bytes, /^ppt\/theme\//u)).toContain(`name="${name}"`);
	});
});

describe('deck tools leave untouched document properties alone', () => {
	it('keeps core properties when only the theme changes', async () => {
		const bytes = await runEdit('update_theme_colors', corpus, { accent2: '#ABCDEF' });

		const [before, after] = await Promise.all([reload(corpus), reload(bytes)]);
		expect(after.coreProperties?.title).toBe(before.coreProperties?.title);
		expect(after.coreProperties?.creator).toBe(before.coreProperties?.creator);
	});
});
