import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PresentationBuilder } from '../../builders/sdk';
import { PptxHandler } from '../../PptxHandler';
import { renameFixtureParts } from '../../../__tests__/helpers/rename-package-parts';

it('allocates an independent notes theme without overwriting occupied parts', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('Slide').build()];
	const initial = await JSZip.loadAsync(await handler.save(data.slides));
	const theme = await initial.file('ppt/theme/theme1.xml')!.async('string');
	initial.file('ppt/theme/theme2.xml', 'occupied opaque theme');
	const themeRelationships =
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/theme.png"/><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/theme" TargetMode="External"/></Relationships>';
	initial.file('ppt/theme/_rels/theme1.xml.rels', themeRelationships);
	initial.file('ppt/media/theme.png', new Uint8Array([1, 2, 3]));
	await handler.load(await initial.generateAsync({ type: 'arraybuffer' }));
	data.slides[0]!.notes = 'Rich notes';
	data.slides[0]!.notesSegments = [{ text: 'Rich notes', style: { bold: true } }];
	const saved = await JSZip.loadAsync(await handler.save(data.slides));
	expect(await saved.file('ppt/theme/theme1.xml')!.async('string')).toBe(theme);
	expect(await saved.file('ppt/theme/theme2.xml')!.async('string')).toBe('occupied opaque theme');
	expect(await saved.file('ppt/theme/theme3.xml')!.async('string')).toBe(theme);
	expect(await saved.file('ppt/theme/_rels/theme3.xml.rels')!.async('string')).toBe(
		themeRelationships,
	);
	expect(await saved.file('ppt/media/theme.png')!.async('uint8array')).toEqual(
		new Uint8Array([1, 2, 3]),
	);
	expect(
		await saved.file('ppt/notesMasters/_rels/notesMaster1.xml.rels')!.async('string'),
	).toContain('../theme/theme3.xml');
	expect(await saved.file('[Content_Types].xml')!.async('string')).toContain(
		'/ppt/theme/theme3.xml',
	);
	expect(await saved.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain('b="1"');
	handler.dispose();
});

it('does not require theme infrastructure for a fresh model with empty notes', async () => {
	const { data, createSlide, handler: builderHandler } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('Empty notes').build()];
	data.slides[0]!.notes = '';
	const handler = new PptxHandler();
	const saved = await JSZip.loadAsync(await handler.save(data.slides));
	expect(saved.file(/^ppt\/notesMasters\//u)).toHaveLength(0);
	handler.dispose();
	builderHandler.dispose();
});

it('builds a notes theme from existing state when a source theme part is absent', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('Theme-free notes').build()];
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	zip.remove('ppt/theme/theme1.xml');
	await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
	data.slides[0]!.notes = 'Authored notes';
	const output = await JSZip.loadAsync(await handler.save(data.slides));
	expect(await output.file('ppt/theme/theme1.xml')!.async('string')).toContain('a:themeElements');
	expect(await output.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain(
		'Authored notes',
	);
	handler.dispose();
});

it('preserves a registered notes master with a nonstandard part name when saving existing notes', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('Nonstandard notes').build()];
	data.slides[0]!.notes = 'Retained notes';
	const source = await renameFixtureParts(
		await handler.save(data.slides),
		new Map([['ppt/notesMasters/notesMaster1.xml', 'ppt/templates/presenter-notes.xml']]),
	);
	const loaded = await handler.load(
		source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer,
	);
	const saved = await JSZip.loadAsync(await handler.save(loaded.slides));
	expect(saved.file('ppt/notesMasters/notesMaster1.xml')).toBeNull();
	expect(saved.file('ppt/templates/presenter-notes.xml')).not.toBeNull();
	expect(await saved.file('ppt/_rels/presentation.xml.rels')!.async('string')).toContain(
		'templates/presenter-notes.xml',
	);
	expect(await saved.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain(
		'Retained notes',
	);
	handler.dispose();
});
