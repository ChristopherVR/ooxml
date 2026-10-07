import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx, createDocument } from './index';
import { parseReviewPreferences } from './settings';
import { WORD_NS } from './xml';

describe('Word review recording preferences', () => {
	it.each(['', ' w:val="1"', ' w:val="true"', ' w:val="on"'])(
		'reads enabled negative flags%s',
		(attribute) => {
			expect(
				parseReviewPreferences(
					`<w:settings xmlns:w="${WORD_NS}"><w:doNotTrackFormatting${attribute}/><w:doNotTrackMoves${attribute}/></w:settings>`,
				),
			).toEqual({ trackFormatting: false, trackMoves: false });
		},
	);
	it.each(['0', 'false', 'off'])('keeps tracking enabled for negative flags set to %s', (value) => {
		expect(
			parseReviewPreferences(
				`<w:settings xmlns:w="${WORD_NS}"><w:doNotTrackFormatting w:val="${value}"/><w:doNotTrackMoves w:val="${value}"/></w:settings>`,
			),
		).toEqual({ trackFormatting: true, trackMoves: true });
	});
	it('defaults to enabled when negative flags are absent', () => {
		expect(parseReviewPreferences(`<w:settings xmlns:w="${WORD_NS}"/>`)).toEqual({
			trackFormatting: true,
			trackMoves: true,
		});
	});
	it('preserves native disabled preferences and can enable them without changing other settings', async () => {
		const bytes = new Uint8Array(
			await readFile(
				new URL('./__fixtures__/review-preferences/preferences.docx', import.meta.url),
			),
		);
		const loaded = await loadDocx(bytes);
		expect(loaded.model).toMatchObject({
			trackChanges: true,
			trackFormatting: false,
			trackMoves: false,
		});
		expect(await loaded.save()).toEqual(bytes);
		const model = { ...loaded.model, trackFormatting: true, trackMoves: true };
		const enabled = await loadDocx(await loaded.save(model));
		expect(enabled.model).toMatchObject({
			trackChanges: true,
			trackFormatting: true,
			trackMoves: true,
		});
		const xml = await (
			await JSZip.loadAsync(await loaded.save(model))
		)
			.file('word/settings.xml')!
			.async('string');
		expect(xml).not.toContain('doNotTrackFormatting');
		expect(xml).not.toContain('doNotTrackMoves');
		expect(xml).toContain('trackRevisions');
		expect(xml).toContain('compat');
		const standalone = await loadDocx(await saveDocx(loaded.model));
		expect(standalone.model).toMatchObject({
			trackChanges: true,
			trackFormatting: false,
			trackMoves: false,
		});
	});
	it('creates the settings part and relationships when preferences are disabled in a new document', async () => {
		const bytes = await saveDocx({
			...createDocument(),
			trackFormatting: false,
			trackMoves: false,
		});
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('word/settings.xml')!.async('string')).toContain('doNotTrackFormatting');
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			'relationships/settings',
		);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('settings+xml');
		expect((await loadDocx(bytes)).model).toMatchObject({
			trackFormatting: false,
			trackMoves: false,
		});
	});
});
