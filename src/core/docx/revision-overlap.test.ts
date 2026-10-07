import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import {
	acceptRevision,
	rejectRevision,
	rejectAllRevisions,
	listRevisions,
} from './revision-commands.js';
import { expectParagraph } from './test-support/access.js';
import { WORD_NS, WORD_DATE_UTC_NS, parseXml } from './xml.js';
import { runHasUnknownProperties } from './write-run-validation.js';

async function fixture(kind: 'ins' | 'del') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}" xmlns:x="urn:historical-format"><w:body><w:p><w:${kind} w:id="text-edit" w:author="Grace"><w:r><w:rPr><w:b/><w:rPrChange w:id="format-edit" w:author="Ada"><w:rPr><w:i/><x:opaque x:value="preserved"/></w:rPr></w:rPrChange></w:rPr><w:${kind === 'del' ? 'delText' : 't'}>Text</w:${kind === 'del' ? 'delText' : 't'}></w:r></w:${kind}></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('overlapping text and run formatting revisions', () => {
	it.each(['bold', 'multiple'])(
		'rejects tracked text overlapping native %s formatting',
		async (name) => {
			for (const kind of ['insert', 'delete'] as const) {
				const loaded = await loadDocx(
					new Uint8Array(
						await readFile(
							new URL(`./__fixtures__/review-formatting/${name}-tracked.docx`, import.meta.url),
						),
					),
				);
				const paragraph = expectParagraph(loaded.model.blocks[0]);
				const original = paragraph.runs[0]!;
				const { revision: formatRevision, ...format } = original;
				paragraph.runs = [
					{ ...original, text: original.text.slice(0, 4) },
					{
						...format,
						text: kind === 'insert' ? '!' : original.text.slice(4, 6),
						formatRevision: formatRevision!,
						revision: {
							id: 'text-review',
							kind,
							author: 'Codex',
							date: '2026-10-07T00:00:00Z',
							dateUtc: '2026-10-07T00:00:00Z',
						},
					},
					{ ...original, text: original.text.slice(kind === 'insert' ? 4 : 6) },
				];
				const reopened = await loadDocx(await loaded.save(loaded.model));
				const rejected = await loadDocx(await reopened.save(rejectAllRevisions(reopened.model)));
				const native = await loadDocx(
					new Uint8Array(
						await readFile(
							new URL(`./__fixtures__/review-formatting/${name}-rejected.docx`, import.meta.url),
						),
					),
				);
				const actual = expectParagraph(rejected.model.blocks[0]);
				const expected = expectParagraph(native.model.blocks[0]);
				expect(actual.runs.map((run) => run.text).join('')).toBe('Format me');
				const { text: _text, ...props } = expected.runs[0]!;
				for (const run of actual.runs) {
					const { text: _piece, ...actualProps } = run;
					expect(actualProps).toEqual(props);
				}
			}
		},
	);
	it('preserves modern UTC and legacy timestamps on nested revisions', async () => {
		const zip = new JSZip();
		const legacy = '2026-10-07T20:36:00Z';
		const utc = '2026-10-07T10:36:00Z';
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}" xmlns:u="${WORD_DATE_UTC_NS}"><w:body><w:p><w:ins w:id="1" w:author="Grace" w:date="${legacy}" u:dateUtc="${utc}"><w:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="Ada" w:date="${legacy}" u:dateUtc="${utc}"><w:rPr/></w:rPrChange></w:rPr><w:t>Text</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const run = expectParagraph(loaded.model.blocks[0]).runs[0]!;
		run.text += '!';
		const reopened = await loadDocx(await loaded.save(loaded.model));
		const actual = expectParagraph(reopened.model.blocks[0]).runs[0]!;
		for (const revision of [actual.revision, actual.formatRevision])
			expect(revision).toMatchObject({ date: legacy, dateUtc: utc });
	});

	it.each(['w:custom="keep"', 'xmlns:x="urn:other" x:dateUtc="keep"'])(
		'keeps unmodeled revision attributes protected: %s',
		(attributes) => {
			const node = parseXml(
				`<w:r xmlns:w="${WORD_NS}"><w:rPr><w:rPrChange w:id="1" w:author="Ada" ${attributes}><w:rPr/></w:rPrChange></w:rPr><w:t>Text</w:t></w:r>`,
			).documentElement;
			expect(runHasUnknownProperties(node)).toBe(true);
		},
	);
	it.each(['ins', 'del'] as const)(
		'preserves %s and formatting history in one run',
		async (kind) => {
			const loaded = await fixture(kind);
			expect(listRevisions(loaded.model).map((entry) => entry.kind)).toEqual([
				kind === 'ins' ? 'insert' : 'delete',
				'formatChange',
			]);
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const original = paragraph.runs[0]!;
			paragraph.runs = [
				{ ...original, text: 'Te' },
				{ ...original, text: 'xt' },
			];
			const reopened = await loadDocx(await loaded.save(loaded.model));
			for (const run of expectParagraph(reopened.model.blocks[0]).runs) {
				expect(run.revision).toMatchObject({
					kind: kind === 'ins' ? 'insert' : 'delete',
					author: 'Grace',
				});
				expect(run.formatRevision).toMatchObject({ kind: 'formatChange', author: 'Ada' });
				expect(run.formatRevision?.previousRunPropertiesXml).toContain('x:value="preserved"');
			}
			const xml = await (
				await JSZip.loadAsync(await loaded.save(loaded.model))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).not.toContain('text-edit');
			expect(xml).not.toContain('format-edit');
		},
	);

	it.each(['ins', 'del'] as const)(
		'resolves formatting independently of %s text history',
		async (kind) => {
			const loaded = await fixture(kind);
			const before = structuredClone(loaded.model);
			const accepted = acceptRevision(loaded.model, 'format-edit');
			const acceptedRun = expectParagraph(accepted.blocks[0]).runs[0]!;
			expect(acceptedRun.bold).toBe(true);
			expect(acceptedRun.formatRevision).toBeUndefined();
			expect(acceptedRun.revision?.id).toBe('text-edit');
			const rejected = rejectRevision(loaded.model, 'format-edit');
			const rejectedRun = expectParagraph(rejected.blocks[0]).runs[0]!;
			expect(rejectedRun.bold).toBeUndefined();
			expect(rejectedRun.italic).toBe(true);
			expect(rejectedRun.formatRevision).toBeUndefined();
			expect(rejectedRun.revision?.id).toBe('text-edit');
			const xml = await (
				await JSZip.loadAsync(await loaded.save(rejected))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).toContain('x:value="preserved"');
			expect(xml).not.toContain('rPrChange');
			expect(xml).toContain(`<w:${kind} `);
			expect(loaded.model).toEqual(before);
		},
	);

	it.each(['ins', 'del'] as const)(
		'retains formatting when the %s text revision is kept',
		async (kind) => {
			const loaded = await fixture(kind);
			const kept = (kind === 'ins' ? acceptRevision : rejectRevision)(loaded.model, 'text-edit');
			const run = expectParagraph(kept.blocks[0]).runs[0]!;
			expect(run.revision).toBeUndefined();
			expect(run.formatRevision?.id).toBe('format-edit');
			const reopened = await loadDocx(await loaded.save(kept));
			expect(listRevisions(reopened.model).map((entry) => entry.kind)).toEqual(['formatChange']);
		},
	);
});
