import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { readCommentParts, W14_NS, W16CID_NS, W16CEX_NS } from './comment-parts';
import { commentParagraphId } from './comments';
import { loadDocx } from './parse';
import { buildXml, parseXml, WORD_NS } from './xml';

const native = () =>
	new Uint8Array(
		readFileSync(new URL('./__fixtures__/word-comments-native.docx', import.meta.url)),
	);

describe('native Word comment preservation', () => {
	it('reads multi-paragraph reply links, resolved state and UTC dates without changing bytes', async () => {
		const bytes = native();
		const loaded = await loadDocx(bytes);
		const parts = await readCommentParts(await JSZip.loadAsync(bytes));
		const date = parseXml(parts.extensible.xml!)
			.getElementsByTagNameNS(W16CEX_NS, 'commentExtensible')[0]!
			.getAttributeNS(W16CEX_NS, 'dateUtc');
		expect(loaded.model.comments).toMatchObject([
			{
				id: '0',
				author: 'Parity Author',
				text: 'Bold text and plain text.\nSecond paragraph.',
				resolved: true,
				date,
			},
			{ id: '1', author: 'Parity Reviewer', text: 'Reply text.', parentId: '0', resolved: true },
		]);
		expect(await loaded.save()).toEqual(bytes);
	});

	it('follows relationships to nonstandard comment and metadata part locations', async () => {
		const zip = await JSZip.loadAsync(native());
		const parts = await readCommentParts(zip);
		let rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		let types = await zip.file('[Content_Types].xml')!.async('string');
		for (const part of Object.values(parts)) {
			const name = part.path.slice('word/'.length);
			zip.file(`word/review/${name}`, part.xml!);
			zip.remove(part.path);
			rels = rels.replace(`Target="${name}"`, `Target="review/${name}"`);
			types = types.replace(`PartName="/${part.path}"`, `PartName="/word/review/${name}"`);
		}
		zip.file('word/_rels/document.xml.rels', rels);
		zip.file('[Content_Types].xml', types);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.model.comments?.[0]?.resolved).toBe(true);
		expect(loaded.model.comments?.[1]?.parentId).toBe('0');
		const saved = await loaded.save({
			...loaded.model,
			comments: loaded.model.comments!.map((comment) => ({ ...comment, resolved: false })),
		});
		const savedZip = await JSZip.loadAsync(saved);
		expect(savedZip.file('word/comments.xml')).toBeNull();
		expect((await loadDocx(saved)).model.comments?.[0]?.resolved).toBe(false);
	});

	it('changes resolved state without rebuilding formatted comment bodies or persistent IDs', async () => {
		const bytes = native();
		const original = await readCommentParts(await JSZip.loadAsync(bytes));
		const loaded = await loadDocx(bytes);
		const saved = await loaded.save({
			...loaded.model,
			comments: loaded.model.comments!.map((comment) => ({ ...comment, resolved: false })),
		});
		const parts = await readCommentParts(await JSZip.loadAsync(saved));
		expect(parts.comments.xml).toBe(original.comments.xml);
		expect(parts.ids.xml).toBe(original.ids.xml);
		expect(parts.extensible.xml).toBe(original.extensible.xml);
		expect((await loadDocx(saved)).model.comments).toMatchObject([
			{ resolved: false },
			{ resolved: false, parentId: '0' },
		]);
	});

	it('edits reply text without changing its root or identity metadata', async () => {
		const bytes = native();
		const original = await readCommentParts(await JSZip.loadAsync(bytes));
		const loaded = await loadDocx(bytes);
		const saved = await loaded.save({
			...loaded.model,
			comments: loaded.model.comments!.map((comment) =>
				comment.id === '1' ? { ...comment, text: 'Updated reply text.' } : comment,
			),
		});
		const parts = await readCommentParts(await JSZip.loadAsync(saved));
		const root = (xml: string) =>
			buildXml(parseXml(xml).getElementsByTagNameNS(WORD_NS, 'comment')[0]!);
		expect(root(parts.comments.xml!)).toBe(root(original.comments.xml!));
		for (const kind of ['ids', 'extended', 'extensible'] as const)
			expect(parts[kind].xml).toBe(original[kind].xml);
		expect((await loadDocx(saved)).model.comments?.[1]?.text).toBe('Updated reply text.');
	});

	it('keeps run formatting and the last-paragraph identity when paragraphs are added and removed', async () => {
		const loaded = await loadDocx(native());
		const original = await readCommentParts(await JSZip.loadAsync(native()));
		const originalId = commentParagraphId(
			parseXml(original.comments.xml!).getElementsByTagNameNS(WORD_NS, 'comment')[0]!,
		);
		for (const text of [
			'Bold revised text and plain text.\nSecond paragraph.\nNew paragraph.',
			'Bold revised text.',
		]) {
			const bytes = await loaded.save({
				...loaded.model,
				comments: loaded.model.comments!.map((comment) =>
					comment.id === '0' ? { ...comment, text } : comment,
				),
			});
			const parts = await readCommentParts(await JSZip.loadAsync(bytes));
			const root = parseXml(parts.comments.xml!).getElementsByTagNameNS(WORD_NS, 'comment')[0]!;
			expect(commentParagraphId(root)).toBe(originalId);
			expect(root.getElementsByTagNameNS(WORD_NS, 'b').length).toBeGreaterThan(0);
			const paragraphs = Array.from(root.getElementsByTagNameNS(WORD_NS, 'p'));
			const ids = paragraphs.map((paragraph) => paragraph.getAttributeNS(W14_NS, 'paraId'));
			expect(new Set(ids).size).toBe(ids.length);
			expect(paragraphs.length).toBe(text.split('\n').length);
			expect(parts.ids.xml).toBe(original.ids.xml);
			expect((await loadDocx(bytes)).model.comments).toMatchObject([{ text }, { parentId: '0' }]);
		}
	});

	it('prunes deleted comment metadata and anchors, promoting retained replies', async () => {
		const loaded = await loadDocx(native());
		const saved = await loaded.save({ ...loaded.model, comments: [loaded.model.comments![1]!] });
		const zip = await JSZip.loadAsync(saved);
		const parts = await readCommentParts(zip);
		expect(parseXml(parts.ids.xml!).getElementsByTagNameNS(W16CID_NS, 'commentId')).toHaveLength(1);
		expect(parts.extended.xml).not.toContain('paraIdParent');
		const document = parseXml(await zip.file('word/document.xml')!.async('string'));
		for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference'])
			expect(
				Array.from(document.getElementsByTagNameNS(WORD_NS, name)).some(
					(node) => node.getAttributeNS(WORD_NS, 'id') === '0',
				),
			).toBe(false);
		expect((await loadDocx(saved)).model.comments?.[0]?.parentId).toBeUndefined();
		const emptied = await loaded.save({ ...loaded.model, comments: [] });
		const emptyParts = await readCommentParts(await JSZip.loadAsync(emptied));
		for (const part of Object.values(emptyParts)) expect(part.xml).toBeDefined();
		expect((await loadDocx(emptied)).model.comments).toEqual([]);
	});

	it('serializes opaque editor IDs as decimal IDs and remaps their document anchors', async () => {
		const loaded = await loadDocx(native());
		const comment = {
			id: 'editor-comment-new',
			author: 'Parity Author',
			text: 'A new comment.',
			date: '2026-10-31T00:00:00Z',
		};
		const blocks = loaded.model.blocks.map((block, index) =>
			block.type === 'paragraph' && index === 0
				? {
						...block,
						runs: block.runs.map((run) => ({
							...run,
							commentIds: [...(run.commentIds ?? []), comment.id],
						})),
					}
				: block,
		);
		const reply = {
			id: 'editor-reply-new',
			parentId: '0',
			author: 'Parity Author',
			text: 'A new reply.',
		};
		const saved = await loaded.save({
			...loaded.model,
			blocks,
			comments: [...loaded.model.comments!, comment, reply],
		});
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.comments?.[2]).toMatchObject({ ...comment, id: '2' });
		expect(reloaded.model.comments?.[3]).toMatchObject({ ...reply, id: '3' });
		const first = reloaded.model.blocks[0]!;
		expect(first.type).toBe('paragraph');
		if (first.type === 'paragraph')
			expect(first.runs.some((run) => run.commentIds?.includes('2'))).toBe(true);
		const document = parseXml(
			await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string'),
		);
		for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference'])
			expect(
				Array.from(document.getElementsByTagNameNS(WORD_NS, name)).some(
					(node) => node.getAttributeNS(WORD_NS, 'id') === '3',
				),
			).toBe(true);
	});

	it('preserves unknown source attributes and metadata extensions across comment edits', async () => {
		const zip = await JSZip.loadAsync(native());
		const parts = await readCommentParts(zip);
		zip.file(
			parts.comments.path,
			parts.comments.xml!.replace(
				'<w:comment ',
				'<w:comment xmlns:custom="urn:parity" custom:flag="keep" ',
			),
		);
		zip.file(
			parts.ids.path,
			parts.ids.xml!.replace(
				'</w16cid:commentsIds>',
				'<custom:future xmlns:custom="urn:parity">keep</custom:future></w16cid:commentsIds>',
			),
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const saved = await loaded.save({
			...loaded.model,
			comments: loaded.model.comments!.map((comment) => ({ ...comment, author: 'Edited Author' })),
		});
		const result = await readCommentParts(await JSZip.loadAsync(saved));
		expect(result.comments.xml).toContain('custom:flag="keep"');
		expect(result.ids.xml).toContain('<custom:future');
	});
});
