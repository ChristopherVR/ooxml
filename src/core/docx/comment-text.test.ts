import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseComments } from './comments';
import { loadDocx } from './parse';
import { patchCommentText } from './write-comment-text';
import { buildXml, parseXml, WORD_NS } from './xml';

const body =
	'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> lead </w:t>' +
	'<w:tab/><w:t>A</w:t><w:br w:clear="all"/><w:t>B</w:t><w:cr/>' +
	'<w:t>C</w:t><w:noBreakHyphen/><w:t>D</w:t><w:softHyphen/><w:t>E</w:t></w:r></w:p>' +
	'<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> end </w:t></w:r></w:p>';
const source = `<w:comments xmlns:w="${WORD_NS}"><w:comment w:id="0" w:author="Ada">${body}</w:comment></w:comments>`;
const original = ' lead \tA\nB\nC\u2011D\u00adE\n end ';

function edit(text: string) {
	const doc = parseXml(source);
	patchCommentText(doc.getElementsByTagNameNS(WORD_NS, 'comment')[0]!, text);
	const xml = buildXml(doc);
	expect(parseComments(xml)[0]!.text).toBe(text);
	return { doc, xml };
}

describe('comment text whitespace and inline characters', () => {
	it('reads native Word spaces, tabs and soft breaks and retains them through edits', async () => {
		const bytes = readFileSync(
			new URL('./__fixtures__/word-comment-whitespace-native.docx', import.meta.url),
		);
		const loaded = await loadDocx(new Uint8Array(bytes));
		expect(loaded.model.comments![0]!.text).toBe(original);
		const comments = loaded.model.comments!.map((item, index) =>
			index === 0 ? { ...item, text: item.text.replace('A', 'Alpha') } : item,
		);
		const saved = await loaded.save({ ...loaded.model, comments });
		expect((await loadDocx(saved)).model.comments).toEqual(comments);
	});

	it('reads tabs, soft breaks and hyphens without trimming significant spaces', () => {
		expect(parseComments(source)[0]!.text).toBe(original);
	});

	it('keeps untouched paragraphs byte-equivalent when another paragraph changes', () => {
		const before = parseXml(source);
		const { doc } = edit(original.replace('end', 'tail'));
		expect(buildXml(doc.getElementsByTagNameNS(WORD_NS, 'p')[0]!)).toBe(
			buildXml(before.getElementsByTagNameNS(WORD_NS, 'p')[0]!),
		);
	});

	it('edits around inline characters without changing their elements or run formatting', () => {
		const { doc } = edit(original.replace('A', 'Alpha').replace('B', 'Beta'));
		expect(doc.getElementsByTagNameNS(WORD_NS, 'p')).toHaveLength(2);
		expect(doc.getElementsByTagNameNS(WORD_NS, 'b')).toHaveLength(1);
		expect(doc.getElementsByTagNameNS(WORD_NS, 'i')).toHaveLength(1);
		expect(doc.getElementsByTagNameNS(WORD_NS, 'br')).toHaveLength(1);
		expect(doc.getElementsByTagNameNS(WORD_NS, 'br')[0]!.getAttributeNS(WORD_NS, 'clear')).toBe(
			'all',
		);
		expect(doc.getElementsByTagNameNS(WORD_NS, 'cr')).toHaveLength(1);
	});

	it.each([
		original.replace('\t', ''),
		original.replace('A\nB', 'AB'),
		original.replace('B\nC', 'BC'),
		original.replace('lead', 'a\tb'),
		original + '\nNew paragraph\twith a tab',
		'',
		'\t',
		'\n',
		'😀\t\u2011\u00ad',
	])('saves edited control characters as Word elements: %j', (text) => {
		const { doc } = edit(text);
		for (const node of Array.from(doc.getElementsByTagNameNS(WORD_NS, 't')))
			expect(node.textContent).not.toMatch(/[\t\n\u00ad\u2011]/);
	});

	it('round-trips comment text edits through a native package while preserving its reply', async () => {
		const bytes = readFileSync(
			new URL('./__fixtures__/word-comments-native.docx', import.meta.url),
		);
		const zip = await JSZip.loadAsync(bytes);
		const doc = parseXml(await zip.file('word/comments.xml')!.async('string'));
		const comment = doc.getElementsByTagNameNS(WORD_NS, 'comment')[0]!;
		const paragraph = comment.getElementsByTagNameNS(WORD_NS, 'p')[0]!;
		const imported = doc.importNode(
			parseXml(source).getElementsByTagNameNS(WORD_NS, 'p')[0]!,
			true,
		);
		comment.replaceChild(imported, paragraph);
		zip.file('word/comments.xml', buildXml(doc));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const edited = loaded.model.comments!.map((item, index) =>
			index === 0 ? { ...item, text: item.text.replace('A\nB', 'Alpha\nBeta') } : item,
		);
		const saved = await loaded.save({ ...loaded.model, comments: edited });
		expect((await loadDocx(saved)).model.comments).toEqual(edited);
		const savedXml = await (
			await JSZip.loadAsync(saved)
		)
			.file('word/comments.xml')!
			.async('string');
		expect(savedXml).toContain('<w:br w:clear="all"/>');
		expect(savedXml).toContain('<w:cr/>');
	});
});
