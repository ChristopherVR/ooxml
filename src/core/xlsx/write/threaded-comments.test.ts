import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS, buildXml, children, first, parseXml } from '../../xml/index';
import { parseRange } from '../address';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from './index';
import {
	bytes,
	mentioned,
	mentionId,
	threadFiles,
	messages,
	id,
	sourceRoot,
	metadataFixture,
} from '../__fixtures__/threaded-comment-metadata';

const at = { row: 1, col: 1 };
describe('threaded comment metadata', () => {
	it('matches differently cased person references without conflating same-name accounts', async () => {
		const zip = await JSZip.loadAsync(await metadataFixture());
		const root = await sourceRoot(zip);
		const nodes = messages(root);
		const personId = nodes[2]!.getAttribute('personId')!;
		nodes[2]!.setAttribute('personId', personId.toLowerCase());
		zip.file(threadFiles(zip)[0]!.name, buildXml(root.ownerDocument!));
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(workbook.sheets[0]!.comments[0]!.replies![1]!.author).toBe('Parity Author');
		createEditSession(workbook).setCommentReply(0, at, 1, 'Edited reply');
		const saved = messages(await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook))));
		expect(saved[2]!.getAttribute('personId')).toBe(personId);
	});

	it('reads replies before their parent and retains a reply-to-reply parent when editing', async () => {
		const zip = await JSZip.loadAsync(bytes);
		const root = await sourceRoot(zip);
		const nodes = messages(root);
		nodes[1]!.setAttribute('dT', '2026-01-01T00:00:01Z');
		nodes[2]!.setAttribute('dT', '2026-01-01T00:00:02Z');
		nodes[2]!.setAttribute('parentId', id(nodes[1]!)!.toLowerCase());
		root.insertBefore(nodes[2]!, nodes[0]!);
		root.insertBefore(nodes[1]!, nodes[0]!);
		zip.file(threadFiles(zip)[0]!.name, buildXml(root.ownerDocument!));
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(workbook.sheets[0]!.comments).toHaveLength(1);
		expect(workbook.sheets[0]!.comments[0]!.replies!.map((reply) => reply.text)).toEqual([
			'First reply',
			'Second reply',
		]);
		createEditSession(workbook).setCommentReply(0, at, 1, 'Edited nested reply');
		const saved = messages(await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook))));
		expect(saved[2]!.getAttribute('parentId')).toBe(id(nodes[1]!));
	});

	it('carries referenced accounts when a thread is pasted into a different workbook', async () => {
		const source = await loadXlsx(await metadataFixture());
		const destination = createWorkbook();
		createEditSession(destination).paste(
			0,
			at,
			createEditSession(source).copy(0, parseRange('B2')!),
		);
		const zip = await JSZip.loadAsync(await saveXlsx(destination));
		const people = children(
			parseXml(await zip.file('xl/persons/person.xml')!.async('string')).documentElement,
			'person',
			NS.tc,
		);
		const ids = new Set(people.map(id));
		expect(people.map((node) => node.getAttribute('userId')).sort()).toEqual([
			'original-account',
			'second-account',
		]);
		const root = await sourceRoot(zip);
		for (const message of messages(root))
			expect(ids.has(message.getAttribute('personId'))).toBe(true);
		for (const mention of Array.from(root.getElementsByTagNameNS(NS.tc, 'mention')))
			expect(ids.has(mention.getAttribute('mentionpersonId'))).toBe(true);
		expect(root.getElementsByTagNameNS('urn:parity', 'kept').length).toBe(1);
	});

	it('duplicates a sheet before its source with new identities and the original envelope', async () => {
		const workbook = await loadXlsx(await metadataFixture());
		const sourceId = id(parseXml(workbook.sheets[0]!.comments[0]!.threadedXml!).documentElement);
		const edit = createEditSession(workbook);
		edit.moveSheet(edit.duplicateSheet(0), 0);
		const saved = await saveXlsx(workbook);
		const back = await loadXlsx(saved);
		const rootIds = back.sheets.map((sheet) =>
			id(parseXml(sheet.comments[0]!.threadedXml!).documentElement),
		);
		expect(rootIds[1]).toBe(sourceId);
		expect(rootIds[0]).not.toBe(sourceId);
		const zip = await JSZip.loadAsync(saved);
		const roots = await Promise.all(
			threadFiles(zip).map(async (file) => parseXml(await file.async('string')).documentElement),
		);
		expect(new Set(roots.flatMap((root) => messages(root).map(id))).size).toBe(6);
		expect(
			new Set(
				roots.flatMap((root) =>
					Array.from(root.getElementsByTagNameNS(NS.tc, 'mention')).map((node) =>
						node.getAttribute('mentionId'),
					),
				),
			).size,
		).toBe(2);
		for (const root of roots)
			expect(root.getElementsByTagNameNS('urn:parity', 'kept').length).toBe(2);
	});

	it('keeps message identities through inserted rows and a cut to another cell', async () => {
		const workbook = await loadXlsx(await metadataFixture());
		const sourceId = id(parseXml(workbook.sheets[0]!.comments[0]!.threadedXml!).documentElement);
		const edit = createEditSession(workbook);
		edit.insertRows(0, 0, 1);
		edit.paste(0, { row: 0, col: 0 }, edit.cut(0, parseRange('B3')!));
		const root = await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook)));
		expect(id(messages(root)[0]!)).toBe(sourceId);
		for (const node of messages(root)) expect(node.getAttribute('ref')).toBe('A1');
	});

	it('carries person records verbatim on an unchanged save', async () => {
		const input = await metadataFixture();
		const before = await JSZip.loadAsync(input);
		const after = await JSZip.loadAsync(await saveXlsx(await loadXlsx(input)));
		expect(await after.file('xl/persons/person.xml')!.async('string')).toBe(
			await before.file('xl/persons/person.xml')!.async('string'),
		);
	});

	it('keeps identities, resolved state, opaque XML and same-name account identities while moving a mention', async () => {
		const input = await metadataFixture();
		const before = await JSZip.loadAsync(input);
		const source = messages(await sourceRoot(before));
		const workbook = await loadXlsx(input);
		const original = structuredClone(workbook.sheets[0]!.comments);
		const edit = createEditSession(workbook);
		const text = `😀 Prefix Hello ${mentioned}! Tail`;
		edit.setComment(0, at, text, 'Parity Author');
		edit.setCommentReply(0, at, 1, 'Edited reply');
		const output = await saveXlsx(workbook);
		const after = await JSZip.loadAsync(output);
		const root = await sourceRoot(after);
		const saved = messages(root);
		expect(saved.map(id)).toEqual(source.map(id));
		expect(saved.map((node) => node.getAttribute('personId'))).toEqual(
			source.map((node) => node.getAttribute('personId')),
		);
		expect(saved[0]!.getAttribute('done')).toBe('1');
		const mention = children(first(saved[0]!, 'mentions', NS.tc)!, 'mention', NS.tc)[0]!;
		expect(mention.getAttribute('mentionId')).toBe(mentionId);
		expect(mention.getAttribute('startIndex')).toBe(String(text.indexOf(mentioned)));
		expect(root.getElementsByTagNameNS('urn:parity', 'kept').length).toBe(2);
		expect(await after.file('xl/persons/person.xml')!.async('string')).toBe(
			await before.file('xl/persons/person.xml')!.async('string'),
		);
		expect((await loadXlsx(output)).sheets[0]!.comments[0]!.resolved).toBe(true);
		expect(workbook.sheets[0]!.comments[0]!.threadedXml).toBe(original[0]!.threadedXml);
		const legacy = await after.file(/^xl\/comments\d+\.xml$/)[0]!.async('string');
		expect(legacy).toContain(`tc=${id(saved[0]!)}`);
		expect(legacy).toContain(`xr:uid="${id(saved[0]!)}"`);
	});

	it('removes a mention whose text changed and restores it with undo', async () => {
		const workbook = await loadXlsx(await metadataFixture());
		const edit = createEditSession(workbook);
		edit.setComment(0, at, 'Hello someone!', 'Parity Author');
		let root = await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook)));
		expect(first(messages(root)[0]!, 'mentions', NS.tc)).toBeUndefined();
		edit.undo();
		root = await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook)));
		expect(first(messages(root)[0]!, 'mentions', NS.tc)).toBeDefined();
	});

	it('appends a new author without changing the source person records', async () => {
		const workbook = await loadXlsx(await metadataFixture());
		createEditSession(workbook).setComment(0, at, 'Changed author', 'New author');
		const zip = await JSZip.loadAsync(await saveXlsx(workbook));
		const root = parseXml(await zip.file('xl/persons/person.xml')!.async('string')).documentElement;
		const persons = children(root, 'person', NS.tc);
		expect(persons.map((person) => person.getAttribute('userId'))).toEqual([
			'original-account',
			'second-account',
			'New author',
		]);
		expect(root.getElementsByTagNameNS('urn:parity', 'kept').length).toBe(2);
	});

	it('pastes a copy before its source without reusing any message or mention identity', async () => {
		const workbook = await loadXlsx(await metadataFixture());
		const sourceId = id(parseXml(workbook.sheets[0]!.comments[0]!.threadedXml!).documentElement);
		const edit = createEditSession(workbook);
		edit.paste(0, { row: 0, col: 0 }, edit.copy(0, parseRange('B2')!));
		const root = await sourceRoot(await JSZip.loadAsync(await saveXlsx(workbook)));
		const all = messages(root);
		expect(new Set(all.map(id)).size).toBe(6);
		expect(
			all
				.find((node) => node.getAttribute('ref') === 'B2' && !node.hasAttribute('parentId'))!
				.getAttribute('id'),
		).toBe(sourceId);
		const mentions = Array.from(root.getElementsByTagNameNS(NS.tc, 'mention'));
		expect(new Set(mentions.map((node) => node.getAttribute('mentionId'))).size).toBe(2);
	});

	it('resolves and reopens a thread with undo without changing its source identities', async () => {
		const workbook = await loadXlsx(bytes);
		const edit = createEditSession(workbook);
		edit.setCommentResolved(0, at, true);
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.comments[0]!.resolved).toBe(true);
		edit.setCommentResolved(0, at, false);
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.comments[0]!.resolved).toBe(false);
		edit.undo();
		expect(workbook.sheets[0]!.comments[0]!.resolved).toBe(true);
	});
});
