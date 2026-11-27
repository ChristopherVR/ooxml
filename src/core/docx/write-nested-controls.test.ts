// Saving blocks inside nested block content controls: a model that carries only a paragraph's own
// control (as an editor that does not track the ones around it) leaves the outer `w:sdt` alone, a
// new paragraph joins the level that holds its control, and no control id is written twice.
import { describe, expect, it } from 'vitest';
import {
	item,
	paragraph,
	run,
	saveAndReload,
	section,
	sdt,
	template,
} from './content-controls/test-fixture';
import type { DocumentModel, Paragraph } from './model';

const strip = (xml: string) => xml.replace(/\sxmlns:w15="[^"]*"/g, '');
const ids = (xml: string) => [...xml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((match) => match[1]);
const sectionBody = () =>
	sdt('sec', sdt(undefined, paragraph(run('A')) + paragraph(run('B')), item), section) +
	paragraph(run('After'));
const groupBody = () =>
	sdt('grp', sdt('inner', paragraph(run('Inner'))), '<w:group/>') + paragraph(run('After'));

async function save(body: string, edit: (blocks: Paragraph[]) => void) {
	const loaded = await template(body);
	const model = structuredClone(loaded.model) as DocumentModel;
	edit(model.blocks as Paragraph[]);
	return strip((await saveAndReload(loaded, model)).xml);
}
const own = (paragraph: Paragraph): Paragraph => {
	const { outerContentControls: _outer, ...rest } = paragraph;
	return rest;
};

describe('blocks inside nested content controls', () => {
	it('keeps the outer controls when the model carries only the paragraph own control', async () => {
		const xml = await save(sectionBody(), (blocks) => {
			blocks[0] = { ...own(blocks[0]!), runs: [{ text: 'A edited' }] };
		});
		expect(xml.match(/<w15:repeatingSection\b/g)).toHaveLength(1);
		expect(xml).toContain('A edited');
		const group = await save(groupBody(), (blocks) => {
			blocks[0] = { ...own(blocks[0]!), runs: [{ text: 'Inner 2' }] };
		});
		expect(group).toContain('<w:group/>');
		expect(group).toContain('Inner 2');
	});

	it('puts a new paragraph with only its own control inside the level that holds it', async () => {
		const xml = await save(sectionBody(), (blocks) => {
			blocks.splice(2, 0, {
				type: 'paragraph',
				id: 'p-new',
				runs: [{ text: 'pasted' }],
				contentControl: structuredClone(blocks[1]!.contentControl!),
			});
		});
		expect(xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(1);
		expect(ids(xml)).toEqual([...new Set(ids(xml))]);
		const item = xml.slice(xml.indexOf('<w15:repeatingSectionItem'), xml.indexOf('After'));
		expect(item).toMatch(/B<\/w:t>.*pasted/);
	});

	it('puts a new first paragraph inside the controls the next one shares', async () => {
		const xml = await save(sectionBody(), (blocks) => {
			blocks.unshift({
				type: 'paragraph',
				id: 'p-first',
				runs: [{ text: 'first' }],
				contentControl: structuredClone(blocks[0]!.contentControl!),
				outerContentControls: structuredClone(blocks[0]!.outerContentControls!),
			});
		});
		expect(xml.match(/<w15:repeatingSection\b/g)).toHaveLength(1);
		expect(xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(1);
		expect(xml).toMatch(/repeatingSectionItem\/?>.*first.*A<\/w:t>/);
	});

	it('gives a new control whose id another control has a new id', async () => {
		const xml = await save(sectionBody(), (blocks) => {
			const control = blocks[0]!.contentControl!;
			blocks.push({
				type: 'paragraph',
				id: 'p-end',
				runs: [{ text: 'end' }],
				contentControl: control,
			});
		});
		expect(ids(xml)).toHaveLength(3);
		expect(new Set(ids(xml)).size).toBe(3);
	});

	it('unwraps the outer controls only when the model says it has none', async () => {
		const xml = await save(sectionBody(), (blocks) => {
			for (const at of [0, 1]) {
				const { contentControl: _control, ...rest } = own(blocks[at]!);
				blocks[at] = { ...rest, outerContentControls: [] };
			}
		});
		expect(xml).not.toContain('<w:sdt>');
	});
});
