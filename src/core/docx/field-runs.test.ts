import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index';
import { fieldDisplayText } from './layout/page-fields';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const marker = (kind: string) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const code = (instr: string) =>
	`<w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const inner = (result: string) =>
	`${marker('begin')}${code(' PA')}${code('GE ')}${marker('separate')}${result}${marker('end')}`;

async function loadParagraph(content: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p>${content}</w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const resultRuns = (paragraph: Paragraph) => paragraph.runs.filter((run) => run.text);

describe('nested complex field results', () => {
	it('tags the innermost result and resumes the enclosing result after its end', async () => {
		const loaded = await loadParagraph(
			`${marker('begin')}${code(' QUOTE ')}${marker('separate')}${text('before')}${inner(text('1'))}${text('after')}${marker('end')}${text('outside')}`,
		);
		expect(resultRuns(loaded.model.blocks[0] as Paragraph)).toEqual([
			{ text: 'before', field: { instr: 'QUOTE' } },
			{ text: '1', field: { instr: 'PAGE' } },
			{ text: 'after', field: { instr: 'QUOTE' } },
			{ text: 'outside' },
		]);
		const pageResult = (loaded.model.blocks[0] as Paragraph).runs.find((run) => run.text === '1');
		expect(fieldDisplayText(pageResult!, { page: '7', numPages: '9', sectionPages: '8' })).toBe(
			'7',
		);
	});

	it('keeps a nested field instruction separate from the enclosing instruction', async () => {
		const loaded = await loadParagraph(
			`${marker('begin')}${code(' IF ')}${inner(text('1'))}${code(' = 1 yes no ')}${marker('separate')}${text('yes')}${marker('end')}`,
		);
		expect(resultRuns(loaded.model.blocks[0] as Paragraph)).toEqual([
			{ text: '1', field: { instr: 'PAGE' } },
			{ text: 'yes', field: { instr: 'IF  = 1 yes no' } },
		]);
	});

	it('tags deleted nested cached results without losing their revision', async () => {
		const deleted = '<w:del w:id="7" w:author="Ann"><w:r><w:delText>1</w:delText></w:r></w:del>';
		const loaded = await loadParagraph(
			`${marker('begin')}${code(' QUOTE ')}${marker('separate')}${inner(deleted)}${text('after')}${marker('end')}`,
		);
		const runs = resultRuns(loaded.model.blocks[0] as Paragraph);
		expect(runs[0]).toMatchObject({
			text: '1',
			field: { instr: 'PAGE' },
			revision: { kind: 'delete', id: '7', author: 'Ann' },
		});
		expect(runs[1]).toEqual({ text: 'after', field: { instr: 'QUOTE' } });
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const editedRuns = paragraph.runs.map((run) => ({
			...run,
			text: run.text === '1' ? '2' : run.text,
		}));
		const saved = await loaded.save({
			...loaded.model,
			blocks: [{ ...paragraph, runs: editedRuns }],
		});
		const reloaded = await loadDocx(saved);
		expect(resultRuns(reloaded.model.blocks[0] as Paragraph)[0]).toMatchObject({
			text: '2',
			field: { instr: 'PAGE' },
			revision: { kind: 'delete', id: '7', author: 'Ann' },
		});
	});

	it('preserves nested field structure and metadata after editing, saving and reparsing', async () => {
		const loaded = await loadParagraph(
			`${marker('begin')}${code(' QUOTE ')}${marker('separate')}${text('before')}${inner(text('1'))}${text('after')}${marker('end')}`,
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const runs = paragraph.runs.map((run) => ({
			...run,
			text: run.text === '1' ? '2' : run.text === 'after' ? 'changed' : run.text,
		}));
		const saved = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
		const reloaded = await loadDocx(saved);
		const next = reloaded.model.blocks[0] as Paragraph;
		expect(resultRuns(next)).toEqual([
			{ text: 'before', field: { instr: 'QUOTE' } },
			{ text: '2', field: { instr: 'PAGE' } },
			{ text: 'changed', field: { instr: 'QUOTE' } },
		]);
		expect(next.runs.filter((run) => run.fieldChar).map((run) => run.fieldChar)).toEqual([
			'begin',
			'separate',
			'begin',
			'separate',
			'end',
			'end',
		]);
	});
});
