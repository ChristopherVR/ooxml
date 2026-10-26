import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { diagramsIn, formatDocxDiagram, loadDocx } from './index';
import { parseDiagramDataModel } from '../diagram/index';
import { NS, buildXml, elements, parseXml } from '../xml/index';

const fixture = () =>
	new Uint8Array(readFileSync(new URL('./__fixtures__/smartart-word.docx', import.meta.url)));

it('replaces a data-owned legacy cache relationship without damaging foreign relationship XML', async () => {
	const zip = await JSZip.loadAsync(fixture());
	const relationships = parseXml(await zip.file('word/_rels/document.xml.rels')!.async('string'));
	const original = elements(relationships.documentElement).find((entry) =>
		entry.getAttribute('Type')?.endsWith('/diagramDrawing'),
	)!;
	const id = original.getAttribute('Id')!;
	relationships.documentElement.removeChild(original);
	zip.file('word/_rels/document.xml.rels', buildXml(relationships));
	const data = (await zip.file('word/diagrams/data1.xml')!.async('string')).replace(
		`relId="${id}"`,
		'relId="legacyCache"',
	);
	zip.file('word/diagrams/data1.xml', data);
	zip.file(
		'word/diagrams/_rels/data1.xml.rels',
		`<Relationships xmlns="${NS.rels}" xmlns:x="urn:kept"><Relationship Id="legacyCache" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/><x:future value="keep"/></Relationships>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const model = structuredClone(loaded.model);
	const run = model.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.diagram)!;
	expect(run.image!.diagram!.parts.drawing?.ownerPart).toBe('word/diagrams/data1.xml');
	run.image!.diagram = formatDocxDiagram(run.image!.diagram!, {
		kind: 'colors',
		value: 'monochromatic2',
	});
	const bytes = await loaded.save(model);
	const reopened = diagramsIn((await loadDocx(bytes)).model.blocks)[0]!;
	expect(reopened.rendering).toBe('generated-drawing');
	const after = await JSZip.loadAsync(bytes);
	const legacyRels = await after.file('word/diagrams/_rels/data1.xml.rels')!.async('string');
	expect(legacyRels).toContain('<x:future value="keep"');
	expect(legacyRels).not.toContain('legacyCache');
	const savedData = parseDiagramDataModel(
		await after.file('word/diagrams/data1.xml')!.async('string'),
	);
	expect(savedData.drawingRelId).toBe(reopened.parts.drawing?.relId);
});
