import { URL as NodeURL } from 'node:url';
import { afterEach, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { loadDocx } from 'ooxml-core/docx';
import { loadXlsx } from 'ooxml-core/xlsx';
import { withDiagramPicture } from '../../../core/diagram/test-support/picture-drawing';
import { smartArtNodeView } from '../docx/smartart-node-view';
import { schema } from '../docx/schema';
import { paintSmartArt } from '../xlsx/grid/smartart';
import { createTestContext } from '../xlsx/grid/test-context';
import { renderDiagramDrawing } from './smartart-svg';
afterEach(() => document.body.replaceChildren());
const url =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP8zwACTGCSAQANHQEDgslx/wAAAABJRU5ErkJggg==';
it('Word and Excel host resolvers feed cropped picture fills to the same SmartArt painter', async () => {
	const loaded = await loadDocx(
		await withDiagramPicture(
			readFileSync(new NodeURL('../../../core/docx/__fixtures__/smartart.docx', import.meta.url)),
			'word',
		),
	);
	const image = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? block.runs.flatMap((run) => (run.image?.diagram ? [run.image] : []))
			: [],
	)[0]!;
	const diagram = image.diagram!;
	const node = schema.nodes.image.create({
		widthPx: 100,
		heightPx: 50,
		diagram: JSON.stringify(diagram),
	});
	const seen: string[] = [];
	const word = smartArtNodeView(node, diagram, loaded.model.theme, 'en', (part) => {
		seen.push(part);
		return url;
	});
	document.body.append(word.dom);
	const wordArt = (word.dom as HTMLElement).querySelector('office-ui-smartart')!;
	const wordImage = wordArt.shadowRoot!.querySelector('pattern image')!;
	expect(wordImage.getAttribute('href')).toBe(url);
	expect(wordImage.getAttribute('x')).toBe('-0.5');
	expect(wordImage.getAttribute('width')).toBe('2');
	expect(seen).toContain('word/media/shared-picture.png');
	const wb = await loadXlsx(
		await withDiagramPicture(
			readFileSync(
				new NodeURL('../../../core/xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url),
			),
			'xl',
		),
	);
	const object = wb.sheets[0]!.drawings.find((drawing) => drawing.kind === 'smartArt')!;
	if (object.kind !== 'smartArt') throw new Error('Expected diagram');
	const host = document.createElement('div');
	document.body.append(host);
	paintSmartArt(createTestContext(wb), host, object, wb.theme, { w: 100, h: 50 }, (part) => {
		seen.push(part);
		return url;
	});
	const excelArt = host.querySelector('office-ui-smartart')!;
	expect(excelArt.shadowRoot!.querySelector('pattern image')?.getAttribute('href')).toBe(url);
	expect(excelArt.shadowRoot!.querySelector('feFuncA')?.getAttribute('slope')).toBe('0.5');
	expect(seen).toContain('xl/media/shared-picture.png');
	const unsupported = structuredClone(object.diagram!);
	const shape = unsupported.shapes[0]!;
	if (shape.fill?.kind !== 'picture') throw new Error('Expected picture fill');
	shape.fill.unsupportedPicture = ['tile'];
	const report = renderDiagramDrawing(document, unsupported, {}, {}, () => url).report;
	expect(report.approximatedFills).toContain('picture:tile');
	shape.fill.unsupportedPicture = [];
	shape.fill.effects = [{ kind: 'unsupported', name: 'extLst', sourceXml: '<a:extLst/>' }];
	expect(
		renderDiagramDrawing(document, unsupported, {}, {}, () => url).report.unappliedEffects,
	).toContain('picture:extLst');
});
