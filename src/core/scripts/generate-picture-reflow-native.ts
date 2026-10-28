import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { withNativePictureLayout } from '../diagram/test-support/picture-layout';
import { PptxHandler } from '../pptx/core/PptxHandler';
import { relayoutSmartArt } from '../pptx/core/utils/smartart-relayout';
import { previewDiagramLayout } from '../diagram/layout-preview';
import { parseDiagramDrawing } from '../diagram/drawing';
import { parseRelationships, resolvePartPath } from '../opc/index';

const output = resolve(process.argv[2] ?? 'artifacts/smartart-picture-reflow');
mkdirSync(output, { recursive: true });
for (const layout of ['picture-accent-list--flat3', 'picture-grid--flat3']) {
	const original = readFileSync(
		new URL(`../pptx/__tests__/fixtures/smartart-gallery/${layout}.pptx`, import.meta.url),
	);
	const bytes = await withNativePictureLayout(original, 'ppt', layout);
	writeFileSync(resolve(output, `${layout}-before.pptx`), bytes);
	const handler = new PptxHandler();
	const { slides } = await handler.load(bytes);
	const element = slides[0]!.elements.find((item) => item.type === 'smartArt');
	if (element?.type !== 'smartArt') throw new Error(`No SmartArt in ${layout}`);
	const zip = await JSZip.loadAsync(bytes);
	for (const part of Object.keys(element.smartArtData!.pictureUrls ?? {})) {
		const image = await zip.file(part)?.async('base64');
		if (image) element.smartArtData!.pictureUrls![part] = `data:image/png;base64,${image}`;
	}
	const before = parseDiagramDrawing(await zip.file('ppt/diagrams/drawing1.xml')!.async('string'));
	before.extentEmu = { width: element.width * 9525, height: element.height * 9525 };
	before.pictureParts = Object.fromEntries(
		[
			...parseRelationships(
				await zip.file('ppt/diagrams/_rels/drawing1.xml.rels')!.async('string'),
			).values(),
		].map((rel) => [rel.id, resolvePartPath('ppt/diagrams/drawing1.xml', rel.target)]),
	);
	writeFileSync(
		resolve(output, `${layout}-before-drawing.json`),
		JSON.stringify(
			{
				width: element.width,
				height: element.height,
				drawing: before,
				pictureUrls: element.smartArtData!.pictureUrls,
				scheme: element.smartArtData!.pictureScheme,
			},
			null,
			2,
		),
	);
	for (const node of element.smartArtData!.nodes) node.text = `Edited ${node.text}`;
	element.smartArtData!.drawingShapes = relayoutSmartArt(
		element.smartArtData!,
		element.width,
		element.height,
	);
	element.smartArtData!.drawingDirty = true;
	writeFileSync(resolve(output, `${layout}-after.pptx`), await handler.save(slides));
	writeFileSync(
		resolve(output, `${layout}-model.json`),
		JSON.stringify(
			{ width: element.width, height: element.height, data: element.smartArtData },
			null,
			2,
		),
	);
	const neutral = previewDiagramLayout({
		nodes: element.smartArtData!.nodes.map((node) => ({
			id: node.id,
			text: node.text,
			parentId: node.parentId,
		})),
		source: element.smartArtData!.pictureLayoutSource,
		scheme: element.smartArtData!.pictureScheme,
		minorFont: element.smartArtData!.themeMinorFont,
		extentEmu: { width: element.width * 9525, height: element.height * 9525 },
	});
	writeFileSync(
		resolve(output, `${layout}-drawing.json`),
		JSON.stringify(
			{
				width: element.width,
				height: element.height,
				drawing: neutral.drawing,
				scheme: element.smartArtData!.pictureScheme,
				fonts: { minor: element.smartArtData!.themeMinorFont },
				pictureUrls: element.smartArtData!.pictureUrls,
				issues: neutral.issues,
			},
			null,
			2,
		),
	);
}
