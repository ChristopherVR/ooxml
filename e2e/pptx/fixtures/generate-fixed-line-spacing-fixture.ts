/**
 * Generates `fixed-line-spacing.pptx` (issue #35): three blocks of five 7pt
 * Arial lines, every paragraph at an exact 11pt line spacing
 * (`a:lnSpc/a:spcPts`), each line opening with a "•" text run:
 *
 *  - `Left`: a two-cell table's left cell, the "•" run at 9pt.
 *  - `Mid`: the same table's right cell, the "•" run at 7pt.
 *  - `Box`: a text box holding the left cell's paragraphs.
 *
 * PowerPoint keeps the exact pitch whatever runs a line holds, so all three
 * blocks put their baselines 11pt apart and level with each other (measured
 * on a PDF export of the reporter's deck). The browser used to grow every line
 * that held the larger run by a pixel.
 *
 * Run with: bun run e2e/fixtures/generate-fixed-line-spacing-fixture.ts
 */
import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

import { writeFixtureDeterministic } from './write-fixture';

const __dirname = dirname(fileURLToPath(import.meta.url));

/** The exact line pitch every paragraph authors, in points. */
export const FIXED_LINE_PITCH_PT = 11;

/** A block's name, the size of its "•" run in points, and its five lines. */
export interface FixedSpacingBlock {
	name: 'Left' | 'Mid' | 'Box';
	bulletPt: number;
	lines: readonly string[];
}

const linesOf = (name: string): string[] => [1, 2, 3, 4, 5].map((n) => `${name} line ${n} text`);

/** The blocks, shared with the spec. */
export const FIXED_SPACING_BLOCKS: readonly FixedSpacingBlock[] = [
	{ name: 'Left', bulletPt: 9, lines: linesOf('Left') },
	{ name: 'Mid', bulletPt: 7, lines: linesOf('Mid') },
	{ name: 'Box', bulletPt: 9, lines: linesOf('Box') },
];

const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
	'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const EMU_PER_PT = 12700;

function runXml(text: string, sizePt: number): string {
	return [
		`<a:r><a:rPr lang="en-US" sz="${sizePt * 100}" dirty="0">`,
		'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>',
		`<a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r>`,
	].join('');
}

/** The block's paragraphs: exact spacing, a "•" run, then the 7pt text. */
function paragraphsXml(block: FixedSpacingBlock): string {
	return block.lines
		.map((line) =>
			[
				`<a:p><a:pPr><a:lnSpc><a:spcPts val="${FIXED_LINE_PITCH_PT * 100}"/></a:lnSpc></a:pPr>`,
				runXml('•', block.bulletPt),
				runXml(` ${line}`, 7),
				'</a:p>',
			].join(''),
		)
		.join('');
}

function cellXml(block: FixedSpacingBlock): string {
	return [
		'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>',
		paragraphsXml(block),
		'</a:txBody><a:tcPr marL="0" marR="0" marT="0" marB="0"><a:noFill/></a:tcPr></a:tc>',
	].join('');
}

function tableXml(left: FixedSpacingBlock, mid: FixedSpacingBlock): string {
	const colPt = 150;
	return [
		'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/>',
		'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>',
		`<p:xfrm><a:off x="${40 * EMU_PER_PT}" y="${60 * EMU_PER_PT}"/>`,
		`<a:ext cx="${2 * colPt * EMU_PER_PT}" cy="${80 * EMU_PER_PT}"/></p:xfrm>`,
		'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>',
		'<a:tblPr firstRow="0" bandRow="0"/>',
		`<a:tblGrid><a:gridCol w="${colPt * EMU_PER_PT}"/><a:gridCol w="${colPt * EMU_PER_PT}"/></a:tblGrid>`,
		`<a:tr h="${80 * EMU_PER_PT}">${cellXml(left)}${cellXml(mid)}</a:tr>`,
		'</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
	].join('');
}

function textBoxXml(block: FixedSpacingBlock): string {
	return [
		'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>',
		`<p:spPr><a:xfrm><a:off x="${380 * EMU_PER_PT}" y="${60 * EMU_PER_PT}"/>`,
		`<a:ext cx="${150 * EMU_PER_PT}" cy="${80 * EMU_PER_PT}"/></a:xfrm>`,
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>',
		'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr>',
		`<a:lstStyle/>${paragraphsXml(block)}</p:txBody></p:sp>`,
	].join('');
}

function slideXml(): string {
	const [left, mid, box] = FIXED_SPACING_BLOCKS;
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		`<p:sld ${NS}><p:cSld><p:spTree>`,
		'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>',
		'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>',
		'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>',
		tableXml(left, mid),
		textBoxXml(box),
		'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
	].join('');
}

export async function generateFixedLineSpacingFixture(): Promise<string> {
	const { handler, data, createSlide } = await PptxHandler.createBlank({
		title: 'Fixed Line Spacing Fixture',
		initialSlideCount: 0,
	});
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	zip.file('ppt/slides/slide1.xml', slideXml());

	const outPath = resolve(__dirname, 'fixed-line-spacing.pptx');
	mkdirSync(dirname(outPath), { recursive: true });
	await writeFixtureDeterministic(outPath, await zip.generateAsync({ type: 'uint8array' }));
	return outPath;
}

if (process.argv[1]?.endsWith('generate-fixed-line-spacing-fixture.ts')) {
	generateFixedLineSpacingFixture()
		.then((path) => console.log(`Wrote ${path}`))
		.catch((error) => {
			console.error(error);
			process.exit(1);
		});
}
