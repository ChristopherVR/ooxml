import { mkdirSync } from 'node:fs';
/**
 * Generates `chart-style-palette.pptx`: one slide with one clustered column
 * chart whose series carry NO explicit colour, a `<c:style val="2"/>` (Office's
 * default chart style) and no chart colour-style part (`colorsN.xml`), so the
 * series fills can only come from the chart-style palette over the deck theme.
 *
 * The theme's accent1 is set to a value no built-in palette uses
 * ({@link CHART_STYLE_PALETTE_ACCENT1}), so a renderer that ignores the theme
 * and paints a hard-coded Office palette fails the spec instead of passing by
 * coincidence. One category and seven series: style 2 paints series 1 to 6 in
 * accent1 to accent6 and series 7 in a darker accent1 (Office's gallery; see
 * `src/ui/src/pptx/render/chart-helpers.ts`), and with one category the seven
 * bars sit left to right in series order.
 *
 * Used by `chart-style-palette.spec.ts`. Re-runnable; globalSetup invokes it.
 */
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import type JSZipType from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

import {
	addChartRel,
	addContentTypeOverride,
	chartGraphicFrameXml,
	injectGraphicFrame,
} from './chart-part';
import { writeFixtureDeterministic } from './write-fixture';

const coreRequire = createRequire(createRequire(import.meta.url).resolve('pptx-viewer-core'));
const JSZip = coreRequire('jszip') as {
	loadAsync: (typeof JSZipType)['loadAsync'];
} & (new () => JSZipType);

const __dirname = dirname(fileURLToPath(import.meta.url));

/** The deck theme's accent1, without `#`. Not a colour any built-in palette uses. */
export const CHART_STYLE_PALETTE_ACCENT1 = '2E86AB';
/** The chart's title, used to find it on the slide. */
export const CHART_STYLE_PALETTE_TITLE = 'Style Palette';
/** Series count: one past the six accents, so series 7 is the first darker shade. */
export const CHART_STYLE_PALETTE_SERIES = 7;

const CAT_AX_ID = 50010;
const VAL_AX_ID = 50020;

function chartXml(): string {
	const series = Array.from({ length: CHART_STYLE_PALETTE_SERIES }, (_, i) => {
		const col = String.fromCodePoint(66 + i);
		return (
			`<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
			`<c:tx><c:strRef><c:f>Sheet1!$${col}$1</c:f><c:strCache><c:ptCount val="1"/>` +
			`<c:pt idx="0"><c:v>S${i + 1}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
			`<c:invertIfNegative val="0"/>` +
			`<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="1"/>` +
			`<c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
			`<c:val><c:numRef><c:f>Sheet1!$${col}$2</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
			`<c:ptCount val="1"/><c:pt idx="0"><c:v>${10 + i}</c:v></c:pt></c:numCache></c:numRef></c:val>` +
			`</c:ser>`
		);
	}).join('');
	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<c:style val="2"/>` +
		`<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${CHART_STYLE_PALETTE_TITLE}</a:t></a:r></a:p>` +
		`</c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` +
		`<c:plotArea><c:layout/>` +
		`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
		series +
		`<c:gapWidth val="50"/><c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/></c:barChart>` +
		`<c:catAx><c:axId val="${CAT_AX_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
		`<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>` +
		`<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
		`<c:crossAx val="${VAL_AX_ID}"/><c:crosses val="autoZero"/><c:auto val="1"/>` +
		`<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
		`<c:valAx><c:axId val="${VAL_AX_ID}"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling>` +
		`<c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/>` +
		`<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
		`<c:crossAx val="${CAT_AX_ID}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` +
		`</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
		`</c:chartSpace>`
	);
}

export async function generateChartStylePaletteFixture(): Promise<string> {
	const { handler, data, createSlide } = await PptxHandler.createBlank({
		title: 'Chart Style Palette Fixture',
		initialSlideCount: 0,
		theme: { colors: { accent1: `#${CHART_STYLE_PALETTE_ACCENT1}` } },
	});
	data.slides.push(
		createSlide('Blank')
			.addShape('rect', { x: 0, y: 0, width: 1, height: 1, fill: { type: 'none' } })
			.build(),
	);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));

	// Fail here, not in the spec, if the theme override did not reach the part.
	const themePath = Object.keys(zip.files).find((p) => /^ppt\/theme\/theme\d+\.xml$/u.test(p));
	const themeXml = themePath ? await zip.file(themePath)!.async('string') : '';
	if (
		!new RegExp(`<a:accent1>\s*<a:srgbClr val="${CHART_STYLE_PALETTE_ACCENT1}"`, 'iu').test(
			themeXml,
		)
	) {
		throw new Error(`theme accent1 is not ${CHART_STYLE_PALETTE_ACCENT1}`);
	}

	const chartPartName = 'ppt/charts/chart1.xml';
	zip.file(chartPartName, chartXml());
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const { xml: rels, rId } = addChartRel(
		await zip.file(relsPath)!.async('string'),
		'../charts/chart1.xml',
	);
	zip.file(relsPath, rels);
	const slidePath = 'ppt/slides/slide1.xml';
	zip.file(
		slidePath,
		injectGraphicFrame(await zip.file(slidePath)!.async('string'), chartGraphicFrameXml(rId)),
	);
	const ctPath = '[Content_Types].xml';
	zip.file(ctPath, addContentTypeOverride(await zip.file(ctPath)!.async('string'), chartPartName));

	const outPath = resolve(__dirname, 'chart-style-palette.pptx');
	mkdirSync(dirname(outPath), { recursive: true });
	await writeFixtureDeterministic(outPath, await zip.generateAsync({ type: 'uint8array' }));
	return outPath;
}

const invokedDirectly =
	typeof process !== 'undefined' &&
	process.argv[1] &&
	process.argv[1].endsWith('generate-chart-style-palette-fixture.ts');
if (invokedDirectly) {
	generateChartStylePaletteFixture()
		.then((p) => console.log(`Wrote ${p}`))
		.catch((err) => {
			console.error(err);
			process.exit(1);
		});
}
