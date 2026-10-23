import { beforeAll, describe, expect, it } from 'vitest';

import { parseDrawingRichTextBody } from '../../../drawingml/index';
import { NS, parseXml } from '../../../xml/index';
import {
	fixtureSlideTextBodies,
	fixtureTextBodies,
	type FixtureTextBody,
} from '../../../drawingml/test-support/text-fixtures';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import type { XmlObject } from '../../core/types';
import {
	neutralBody,
	neutralParagraph,
	neutralRun,
	pptxBody,
	pptxParagraph,
	pptxRun,
} from './smartart-text-neutral-parity-project';

/**
 * SmartArt text parity harness (docs/agnostic-core-plan.md, SmartArt step S1): every text body of
 * the SmartArt data (`dgm:t`) and cached drawing (`dsp:txBody`) parts of the committed pptx, docx
 * and xlsx fixtures is read by BOTH the pptx object-tree parsers (`text-run-properties-parser`,
 * `paragraph-properties-parser`, `body-properties-parser`) and the neutral DrawingML reader
 * (`parseDrawingRichTextBody`), and the fields they share must agree once converted to one unit.
 *
 * Fields only one side reads (the neutral model's non-sRGB fills, outlines, list styles; the pptx
 * resolved theme colours) are not compared. Tab-stop leaders are a pptx extension outside
 * CT_TextTabStop, which the neutral model does not read.
 */

const parser = new PptxRuntimeDependencyFactory().createParser();

function local(key: string): string {
	return key.split(':').pop() ?? key;
}

function list(value: unknown): XmlObject[] {
	if (value === '') return [{}];
	return (Array.isArray(value) ? value : value ? [value] : []) as XmlObject[];
}

function entry(node: XmlObject, name: string): unknown {
	const key = Object.keys(node).find((candidate) => local(candidate) === name);
	return key ? node[key] : undefined;
}

/** `modelId` -> text body object for every `dgm:pt`/`dsp:sp` of a parsed part. */
function objectBodies(tree: unknown, into = new Map<string, XmlObject>()): Map<string, XmlObject> {
	if (!tree || typeof tree !== 'object') return into;
	for (const node of Array.isArray(tree) ? tree : [tree]) {
		if (!node || typeof node !== 'object') continue;
		const object = node as XmlObject;
		const nvSpPr = entry(object, 'nvSpPr') as XmlObject | undefined;
		const cNvPr = nvSpPr ? (entry(nvSpPr, 'cNvPr') as XmlObject | undefined) : undefined;
		const modelId = object['@_modelId'] ?? cNvPr?.['@_id'];
		const body = entry(object, 'txBody') ?? entry(object, 't');
		if (modelId !== undefined && body !== undefined && typeof body === 'object')
			into.set(String(modelId), body as XmlObject);
		for (const [key, value] of Object.entries(object))
			if (!key.startsWith('@_')) objectBodies(value, into);
	}
	return into;
}

/** The `modelId` (SmartArt) or `p:cNvPr/@id` (slide shape) of a text body's owner. */
function bodyKey(element: Element): string | undefined {
	const owner = element.parentNode as Element | null;
	const modelId = owner?.getAttribute('modelId');
	if (modelId) return modelId;
	const nvSpPr = owner
		? Array.from(owner.childNodes).find((n) => (n as Element).localName === 'nvSpPr')
		: undefined;
	const cNvPr = nvSpPr
		? Array.from(nvSpPr.childNodes).find((n) => (n as Element).localName === 'cNvPr')
		: undefined;
	return (cNvPr as Element | undefined)?.getAttribute('id') ?? undefined;
}

function compareAll(bodies: FixtureTextBody[]): Comparison[] {
	const objectsByPart = new Map<string, Map<string, XmlObject>>();
	const differences: Comparison[] = [];
	for (const body of bodies) {
		const key = `${body.file}!${body.part}`;
		let objects = objectsByPart.get(key);
		if (!objects) {
			objects = objectBodies(parser.parse(body.partXml));
			objectsByPart.set(key, objects);
		}
		const difference = compare(body, objects);
		if (difference) differences.push(difference);
	}
	return differences;
}

interface Comparison {
	label: string;
	pptx: unknown;
	neutral: unknown;
}

/** Number of non-empty fields compared, so the harness cannot pass vacuously. */
let fieldsCompared = 0;

function countFields(value: unknown): number {
	if (!value || typeof value !== 'object') return 1;
	return Object.values(value).reduce<number>((sum, entry) => sum + countFields(entry), 0);
}

function compare(body: FixtureTextBody, objects: Map<string, XmlObject>): Comparison | undefined {
	const object = objects.get(bodyKey(body.element) ?? '');
	if (!object) return { label: body.label, pptx: 'missing', neutral: bodyKey(body.element) };
	const neutral = parseDrawingRichTextBody(body.element)!;
	const paragraphs = list(entry(object, 'p'));
	const pptxSide = {
		body: pptxBody(entry(object, 'bodyPr') as XmlObject | undefined),
		paragraphs: paragraphs.map((paragraph) => ({
			properties: pptxParagraph(entry(paragraph, 'pPr') as XmlObject | undefined),
			runs: list(entry(paragraph, 'r')).map((run) =>
				pptxRun(entry(run, 'rPr') as XmlObject | undefined),
			),
		})),
	};
	const neutralSide = {
		body: neutralBody(neutral.bodyProperties),
		paragraphs: neutral.paragraphs.map((paragraph) => ({
			properties: neutralParagraph(paragraph.properties),
			runs: paragraph.items.flatMap((item) =>
				item.kind === 'run' ? [neutralRun(item.properties)] : [],
			),
		})),
	};
	fieldsCompared += countFields(pptxSide);
	return JSON.stringify(pptxSide) === JSON.stringify(neutralSide)
		? undefined
		: { label: body.label, pptx: pptxSide, neutral: neutralSide };
}

describe('SmartArt text: pptx parsers and the neutral DrawingML text model agree', () => {
	let bodies: FixtureTextBody[] = [];
	beforeAll(async () => {
		bodies = await fixtureTextBodies();
	}, 120_000);

	it('covers the committed SmartArt parts of every product', () => {
		expect(bodies.length).toBeGreaterThan(500);
		expect(new Set(bodies.map((body) => body.file.split('.').pop()))).toEqual(
			new Set(['pptx', 'docx', 'xlsx']),
		);
	});

	it('reads the same body, paragraph and run properties', () => {
		fieldsCompared = 0;
		const differences = compareAll(bodies);
		expect(differences.slice(0, 3)).toEqual([]);
		expect(differences.length).toBe(0);
		expect(fieldsCompared).toBeGreaterThan(5000);
	});

	it('agrees on the richer slide text of the top-level decks too', async () => {
		// Shapes without a `p:cNvPr` id (table cells) cannot be paired and are skipped.
		const slides = (await fixtureSlideTextBodies()).filter((body) => bodyKey(body.element));
		fieldsCompared = 0;
		const differences = compareAll(slides);
		expect(differences.slice(0, 3)).toEqual([]);
		expect(differences.length).toBe(0);
		expect(fieldsCompared).toBeGreaterThan(1000);
	}, 120_000);

	it('agrees on the edge cases the corpus does not carry', () => {
		const partXml = `<dgm:dataModel xmlns:dgm="${NS.dgm}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><dgm:ptLst><dgm:pt modelId="1"><dgm:t>
			<a:bodyPr rot="-5400000" upright="true" compatLnSpc="0"/>
			<a:p><a:pPr lvl="12" algn="just" rtl="true" marR="9525" eaLnBrk="0" fontAlgn="ctr" hangingPunct="1">
				<a:lnSpc><a:spcPct val="650%"/></a:lnSpc><a:spcAft><a:spcPts val="1200"/></a:spcAft>
				<a:buClr><a:srgbClr val="00FF00"/></a:buClr><a:buSzPts val="1400"/><a:buFont typeface="Wingdings"/><a:buNone/>
			</a:pPr>
			<a:r><a:rPr b="true" i="1" u="dbl" strike="dblStrike" cap="all" baseline="-25000" spc="300" err="1" noProof="true">
				<a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface=" Segoe UI " pitchFamily="34" charset="0"/>
				<a:sym typeface="Symbol"/><a:hlinkClick r:id="rId9"/></a:rPr><a:t>x</a:t></a:r>
			<a:r><a:rPr u="none" strike="noStrike" baseline="0"/><a:t>y</a:t></a:r></a:p>
			<a:p><a:pPr><a:buAutoNum type="romanUcPeriod"/></a:pPr></a:p>
			<a:p><a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:buChar char="o"/></a:pPr></a:p>
		</dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>`;
		const element = parseXml(partXml).getElementsByTagNameNS(NS.dgm, 't')[0]!;
		const body: FixtureTextBody = {
			label: 'inline',
			partXml,
			part: 'inline',
			file: 'inline',
			index: 0,
			element,
		};
		expect(compareAll([body])).toEqual([]);
	});
});
