import { beforeAll, describe, expect, it } from 'vitest';

import { NS, parseXml } from '../xml/index';
import { drawingParagraphText, parseDrawingRichTextBody } from './read-text-body';
import { drawingRichTextBodyXml } from './write-text-body';
import type { DrawingRichTextBody } from './text-model';
import {
	fixtureSlideTextBodies,
	fixtureTextBodies,
	type FixtureTextBody,
} from './test-support/text-fixtures';

/** Writes a body and reads it back. */
function roundTrip(body: DrawingRichTextBody): DrawingRichTextBody {
	const xml = drawingRichTextBodyXml(body);
	return parseDrawingRichTextBody(parseXml(xml).documentElement)!;
}

/** The kept XML differs only by the namespace declarations the writer drops. */
function normaliseKept<T>(value: T): T {
	return JSON.parse(JSON.stringify(value).replace(/ xmlns:(?:a|r)=\\"[^"\\]*\\"/gu, '')) as T;
}

describe('DrawingML rich text body round trip on committed SmartArt parts', () => {
	let bodies: FixtureTextBody[] = [];
	beforeAll(async () => {
		bodies = await fixtureTextBodies();
	}, 120_000);

	it('finds text bodies in pptx, docx and xlsx SmartArt parts', () => {
		expect(bodies.length).toBeGreaterThan(500);
		expect(bodies.some((body) => body.file.endsWith('.docx'))).toBe(true);
		expect(bodies.some((body) => body.file.endsWith('.xlsx'))).toBe(true);
		expect(bodies.some((body) => body.part.includes('/drawing'))).toBe(true);
		expect(bodies.some((body) => body.part.includes('/data'))).toBe(true);
	});

	it('reads, writes and reads every body to the same model', () => {
		const failures: string[] = [];
		for (const body of bodies) {
			const parsed = parseDrawingRichTextBody(body.element)!;
			const again = roundTrip(parsed);
			try {
				expect(normaliseKept(again)).toEqual(normaliseKept(parsed));
			} catch {
				failures.push(body.label);
			}
		}
		expect(failures).toEqual([]);
	});

	it('keeps the paragraph text of every body', () => {
		for (const body of bodies) {
			const parsed = parseDrawingRichTextBody(body.element)!;
			const expected = Array.from(body.element.getElementsByTagNameNS(NS.a, 'p')).length;
			expect(parsed.paragraphs.length, body.label).toBe(expected);
			const text = roundTrip(parsed).paragraphs.map(drawingParagraphText);
			expect(text, body.label).toEqual(parsed.paragraphs.map(drawingParagraphText));
		}
	});
});

describe('DrawingML rich text body round trip on committed slide text', () => {
	it('reads, writes and reads every slide, layout and master body to the same model', async () => {
		const bodies = await fixtureSlideTextBodies();
		expect(bodies.length).toBeGreaterThan(100);
		const failures: string[] = [];
		for (const body of bodies) {
			const parsed = parseDrawingRichTextBody(body.element)!;
			if (
				JSON.stringify(normaliseKept(roundTrip(parsed))) !== JSON.stringify(normaliseKept(parsed))
			)
				failures.push(body.label);
		}
		expect(failures).toEqual([]);
	}, 120_000);
});

describe('DrawingML rich text body model', () => {
	const xml = `<a:txBody xmlns:a="${NS.a}" xmlns:r="${NS.r}">
		<a:bodyPr rot="5400000" vert="vert270" wrap="square" lIns="91440" anchor="ctr" anchorCtr="0" upright="1">
			<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>
		</a:bodyPr>
		<a:lstStyle><a:lvl2pPr marL="457200"><a:buChar char="-"/></a:lvl2pPr></a:lstStyle>
		<a:p>
			<a:pPr marL="228600" lvl="1" indent="-228600" algn="ctr" defTabSz="914400" rtl="0" fontAlgn="base">
				<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
				<a:spcBef><a:spcPts val="600"/></a:spcBef>
				<a:buClr><a:schemeClr val="accent1"/></a:buClr>
				<a:buSzPct val="75000"/>
				<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
				<a:buAutoNum type="arabicPeriod" startAt="3"/>
				<a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst>
				<a:defRPr sz="1800"/>
			</a:pPr>
			<a:r>
				<a:rPr lang="en-US" sz="2400" b="1" i="0" u="sng" strike="sngStrike" kern="1200" cap="small" spc="-50" baseline="30000" dirty="0">
					<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
					<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
					<a:effectLst><a:glow rad="63500"><a:srgbClr val="00FF00"/></a:glow></a:effectLst>
					<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
					<a:latin typeface="Calibri" panose="020F0502020204030204"/>
					<a:ea typeface="+mn-ea"/>
					<a:cs typeface="+mn-cs"/>
					<a:sym typeface="Wingdings" pitchFamily="2" charset="2"/>
					<a:hlinkClick r:id="rId3" tooltip="Go" history="0"/>
				</a:rPr>
				<a:t>Hello </a:t>
			</a:r>
			<a:br><a:rPr lang="en-US"/></a:br>
			<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld>
			<a:endParaRPr lang="en-US" sz="2400"/>
		</a:p>
	</a:txBody>`;
	const body = parseDrawingRichTextBody(parseXml(xml).documentElement)!;

	it('reads body, list style, paragraph and run properties in file units', () => {
		expect(body.bodyProperties).toMatchObject({
			rotation: 90,
			vertical: 'vert270',
			wrap: 'square',
			insetLeftEmu: 91440,
			anchor: 'ctr',
			anchorCenter: false,
			upright: true,
			autofit: { kind: 'normal', fontScale: 0.625, lineSpacingReduction: 0.2 },
		});
		expect(body.listStyle?.levels[1]?.bullet).toEqual({ kind: 'char', char: '-' });
		expect(body.listStyle?.levels[0]).toBeUndefined();
		const paragraph = body.paragraphs[0]!;
		expect(paragraph.properties).toMatchObject({
			level: 1,
			align: 'ctr',
			marginLeftEmu: 228600,
			indentEmu: -228600,
			defaultTabSizeEmu: 914400,
			rightToLeft: false,
			fontAlign: 'base',
			lineSpacing: { unit: 'percent', value: 0.9 },
			spaceBefore: { unit: 'points', value: 6 },
			bulletColor: { kind: 'scheme', value: 'accent1', transforms: [] },
			bulletSize: { unit: 'percent', value: 0.75 },
			bulletFont: { typeface: 'Arial', pitchFamily: 34, charset: 0 },
			bullet: { kind: 'autoNumber', scheme: 'arabicPeriod', startAt: 3 },
			tabStops: [{ positionEmu: 914400, align: 'l' }],
			defaultRunProperties: { sizePt: 18 },
		});
		const run = paragraph.items[0]!;
		expect(run.kind).toBe('run');
		if (run.kind !== 'run') return;
		expect(run.text).toBe('Hello ');
		expect(run.properties).toMatchObject({
			language: 'en-US',
			sizePt: 24,
			bold: true,
			italic: false,
			underline: 'sng',
			strike: 'sngStrike',
			kernPt: 12,
			caps: 'small',
			spacingPt: -0.5,
			baseline: 0.3,
			dirty: false,
			line: { widthEmu: 12700 },
			fill: { kind: 'solid', color: { kind: 'scheme', value: 'tx1' } },
			highlight: { kind: 'srgb', value: 'FFFF00' },
			latin: { typeface: 'Calibri', panose: '020F0502020204030204' },
			eastAsian: { typeface: '+mn-ea' },
			complexScript: { typeface: '+mn-cs' },
			symbol: { typeface: 'Wingdings', pitchFamily: 2, charset: 2 },
			hyperlinkClick: { relId: 'rId3', tooltip: 'Go', history: false },
		});
		expect(run.properties?.extraXml?.[0]).toContain('effectLst');
		expect(paragraph.items.map((item) => item.kind)).toEqual(['run', 'break', 'field']);
		expect(paragraph.items[2]).toMatchObject({ id: '{1}', type: 'slidenum', text: '3' });
		expect(paragraph.endProperties).toEqual({ language: 'en-US', sizePt: 24 });
	});

	it('writes children in schema order, kept ones in place', () => {
		const written = drawingRichTextBodyXml(body);
		const order = (name: string) => written.indexOf(`<a:${name}`);
		expect(order('ln')).toBeLessThan(order('solidFill'));
		expect(order('effectLst')).toBeGreaterThan(order('solidFill'));
		expect(order('effectLst')).toBeLessThan(order('highlight'));
		expect(order('hlinkClick')).toBeGreaterThan(order('sym'));
		expect(written).toContain('r:id="rId3"');
		expect(normaliseKept(roundTrip(body))).toEqual(normaliseKept(body));
	});

	it('writes one empty paragraph for a body without paragraphs', () => {
		expect(drawingRichTextBodyXml({ paragraphs: [] })).toContain('<a:p/>');
	});
});
