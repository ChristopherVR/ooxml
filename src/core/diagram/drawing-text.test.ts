import { expect, it } from 'vitest';
import { NS, parseXml } from '../xml';
import { parseDrawingTextBody } from './drawing-text';

it('retains run, cached field and break order with paragraph defaults', () => {
	const body = parseXml(
		`<a:txBody xmlns:a="${NS.a}"><a:bodyPr/><a:p><a:pPr algn="ctr"><a:defRPr sz="1800" b="0"><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:r><a:t>Before</a:t></a:r><a:fld id="1"><a:rPr b="1"/><a:t>Field</a:t></a:fld><a:br/><a:r><a:t>After</a:t></a:r></a:p></a:txBody>`,
	).documentElement;
	const parsed = parseDrawingTextBody(body)!;
	expect(parsed.text).toBe('BeforeField\nAfter');
	expect(parsed.paragraphs[0]).toMatchObject({
		align: 'ctr',
		defaultProperties: { sizePt: 18, bold: false, typeface: 'Arial' },
		runs: [{ text: 'Before' }, { text: 'Field', bold: true }, { text: '\n' }, { text: 'After' }],
	});
});

it('retains fractional and absolute DrawingML paragraph spacing, including zero', () => {
	const body = parseXml(
		`<a:txBody xmlns:a="${NS.a}"><a:bodyPr/><a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPct val="25000"/></a:spcAft></a:pPr><a:r><a:t>One</a:t></a:r></a:p><a:p><a:pPr><a:lnSpc><a:spcPts val="0"/></a:lnSpc></a:pPr></a:p></a:txBody>`,
	).documentElement;
	expect(parseDrawingTextBody(body)!.paragraphs).toMatchObject([
		{
			lineSpacing: { unit: 'percent', value: 1.5 },
			spaceBefore: { unit: 'points', value: 6 },
			spaceAfter: { unit: 'percent', value: 0.25 },
		},
		{ lineSpacing: { unit: 'points', value: 0 } },
	]);
});

it.each(['-100', '100x', 'Infinity', ''])(
	'does not turn malformed spacing %s into layout values',
	(value) => {
		const body = parseXml(
			`<a:txBody xmlns:a="${NS.a}"><a:p><a:pPr><a:lnSpc><a:spcPts val="${value}"/></a:lnSpc></a:pPr></a:p></a:txBody>`,
		).documentElement;
		expect(parseDrawingTextBody(body)!.paragraphs[0]!.lineSpacing).toBeUndefined();
	},
);
