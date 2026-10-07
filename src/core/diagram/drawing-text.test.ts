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
