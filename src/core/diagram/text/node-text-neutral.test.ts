import { describe, expect, it } from 'vitest';

import { NS, parseXml } from '../../xml/index';
import { diagramParagraphsText } from './node-text-paragraphs';
import { projectDiagramNodeTextNeutral, readDiagramNodeParagraphs } from './node-text-neutral';

const T = `<dgm:t xmlns:dgm="${NS.dgm}" xmlns:a="${NS.a}"><a:bodyPr/><a:lstStyle/>
	<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" b="1"/><a:t>Title</a:t></a:r><a:br/><a:r><a:t>more</a:t></a:r></a:p>
	<a:p><a:pPr lvl="1"/><a:r><a:rPr sz="1200"/><a:t>Child</a:t></a:r></a:p>
	<a:p><a:pPr lvl="1"><a:buNone/></a:pPr><a:fld id="{F}" type="datetime"><a:t>Today</a:t></a:fld></a:p>
	<a:p><a:endParaRPr sz="900"/></a:p>
</dgm:t>`;

const body = parseXml(T).documentElement;

describe('neutral SmartArt node text', () => {
	it('reads paragraphs with neutral properties', () => {
		const paragraphs = readDiagramNodeParagraphs(body)!;
		expect(paragraphs).toHaveLength(4);
		expect(diagramParagraphsText(paragraphs)).toBe('Title\nmore\nChild\nToday\n');
		expect(paragraphs[0]!.properties).toEqual({ align: 'ctr' });
		expect(paragraphs[0]!.items[0]).toEqual({
			kind: 'run',
			run: { text: 'Title', properties: { language: 'en-US', bold: true } },
		});
		expect(paragraphs[2]!.items[0]).toMatchObject({
			kind: 'field',
			id: '{F}',
			fieldType: 'datetime',
		});
		expect(paragraphs[3]!.endProperties).toEqual({ sizePt: 9 });
	});

	it('returns undefined without paragraphs', () => {
		expect(readDiagramNodeParagraphs(undefined)).toBeUndefined();
		expect(
			readDiagramNodeParagraphs(parseXml(`<dgm:t xmlns:dgm="${NS.dgm}"/>`).documentElement),
		).toBeUndefined();
	});

	it('projects segments with inherited styles, levels and synthesized bullets', () => {
		const paragraphs = readDiagramNodeParagraphs(body)!;
		const segments = projectDiagramNodeTextNeutral(
			{ text: diagramParagraphsText(paragraphs), paragraphs },
			{ paragraph: {}, run: { sizePt: 18 } },
			{ bulletEnabled: true },
		);
		expect(segments.map((segment) => segment.text)).toEqual([
			'Title',
			'\n',
			'more',
			'',
			'Child',
			'',
			'Today',
			'',
			'',
		]);
		expect(segments[0]).toMatchObject({
			style: { paragraph: { align: 'ctr' }, run: { sizePt: 18, bold: true, language: 'en-US' } },
			paragraphLevel: 0,
		});
		expect(segments[0]!.bulletInfo).toBeUndefined();
		expect(segments[1]!.isLineBreak).toBe(true);
		expect(segments[3]!.isParagraphBreak).toBe(true);
		expect(segments[4]).toMatchObject({
			bulletInfo: { kind: 'char', char: '•' },
			paragraphLevel: 1,
			style: { run: { sizePt: 12 } },
		});
		// An explicit `a:buNone` wins over `bulletEnabled`.
		expect(segments[6]).toMatchObject({ bulletInfo: { kind: 'none' }, fieldType: 'datetime' });
		// A paragraph without runs projects one empty segment in its end-paragraph style.
		expect(segments[8]!.style.run).toEqual({ sizePt: 9 });
	});

	it('reconciles an edited flat text into the paragraphs', () => {
		const paragraphs = readDiagramNodeParagraphs(body)!;
		const segments = projectDiagramNodeTextNeutral({
			text: 'Renamed\nmore\nChild\nToday\n',
			paragraphs,
		});
		expect(segments[0]).toMatchObject({ text: 'Renamed', style: { run: { bold: true } } });
	});
});
