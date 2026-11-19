import { describe, expect, it } from 'vitest';
import { xmlTags } from './tags';

const tags = (xml: string) =>
	[...xmlTags(xml)].map(
		(tag) => `${tag.closing ? '/' : ''}${tag.name}${tag.selfClosing ? '/' : ''}`,
	);

/** Best of three runs, in milliseconds. */
const timed = (run: () => void): number => {
	let best = Number.POSITIVE_INFINITY;
	for (let attempt = 0; attempt < 3; attempt++) {
		const start = performance.now();
		run();
		best = Math.min(best, performance.now() - start);
	}
	return best;
};

describe('xmlTags', () => {
	it('reads start, end and empty tags in order, with their attributes', () => {
		const [first, second] = xmlTags('<a:path w="1" t=\'>\'><a:close/></a:path>');
		expect(first).toEqual({
			name: 'a:path',
			local: 'path',
			attributes: ' w="1" t=\'>\'',
			closing: false,
			selfClosing: false,
		});
		expect(second).toMatchObject({ name: 'a:close', local: 'close', selfClosing: true });
		expect(tags('<a:path w="1" t=">"><a:close/></a:path >')).toEqual([
			'a:path',
			'a:close/',
			'/a:path',
		]);
	});

	it('skips comments, CDATA, processing instructions and declarations', () => {
		expect(
			tags(
				'<?xml version="1.0"?><!DOCTYPE x><r><!-- <a:custGeom> --><![CDATA[<b>]]><?pi <c>?><d/></r>',
			),
		).toEqual(['r', 'd/', '/r']);
	});

	it('stops at markup that never ends and passes over a stray <', () => {
		expect(tags('<a>1 < 2<b/><c x="unterminated>')).toEqual(['a', 'b/']);
		expect(tags('<a><!-- open')).toEqual(['a']);
	});

	it('is linear on input that never closes its tags', () => {
		const run = (count: number) => () => {
			for (const tag of xmlTags('<a:custGeom/>'.repeat(count) + '<v:shape '.repeat(count)))
				void tag;
		};
		expect(timed(run(160_000)) / timed(run(40_000))).toBeLessThan(8);
	});
});
