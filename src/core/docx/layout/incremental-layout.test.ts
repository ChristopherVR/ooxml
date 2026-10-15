import { describe, expect, it } from 'vitest';
import { twips } from '../index';
import { layoutDocument } from './layout';
import { createLayoutSession } from './layout-session';
import { createFakeMeasurer } from './measure';
import type { LayoutDocumentInput, LayoutParagraph } from './input';

/** Small deterministic generator so a failing seed reproduces. */
function random(seed: number): () => number {
	let state = seed >>> 0 || 1;
	return () => {
		state ^= state << 13;
		state >>>= 0;
		state ^= state >>> 17;
		state ^= state << 5;
		state >>>= 0;
		return state / 0x100000000;
	};
}

const WORDS = ['alpha', 'be', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta', 'iota', 'kappa'];

function build(rand: () => number, serial: { next: number }): LayoutParagraph {
	const words = Math.floor(rand() * rand() * 70);
	const text = Array.from({ length: words }, () => WORDS[Math.floor(rand() * WORDS.length)]!)
		.join(' ')
		.replace(/ (?=be )/g, rand() < 0.1 ? '\n' : ' ');
	const paragraph: LayoutParagraph = {
		kind: 'paragraph',
		id: `p${serial.next++}`,
		runs: [{ text }],
	};
	if (rand() < 0.1) paragraph.keepNext = true;
	if (rand() < 0.06) paragraph.keepLines = true;
	if (rand() < 0.03) paragraph.pageBreakBefore = true;
	if (rand() < 0.05) paragraph.widowControl = false;
	if (rand() < 0.4) paragraph.spacingAfterTwips = twips(Math.floor(rand() * 400));
	if (rand() < 0.2) paragraph.spacingBeforeTwips = twips(Math.floor(rand() * 400));
	if (rand() < 0.1) {
		paragraph.contextualSpacing = true;
		paragraph.styleId = 'Same';
	}
	if (rand() < 0.1) paragraph.runs[0]!.fontSizePt = 8 + Math.floor(rand() * 24);
	return paragraph;
}

function document(blocks: LayoutParagraph[]): LayoutDocumentInput {
	return {
		sections: [
			{
				page: {
					widthPx: 400,
					heightPx: 300,
					marginTopPx: 20,
					marginRightPx: 20,
					marginBottomPx: 20,
					marginLeftPx: 20,
				},
				blocks: structuredClone(blocks),
			},
		],
	};
}

/** One random edit: retype, insert, delete, split, merge or toggle a flow property. */
function edit(blocks: LayoutParagraph[], rand: () => number, serial: { next: number }): void {
	const at = Math.floor(rand() * blocks.length);
	const target = blocks[at]!;
	const kind = Math.floor(rand() * 7);
	if (kind === 0) target.runs = [{ text: `${target.runs[0]!.text} more words appended here` }];
	else if (kind === 1) target.runs = [{ text: target.runs[0]!.text.slice(0, 12) }];
	else if (kind === 2) blocks.splice(at, 0, build(rand, serial));
	else if (kind === 3 && blocks.length > 3) blocks.splice(at, 1);
	else if (kind === 4) {
		const text = target.runs[0]!.text;
		const cut = Math.floor(text.length / 2);
		target.runs = [{ text: text.slice(0, cut) }];
		blocks.splice(at + 1, 0, { ...build(rand, serial), runs: [{ text: text.slice(cut) }] });
	} else if (kind === 5) {
		if (target.keepNext) delete target.keepNext;
		else target.keepNext = true;
	} else if (target.pageBreakBefore) delete target.pageBreakBefore;
	else target.pageBreakBefore = true;
}

describe('incremental pagination', () => {
	it('matches a full layout after every random edit', () => {
		const modes = { full: 0, incremental: 0, unchanged: 0 };
		let resumedLate = 0;
		const seeds = Number(process.env.INCREMENTAL_SEEDS ?? 30);
		for (let seed = 1; seed <= seeds; seed++) {
			const rand = random(seed * 7919);
			const serial = { next: 0 };
			const blocks = Array.from({ length: 40 + Math.floor(rand() * 260) }, () =>
				build(rand, serial),
			);
			const session = createLayoutSession(createFakeMeasurer(), { incremental: true });
			expect(session.layoutDocument(document(blocks))).toEqual(
				layoutDocument(document(blocks), createFakeMeasurer()),
			);
			for (let step = 0; step < 8; step++) {
				const batch = 1 + Math.floor(rand() * 3);
				for (let count = 0; count < batch; count++) edit(blocks, rand, serial);
				const actual = session.layoutDocument(document(blocks));
				const expected = layoutDocument(document(blocks), createFakeMeasurer());
				expect(actual, `seed ${seed} step ${step}`).toEqual(expected);
				const pass = session.lastPass()!;
				modes[pass.mode]++;
				if (pass.mode === 'incremental' && pass.fromPage > 0) resumedLate++;
			}
		}
		// The property is only meaningful if the incremental path actually ran.
		expect(modes.incremental).toBeGreaterThan(seeds * 3);
		expect(resumedLate).toBeGreaterThan(seeds / 2);
	}, 600_000);

	it('lays out only the pages around an edit in a long document', () => {
		const rand = random(42);
		const serial = { next: 0 };
		const blocks = Array.from({ length: 1500 }, () => build(rand, serial));
		const session = createLayoutSession(createFakeMeasurer(), { incremental: true });
		session.layoutDocument(document(blocks));
		expect(session.lastPass()).toMatchObject({ mode: 'full', placed: 1500 });
		blocks[700]!.runs = [{ text: 'a short replacement' }];
		const actual = session.layoutDocument(document(blocks));
		expect(actual).toEqual(layoutDocument(document(blocks), createFakeMeasurer()));
		const pass = session.lastPass()!;
		expect(pass.mode).toBe('incremental');
		expect(pass.fromPage).toBeGreaterThan(5);
		// The edit shifts page breaks until the layout settles, then the old pages are reused.
		expect(pass.placed).toBeLessThan(400);
		// A second, unrelated edit resumes from the new snapshot.
		blocks[100]!.runs = [{ text: `${blocks[100]!.runs[0]!.text} and a longer ending for it` }];
		expect(session.layoutDocument(document(blocks))).toEqual(
			layoutDocument(document(blocks), createFakeMeasurer()),
		);
		expect(session.lastPass()!.mode).toBe('incremental');
	});

	it('lays out again when only the page setup changes', () => {
		const rand = random(11);
		const serial = { next: 0 };
		const blocks = Array.from({ length: 120 }, () => build(rand, serial));
		const session = createLayoutSession(createFakeMeasurer(), { incremental: true });
		session.layoutDocument(document(blocks));
		const narrow = document(blocks);
		narrow.sections[0]!.page.widthPx = 300;
		narrow.sections[0]!.page.marginBottomPx = 60;
		expect(session.layoutDocument(narrow)).toEqual(layoutDocument(narrow, createFakeMeasurer()));
		expect(session.lastPass()!.mode).toBe('full');
	});

	it('takes the full pass for documents with tables, columns or footnotes', () => {
		const session = createLayoutSession(createFakeMeasurer(), { incremental: true });
		const base = document([build(random(3), { next: 0 })]);
		const columns: LayoutDocumentInput = {
			sections: [{ ...base.sections[0]!, columns: { count: 2, gapPx: 10 } }],
		};
		expect(session.layoutDocument(columns)).toEqual(layoutDocument(columns, createFakeMeasurer()));
		expect(session.lastPass()!.mode).toBe('full');
		expect(session.lastPass()!.placed).toBe(0);
	});
});
