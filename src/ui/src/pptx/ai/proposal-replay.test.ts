import type { PptxData, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { PptxAiConfig, PptxAiToolName, PptxAiWritePolicy } from './config';
import { makeMockBridge, makeSlide, textElement } from './mock-bridge';
import { ProposalStore } from './proposals';
import { allToolNames, buildToolExecutors } from './tools';

const CONNECTION = { kind: 'transport', transport: {} } as unknown as PptxAiConfig['connection'];

function setup(writePolicy: PptxAiWritePolicy = 'stage') {
	const bridge = makeMockBridge();
	const proposals = new ProposalStore(bridge);
	const executors = buildToolExecutors(bridge, proposals, {
		connection: CONNECTION,
		writePolicy,
	});
	const call = async (name: string, input: unknown) => {
		const run = executors.get(name as PptxAiToolName);
		if (!run) {
			throw new Error(`No tool ${name}`);
		}
		return run(input) as Promise<{ proposalId?: string; applied?: boolean; staged?: boolean }>;
	};
	const edit = (slideIndex: number, elementId: string, text: string) =>
		call('update_element', { slideIndex, elementId, text });
	return { bridge, proposals, call, edit };
}

const textOf = (slides: PptxSlide[], slide: number, element: number) =>
	(slides[slide]?.elements[element] as { text?: string } | undefined)?.text;

describe('replaying staged proposals', () => {
	it('keeps every proposal when they are accepted together', async () => {
		const { bridge, proposals, edit } = setup();
		await edit(0, 'el-1', 'EDIT A');
		await edit(1, 'el-3', 'EDIT B');

		expect(proposals.acceptAll()).toBe(2);

		expect(textOf(bridge.getSlides(), 0, 0)).toBe('EDIT A');
		expect(textOf(bridge.getSlides(), 1, 0)).toBe('EDIT B');
		expect(bridge.edits).toHaveLength(2);
	});

	it('keeps a user edit made after staging', async () => {
		const { bridge, proposals, edit } = setup();
		const { proposalId } = await edit(0, 'el-1', 'AI EDIT');
		bridge.applySlidesUpdate((slides) => {
			(slides[1].elements[0] as { text?: string }).text = 'USER EDIT';
			return slides;
		}, 'user');

		expect(proposals.apply(proposalId as string)).toBe(true);

		expect(textOf(bridge.getSlides(), 0, 0)).toBe('AI EDIT');
		expect(textOf(bridge.getSlides(), 1, 0)).toBe('USER EDIT');
	});

	it('keeps a user edit to another property of the same element', async () => {
		const { bridge, proposals, edit } = setup();
		const { proposalId } = await edit(0, 'el-1', 'AI EDIT');
		bridge.applySlidesUpdate((slides) => {
			slides[0].elements[0].x = 321;
			return slides;
		}, 'user');

		expect(proposals.apply(proposalId as string)).toBe(true);

		expect(bridge.getSlides()[0].elements[0]).toMatchObject({ text: 'AI EDIT', x: 321 });
	});

	it('finds its target after the user inserted a slide before it', async () => {
		const { bridge, proposals, edit } = setup();
		const { proposalId } = await edit(1, 'el-3', 'AI EDIT');
		bridge.applySlidesUpdate((slides) => {
			slides.unshift(makeSlide(9, [textElement('el-9', 'New first slide')]));
			return slides;
		}, 'user');

		expect(proposals.apply(proposalId as string)).toBe(true);

		expect(textOf(bridge.getSlides(), 2, 0)).toBe('AI EDIT');
		expect(textOf(bridge.getSlides(), 0, 0)).toBe('New first slide');
	});

	it('reports a conflict when two proposals change the same text', async () => {
		const { bridge, proposals, edit } = setup();
		await edit(0, 'el-1', 'FIRST');
		const second = await edit(0, 'el-1', 'SECOND');

		expect(proposals.acceptAll()).toBe(1);

		expect(textOf(bridge.getSlides(), 0, 0)).toBe('FIRST');
		expect(proposals.get(second.proposalId as string)?.conflict).toContain('el-1');
		expect(proposals.apply(second.proposalId as string)).toBe(false);
		expect(proposals.size).toBe(1);
		expect(proposals.revert(second.proposalId as string)).toBe(true);
	});

	it('reports a conflict when the user changed or removed the target', async () => {
		const { bridge, proposals, edit } = setup();
		const changed = await edit(0, 'el-1', 'AI');
		const removed = await edit(1, 'el-3', 'AI');
		bridge.applySlidesUpdate((slides) => {
			(slides[0].elements[0] as { text?: string }).text = 'USER';
			slides[1].elements = [];
			return slides;
		}, 'user');

		expect(proposals.apply(changed.proposalId as string)).toBe(false);
		expect(proposals.apply(removed.proposalId as string)).toBe(false);
		expect(textOf(bridge.getSlides(), 0, 0)).toBe('USER');
		expect(bridge.getSlides()[1].elements).toHaveLength(0);
	});

	it('can retry a conflict once the deck fits again', async () => {
		const { bridge, proposals, edit } = setup();
		const { proposalId } = await edit(0, 'el-1', 'AI');
		const setText = (text: string) =>
			bridge.applySlidesUpdate((slides) => {
				(slides[0].elements[0] as { text?: string }).text = text;
				return slides;
			}, 'user');
		setText('USER');
		expect(proposals.apply(proposalId as string)).toBe(false);

		setText('Title One');

		expect(proposals.apply(proposalId as string)).toBe(true);
		expect(textOf(bridge.getSlides(), 0, 0)).toBe('AI');
	});

	it('adds two staged slides and elements without losing either', async () => {
		const { bridge, proposals, call } = setup();
		await call('add_slide', {});
		await call('add_slide', {});
		const element = { slideIndex: 0, type: 'text', x: 1, y: 1, width: 10, height: 10 };
		await call('add_element', { ...element, text: 'One' });
		await call('add_element', { ...element, text: 'Two' });

		expect(proposals.acceptAll()).toBe(4);

		const slides = bridge.getSlides();
		expect(slides).toHaveLength(4);
		expect(slides.map((slide) => slide.slideNumber)).toEqual([1, 2, 3, 4]);
		expect(slides[0].elements.map((el) => (el as { text?: string }).text)).toEqual([
			'Title One',
			'Body copy',
			'One',
			'Two',
		]);
	});

	it('deletes a staged slide by identity after the deck shifted', async () => {
		const { bridge, proposals, call } = setup();
		const { proposalId } = await call('delete_slides', { slideIndexes: [1] });
		bridge.applySlidesUpdate((slides) => {
			slides.unshift(makeSlide(9, []));
			return slides;
		}, 'user');

		expect(proposals.apply(proposalId as string)).toBe(true);

		expect(bridge.getSlides().map((slide) => slide.id)).toEqual(['slide-9', 'slide-0']);
	});
});

/** One representative call for each way a write reaches the deck. */
const WRITES: [tool: string, input: unknown, landed: (deck: PptxData) => unknown][] = [
	[
		'update_element',
		{ slideIndex: 0, elementId: 'el-1', text: 'Written' },
		(d) => textOf(d.slides, 0, 0) === 'Written',
	],
	['apply_theme_preset', { presetName: 'DARK' }, (d) => d.theme?.colorScheme?.dk1 === '#E4E4E7'],
	[
		'update_theme_colors',
		{ accent1: '#123456' },
		(d) => d.theme?.colorScheme?.accent1 === '#123456',
	],
	[
		'update_theme_fonts',
		{ majorFont: 'Written Sans' },
		(d) => d.theme?.fontScheme?.majorFont?.latin === 'Written Sans',
	],
	['update_metadata', { title: 'Written' }, (d) => d.coreProperties?.title === 'Written'],
	['manage_sections', { action: 'add', name: 'Written' }, (d) => d.sections?.length === 1],
];

describe('every write follows the active policy', () => {
	it.each(WRITES)('stage holds %s until it is accepted', async (tool, input, landed) => {
		const { bridge, proposals, call } = setup('stage');
		const before = structuredClone(bridge.getDeckData?.());

		const result = await call(tool, input);

		expect(result.staged).toBe(true);
		expect(bridge.edits).toHaveLength(0);
		expect(bridge.getDeckData?.()).toEqual(before);

		expect(proposals.acceptAll()).toBe(1);
		expect(landed(bridge.getDeckData?.() as PptxData)).toBe(true);
	});

	it.each(WRITES)('approve holds %s for its own acceptance', async (tool, input, landed) => {
		const { bridge, proposals, call } = setup('approve');

		const { proposalId } = await call(tool, input);

		expect(proposals.get(proposalId as string)?.requiresApproval).toBe(true);
		expect(proposals.acceptAll()).toBe(0);
		expect(bridge.edits).toHaveLength(0);

		expect(proposals.apply(proposalId as string)).toBe(true);
		expect(landed(bridge.getDeckData?.() as PptxData)).toBe(true);
	});

	it.each(WRITES)('auto applies %s at once', async (tool, input, landed) => {
		const { bridge, proposals, call } = setup('auto');

		const result = await call(tool, input);

		expect(result.applied).toBe(true);
		expect(proposals.size).toBe(0);
		expect(landed(bridge.getDeckData?.() as PptxData)).toBe(true);
	});

	it.each(['stage', 'approve', 'auto'] as const)(
		'forced-approval tools always ask under %s',
		async (policy) => {
			const { bridge, proposals, call } = setup(policy);
			const exported = (await call('export_to_json', {})) as unknown as { json: string };

			const deleted = await call('delete_slides', { slideIndexes: [1] });
			const imported = await call('import_from_json', { json: exported.json });

			for (const { proposalId } of [deleted, imported]) {
				expect(proposals.get(proposalId as string)?.requiresApproval).toBe(true);
			}
			expect(proposals.acceptAll()).toBe(0);
			expect(bridge.edits).toHaveLength(0);
			expect(bridge.getSlides()).toHaveLength(2);

			expect(proposals.apply(deleted.proposalId as string)).toBe(true);
			expect(bridge.getSlides()).toHaveLength(1);
		},
	);
});

describe('tool input validation', () => {
	it.each([null, 'text', 42, []])('rejects non-object input %j for every tool', async (input) => {
		const { bridge, proposals, call } = setup('auto');
		const before = structuredClone(bridge.getDeckData?.());
		let rejected = 0;
		for (const name of allToolNames()) {
			try {
				await call(name, input);
			} catch (error) {
				rejected += 1;
				expect((error as Error).message).toContain(`Invalid input for ${name}`);
			}
		}

		expect(rejected).toBe(allToolNames().length);
		expect(bridge.edits).toHaveLength(0);
		expect(proposals.size).toBe(0);
		expect(bridge.getDeckData?.()).toEqual(before);
	});

	it('never lets a wrong type reach a staged change', async () => {
		const { bridge, proposals, call } = setup();

		await expect(
			call('update_element', { slideIndex: 0, elementId: 'el-1', x: 'nope', rotation: {} }),
		).rejects.toThrow(/x: .*rotation: /u);

		expect(proposals.size).toBe(0);
		expect(bridge.getSlides()[0].elements[0].x).toBe(40);
	});
});
