import type { PptxElement, PptxLayoutOption, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { BuildingBlockAvailabilityInput } from '../building-block-actions';
import type { BuildingBlockFindReplacePort } from './find-layout';
import { buildingBlockRequestAction } from './requests';
import { createPptxBuildingBlockSession, selectionEquation } from './session';
import type { BuildingBlockSessionPort } from './session';

const READY: BuildingBlockAvailabilityInput = {
	ready: true,
	hasActiveSlide: true,
	canEdit: true,
	editing: true,
	canInsert: true,
	hasSelection: true,
	selectionHasHyperlink: true,
	selectionHasEquation: false,
	layoutCount: 2,
	hasCurrentLayout: true,
	hasHandler: true,
};

function findPort(overrides: Partial<BuildingBlockFindReplacePort> = {}) {
	return {
		open: false,
		setOpen: vi.fn(),
		query: '',
		replacement: '',
		options: {
			matchCase: false,
			includeNotes: false,
			includeChartTitles: true,
			includeChartData: false,
		},
		matches: [],
		currentIndex: -1,
		setQuery: vi.fn(),
		setReplacement: vi.fn(),
		setOptions: vi.fn(),
		find: vi.fn(),
		navigate: vi.fn(),
		replaceCurrent: vi.fn(),
		replaceAll: vi.fn(),
		...overrides,
	} satisfies BuildingBlockFindReplacePort;
}

function setup(input: Partial<BuildingBlockAvailabilityInput> = {}) {
	let availability: BuildingBlockAvailabilityInput = { ...READY, ...input };
	const slide = { id: 's1', elements: [], layoutPath: 'l1' } as unknown as PptxSlide;
	const layouts = [
		{ path: 'l1', name: 'Title' },
		{ path: 'l2', name: 'Blank' },
	] as PptxLayoutOption[];
	let target = { deck: {}, slideId: 's1' as string | undefined, slideIndex: 0 };
	const port = {
		availabilityInput: () => availability,
		fileLimits: () => undefined,
		onRequest: vi.fn(),
		target: () => target,
		canvasSize: () => ({ width: 960, height: 540 }),
		selectedElement: () =>
			({ id: 'e1', actionClick: { url: 'https://example.com' } }) as unknown as PptxElement,
		slideCount: () => 3,
		activeSlide: () => slide,
		scopedLayouts: () => layouts,
		addElement: vi.fn(),
		insertMediaFile: vi.fn(async (_file: File, same: () => boolean) =>
			same() ? { ok: true as const } : { ok: false as const, code: 'stale-target' as const },
		),
		applyHyperlink: vi.fn(() => ({ ok: true as const })),
		removeHyperlink: vi.fn(() => ({ ok: true as const })),
		insertSmartArt: vi.fn(),
		insertChart: vi.fn(),
		insertEquation: vi.fn(),
		updateEquation: vi.fn(),
		findReplace: () => findPort(),
		applyLayout: vi.fn(async () => undefined),
		decodeImage: vi.fn(async () => ({ id: 'img' }) as unknown as PptxElement),
	} satisfies BuildingBlockSessionPort;
	const session = createPptxBuildingBlockSession(port);
	return {
		port,
		session,
		setAvailability: (next: Partial<BuildingBlockAvailabilityInput>) => {
			availability = { ...READY, ...next };
		},
		moveTo: (slideId: string) => {
			target = { ...target, slideId };
		},
	};
}

const png = () => new File([new Uint8Array([1, 2, 3])], 'a.png', { type: 'image/png' });

describe('createPptxBuildingBlockSession', () => {
	it('raises a request only when its action is available, and completes it on success', async () => {
		const { session, port, setAvailability } = setup();
		const notified = vi.fn();
		session.requests.subscribe(notified);
		expect(session.request({ kind: 'image' })).toBe(true);
		expect(session.actions().pending).toEqual({ kind: 'image' });
		expect(port.onRequest).toHaveBeenCalledWith({ kind: 'image' });
		expect(notified).toHaveBeenCalledTimes(1);
		expect(await session.actions().insertImage(png())).toEqual({ ok: true });
		expect(port.addElement).toHaveBeenCalledWith({ id: 'img' });
		expect(session.actions().pending).toBeNull();

		setAvailability({ canEdit: false });
		expect(session.request({ kind: 'image' })).toBe(false);
		expect(session.actions().pending).toBeNull();
	});

	it('refuses an image whose slide changed while it decoded', async () => {
		const { session, port, moveTo } = setup();
		port.decodeImage.mockImplementationOnce(async () => {
			moveTo('s2');
			return { id: 'img' } as unknown as PptxElement;
		});
		expect(await session.actions().insertImage(png())).toEqual({
			ok: false,
			code: 'stale-target',
		});
		expect(port.addElement).not.toHaveBeenCalled();
	});

	it('validates files before any decoding', async () => {
		const { session, port } = setup();
		const text = new File(['x'], 'a.txt', { type: 'text/plain' });
		expect(await session.actions().insertImage(text)).toEqual({
			ok: false,
			code: 'unsupported-type',
		});
		expect(await session.actions().insertMedia(text)).toEqual({
			ok: false,
			code: 'unsupported-type',
		});
		expect(port.decodeImage).not.toHaveBeenCalled();
	});

	it('gates structured inserts and the hyperlink by availability', () => {
		const { session, port, setAvailability } = setup();
		expect(session.actions().hyperlink.current?.url).toBe('https://example.com');
		expect(session.actions().hyperlink.slideCount).toBe(3);
		expect(session.actions().smartArt.insert('basicBlockList' as never, ['a'])).toEqual({
			ok: true,
		});
		expect(port.insertSmartArt).toHaveBeenCalledOnce();
		const blank = { ...session.actions().hyperlink.current!, url: '' };
		expect(session.actions().hyperlink.apply(blank)).toEqual({
			ok: false,
			code: 'invalid-target',
		});
		expect(
			session.actions().chart.insert('column', { data: { categories: [], series: [] } }),
		).toEqual({ ok: false, code: 'invalid-data' });
		setAvailability({ canEdit: false });
		expect(session.actions().equation.insert({})).toEqual({ ok: false, code: 'read-only' });
		expect(port.insertEquation).not.toHaveBeenCalled();
	});

	it('runs layout apply only for a listed layout, and reset to the slide own layout', async () => {
		const { session, port } = setup();
		expect(session.actions().layout.currentPath).toBe('l1');
		expect(await session.actions().layout.apply('nope')).toEqual({
			ok: false,
			code: 'unknown-layout',
		});
		expect(await session.actions().layout.apply('l2')).toEqual({ ok: true });
		expect(await session.resetSlide()).toEqual({ ok: true });
		expect(port.applyLayout.mock.calls).toEqual([['l2'], ['l1']]);
	});

	it('adds registered groups under their key and removes them again', () => {
		const { session } = setup();
		const create = vi.fn(() => ({ ping: () => 'pong' }));
		const stop = session.register({ id: 'demo' as never, create: create as never }, {});
		expect(Reflect.get(session.actions(), 'demo')).toEqual({ ping: expect.any(Function) });
		stop();
		expect(Reflect.get(session.actions(), 'demo')).toBeUndefined();
	});
});

describe('helpers', () => {
	it('maps an equation request to edit or insert', () => {
		expect(buildingBlockRequestAction({ kind: 'equation', existingOmml: null })).toBe(
			'insertEquation',
		);
		expect(buildingBlockRequestAction({ kind: 'equation', existingOmml: {} })).toBe('editEquation');
	});

	it('reads the first equation of an element', () => {
		const omml = { m: 1 };
		const element = { textSegments: [{ text: 'a' }, { equationXml: omml }] } as never;
		expect(selectionEquation(element)).toBe(omml);
		expect(selectionEquation(null)).toBeNull();
	});
});
