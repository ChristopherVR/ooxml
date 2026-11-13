import type { PptxCustomShow, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { BuildingBlockAvailabilityInput } from '../../building-block-actions';
import { createFindBarController } from '../find-bar-controller';
import { createPptxBuildingBlockSession } from '../session';
import type { BuildingBlockSessionPort } from '../session';
import { BUILDING_BLOCK_GROUP_IDS, registerBuildingBlockGroups } from './index';
import type { BuildingBlockGroupPorts } from './index';

const READY: BuildingBlockAvailabilityInput = {
	ready: true,
	hasActiveSlide: true,
	canEdit: true,
	editing: true,
	canInsert: true,
	hasSelection: false,
	selectionHasHyperlink: false,
	selectionHasEquation: false,
	layoutCount: 0,
	hasCurrentLayout: false,
	hasHandler: true,
};

function setup() {
	let input = READY;
	const onRequest = vi.fn();
	const onCancel = vi.fn();
	const port: BuildingBlockSessionPort = {
		availabilityInput: () => input,
		onRequest,
		onCancel,
		fileLimits: () => undefined,
		target: () => ({ deck: null, slideId: 's1', slideIndex: 0 }),
		canvasSize: () => ({ width: 960, height: 540 }),
		selectedElement: () => null,
		slideCount: () => 1,
		activeSlide: () => undefined,
		scopedLayouts: () => [],
		addElement: vi.fn(),
		insertMediaFile: vi.fn(async () => ({ ok: true as const })),
		applyHyperlink: vi.fn(() => ({ ok: true as const })),
		removeHyperlink: vi.fn(() => ({ ok: true as const })),
		insertSmartArt: vi.fn(),
		insertChart: vi.fn(),
		insertEquation: vi.fn(),
		updateEquation: vi.fn(),
		findReplace: () =>
			createFindBarController({ slides: () => [], commit: vi.fn(), navigate: vi.fn() }).port(),
		applyLayout: vi.fn(async () => undefined),
	};
	const session = createPptxBuildingBlockSession(port);
	const slides = [{ id: 's1', rId: 'rId2', elements: [] }] as unknown as PptxSlide[];
	let shows: PptxCustomShow[] = [];
	let shortcutsOpen = false;
	const ports: BuildingBlockGroupPorts = {
		passwordProtection: {
			input: () => ({
				openPasswordSet: false,
				recommendation: { kind: null },
				hasModifyVerifier: false,
				customProperties: [],
			}),
			customProperties: () => [],
			applyOpenPassword: vi.fn(),
			applyCustomProperties: vi.fn(),
			markDirty: vi.fn(),
		},
		documentProperties: {
			current: () => ({ core: { title: 'Old' }, custom: [], app: {} }),
			applyUpdate: vi.fn(),
		},
		fontEmbedding: {
			input: () => ({ embedFonts: false, embeddedFamilies: [], usedFamilies: ['Arial'] }),
			apply: vi.fn(),
			markDirty: vi.fn(),
		},
		digitalSignatures: {
			signed: () => false,
			count: () => 0,
			content: () => null,
			markDirty: vi.fn(),
		},
		setUpSlideShow: {
			presentationProperties: () => ({}),
			slideCount: () => 1,
			customShows: () => shows,
			save: vi.fn(),
		},
		customShows: {
			customShows: () => shows,
			slides: () => slides,
			activeId: () => null,
			commit: vi.fn((next: PptxCustomShow[]) => {
				shows = next;
			}),
			start: vi.fn(),
		},
		shortcuts: {
			open: () => shortcutsOpen,
			setOpen: (open) => {
				shortcutsOpen = open;
			},
			toggle: () => {
				shortcutsOpen = !shortcutsOpen;
			},
			translate: (key) => key,
		},
	};
	const stop = registerBuildingBlockGroups(session, ports);
	return {
		session,
		ports,
		stop,
		onRequest,
		onCancel,
		setInput: (next: Partial<BuildingBlockAvailabilityInput>) => {
			input = { ...READY, ...next };
		},
	};
}

describe('registerBuildingBlockGroups', () => {
	it('adds every built-in group to actions and removes them again', () => {
		const { session, stop } = setup();
		for (const id of BUILDING_BLOCK_GROUP_IDS) {
			expect(session.actions()[id], id).toBeDefined();
		}
		expect(session.actions().broadcast).toEqual({
			supported: false,
			reason: 'unsupported',
			requires: 'collaboration-service',
		});
		stop();
		expect(Reflect.get(session.actions(), 'passwordProtection')).toBeUndefined();
	});

	it('gates file dialog operations and completes their request', () => {
		const { session, ports, setInput } = setup();
		expect(session.request({ kind: 'documentProperties' })).toBe(true);
		expect(session.actions().documentProperties.apply({ core: { title: 'New' } })).toEqual({
			ok: true,
		});
		expect(ports.documentProperties.applyUpdate).toHaveBeenCalledOnce();
		expect(session.actions().pending).toBeNull();
		expect(session.actions().passwordProtection.setOpenPassword('ab', 'ab')).toEqual({
			ok: false,
			code: 'password-too-short',
		});
		// Viewable: raised read-only, but applying is refused.
		setInput({ canEdit: false });
		expect(session.request({ kind: 'fontEmbedding' })).toBe(true);
		expect(session.actions().fontEmbedding.setEmbedFonts(false)).toEqual({
			ok: false,
			code: 'read-only',
		});
		setInput({ ready: false });
		expect(session.request({ kind: 'digitalSignatures' })).toBe(false);
	});

	it('runs custom show edits on the latest list and starts a show', () => {
		const { session, ports } = setup();
		expect(session.actions().customShows.create('Short', ['rId2'])).toEqual({ ok: true });
		const id = session.actions().customShows.shows[0]!.id;
		expect(session.actions().customShows.rename(id, 'Shorter')).toEqual({ ok: true });
		expect(session.actions().customShows.shows[0]?.name).toBe('Shorter');
		expect(session.actions().customShows.start(id)).toEqual({ ok: true });
		expect(ports.customShows.start).toHaveBeenCalledWith(id);
		expect(session.actions().customShows.start('nope')).toEqual({
			ok: false,
			code: 'unknown-show',
		});
		expect(session.actions().setUpSlideShow.apply({ loopContinuously: true })).toEqual({
			ok: true,
		});
		expect(ports.setUpSlideShow.save).toHaveBeenCalledOnce();
	});

	it('mirrors the shortcut help as announced requests and closes it on cancel', () => {
		const { session, onRequest, onCancel } = setup();
		session.announce({ kind: 'shortcuts', open: true }, true);
		expect(session.actions().pending).toEqual({ kind: 'shortcuts', open: true });
		session.announce({ kind: 'shortcuts', open: false }, false);
		expect(session.actions().pending).toBeNull();
		expect(onRequest).toHaveBeenCalledTimes(2);
		session.actions().shortcuts.setOpen(true);
		expect(session.actions().shortcuts.open).toBe(true);
		expect(session.actions().shortcuts.catalogue.length).toBeGreaterThan(0);
		session.actions().cancel();
		expect(onCancel).toHaveBeenCalledOnce();
	});
});
