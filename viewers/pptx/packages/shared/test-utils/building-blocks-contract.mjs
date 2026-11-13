/**
 * The parity contract every binding's PowerPoint building blocks pass (React
 * `useViewerBuildingBlocks`, vanilla `createViewerBuildingBlocks`, and the Vue, Svelte and
 * Angular ports). Each binding supplies a harness (see `building-blocks-contract.d.mts`) and
 * calls {@link runBuildingBlocksContract} from one test; the assertions are the same for all of
 * them, so a port that drifts from the shared session fails here. See
 * `docs/pptx-building-blocks.md` at the repository root.
 */
import { runShellContract } from './building-blocks-contract-shell.mjs';

/** The compact toolbar's controls, in order, for ports that render it. */
export const COMPACT_TOOLBAR_CONTROLS = [
	'undo',
	'redo',
	'insertImage',
	'insertMedia',
	'insertChart',
	'insertSmartArt',
	'insertEquation',
	'hyperlink',
	'find',
	'resetLayout',
	'zoomOut',
	'zoomIn',
	'zoomToFit',
	'previousSlide',
	'nextSlide',
];

/** The `actions` keys of the built-in groups (`BUILDING_BLOCK_GROUP_IDS` in `ooxml-ui/pptx`). */
export const BUILDING_BLOCK_GROUP_IDS = [
	'passwordProtection',
	'documentProperties',
	'fontEmbedding',
	'digitalSignatures',
	'setUpSlideShow',
	'customShows',
	'broadcast',
	'shortcuts',
];

/** A two-slide deck whose first slide reads "Alpha beta alpha". */
export async function buildingBlocksContractDeck(Handler) {
	const { handler, data } = await Handler.create({ initialSlideCount: 2 });
	try {
		data.slides[0].elements = [
			{ id: 'title', type: 'text', x: 67, y: 40, width: 400, height: 60, text: 'Alpha beta alpha' },
		];
		data.slides[0].isDirty = true;
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

const png = () => new File([new Uint8Array([137, 80, 78, 71])], 'a.png', { type: 'image/png' });

async function until(harness, expect, check, label) {
	const deadline = Date.now() + 15000;
	for (;;) {
		await harness.settle();
		let failure;
		try {
			check();
			return;
		} catch (error) {
			failure = error;
		}
		if (Date.now() > deadline) {
			throw new Error(`${harness.name}: timed out waiting for ${label}: ${String(failure)}`);
		}
		await new Promise((resolve) => {
			setTimeout(resolve, 10);
		});
	}
}

/**
 * Run the contract against `harness`. `expect` and `vi` are the caller's Vitest, so the contract
 * has no dependency of its own.
 */
export async function runBuildingBlocksContract({ harness, expect, vi, Handler, Y }) {
	const content = await buildingBlocksContractDeck(Handler);
	const onRequest = vi.fn();
	const onDirtyChange = vi.fn();
	const host = document.createElement('div');
	document.body.append(host);
	const options = (canEdit, extra = {}) => ({
		content,
		canEdit,
		onRequest,
		onDirtyChange,
		...extra,
	});
	try {
		// 1. View-only by default: every edit workflow is refused with a code, nothing changes.
		await harness.mount(host, options(false));
		await until(harness, expect, () => expect(harness.state().slideCount).toBe(2), 'the deck');
		expect(harness.state().error).toBeNull();
		// The effective mode: `preview` while the blocks may not edit.
		expect(harness.state().mode).toBe('preview');
		// The shell state: view-only, not sharing.
		expect(harness.state().shellState).toMatchObject({ canEdit: false, status: 'disconnected' });
		expect(harness.state().collaborationState).toMatchObject({ active: false, status: 'off' });
		// Every built-in action group is registered.
		for (const id of BUILDING_BLOCK_GROUP_IDS) {
			expect(harness.actions()[id], `actions.${id}`).toBeDefined();
		}
		// Read-only, the dialog groups still describe the deck but refuse to change it.
		expect(harness.actions().passwordProtection.state.openPassword).toBe(false);
		expect(harness.actions().passwordProtection.setOpenPassword('abcd', 'abcd')).toEqual({
			ok: false,
			code: 'read-only',
		});
		expect(harness.actions().customShows.create('Contract')).toEqual({
			ok: false,
			code: 'read-only',
		});
		expect(harness.actions().availability.shortcuts.available).toBe(true);
		expect(harness.actions().availability.broadcast).toEqual({
			available: false,
			reason: 'unsupported',
			needsHostUi: true,
		});
		expect(harness.actions().availability.insertImage).toEqual({
			available: false,
			reason: 'read-only',
			needsHostUi: true,
		});
		expect(await harness.actions().insertImage(png())).toEqual({ ok: false, code: 'read-only' });
		expect(harness.actions().smartArt.insert('basicBlockList', ['A'])).toEqual({
			ok: false,
			code: 'read-only',
		});
		expect(harness.actions().unmountedDialogActions).toEqual(['onOpenBroadcastDialog']);

		// 2. Granting edit permission opens the workflows; selection-bound ones still say why not.
		await harness.update(options(true));
		await until(
			harness,
			expect,
			() => expect(harness.actions().availability.insertImage.available).toBe(true),
			'edit permission',
		);
		expect(harness.actions().availability.hyperlink).toEqual({
			available: false,
			reason: 'no-selection',
			needsHostUi: true,
		});
		expect(harness.actions().hyperlink.current).toBeNull();
		expect(harness.actions().hyperlink.slideCount).toBe(2);
		expect(harness.actions().equation.current).toBeNull();
		expect(harness.state().mode).toBe('edit');

		// 3. The toolbar's Pictures control raises a host request instead of a picker.
		await harness.clickPictures();
		await until(
			harness,
			expect,
			() => expect(onRequest).toHaveBeenCalledWith({ kind: 'image' }),
			'the request',
		);
		expect(harness.actions().pending).toEqual({ kind: 'image' });
		harness.actions().cancel();
		await until(harness, expect, () => expect(harness.actions().pending).toBeNull(), 'cancel');

		// 4. Validation before anything changes.
		const text = new File(['x'], 'a.txt', { type: 'text/plain' });
		expect(await harness.actions().insertImage(text)).toEqual({
			ok: false,
			code: 'unsupported-type',
		});
		expect(
			harness.actions().chart.insert('column', { data: { categories: [], series: [] } }),
		).toEqual({ ok: false, code: 'invalid-data' });

		// 5. An insert runs the viewer's own code and marks the deck dirty.
		expect(harness.actions().smartArt.insert('basicBlockList', ['A', 'B'])).toEqual({ ok: true });
		await until(harness, expect, () => expect(onDirtyChange).toHaveBeenCalledWith(true), 'dirty');

		// 6. Find / Replace over the host's own bar.
		harness.actions().findReplace.setOpen(true);
		harness.actions().findReplace.setQuery('alpha');
		await until(
			harness,
			expect,
			() => expect(harness.actions().findReplace.matches).toHaveLength(2),
			'the matches',
		);
		expect(harness.actions().findReplace.open).toBe(true);
		expect(harness.actions().findReplace.currentIndex).toBe(0);
		harness.actions().findReplace.next();
		await until(
			harness,
			expect,
			() => expect(harness.actions().findReplace.currentIndex).toBe(1),
			'the next match',
		);
		harness.actions().findReplace.setOpen(false);
		await until(
			harness,
			expect,
			() => expect(harness.actions().findReplace.matches).toHaveLength(0),
			'the closed bar',
		);
		expect(harness.actions().findReplace.replaceAll()).toEqual({ ok: false, code: 'no-match' });

		// 7. One operation per dialog group, with the same codes everywhere.
		const actions = () => harness.actions();
		expect(actions().passwordProtection.setOpenPassword('ab', 'ab')).toEqual({
			ok: false,
			code: 'password-too-short',
		});
		expect(actions().passwordProtection.setModifyPassword('abcd')).toEqual({
			ok: false,
			code: 'unsupported',
		});
		expect(actions().passwordProtection.setMarkedFinal(true)).toEqual({ ok: true });
		await until(
			harness,
			expect,
			() => expect(actions().passwordProtection.state.markedFinal).toBe(true),
			'Mark as Final',
		);
		expect(actions().passwordProtection.setMarkedFinal(false)).toEqual({ ok: true });
		expect(actions().documentProperties.apply({ core: { title: 'Contract title' } })).toEqual({
			ok: true,
		});
		await until(
			harness,
			expect,
			() => expect(actions().documentProperties.current.core.title).toBe('Contract title'),
			'the document title',
		);
		expect(actions().fontEmbedding.state.canEmbed).toBe(false);
		expect(actions().fontEmbedding.setEmbedFonts(true)).toEqual({ ok: false, code: 'unsupported' });
		expect(actions().digitalSignatures.state.signed).toBe(false);
		expect(actions().digitalSignatures.remove()).toEqual({ ok: false, code: 'no-signatures' });
		expect(actions().digitalSignatures.sign()).toEqual({ ok: false, code: 'unsupported' });
		expect(actions().setUpSlideShow.current.loop).toBe(false);
		expect(actions().setUpSlideShow.apply({ loopContinuously: true })).toEqual({ ok: true });
		expect(actions().setUpSlideShow.apply({ showSlidesMode: 'range', showSlidesFrom: 9 })).toEqual({
			ok: false,
			code: 'invalid-range',
		});
		await until(
			harness,
			expect,
			() => expect(actions().setUpSlideShow.current.loop).toBe(true),
			'the slide show settings',
		);
		expect(actions().customShows.create('Contract')).toEqual({ ok: true });
		await until(
			harness,
			expect,
			() => expect(actions().customShows.shows).toHaveLength(1),
			'the show',
		);
		expect(actions().customShows.slides).toHaveLength(2);
		expect(actions().customShows.start('not-a-show')).toEqual({ ok: false, code: 'unknown-show' });
		expect(actions().broadcast.supported).toBe(false);
		expect(actions().shortcuts.catalogue.length).toBeGreaterThan(0);
		actions().shortcuts.setOpen(true);
		await until(
			harness,
			expect,
			() => expect(actions().shortcuts.open).toBe(true),
			'the shortcut help',
		);
		expect(onRequest).toHaveBeenCalledWith({ kind: 'shortcuts', open: true });
		expect(actions().pending).toEqual({ kind: 'shortcuts', open: true });
		actions().cancel();
		await until(harness, expect, () => expect(actions().shortcuts.open).toBe(false), 'cancel');
		expect(actions().pending).toBeNull();

		// 8. Layouts: only the active slide's master's layouts apply.
		const layout = harness.actions().layout;
		const unknown = await layout.apply('ppt/slideLayouts/not-a-layout.xml');
		expect(unknown).toEqual({
			ok: false,
			code: layout.available.length > 0 ? 'unknown-layout' : 'no-layout',
		});

		// 8b. Shell state, a host-owned room joined and left, and live options (no remount).
		await runShellContract({
			harness,
			expect,
			until: (check, label) => until(harness, expect, check, label),
			options: (extra) => options(true, extra),
			Y,
			onDirtyChange,
		});

		// 9. Taking permission away again refuses at once.
		await harness.update(options(false));
		await until(
			harness,
			expect,
			() => expect(harness.actions().availability.insertChart.available).toBe(false),
			'view-only again',
		);
		expect(harness.actions().equation.insert({})).toEqual({ ok: false, code: 'read-only' });
		expect(harness.state().mode).toBe('preview');

		// 10. Ports with the compact toolbar render the same controls in the same order.
		if (harness.toolbar) {
			const toolbar = host.querySelector('[data-pptx-block="toolbar"]');
			expect(toolbar).not.toBeNull();
			const ids = [...toolbar.querySelectorAll('[data-control]')].map(
				(node) => node.dataset.control,
			);
			expect(ids).toEqual(COMPACT_TOOLBAR_CONTROLS);
			expect(host.querySelector('[data-pptx-block="canvas"]')).not.toBeNull();
			expect(host.querySelector('[data-pptx-block="slides-pane"]')).not.toBeNull();
		}
	} finally {
		harness.destroy();
		host.remove();
	}
}
