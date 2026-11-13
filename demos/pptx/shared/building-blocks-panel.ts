/**
 * The host controls of the building-blocks demo page (`?blocks=1`), shared by every framework
 * demo that has building blocks: plain DOM over the binding's `actions` and state, so the
 * framework-neutral spec (`e2e/pptx/building-blocks.spec.ts`) reads the same labels everywhere.
 * A binding only supplies the adapter and re-renders by calling `refresh()`.
 */

type Result = { ok: true } | { ok: false; code: string };

/** The slice of a binding's building blocks the panel drives (structural, no package import). */
export interface BuildingBlocksPanelAdapter {
	actions(): {
		pending: { kind: string } | null;
		cancel(): void;
		insertImage(file: File): Promise<Result>;
		smartArt: { insert(layout: never, items: string[]): Result };
		hyperlink: { apply(data: never): Result; current: unknown };
		findReplace: {
			open: boolean;
			matches: readonly unknown[];
			currentIndex: number;
			setOpen(open: boolean): void;
			setQuery(query: string): void;
			next(): void;
		};
		documentProperties: {
			current: { core: { title?: string } };
			apply(patch: { core: { title: string } }): Result;
		};
		setUpSlideShow: {
			current: { loop: boolean };
			apply(patch: { loopContinuously: boolean }): Result;
		};
		customShows: { shows: readonly unknown[]; create(name: string): Result };
		shortcuts: { open: boolean; toggle(): void };
	};
	state(): { mode: string; slideCount: number };
	setCanEdit(canEdit: boolean): void;
}

const LINK = {
	targetType: 'url',
	url: 'https://example.com/',
	tooltip: '',
	emailAddress: '',
	emailSubject: '',
	slideNumber: 1,
	filePath: '',
	actionVerb: 'nextSlide',
};

function labelled<K extends 'output' | 'button' | 'input'>(
	doc: Document,
	tag: K,
	label: string,
): HTMLElementTagNameMap[K] {
	const node = doc.createElement(tag);
	node.setAttribute('aria-label', label);
	if (node instanceof HTMLButtonElement) {
		node.type = 'button';
		node.textContent = label;
	}
	return node;
}

/** Mount the panel into `host`; returns `refresh` (call it after every binding update). */
export function mountBuildingBlocksPanel(
	host: HTMLElement,
	adapter: BuildingBlocksPanelAdapter,
): { refresh(): void } {
	const doc = host.ownerDocument;
	host.dataset.buildingBlocksPanel = '';
	const outputs = {
		mode: labelled(doc, 'output', 'Blocks mode'),
		slides: labelled(doc, 'output', 'Blocks slides'),
		pending: labelled(doc, 'output', 'Blocks pending'),
		result: labelled(doc, 'output', 'Blocks result'),
		matches: labelled(doc, 'output', 'Blocks matches'),
		current: labelled(doc, 'output', 'Blocks current match'),
		title: labelled(doc, 'output', 'Blocks title'),
		loop: labelled(doc, 'output', 'Blocks loop'),
		shows: labelled(doc, 'output', 'Blocks custom shows'),
		shortcuts: labelled(doc, 'output', 'Blocks shortcuts'),
	};
	const show = (result: Result) => {
		outputs.result.textContent = result.ok ? 'ok' : result.code;
	};
	const button = (label: string, run: () => void) => {
		const node = labelled(doc, 'button', label);
		node.onclick = run;
		return node;
	};
	const picture = labelled(doc, 'input', 'Host picture file');
	picture.type = 'file';
	picture.accept = 'image/*';
	picture.onchange = () => {
		const file = picture.files?.[0];
		if (file) {
			void adapter
				.actions()
				.insertImage(file)
				.then((result) => show(result));
		}
		picture.value = '';
	};
	const query = labelled(doc, 'input', 'Host find query');
	const actions = () => adapter.actions();
	host.replaceChildren(
		...Object.values(outputs),
		button('Grant editing', () => adapter.setCanEdit(true)),
		button('Revoke editing', () => adapter.setCanEdit(false)),
		picture,
		button('Host SmartArt', () =>
			show(actions().smartArt.insert('basicBlockList' as never, ['A', 'B'])),
		),
		button('Host link', () => show(actions().hyperlink.apply(LINK as never))),
		query,
		button('Host find', () => {
			actions().findReplace.setOpen(true);
			actions().findReplace.setQuery(query.value);
		}),
		button('Host next match', () => actions().findReplace.next()),
		button('Host close find', () => actions().findReplace.setOpen(false)),
		button('Host set title', () =>
			show(actions().documentProperties.apply({ core: { title: 'Blocks title' } })),
		),
		button('Host loop show', () =>
			show(actions().setUpSlideShow.apply({ loopContinuously: true })),
		),
		button('Host add show', () => show(actions().customShows.create('Host show'))),
		button('Host shortcuts', () => actions().shortcuts.toggle()),
		button('Host cancel', () => actions().cancel()),
	);
	const refresh = () => {
		const now = actions();
		const state = adapter.state();
		outputs.mode.textContent = state.mode;
		outputs.slides.textContent = String(state.slideCount);
		outputs.pending.textContent = now.pending?.kind ?? 'none';
		outputs.matches.textContent = now.findReplace.open
			? String(now.findReplace.matches.length)
			: 'closed';
		outputs.current.textContent = String(now.findReplace.currentIndex);
		outputs.title.textContent = now.documentProperties.current.core.title ?? '';
		outputs.loop.textContent = String(now.setUpSlideShow.current.loop);
		outputs.shows.textContent = String(now.customShows.shows.length);
		outputs.shortcuts.textContent = now.shortcuts.open ? 'open' : 'closed';
	};
	refresh();
	return { refresh };
}
