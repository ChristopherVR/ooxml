import {
	defineTeamsPresentationPreview,
	type TeamsPresentationPreview,
} from './presentation-preview';

const engine = vi.hoisted(() => ({
	load: vi.fn(),
	dispose: vi.fn(),
	disposeStage: vi.fn(),
	revoke: vi.fn(),
	fieldContext: vi.fn(() => ({})),
	outlineListeners: new Set<() => void>(),
	registerFonts: vi.fn((_fonts: unknown) => vi.fn()),
}));
vi.mock('ooxml-ui/pptx/dom', () => ({
	loadPresentation: engine.load,
	revokeBlobUrls: engine.revoke,
	disposePptxRenderStage: engine.disposeStage,
	createDefaultRegistry: () => ({}),
	createTranslator: () => (key: string) => key,
	buildFieldSubstitutionContext: engine.fieldContext,
	buildEmbeddedFontStyles: () => ({ fontFaceCss: '', objectUrls: [] }),
	glyphOutlineFontCache: {
		registerEmbeddedFonts: engine.registerFonts,
		subscribe: (listener: () => void) => {
			engine.outlineListeners.add(listener);
			return () => engine.outlineListeners.delete(listener);
		},
	},
	renderSlideStage: ({ slide }: { slide: { id: string } }) => {
		const node = document.createElement('div');
		node.setAttribute('role', 'region');
		node.textContent = slide.id;
		const audio = document.createElement('audio');
		audio.src = 'blob:media';
		node.append(audio);
		return node;
	},
}));
beforeAll(() => defineTeamsPresentationPreview());
beforeEach(() => {
	engine.load.mockReset();
	engine.dispose.mockClear();
	engine.disposeStage.mockClear();
	engine.revoke.mockClear();
	engine.registerFonts.mockClear();
	vi.spyOn(HTMLMediaElement.prototype, 'pause').mockImplementation(() => {});
	vi.spyOn(HTMLMediaElement.prototype, 'load').mockImplementation(() => {});
});
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
const deck = (id = 'one') => ({
	canvasSize: { width: 960, height: 540 },
	handler: { dispose: engine.dispose },
	blobUrls: [`blob:${id}`],
	embeddedFonts: [],
	warnings: [],
	headerFooter: {},
	customProperties: [],
	mediaDataUrls: new Map(),
	slides: [
		{
			id,
			elements: [],
			notes: '<script>literal notes</script>',
			warnings: [{ message: 'Approximate element' }],
		},
		{ id: 'two', elements: [], hidden: true },
	],
});
async function mount() {
	const preview = document.createElement('teams-presentation-preview') as TeamsPresentationPreview;
	document.body.append(preview);
	await preview.updateComplete;
	return preview;
}

describe('presentation reading preview', () => {
	it('retires each rendered stage before navigation, reload and disconnect', async () => {
		engine.load.mockResolvedValue(deck());
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		const first = preview.shadowRoot!.querySelector('[role=region]')!;
		preview
			.shadowRoot!.querySelector('.stage-frame')!
			.dispatchEvent(
				new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, composed: true }),
			);
		await preview.updateComplete;
		expect(engine.disposeStage).toHaveBeenCalledExactlyOnceWith(first);
		const second = preview.shadowRoot!.querySelector('[role=region]')!;
		await preview.load(new Uint8Array([2]));
		expect(engine.disposeStage).toHaveBeenNthCalledWith(2, second);
		const reloaded = preview.shadowRoot!.querySelector('[role=region]')!;
		preview.remove();
		expect(engine.disposeStage).toHaveBeenNthCalledWith(3, reloaded);
	});

	it('navigates, renders safe notes, stops media and releases the archive and URLs', async () => {
		engine.load.mockResolvedValue(deck());
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		expect(engine.load.mock.calls[0]?.[1]).toMatchObject({
			allowExternalImages: false,
			maxUncompressedBytes: 134_217_728,
		});
		expect(preview.shadowRoot!.querySelector('[role=region]')!.getAttribute('aria-label')).toBe(
			'Slide 1',
		);
		expect(preview.shadowRoot!.querySelector('script')).toBeNull();
		expect(preview.shadowRoot!.textContent).toContain('<script>literal notes</script>');
		expect(preview.shadowRoot!.textContent).toContain('Approximate element');
		const firstAudio = preview.shadowRoot!.querySelector('audio')!;
		const frame = preview.shadowRoot!.querySelector('.stage-frame')!;
		frame.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, composed: true }),
		);
		await preview.updateComplete;
		expect(preview.shadowRoot!.textContent).toContain('Slide 2 of 2 (hidden)');
		expect(firstAudio.pause).toHaveBeenCalled();
		expect(firstAudio.hasAttribute('src')).toBe(false);
		preview.remove();
		expect(engine.dispose).toHaveBeenCalledTimes(1);
		expect(engine.revoke).toHaveBeenCalledWith(['blob:one']);
	});
	it('repaints the slide when deferred glyph outlines arrive, until it closes', async () => {
		engine.load.mockResolvedValue(deck());
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		const first = preview.shadowRoot!.querySelector('[role=region]')!;
		expect(engine.outlineListeners.size).toBe(1);
		for (const listener of engine.outlineListeners) listener();
		await preview.updateComplete;
		const repainted = preview.shadowRoot!.querySelector('[role=region]')!;
		expect(repainted).not.toBe(first);
		expect(repainted.getAttribute('aria-label')).toBe('Slide 1');
		expect(engine.disposeStage).toHaveBeenCalledWith(first);
		preview.remove();
		expect(engine.outlineListeners.size).toBe(0);
	});
	it("releases each deck's glyph-outline fonts on reload and disconnect", async () => {
		engine.load.mockResolvedValueOnce(deck('one')).mockResolvedValueOnce(deck('two'));
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		await preview.load(new Uint8Array([2]));
		const [first, second] = engine.registerFonts.mock.results.map((result) => result.value);
		expect(engine.registerFonts).toHaveBeenCalledTimes(2);
		expect(first).toHaveBeenCalledTimes(1);
		expect(second).not.toHaveBeenCalled();
		preview.remove();
		expect(first).toHaveBeenCalledTimes(1);
		expect(second).toHaveBeenCalledTimes(1);
	});
	it('passes the deck start slide number to the field context', async () => {
		engine.load.mockResolvedValue({ ...deck(), firstSlideNumber: 45 });
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		expect(engine.fieldContext).toHaveBeenLastCalledWith(
			expect.objectContaining({ firstSlideNumber: 45 }),
		);
		preview.remove();
	});
	it('disposes a superseded load without replacing the current slide', async () => {
		let finish!: (value: ReturnType<typeof deck>) => void;
		engine.load
			.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)))
			.mockResolvedValueOnce(deck('current'));
		const preview = await mount();
		const first = preview.load(new Uint8Array([1]));
		await vi.waitFor(() => expect(engine.load).toHaveBeenCalledTimes(1));
		await preview.load(new Uint8Array([2]));
		finish(deck('stale'));
		await first;
		expect(preview.data?.slides[0]?.id).toBe('current');
		expect(engine.dispose).toHaveBeenCalledTimes(1);
		expect(engine.revoke).toHaveBeenCalledWith(['blob:stale']);
	});
	it('releases a load that completes after the preview closes', async () => {
		let finish!: (value: ReturnType<typeof deck>) => void;
		engine.load.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)));
		const preview = await mount();
		const pending = preview.load(new Uint8Array([1]));
		await vi.waitFor(() => expect(engine.load).toHaveBeenCalledTimes(1));
		preview.remove();
		finish(deck('closed'));
		await pending;
		expect(preview.data).toBeNull();
		expect(engine.dispose).toHaveBeenCalledTimes(1);
		expect(engine.revoke).toHaveBeenCalledWith(['blob:closed']);
	});
});
