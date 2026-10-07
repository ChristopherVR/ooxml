import {
	defineTeamsPresentationPreview,
	type TeamsPresentationPreview,
} from './presentation-preview';

const engine = vi.hoisted(() => ({ load: vi.fn(), dispose: vi.fn(), revoke: vi.fn() }));
vi.mock('ooxml-ui/pptx/dom', () => ({
	loadPresentation: engine.load,
	revokeBlobUrls: engine.revoke,
	createDefaultRegistry: () => ({}),
	createTranslator: () => (key: string) => key,
	buildFieldSubstitutionContext: () => ({}),
	buildEmbeddedFontStyles: () => ({ fontFaceCss: '', objectUrls: [] }),
	glyphOutlineFontCache: { registerEmbeddedFonts: vi.fn() },
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
	engine.revoke.mockClear();
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
