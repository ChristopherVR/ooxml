import {
	defineTeamsPresentationPreview,
	type TeamsPresentationPreview,
} from './presentation-preview';

const engine = vi.hoisted(() => ({
	load: vi.fn(),
	dispose: vi.fn(),
	svg: vi.fn(() => '<svg xmlns="http://www.w3.org/2000/svg"><text>Slide content</text></svg>'),
}));
vi.mock('ooxml-core/pptx', () => ({
	PptxHandler: class {
		load = engine.load;
		dispose = engine.dispose;
	},
	SvgExporter: { exportSlide: engine.svg },
}));
beforeAll(() => defineTeamsPresentationPreview());
beforeEach(() => {
	vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
	const NativeURL = URL;
	vi.stubGlobal(
		'URL',
		class extends NativeURL {
			static override createObjectURL = vi.fn(() => `blob:slide-${Math.random()}`);
			static override revokeObjectURL = vi.fn();
		},
	);
	engine.load.mockReset();
	engine.dispose.mockClear();
	engine.svg.mockClear();
});
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
	vi.restoreAllMocks();
});
const deck = () => ({
	width: 960,
	height: 540,
	slides: [
		{
			id: 'one',
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

describe('static presentation preview', () => {
	it('navigates slides, shows safe notes and warnings, and releases images and the handler', async () => {
		engine.load.mockResolvedValue(deck());
		const preview = await mount();
		await preview.load(new Uint8Array([1]));
		expect(engine.load.mock.calls[0]?.[1]).toMatchObject({
			allowExternalImages: false,
			maxUncompressedBytes: 134_217_728,
		});
		expect(preview.shadowRoot!.querySelector('img')!.src).toContain('blob:slide-');
		expect(preview.shadowRoot!.querySelector('svg')).toBeNull();
		expect(preview.shadowRoot!.querySelector('script')).toBeNull();
		expect(preview.shadowRoot!.textContent).toContain('<script>literal notes</script>');
		expect(preview.shadowRoot!.textContent).toContain('Approximate element');
		const image = preview.image;
		(preview.shadowRoot!.querySelectorAll('button')[1] as HTMLButtonElement).click();
		await preview.updateComplete;
		await preview.updateComplete;
		expect(preview.shadowRoot!.textContent).toContain('Slide 2 of 2 (hidden)');
		expect(URL.revokeObjectURL).toHaveBeenCalledWith(image);
		preview.remove();
		expect(engine.dispose).toHaveBeenCalled();
	});
	it('ignores a superseded parse and disposes its resources', async () => {
		let finish!: (value: ReturnType<typeof deck>) => void;
		engine.load
			.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)))
			.mockResolvedValueOnce(deck());
		const preview = await mount();
		const first = preview.load(new Uint8Array([1]));
		await vi.waitFor(() => expect(engine.load).toHaveBeenCalledTimes(1));
		await preview.load(new Uint8Array([2]));
		const image = preview.image;
		finish(deck());
		await first;
		expect(preview.image).toBe(image);
		expect(engine.dispose).toHaveBeenCalledTimes(1);
	});
});
