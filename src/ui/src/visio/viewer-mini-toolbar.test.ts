import { afterEach, describe, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { mountViewer } from './binding';
import type { EditWorkerRequest } from './worker-editor';
import { registerViewerControls } from './office-ui';
import { ViewerMiniToolbar } from './viewer-mini-toolbar';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

/** Transport shim exercises the actual core backend without browser worker scheduling. */
class CoreWorker {
	onmessage: ((event: MessageEvent) => void) | null = null;
	onerror: ((event: ErrorEvent) => void) | null = null;
	#terminated = false;
	terminate(): void {
		this.#terminated = true;
	}
	postMessage(request: ArrayBuffer | EditWorkerRequest): void {
		void (async () => {
			try {
				const result =
					request instanceof ArrayBuffer
						? { document: await parseVsdx(request) }
						: await (async () => {
								const saved = await editVsdx(request.bytes, request.edits);
								return { ...saved, document: await parseVsdx(saved.bytes) };
							})();
				if (!this.#terminated) this.onmessage?.({ data: { ok: true, ...result } } as MessageEvent);
			} catch (error) {
				if (!this.#terminated)
					this.onmessage?.({ data: { ok: false, message: String(error) } } as MessageEvent);
			}
		})();
	}
}

afterEach(() => {
	vi.unstubAllGlobals();
	document.body.replaceChildren();
	localStorage.clear();
});

/** A ribbon stand-in holding the Home controls the bar mirrors. */
function bar() {
	registerViewerControls();
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const ribbon = document.createElement('div');
	// Every command the bar mirrors starts disabled, as Home's do with nothing selected.
	const ids = [
		'grow-font',
		'shrink-font',
		'bold',
		'italic',
		'underline',
		'align-left',
		'align-center',
		'align-right',
	];
	ribbon.innerHTML =
		ids
			.map((id) => {
				const title = id === 'bold' ? ' title="Bold: source formulas can refuse this edit."' : '';
				const pressed = id === 'italic' ? ' pressed="true"' : '';
				return `<office-ui-button command="${id}"${title}${pressed} disabled></office-ui-button>`;
			})
			.join('') +
		'<office-ui-select data-combo="font" disabled></office-ui-select>' +
		'<office-ui-select data-combo="font-size" disabled></office-ui-select>';
	root.append(ribbon);
	const toolbar = new ViewerMiniToolbar(root, ribbon);
	const viewport = document.createElement('div');
	root.append(viewport);
	return { root, ribbon, toolbar, viewport };
}
const control = (toolbar: ViewerMiniToolbar, id: string) =>
	toolbar.element.querySelector<HTMLElement & { disabled: boolean }>(`[command="${id}"]`)!;

describe('Mini Toolbar', () => {
	it('is hidden until shown and offers the Home text formatting commands', () => {
		const { toolbar } = bar();
		expect(toolbar.visible).toBe(false);
		for (const id of [
			'mini-grow',
			'mini-shrink',
			'mini-bold',
			'mini-italic',
			'mini-underline',
			'mini-align-left',
			'mini-align-center',
			'mini-align-right',
		])
			expect(control(toolbar, id), id).not.toBeNull();
		expect(toolbar.element.querySelector('[data-combo="mini-font"]')).not.toBeNull();
		expect(toolbar.element.getAttribute('role')).toBe('toolbar');
	});

	it('copies state from the ribbon rather than working it out again', () => {
		const { ribbon, toolbar } = bar();
		toolbar.sync();
		expect(control(toolbar, 'mini-bold').disabled).toBe(true);
		expect(control(toolbar, 'mini-italic').getAttribute('pressed')).toBe('true');
		const bold = ribbon.querySelector<HTMLElement & { disabled: boolean }>('[command="bold"]')!;
		bold.removeAttribute('disabled');
		toolbar.sync();
		expect(control(toolbar, 'mini-bold').disabled).toBe(false);
		// The tooltip keeps the command's name, not the ribbon's long explanation.
		expect(control(toolbar, 'mini-bold').title).toBe('Bold');
	});

	it('does not show when no command can act, or when the option is off', () => {
		const { ribbon, toolbar } = bar();
		for (const element of ribbon.querySelectorAll('office-ui-button, office-ui-select'))
			element.setAttribute('disabled', '');
		toolbar.showAt(100, 100);
		expect(toolbar.visible).toBe(false);
		ribbon.querySelector('[command="bold"]')!.removeAttribute('disabled');
		toolbar.showAt(100, 100);
		expect(toolbar.visible).toBe(true);
		toolbar.enabled = false;
		expect(toolbar.visible).toBe(false);
		toolbar.showAt(100, 100);
		expect(toolbar.visible).toBe(false);
	});

	it('sits above the point, and below it when there is no room', () => {
		const { ribbon, toolbar } = bar();
		ribbon.querySelector('[command="bold"]')!.removeAttribute('disabled');
		toolbar.element.getBoundingClientRect = () => ({ width: 200, height: 36 }) as DOMRect;
		toolbar.showAt(300, 200);
		expect(toolbar.element.style.top).toBe('156px');
		toolbar.showAt(300, 10);
		expect(Number.parseInt(toolbar.element.style.top)).toBeGreaterThan(10);
	});

	it('dismisses on a press elsewhere, Escape and scrolling, but not on a press inside', () => {
		const { ribbon, toolbar, viewport } = bar();
		ribbon.querySelector('[command="bold"]')!.removeAttribute('disabled');
		const dispose = toolbar.wire(viewport);
		toolbar.showAt(100, 100);
		toolbar.element.dispatchEvent(
			new PointerEvent('pointerdown', { bubbles: true, composed: true }),
		);
		expect(toolbar.visible).toBe(true);
		document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
		expect(toolbar.visible).toBe(false);
		toolbar.showAt(100, 100);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(toolbar.visible).toBe(false);
		toolbar.showAt(100, 100);
		viewport.dispatchEvent(new Event('scroll'));
		expect(toolbar.visible).toBe(false);
		dispose();
		toolbar.showAt(100, 100);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(toolbar.visible).toBe(true);
	});

	it('keeps focus where it is when a button is pressed', () => {
		const { toolbar } = bar();
		const press = new PointerEvent('pointerdown', { bubbles: true, cancelable: true });
		control(toolbar, 'mini-bold').dispatchEvent(press);
		expect(press.defaultPrevented).toBe(true);
	});
});

describe('Mini Toolbar in the viewer', () => {
	async function editable() {
		vi.stubGlobal('Worker', CoreWorker);
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host);
		await viewer.load(await createVsdxFixture('Formatted shape'));
		const root = viewer.element.shadowRoot!;
		const group = () => root.querySelector<SVGGElement>('[data-shape-id]')!;
		const toolbar = () => root.querySelector<HTMLElement>('.mini-toolbar')!;
		const rightClick = () =>
			group().dispatchEvent(
				new MouseEvent('contextmenu', {
					bubbles: true,
					cancelable: true,
					clientX: 400,
					clientY: 300,
				}),
			);
		return { viewer, root, toolbar, rightClick };
	}

	it('appears with the shape menu and formats the shape through the usual path', async () => {
		const { viewer, toolbar, rightClick } = await editable();
		expect(toolbar().hidden).toBe(true);
		rightClick();
		expect(toolbar().hidden).toBe(false);
		const bold = toolbar().querySelector<HTMLElement>('[command="mini-bold"]')!;
		expect(bold.hasAttribute('disabled')).toBe(false);
		bold.shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
		await vi.waitFor(() =>
			expect(viewer.element.document!.pages[0]!.shapes[0]!.text.bold).toBe(true),
		);
		viewer.destroy();
	});

	it('does not appear when the Mini Toolbar option is off', async () => {
		localStorage.setItem('ooxml-visio-options', JSON.stringify({ miniToolbar: false }));
		const { viewer, toolbar, rightClick } = await editable();
		rightClick();
		expect(toolbar().hidden).toBe(true);
		viewer.destroy();
	});
});
