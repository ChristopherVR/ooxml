// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { choosePdfExportSelection, registerPptxWebControls } from './index';
import type { PptxUiPdfExportDialogElement } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

type Shadowed = HTMLElement & { shadowRoot: ShadowRoot };
type Checkable = HTMLElement & { checked: boolean };

function dialog(): PptxUiPdfExportDialogElement {
	const host = document.createElement('pptx-ui-pdf-export-dialog') as PptxUiPdfExportDialogElement;
	host.slides = [{}, { hidden: true }, {}, {}];
	host.currentSlide = 2;
	document.body.append(host);
	return host;
}

const footerButton = (host: HTMLElement, id: string) =>
	(
		host.querySelector('office-ui-dialog-footer') as Shadowed
	).shadowRoot.querySelector<HTMLButtonElement>(`button[data-action="${id}"]`)!;
const byTestId = <T extends HTMLElement>(host: HTMLElement, id: string) =>
	host.querySelector<T>(`[data-testid="${id}"]`)!;
const pick = (host: HTMLElement, range: string) =>
	byTestId<Checkable>(host, `pptx-pdf-export-range-${range}`).click();

describe('pptx-ui-pdf-export-dialog', () => {
	it('opens on All with the hidden-slide setting, translated', () => {
		const host = dialog();
		host.includeHiddenSlides = true;
		host.translateLabel = (key) => (key === 'pptx.pdfExport.dialog.title' ? 'PDF-Optionen' : key);
		host.open = true;
		const shell = host.querySelector('office-ui-dialog') as HTMLElement & {
			open: boolean;
			heading: string;
		};
		expect(shell.open).toBe(true);
		expect(shell.heading).toBe('PDF-Optionen');
		expect(byTestId<Checkable>(host, 'pptx-pdf-export-range-all').checked).toBe(true);
		expect(byTestId<Checkable>(host, 'pptx-pdf-export-include-hidden').checked).toBe(true);
		expect(byTestId<HTMLInputElement>(host, 'pptx-pdf-export-from').disabled).toBe(true);
		expect(byTestId<HTMLInputElement>(host, 'pptx-pdf-export-to').value).toBe('4');
	});

	it('confirms the current slide as an explicit selection and closes', () => {
		const host = dialog();
		host.open = true;
		const confirmed = vi.fn();
		const closed = vi.fn();
		host.addEventListener('pdf-export-confirm', (event) =>
			confirmed((event as CustomEvent).detail),
		);
		host.addEventListener('dialog-close', closed);
		pick(host, 'current');
		footerButton(host, 'insert').click();
		expect(confirmed).toHaveBeenCalledExactlyOnceWith({
			choice: { range: 'current', from: 1, to: 4, includeHiddenSlides: false },
			selection: { slideIndexes: [2], includeHiddenSlides: true },
		});
		expect(host.open).toBe(false);
		expect(closed).toHaveBeenCalledOnce();
	});

	it('disables Export and explains an invalid range, then confirms a valid one', () => {
		const host = dialog();
		host.open = true;
		const confirmed = vi.fn();
		host.addEventListener('pdf-export-confirm', (event) =>
			confirmed((event as CustomEvent).detail),
		);
		pick(host, 'slides');
		const from = byTestId<HTMLInputElement>(host, 'pptx-pdf-export-from');
		const to = byTestId<HTMLInputElement>(host, 'pptx-pdf-export-to');
		expect(from.disabled).toBe(false);
		from.value = '3';
		to.value = '9';
		to.dispatchEvent(new Event('input'));
		expect(footerButton(host, 'insert').disabled).toBe(true);
		expect(byTestId(host, 'pptx-pdf-export-error').textContent).toBe(
			'Enter slide numbers from 1 to 4.',
		);
		to.value = '4';
		to.dispatchEvent(new Event('input'));
		expect(byTestId(host, 'pptx-pdf-export-error').textContent).toBe('');
		to.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(confirmed).toHaveBeenCalledExactlyOnceWith(
			expect.objectContaining({ selection: { from: 3, to: 4, includeHiddenSlides: false } }),
		);
	});

	it('reports Cancel without confirming', () => {
		const host = dialog();
		host.open = true;
		const confirmed = vi.fn();
		const closed = vi.fn();
		host.addEventListener('pdf-export-confirm', confirmed);
		host.addEventListener('dialog-close', closed);
		footerButton(host, 'cancel').click();
		return Promise.resolve().then(() => {
			expect(confirmed).not.toHaveBeenCalled();
			expect(closed).toHaveBeenCalledOnce();
			expect(host.open).toBe(false);
		});
	});
});

describe('pptx-ui-pdf-export-dialog inputs', () => {
	it('takes slides and the hidden-slide setting set after open', () => {
		const host = document.createElement(
			'pptx-ui-pdf-export-dialog',
		) as PptxUiPdfExportDialogElement;
		document.body.append(host);
		host.open = true;
		host.slides = [{}, {}, {}];
		host.includeHiddenSlides = true;
		expect(byTestId<HTMLInputElement>(host, 'pptx-pdf-export-to').value).toBe('3');
		expect(byTestId<Checkable>(host, 'pptx-pdf-export-include-hidden').checked).toBe(true);
		// Once the user changed something, a later input no longer resets the choice.
		pick(host, 'current');
		host.slides = [{}, {}, {}, {}];
		expect(byTestId<Checkable>(host, 'pptx-pdf-export-range-current').checked).toBe(true);
	});
});

describe('choosePdfExportSelection', () => {
	const settings = { mode: 'vector' as const, includeHiddenSlides: false };

	it('mounts the dialog in the container and resolves to the confirmed selection', async () => {
		const container = document.createElement('div');
		document.body.append(container);
		const chosen = choosePdfExportSelection(container, {
			settings,
			slides: [{}, {}],
			currentSlide: 1,
		});
		const host = container.querySelector('pptx-ui-pdf-export-dialog') as HTMLElement;
		expect(host).not.toBeNull();
		pick(host, 'current');
		footerButton(host, 'insert').click();
		await expect(chosen).resolves.toEqual({ slideIndexes: [1], includeHiddenSlides: true });
		expect(container.querySelector('pptx-ui-pdf-export-dialog')).toBeNull();
	});

	it('resolves to null on Cancel', async () => {
		const container = document.createElement('div');
		document.body.append(container);
		const chosen = choosePdfExportSelection(container, { settings, slides: [{}], currentSlide: 0 });
		footerButton(container.querySelector('pptx-ui-pdf-export-dialog')!, 'cancel').click();
		await expect(chosen).resolves.toBeNull();
	});

	it('asks nothing for the picture export', async () => {
		const container = document.createElement('div');
		await expect(
			choosePdfExportSelection(container, {
				settings: { mode: 'pictures', includeHiddenSlides: false },
				slides: [{}],
				currentSlide: 0,
			}),
		).resolves.toEqual({});
		expect(container.childElementCount).toBe(0);
	});
});
