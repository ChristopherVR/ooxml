import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { buildViewerCss } from '../../styles/css';
import { createFormatBackgroundPanel } from './format-background-panel';

describe('createFormatBackgroundPanel', () => {
	it('drives the background live on input and records the committed pick as a recent colour', () => {
		const t = createTranslator();
		const setSlideBackgroundColor = vi.fn();
		const pushRecentColor = vi.fn();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor,
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
			pushRecentColor,
		});
		const input = panel.el.querySelector<HTMLInputElement>('input[type="color"]')!;

		input.value = '#123456';
		input.dispatchEvent(new Event('input', { bubbles: true }));
		expect(setSlideBackgroundColor).toHaveBeenCalledWith('#123456');
		// Dragging inside the native dialog streams `input` events: none of
		// those is a pick.
		expect(pushRecentColor).not.toHaveBeenCalled();

		input.dispatchEvent(new Event('change', { bubbles: true }));
		expect(pushRecentColor).toHaveBeenCalledWith('#123456');
	});

	it('tolerates a caller without a recent-colours sink', () => {
		const t = createTranslator();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
		});
		const input = panel.el.querySelector<HTMLInputElement>('input[type="color"]')!;
		expect(() => input.dispatchEvent(new Event('change', { bubbles: true }))).not.toThrow();
	});

	it('toggles Hide Background Graphics via the checkbox', () => {
		const t = createTranslator();
		const setHideBackgroundGraphics = vi.fn();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics,
		});
		const checkbox = panel.el.querySelector<HTMLInputElement>('pptx-ui-checkbox')!;

		checkbox.checked = true;
		checkbox.dispatchEvent(new Event('change', { bubbles: true }));
		expect(setHideBackgroundGraphics).toHaveBeenCalledWith(true);

		checkbox.checked = false;
		checkbox.dispatchEvent(new Event('change', { bubbles: true }));
		expect(setHideBackgroundGraphics).toHaveBeenCalledWith(false);
	});

	it('names its close button for itself, not for the find bar', () => {
		const t = createTranslator();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
		});
		const labels = [...panel.el.querySelectorAll('button')].map((button) =>
			button.getAttribute('aria-label'),
		);
		expect(labels).toContain(t('pptx.common.close'));
		expect(labels).not.toContain(t('pptx.findReplace.closeAriaLabel'));
	});

	it('takes focus when opened and closes on Escape, handing focus back', () => {
		const t = createTranslator();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
		});
		const opener = document.createElement('button');
		const editorKeys = vi.fn();
		document.body.append(opener, panel.el);
		document.body.addEventListener('keydown', editorKeys);
		try {
			opener.focus();
			panel.setOpen(true);
			expect(panel.el.contains(document.activeElement)).toBeTruthy();

			const escape = new KeyboardEvent('keydown', {
				key: 'Escape',
				bubbles: true,
				cancelable: true,
			});
			document.activeElement!.dispatchEvent(escape);
			expect(panel.isOpen()).toBeFalsy();
			expect(panel.el.hidden).toBeTruthy();
			expect(escape.defaultPrevented).toBeTruthy();
			// The editor's own Escape (clear the selection) must not also run.
			expect(editorKeys).not.toHaveBeenCalled();
			expect(document.activeElement).toBe(opener);
		} finally {
			document.body.removeEventListener('keydown', editorKeys);
			opener.remove();
			panel.el.remove();
		}
	});

	it('hands focus back even when hiding the panel blurs its focused control', () => {
		const t = createTranslator();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
		});
		const opener = document.createElement('button');
		document.body.append(opener, panel.el);
		// A browser drops focus from a control the moment its container is hidden;
		// happy-dom does not, so mimic it on this one element.
		const hidden = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'hidden')!;
		Object.defineProperty(panel.el, 'hidden', {
			configurable: true,
			get() {
				return hidden.get!.call(this);
			},
			set(value: boolean) {
				hidden.set!.call(this, value);
				if (value && this.contains(document.activeElement)) {
					(document.activeElement as HTMLElement).blur();
				}
			},
		});
		try {
			opener.focus();
			panel.setOpen(true);
			expect(panel.el.contains(document.activeElement)).toBeTruthy();

			panel.setOpen(false);
			expect(document.activeElement).toBe(opener);
		} finally {
			opener.remove();
			panel.el.remove();
		}
	});

	it('has a style rule for every class it renders', () => {
		const t = createTranslator();
		const panel = createFormatBackgroundPanel(document, t, {
			setSlideBackgroundColor: vi.fn(),
			clearSlideBackground: vi.fn(),
			setHideBackgroundGraphics: vi.fn(),
		});
		const css = buildViewerCss();
		const classes = new Set<string>();
		for (const node of [panel.el, ...panel.el.querySelectorAll<HTMLElement>('[class]')]) {
			for (const name of node.classList) {
				if (name.startsWith('pptxv-')) {
					classes.add(name);
				}
			}
		}
		expect(classes).toContain('pptxv-format-background-panel');
		const missing = [...classes].filter(
			(name) => !new RegExp(`\\.${name}(?![\\w-])`, 'u').test(css),
		);
		expect(missing).toStrictEqual([]);
	});
});
