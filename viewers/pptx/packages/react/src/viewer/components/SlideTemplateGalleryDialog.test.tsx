import { SLIDE_TEMPLATES } from 'ooxml-ui/pptx';
import { keyToLabel, translationsEn } from 'ooxml-ui/pptx/i18n';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? keyToLabel(key),
	}),
}));

// The shared footer renders its buttons client-side in a shadow root, which
// static markup cannot show. Stand in with plain buttons carrying the same
// action model so these markup tests keep covering label, order and gating.
vi.mock<typeof import('./DialogFooter')>(import('./DialogFooter'), () => ({
	DialogFooter: ({
		actions,
	}: {
		actions: readonly { id: string; label: string; disabled?: boolean }[];
	}) =>
		React.createElement(
			'div',
			{ 'data-dialog-footer': '' },
			actions.map((action) =>
				React.createElement('button', { key: action.id, disabled: action.disabled }, action.label),
			),
		),
}));

const { SlideTemplateGalleryDialog } = await import('./SlideTemplateGalleryDialog');

describe('slide template gallery dialog', () => {
	it('renders nothing when closed', () => {
		const html = renderToStaticMarkup(
			<SlideTemplateGalleryDialog isOpen={false} onClose={vi.fn()} onInsert={vi.fn()} />,
		);
		expect(html).toBe('');
	});

	it('renders a dialog with one option per catalog template', () => {
		const html = renderToStaticMarkup(
			<SlideTemplateGalleryDialog isOpen onClose={vi.fn()} onInsert={vi.fn()} />,
		);
		expect(html).toContain('role="dialog"');
		expect(html).toContain('aria-label="Slide Templates"');
		const optionCount = (html.match(/role="option"/gu) ?? []).length;
		expect(optionCount).toBe(SLIDE_TEMPLATES.length);
		expect(html).toContain('aria-label="Title Slide"');
		expect(html).toContain('aria-label="Agenda"');
		expect(html).toContain('>Insert</button>');
		expect(html).toContain('>Cancel</button>');
	});

	it('carries the scoped stylesheet reset on both fixed layers and the previews', () => {
		// Used on its own the dialog sits outside any viewer root, so the backdrop and the
		// dialog layer mark themselves for the stylesheet's scoped preflight.
		const html = renderToStaticMarkup(
			<SlideTemplateGalleryDialog isOpen onClose={vi.fn()} onInsert={vi.fn()} />,
		);
		const layers = html.match(/<(button|div) [^>]*class="fixed inset-0[^"]*"[^>]*>/gu) ?? [];
		expect(layers).toHaveLength(2);
		for (const layer of layers) {
			expect(layer).toContain('data-pptx-scope=""');
		}
		const previews = (html.match(/aria-hidden="true" data-pptx-scope=""/gu) ?? []).length;
		expect(previews).toBe(SLIDE_TEMPLATES.length);
	});

	it('renders live previews with the provided theme scheme colour', () => {
		const html = renderToStaticMarkup(
			<SlideTemplateGalleryDialog
				isOpen
				onClose={vi.fn()}
				onInsert={vi.fn()}
				scheme={{ accent1: '#ba0021' }}
			/>,
		);
		expect(html.toLowerCase()).toContain('#ba0021');
	});
});
