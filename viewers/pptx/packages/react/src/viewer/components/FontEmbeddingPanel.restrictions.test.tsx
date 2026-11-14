// @vitest-environment happy-dom
/**
 * FontEmbeddingPanel: a typeface whose licence (OS/2 fsType) restricts
 * embedding shows its restriction and the reason. The decision is the shared
 * `describeEmbeddedFontRestrictions`; this pins the React mapping onto it.
 */
import { describeEmbeddedFontRestrictions } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { FontEmbeddingPanel } from './FontEmbeddingPanel';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

vi.mock<typeof import('ooxml-ui/pptx')>(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	scanAvailableFontFamilies: async () => new Set<string>(),
}));

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	delete (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT;
});

describe('FontEmbeddingPanel licence restrictions', () => {
	it('shows a restricted and a preview-only badge with their reasons', async () => {
		const fonts = [
			{ name: 'Brand', fsType: 0x0002 },
			{ name: 'Proof', fsType: 0x0004 },
			{ name: 'Open', fsType: 0 },
		];
		await act(async () => {
			root.render(
				<FontEmbeddingPanel
					isOpen
					embedFontsEnabled
					usedFontFamilies={['Brand', 'Proof', 'Open']}
					embeddedFonts={fonts.map((font) => font.name)}
					embeddedFontRestrictions={describeEmbeddedFontRestrictions(fonts)}
					onClose={() => undefined}
					onToggleEmbedFonts={() => undefined}
				/>,
			);
		});
		const restricted = container.querySelector('[data-font-restriction="restricted"]');
		expect(restricted?.textContent).toBe(translationsEn['pptx.fonts.restricted']);
		expect(restricted?.getAttribute('title')).toBe(translationsEn['pptx.fonts.restrictedReason']);
		expect(
			container.querySelector('[data-font-restriction-reason="restricted"]')?.textContent,
		).toBe(translationsEn['pptx.fonts.restrictedReason']);
		expect(container.querySelector('[data-font-restriction="preview-print"]')?.textContent).toBe(
			translationsEn['pptx.fonts.previewOnly'],
		);
		expect(container.querySelectorAll('[data-font-restriction]')).toHaveLength(2);
	});
});
