import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React from 'react';
/**
 * Tests for TabRowActions: the Record + Share actions on the ribbon tab row.
 * Uses react-dom/server renderToStaticMarkup, matching the convention used
 * by Toolbar.test.tsx and StatusBar.test.tsx.
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it, vi } from 'vitest';

import type { TabRowActionsProps } from './TabRowActions';

vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string, opts?: Record<string, unknown>) => {
			const translations: Record<string, string> = {
				'pptx.titleBar.record': 'Record',
				'pptx.toolbar.share': 'Share',
				'pptx.toolbar.comments': 'Comments',
			};
			const v = translations[key];
			if (typeof v === 'string') {
				return v;
			}
			const fallback = translationsEn[key];
			if (fallback === undefined) {
				return key;
			}
			return opts
				? fallback.replace(/\{\{(\w+)\}\}/gu, (_m, name: string) => String(opts[name] ?? ''))
				: fallback;
		},
	}),
}));

const { TabRowActions } = await import('./TabRowActions');

function render(el: React.ReactElement): string {
	return renderToStaticMarkup(el);
}

function createProps(overrides: Partial<TabRowActionsProps> = {}): TabRowActionsProps {
	return {
		onEnterRehearsalMode: vi.fn<() => void>(),
		onOpenShareDialog: vi.fn<() => void>(),
		...overrides,
	};
}

describe('tabRowActions - default (backward compatible)', () => {
	it('renders Record and the shared Comments / Share element when hiddenActions is omitted', () => {
		const html = render(React.createElement(TabRowActions, createProps()));
		expect(html).toContain('aria-label="Record"');
		expect(html).toContain('<pptx-ui-ribbon-actions');
		expect(html).toContain('share-label="Share"');
		expect(html).not.toContain('no-share');
		// No comments handler: Comments is left out.
		expect(html).toContain('no-comments=""');
	});

	it('shows Comments with its pressed state and slide count when a handler is given', () => {
		const html = render(
			React.createElement(
				TabRowActions,
				createProps({
					onToggleComments: vi.fn<() => void>(),
					isCommentsPanelOpen: true,
					slideCommentCount: 3,
				}),
			),
		);
		expect(html).toContain('comments-label="Comments"');
		expect(html).toContain('comments-pressed=""');
		expect(html).toContain('comments-count="3"');
		expect(html).not.toContain('no-comments');
	});
});

describe('tabRowActions - hiddenActions', () => {
	it('hides Share when "share" is hidden', () => {
		const html = render(
			React.createElement(TabRowActions, createProps({ hiddenActions: ['share'] })),
		);
		expect(html).toContain('no-share=""');
		expect(html).toContain('aria-label="Record"');
	});

	it('omits the Record button when "record" is hidden', () => {
		const html = render(
			React.createElement(TabRowActions, createProps({ hiddenActions: ['record'] })),
		);
		expect(html).not.toContain('aria-label="Record"');
		expect(html).not.toContain('no-share');
	});

	it('hides both when both ids are hidden', () => {
		const html = render(
			React.createElement(TabRowActions, createProps({ hiddenActions: ['record', 'share'] })),
		);
		expect(html).not.toContain('aria-label="Record"');
		expect(html).toContain('no-share=""');
	});
});
