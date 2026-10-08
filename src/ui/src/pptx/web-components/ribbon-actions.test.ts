// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { buildTabRowActionsState } from '../render';
import type { TabRowActionsInput } from '../render';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const input: TabRowActionsInput = {
	translate: (key, params) => (params ? `${key}:${params.count}` : key),
	showComments: true,
	commentsOpen: false,
	commentCount: 0,
	showShare: true,
	isCollaborating: false,
};

function mount(overrides: Partial<TabRowActionsInput> = {}) {
	const host = document.createElement('pptx-ui-ribbon-actions') as HTMLElement & {
		state: unknown;
	};
	host.state = buildTabRowActionsState({ ...input, ...overrides });
	document.body.append(host);
	return host;
}
const part = (host: HTMLElement, name: string) =>
	host.shadowRoot!.querySelector<HTMLElement>(`[part="${name}"]`)!;

describe('buildTabRowActionsState', () => {
	it('labels Comments and Share and reflects the session', () => {
		expect(buildTabRowActionsState(input)).toEqual({
			commentsLabel: 'pptx.toolbar.comments',
			commentsTitle: 'pptx.toolbar.comments',
			commentsPressed: false,
			commentsCount: 0,
			noComments: false,
			shareLabel: 'pptx.toolbar.share',
			shareTitle: 'pptx.toolbar.share',
			sharePressed: false,
			noShare: false,
		});
		const live = buildTabRowActionsState({
			...input,
			isCollaborating: true,
			collaboratorCount: 3,
			showComments: false,
		});
		expect(live.sharePressed).toBe(true);
		expect(live.shareTitle).toBe('pptx.toolbar.sharingUsers:3');
		expect(live.noComments).toBe(true);
	});
});

describe('pptx-ui-ribbon-actions', () => {
	it('is the shared element: Comments then Share, no mode selector', () => {
		const host = mount({ commentsOpen: true, commentCount: 2 });
		expect(host).toBeInstanceOf(customElements.get('office-ui-ribbon-actions')!);
		expect(part(host, 'mode').hidden).toBe(true);
		expect(part(host, 'comments').textContent).toBe('pptx.toolbar.comments2');
		expect(part(host, 'comments').getAttribute('aria-label')).toBe('pptx.toolbar.comments');
		expect(part(host, 'comments').getAttribute('aria-pressed')).toBe('true');
		expect(part(host, 'share').textContent).toBe('pptx.toolbar.share');
		expect(host.shadowRoot!.textContent).toContain('--office-accent: var(--pptx-primary');
	});

	it('keeps the pptx event names and hides Share when the host does', () => {
		const host = mount({ showShare: false });
		const events: string[] = [];
		for (const type of ['comments-toggle', 'share-request'])
			host.addEventListener(type, () => events.push(type));
		part(host, 'comments').click();
		expect(events).toEqual(['comments-toggle']);
		expect(part(host, 'share').hidden).toBe(true);
	});
});
