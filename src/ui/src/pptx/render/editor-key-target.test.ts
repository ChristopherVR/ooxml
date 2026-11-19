// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';

import { isEditorTextInputTarget } from './editor-key-target';

function mount(html: string): HTMLElement {
	const host = document.createElement('div');
	host.innerHTML = html;
	document.body.appendChild(host);
	return host;
}

afterEach(() => {
	document.body.innerHTML = '';
});

describe('isEditorTextInputTarget across a shadow root', () => {
	it('sees the focused field inside a custom element host (SmartArt text pane Backspace)', () => {
		const host = mount('<div id="pane"></div>');
		const pane = host.querySelector('#pane') as HTMLElement;
		const root = pane.attachShadow({ mode: 'open' });
		root.innerHTML = '<input data-node="n1" />';
		(root.querySelector('input') as HTMLInputElement).focus();
		expect(isEditorTextInputTarget(pane)).toBe(true);
	});

	it('does not treat a shadow host without a focused field as text input', () => {
		const host = mount('<div id="pane"></div>');
		const pane = host.querySelector('#pane') as HTMLElement;
		pane.attachShadow({ mode: 'open' }).innerHTML = '<button>x</button>';
		expect(isEditorTextInputTarget(pane)).toBe(false);
	});
});
