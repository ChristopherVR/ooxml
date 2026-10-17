import { describe, expect, it } from 'vitest';
import { applyScreenTips, restoreScreenTips } from './viewer-screen-tips';

function ribbon(): HTMLElement {
	const root = document.createElement('div');
	root.innerHTML = `<button id="a" title="Bold (Ctrl+B)"></button>
		<button id="b" title="Chart: not available yet. Embedded charts are not supported."></button>`;
	return root;
}

describe('ScreenTip style', () => {
	it('drops the shortcut in plain style and keeps the reason a command is unavailable', () => {
		const root = ribbon();
		applyScreenTips(root, 'plain');
		expect(root.querySelector('#a')!.getAttribute('title')).toBe('Bold');
		expect(root.querySelector('#b')!.getAttribute('title')).toMatch(/not available yet/);
	});

	it('removes tips when off and restores them on demand', () => {
		const root = ribbon();
		applyScreenTips(root, 'off');
		expect(root.querySelector('#a')!.hasAttribute('title')).toBe(false);
		restoreScreenTips(root);
		expect(root.querySelector('#a')!.getAttribute('title')).toBe('Bold (Ctrl+B)');
		applyScreenTips(root, 'off');
		applyScreenTips(root, 'descriptions');
		expect(root.querySelector('#a')!.getAttribute('title')).toBe('Bold (Ctrl+B)');
	});
});
