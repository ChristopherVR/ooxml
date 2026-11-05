import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import EquationView from './EquationView.svelte';

let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

describe('EquationView', () => {
	it('anchors the maths with the shape text body instead of centring it vertically', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		const element = {
			type: 'text',
			id: 'eq1',
			x: 10,
			y: 20,
			width: 200,
			height: 50,
			textStyle: { textAnchor: 'top' },
			textSegments: [
				{ text: '', style: {}, equationXml: { 'm:oMath': { 'm:r': { 'm:t': 'x' } } } },
			],
		} as unknown as PptxElement;
		mounted = mount(EquationView, {
			target,
			props: { element, mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		const style =
			target.querySelector<HTMLElement>('[data-element-id="eq1"]')?.getAttribute('style') ?? '';
		expect(style).toContain('flex-direction: column');
		expect(style).toContain('justify-content: flex-start');
	});
});
