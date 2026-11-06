import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ConnectorView from './ConnectorView.svelte';

/** PowerPoint's dash lengths and cap handling reach the Svelte connector. */
let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

function dashOf(shapeStyle: Record<string, unknown>): Element | null {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const element = {
		type: 'connector',
		id: 'conn-1',
		x: 0,
		y: 0,
		width: 120,
		height: 0,
		shapeStyle: { strokeColor: '#000000', strokeWidth: 6, ...shapeStyle },
	} as unknown as PptxElement;
	mounted = mount(ConnectorView, {
		target,
		props: { element, mediaDataUrls: new Map<string, string>(), zIndex: 2 },
	});
	flushSync();
	return target.querySelector('[stroke-dasharray]');
}

describe('connectorView dash patterns', () => {
	it('keeps the device-grid snap on the square-cap path', () => {
		expect(
			(dashOf({ strokeDash: 'dash', lineCap: 'sq' }) as HTMLElement).style.transform,
		).toContain('calc(');
	});

	it("draws a dashed square-cap line with PowerPoint's overshoot at its start", () => {
		expect(dashOf({ strokeDash: 'dash', lineCap: 'sq' })?.getAttribute('d')).toMatch(/^M-3,/u);
	});

	it('uses the measured dash and gap lengths', () => {
		expect(dashOf({ strokeDash: 'dash' })?.getAttribute('stroke-dasharray')).toBe('24 18');
	});

	it('folds a round cap into the dash and keeps square dashes flat', () => {
		const round = dashOf({ strokeDash: 'dash', lineCap: 'rnd' });
		expect(round?.getAttribute('stroke-dasharray')).toBe('18 24');
		expect(round?.getAttribute('stroke-linecap')).toBe('round');
		expect(dashOf({ strokeDash: 'dash', lineCap: 'sq' })?.getAttribute('stroke-linecap')).toBe(
			'butt',
		);
	});
});
