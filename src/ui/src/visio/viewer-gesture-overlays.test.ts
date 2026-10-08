import { afterEach, expect, it } from 'vitest';
import { hideGestureOverlays } from './viewer-gesture-overlays';
afterEach(() => document.body.replaceChildren());
it('restores exact captured visibility without changing replacement overlays', () => {
	const viewport = document.createElement('div');
	document.body.append(viewport);
	const original = document.createElement('div');
	original.className = 'rotation-overlay';
	original.style.visibility = 'collapse';
	viewport.append(original);
	const restore = hideGestureOverlays(viewport, '.rotation-overlay');
	expect(original.style.visibility).toBe('hidden');
	const replacement = document.createElement('div');
	replacement.className = 'rotation-overlay';
	original.replaceWith(replacement);
	restore();
	expect(original.style.visibility).toBe('collapse');
	expect(replacement.style.visibility).toBe('');
});
