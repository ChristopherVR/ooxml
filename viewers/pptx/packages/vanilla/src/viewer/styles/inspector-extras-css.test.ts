import { describe, expect, it } from 'vitest';

import { createAccessibilitySection } from '../ui/inspector/accessibility-section';
import { createAfterAnimationRow } from '../ui/inspector/after-animation-row';
import { createEffectSoundRow } from '../ui/inspector/effect-sound-row';
import { createMediaSection } from '../ui/inspector/media-section';
import { createSlideSizeRescalePrompt } from '../ui/inspector/slide-size-rescale-prompt';
import type { InspectorHandlers } from '../ui/inspector/types';
import { buildViewerCss } from './css';

const identity = (key: string): string => key;
const section = (): HTMLElement => document.createElement('div');
const noop = (): void => {};

/** Whether `css` has a selector naming `.name` itself, not just a longer class it prefixes. */
function hasRuleFor(css: string, name: string): boolean {
	return new RegExp(`\\.${name}(?![\\w-])`, 'u').test(css);
}

function renderedClasses(root: HTMLElement): string[] {
	const classes = new Set<string>();
	for (const node of [root, ...root.querySelectorAll<HTMLElement>('[class]')]) {
		for (const name of node.classList) {
			if (name.startsWith('pptxv-')) {
				classes.add(name);
			}
		}
	}
	return [...classes];
}

describe('inspector panel classes that used to have no rule', () => {
	it('styles every class the media, sound, after-animation, rescale and alt-text parts render', () => {
		const css = buildViewerCss();
		const handlers = { setMediaProperties: noop } as unknown as InspectorHandlers;
		const roots = [
			createMediaSection(document, identity, section, handlers).el,
			createEffectSoundRow(document, identity, noop, noop).el,
			createAfterAnimationRow(document, identity, noop, noop).el,
			createSlideSizeRescalePrompt(document, identity, noop).el,
			createAccessibilitySection(document, identity, section, handlers).el,
		];
		const missing = roots.flatMap(renderedClasses).filter((name) => !hasRuleFor(css, name));

		expect(missing).toStrictEqual([]);
	});

	it('styles the SmartArt hierarchy actions and the OLE nested-deck tab by the classes they render', () => {
		const css = buildViewerCss();
		const names = [
			'pptxv-smartart-hierarchy',
			'pptxv-ole-edit-content',
			'pptxv-ole-edit-deck-slide',
			'pptxv-ole-edit-deck-caption',
			'pptxv-inspector-pattern',
			'pptxv-tags-toggle-label',
		];
		expect(names.filter((name) => !hasRuleFor(css, name))).toStrictEqual([]);
		// The dialog renders `-deck-slide`; a rule for the never-rendered `-deck-row` is dead.
		expect(hasRuleFor(css, 'pptxv-ole-edit-deck-row')).toBeFalsy();
	});
});
